-- 現場マップ（Supabase 版）2026-11-10〜
-- 列名はスプレッドシート版の見出しと同じ日本語にしてある（画面が受け取るJSONの形をGAS版と同じにするため）。
-- 表は外から直接読めない（自動公開なし＋RLS有効・ポリシーなし）。画面はこのファイルの関数だけを呼ぶ。
-- パスワードは毎回の呼び出しで確かめる（GAS版と同じ）。金額は会議用のときだけ返す。
-- ※ パスワードの実際の値はここに書かない（GASの bootstrap で移す）

create extension if not exists pgcrypto;

-- ---------- 表 ----------
create table if not exists cfg (k text primary key, v text not null default '');
-- k: pw_normal / pw_meeting / users / radius / gas_token / imported_at / excel_at

create table if not exists accounts (
  "名前" text primary key, "パスワード" text not null default '', "担当名" text not null default '', "メモ" text not null default ''
);

-- Excel から取り込む物件（毎朝 GAS が入れ替える）。ord＝Excelの並び
create table if not exists sites (
  "現場名" text primary key, ord int not null default 0,
  "元請" text not null default '', "職人" text not null default '', "消検月" text not null default '',
  "図面" text not null default '', "ELV" text not null default '', "見積" text not null default '', "受注" text not null default '',
  "発注" text not null default '', "着工" text not null default '', "設置" text not null default '', "下検査" text not null default '',
  "注意点" text not null default '', "住所" text not null default '', "緯度" float8, "経度" float8,
  "要確認" text not null default '', "取込日時" timestamp
);

-- アプリで追加した物件
create table if not exists manual (
  "現場名" text primary key, "元請" text not null default '', "職人" text not null default '', "消検月" text not null default '',
  "注意点" text not null default '', "住所" text not null default '', "緯度" float8, "経度" float8,
  "登録日時" timestamp, "登録者" text not null default '', "更新日時" timestamp, "更新者" text not null default ''
);

-- アプリでの項目の変更（GAS版の「フラグ変更」シート）。「元の値」はそのときのExcelの値
create table if not exists edits (
  "現場名" text not null, "項目" text not null, "値" text not null default '', "元の値" text not null default '',
  "日時" timestamp, "更新者" text not null default '', primary key ("現場名", "項目")
);

create table if not exists amounts ("現場名" text primary key, "金額（千円）" numeric, "取込日時" timestamp);

create table if not exists extra (
  "現場名" text primary key, "消検日" timestamp, "建築確認検査日" timestamp, "搬入" text not null default '',
  "工事進捗メモ" text not null default '', "更新日時" timestamp, "更新者" text not null default '', "正式名称" text not null default ''
);

create table if not exists visits (
  id bigserial primary key, "日時" timestamp not null, "記録者" text not null default '', "現場名" text not null,
  "建物進捗" text not null default '', "メモ" text not null default '', "写真" text not null default '',
  "緯度" float8, "経度" float8, "距離m" int
);
create index if not exists visits_site on visits ("現場名");

create table if not exists drawings (
  "ファイルID" text primary key, "現場名" text not null, "ファイル名" text not null default '', "サイズKB" int,
  "登録日時" timestamp, "登録者" text not null default ''
);

alter table cfg enable row level security;
alter table accounts enable row level security;
alter table sites enable row level security;
alter table manual enable row level security;
alter table edits enable row level security;
alter table amounts enable row level security;
alter table extra enable row level security;
alter table visits enable row level security;
alter table drawings enable row level security;
revoke all on all tables in schema public from anon, authenticated;
revoke all on all sequences in schema public from anon, authenticated;

-- ---------- 小道具 ----------
create or replace function _now() returns timestamp language sql stable as
$$ select date_trunc('second', now() at time zone 'Asia/Tokyo')::timestamp $$;
create or replace function _cfg(p text) returns text language sql stable as
$$ select coalesce((select v from cfg where k = p), '') $$;
create or replace function _t(x jsonb, k text) returns text language sql immutable as
$$ select btrim(coalesce(x ->> k, '')) $$;
create or replace function _num(x jsonb, k text) returns float8 language sql immutable as
$$ select nullif(btrim(coalesce(x ->> k, '')), '')::float8 $$;
create or replace function _ts(x jsonb, k text) returns timestamp language sql immutable as
$$ select nullif(btrim(coalesce(x ->> k, '')), '')::timestamp $$;
create or replace function _split(v text) returns text[] language sql immutable as
$$ select coalesce(array_remove(regexp_split_to_array(coalesce(v, ''), '[、,，・\s]+'), ''), '{}') $$;
create or replace function _norm_circle(v text) returns text language sql immutable as
$$ select case when v = '○' then '〇' else coalesce(v, '') end $$;
create or replace function _flag_keys() returns text[] language sql immutable as
$$ select array['図面','ELV','見積','受注','発注','着工','設置','下検査'] $$;
create or replace function _edit_keys() returns text[] language sql immutable as
$$ select array['元請','職人','消検月','注意点'] || _flag_keys() $$;

-- だれか（パスワードから）。role: meeting / normal / worker / ''
create or replace function _auth(p_pw text) returns jsonb language sql stable as $$
  select case
    when coalesce(p_pw, '') = '' then jsonb_build_object('role', '')
    when _cfg('pw_meeting') <> '' and p_pw = _cfg('pw_meeting') then jsonb_build_object('role', 'meeting')
    when _cfg('pw_normal') <> '' and p_pw = _cfg('pw_normal') then jsonb_build_object('role', 'normal')
    else coalesce((select jsonb_build_object('role', 'worker', 'name', "名前", 'worker', coalesce(nullif("担当名", ''), "名前"))
                   from accounts where "パスワード" <> '' and "パスワード" = p_pw limit 1), jsonb_build_object('role', ''))
  end $$;

-- 物件一覧（Excel＋アプリ追加。アプリでの変更を重ねる。Excel側がその後変わった項目はExcelの値）
create or replace function _sites() returns table (j jsonb, o float8, n text) language sql stable as $$
  with x as (
    select to_jsonb(s) - 'ord' as j, s.ord::float8 as o, s."現場名" as n from sites s
    union all
    select to_jsonb(m) || jsonb_build_object('手入力', 1, '要確認', case when m."緯度" is null then '位置未設定' else '' end),
           1e9 + coalesce(extract(epoch from m."登録日時"), 0), m."現場名"
    from manual m where not exists (select 1 from sites s where s."現場名" = m."現場名")
  )
  select x.j || coalesce((
      select jsonb_object_agg(e."項目", e."値") || jsonb_object_agg('元_' || e."項目", coalesce(x.j ->> e."項目", '')) || jsonb_object_agg('アプリ_' || e."項目", 1)
      from edits e
      where e."現場名" = x.n and e."項目" = any(_edit_keys()) and coalesce(x.j ->> e."項目", '') = e."元の値"
    ), '{}'::jsonb), x.o, x.n
  from x $$;

-- 会議用の金額（Excel＋アプリで入れた金額）
create or replace function _amounts() returns jsonb language sql stable as $$
  select coalesce(jsonb_agg(jsonb_build_object('現場名', n, '金額（千円）', v)), '[]'::jsonb) from (
    select a."現場名" n,
           case when e."現場名" is not null and coalesce(a."金額（千円）"::text, '') = e."元の値" then nullif(e."値", '')::numeric else a."金額（千円）" end v
    from amounts a left join edits e on e."現場名" = a."現場名" and e."項目" = '金額'
    union all
    select e."現場名", nullif(e."値", '')::numeric from edits e
    where e."項目" = '金額' and e."元の値" = '' and not exists (select 1 from amounts a where a."現場名" = e."現場名")
  ) t $$;

-- ---------- 画面から呼ぶ関数 ----------
-- 記録者の名前（パスワードなしで返す。ログイン画面用）
create or replace function users() returns jsonb language sql stable security definer set search_path = public as $$
  select jsonb_build_object('ok', true, 'users', to_jsonb(_split(replace(_cfg('users'), ' ', ''))))
$$;

-- 全データ（GAS版の doGet と同じ形）
create or replace function load(p_pw text) returns jsonb language plpgsql stable security definer set search_path = public as $$
declare a jsonb := _auth(p_pw); r text := a ->> 'role'; w text := a ->> 'worker'; mine text[];
begin
  if r = '' then return jsonb_build_object('ok', false, 'error', 'パスワードが違います'); end if;
  if r = 'worker' then
    select coalesce(array_agg(n), '{}') into mine from _sites() s where w = any(_split(s.j ->> '職人'));
  end if;
  return jsonb_build_object(
    'ok', true, 'role', r,
    'me', case when r = 'worker' then a ->> 'name' end,
    'staff', case when r <> 'worker' then (select coalesce(jsonb_agg(coalesce(nullif("担当名", ''), "名前") order by "名前"), '[]') from accounts) end,
    'amounts', case when r = 'meeting' then _amounts() end,
    'sites', (select coalesce(jsonb_agg(s.j order by s.o, s.n), '[]') from _sites() s where r <> 'worker' or s.n = any(mine)),
    'extra', (select coalesce(jsonb_agg(to_jsonb(e)), '[]') from extra e where r <> 'worker' or e."現場名" = any(mine)),
    'visits', (select coalesce(jsonb_agg(to_jsonb(v) - 'id' order by v.id), '[]') from visits v where r <> 'worker' or v."現場名" = any(mine)),
    'drawings', (select coalesce(jsonb_agg(to_jsonb(d) order by d."登録日時"), '[]') from drawings d where r <> 'worker' or d."現場名" = any(mine)),
    'users', case when r = 'worker' then jsonb_build_array(a ->> 'name') else to_jsonb(_split(replace(_cfg('users'), ' ', ''))) end,
    'radius', coalesce(nullif(_cfg('radius'), '')::int, 300),
    'updated', to_jsonb(_now()),
    'importedAt', _cfg('imported_at'), 'excelAt', _cfg('excel_at'));
end $$;

-- GAS が写真・図面を保存する前に「この人は保存してよいか」を聞く
create or replace function whoami(p_pw text, p_site text) returns jsonb language plpgsql stable security definer set search_path = public as $$
declare a jsonb := _auth(p_pw); r text := a ->> 'role';
begin
  if r = '' then return jsonb_build_object('ok', false, 'error', 'パスワードが違います'); end if;
  if r = 'worker' and not exists (select 1 from _sites() s where s.n = p_site and (a ->> 'worker') = any(_split(s.j ->> '職人'))) then
    return jsonb_build_object('ok', false, 'error', '担当の物件ではありません');
  end if;
  return jsonb_build_object('ok', true, 'role', r, 'name', a ->> 'name');
end $$;

-- 1項目の変更を記録（Excelと同じ値に戻したら記録を消す）
create or replace function _set_edit(p_site text, p_key text, p_val text, p_base text, p_user text) returns void language plpgsql as $$
begin
  if p_val = case when p_key = any(_flag_keys()) then _norm_circle(p_base) else p_base end then
    delete from edits where "現場名" = p_site and "項目" = p_key;
  else
    insert into edits values (p_site, p_key, p_val, p_base, _now(), p_user)
    on conflict ("現場名", "項目") do update set "値" = excluded."値", "元の値" = excluded."元の値", "日時" = excluded."日時", "更新者" = excluded."更新者";
  end if;
end $$;

-- いまのExcel側の値
create or replace function _base(p_site text, p_key text) returns text language plpgsql stable as $$
declare v text;
begin
  if p_key = '金額' then
    select coalesce("金額（千円）"::text, '') into v from amounts where "現場名" = p_site;
    return coalesce(v, '');
  end if;
  if not (p_key = any(_edit_keys())) then return ''; end if;
  execute format('select coalesce(%I::text, '''') from sites where "現場名" = $1', p_key) into v using p_site;
  return coalesce(v, '');
end $$;

create or replace function _check_edit(k text, v text) returns text language sql immutable as $$
  select case
    when k = any(_flag_keys()) then case when v = any(array['', '〇', '／', '△']) then '' else k || 'の値が不正です' end
    when k = '消検月' then case when v = '' or v ~ '^\d{4}-\d{2}$' then '' else '消検月は 2026-12 の形で' end
    when k = '金額' then case when v = '' or v ~ '^\d+(\.\d+)?$' then '' else '金額は数字（千円）で' end
    when k = any(_edit_keys()) then ''
    else '変えられない項目です：' || k end $$;

create or replace function _pw_problem(pw text, except_name text) returns text language sql stable as $$
  select case
    when pw = '' then ''
    when length(pw) < 6 then 'パスワードは6文字以上にしてください'
    when pw = _cfg('pw_normal') or pw = _cfg('pw_meeting') then '普段用・会議用と同じパスワードは使えません'
    when exists (select 1 from accounts where "パスワード" = pw and "名前" <> coalesce(except_name, '')) then 'ほかの職人と同じパスワードです'
    else '' end $$;

-- 書き込み（GAS版の doPost と同じ action 名・同じ中身）
create or replace function act(p_pw text, p_user text, p_action text, p_body jsonb) returns jsonb language plpgsql security definer set search_path = public as $$
declare
  a jsonb := _auth(p_pw); r text := a ->> 'role'; b jsonb := coalesce(p_body, '{}'); u text := btrim(coalesce(p_user, ''));
  site text := _t(b, 'site'); k text; v text; e text; f jsonb; in_excel boolean; is_manual boolean; n int; nm text; addr text;
begin
  if r = '' then return jsonb_build_object('ok', false, 'error', 'パスワードが違います'); end if;
  -- 金額を書けるのは会議用だけ
  if r <> 'meeting' and b ? 'fields' then b := jsonb_set(b, '{fields}', (b -> 'fields') - '金額'); end if;
  if p_action in ('listAccounts', 'saveAccount', 'deleteAccount', 'setSharedPw') and r <> 'meeting' then
    return jsonb_build_object('ok', false, 'error', '会議用パスワードで入ったときだけ使えます');
  end if;
  if r = 'worker' then
    if p_action not in ('visit', 'deleteVisit') then return jsonb_build_object('ok', false, 'error', 'この操作はできません（業務部に頼んでください）'); end if;
    u := a ->> 'name';
    if not exists (select 1 from _sites() s where s.n = site and (a ->> 'worker') = any(_split(s.j ->> '職人'))) then
      return jsonb_build_object('ok', false, 'error', '担当の物件ではありません');
    end if;
    if p_action = 'deleteVisit' and _t(b, 'by') <> u then return jsonb_build_object('ok', false, 'error', '自分の記録だけ削除できます'); end if;
  end if;
  perform pg_advisory_xact_lock(4711); -- 書き込みは1件ずつ（GAS版のロックと同じ）

  if p_action = 'listAccounts' then
    return jsonb_build_object('ok', true,
      'accounts', (select coalesce(jsonb_agg(jsonb_build_object('name', "名前", 'pw', "パスワード", 'worker', "担当名") order by "名前"), '[]') from accounts),
      'normalPw', _cfg('pw_normal'), 'meetingPw', _cfg('pw_meeting'));

  elsif p_action = 'visit' then
    insert into visits ("日時", "記録者", "現場名", "建物進捗", "メモ", "写真", "緯度", "経度", "距離m")
    values (_now(), u, site, _t(b, 'stage'), _t(b, 'memo'),
            coalesce((select string_agg(x, E'\n') from jsonb_array_elements_text(coalesce(b -> 'photoIds', '[]')) x where x <> ''), ''),
            _num(b, 'lat'), _num(b, 'lng'), round(_num(b, 'dist'))::int);
    return jsonb_build_object('ok', true);

  elsif p_action = 'deleteVisit' then
    delete from visits where id = (select id from visits where "現場名" = site and "記録者" = _t(b, 'by')
      and to_char("日時", 'YYYY-MM-DD"T"HH24:MI:SS') = left(_t(b, 'at'), 19) order by id desc limit 1);
    get diagnostics n = row_count;
    if n = 0 then return jsonb_build_object('ok', false, 'error', '該当する記録が見つかりません（更新してからもう一度）'); end if;
    return jsonb_build_object('ok', true);

  elsif p_action = 'extra' then
    insert into extra ("現場名", "更新日時", "更新者") values (site, _now(), u)
      on conflict ("現場名") do update set "更新日時" = excluded."更新日時", "更新者" = excluded."更新者";
    if b ? 'fireDate' then update extra set "消検日" = nullif(btrim(_t(b, 'fireDate') || ' ' || _t(b, 'fireTime')), '')::timestamp where "現場名" = site; end if;
    if b ? 'bldgDate' then update extra set "建築確認検査日" = nullif(btrim(_t(b, 'bldgDate') || ' ' || _t(b, 'bldgTime')), '')::timestamp where "現場名" = site; end if;
    if b ? 'delivery' then update extra set "搬入" = _t(b, 'delivery') where "現場名" = site; end if;
    if b ? 'progress' then update extra set "工事進捗メモ" = _t(b, 'progress') where "現場名" = site; end if;
    if b ? 'alias' then update extra set "正式名称" = _t(b, 'alias') where "現場名" = site; end if;
    return jsonb_build_object('ok', true);

  elsif p_action = 'setFlag' then
    k := _t(b, 'key'); v := _t(b, 'value');
    if not (k = any(_flag_keys())) or not (v = any(array['', '〇', '／', '△'])) then return jsonb_build_object('ok', false, 'error', '項目か値が不正です'); end if;
    perform _set_edit(site, k, v, _base(site, k), u);
    return jsonb_build_object('ok', true);

  elsif p_action = 'editSite' then
    f := coalesce(b -> 'fields', '{}');
    for k in select jsonb_object_keys(f) loop
      e := _check_edit(k, _t(f, k)); if e <> '' then return jsonb_build_object('ok', false, 'error', e); end if;
    end loop;
    in_excel := exists (select 1 from sites where "現場名" = site);
    is_manual := not in_excel and exists (select 1 from manual where "現場名" = site);
    if not in_excel and not is_manual then return jsonb_build_object('ok', false, 'error', '物件が見つかりません'); end if;
    for k in select jsonb_object_keys(f) loop
      v := _t(f, k);
      if is_manual and k in ('元請', '職人', '消検月', '注意点') then
        execute format('update manual set %I = $1 where "現場名" = $2', k) using v, site;
      else
        perform _set_edit(site, k, v, case when in_excel then _base(site, k) else '' end, u);
      end if;
    end loop;
    if is_manual then update manual set "更新日時" = _now(), "更新者" = u where "現場名" = site; end if;
    return jsonb_build_object('ok', true);

  elsif p_action = 'addSite' then
    nm := _t(b, 'name');
    if nm = '' then return jsonb_build_object('ok', false, 'error', '現場名を入れてください'); end if;
    if exists (select 1 from sites where "現場名" = nm) or exists (select 1 from manual where "現場名" = nm) then
      return jsonb_build_object('ok', false, 'error', '同じ現場名がすでにあります');
    end if;
    insert into manual ("現場名", "元請", "職人", "消検月", "注意点", "住所", "緯度", "経度", "登録日時", "登録者")
    values (nm, _t(b, 'moto'), _t(b, 'worker'), _t(b, 'ym'), _t(b, 'note'), _t(b, 'addr'), _num(b, 'lat'), _num(b, 'lng'), _now(), u);
    f := coalesce(b -> 'fields', '{}');
    for k in select jsonb_object_keys(f) loop
      v := _t(f, k);
      if v <> '' and (k = any(_flag_keys()) or k = '金額') and _check_edit(k, v) = '' then perform _set_edit(nm, k, v, '', u); end if;
    end loop;
    return jsonb_build_object('ok', true, 'name', nm, 'lat', coalesce(to_jsonb(_num(b, 'lat')), '""'), 'lng', coalesce(to_jsonb(_num(b, 'lng')), '""'), 'addr', _t(b, 'addr'));

  elsif p_action = 'deleteSite' then
    delete from manual where "現場名" = site;
    get diagnostics n = row_count;
    if n = 0 then return jsonb_build_object('ok', false, 'error', 'アプリで追加した物件だけ削除できます'); end if;
    return jsonb_build_object('ok', true);

  elsif p_action = 'setLocation' then
    if _num(b, 'lat') is null or _num(b, 'lng') is null then
      return jsonb_build_object('ok', false, 'error', '位置が分かりませんでした。地図をタップして指定してください');
    end if;
    update sites set "住所" = _t(b, 'addr'), "緯度" = _num(b, 'lat'), "経度" = _num(b, 'lng'),
      "要確認" = case when "要確認" like '同名%' then "要確認" else '' end where "現場名" = site;
    get diagnostics n = row_count;
    if n = 0 then
      update manual set "住所" = _t(b, 'addr'), "緯度" = _num(b, 'lat'), "経度" = _num(b, 'lng') where "現場名" = site;
      get diagnostics n = row_count;
      if n = 0 then return jsonb_build_object('ok', false, 'error', '物件が見つかりません'); end if;
    end if;
    return jsonb_build_object('ok', true, 'lat', _num(b, 'lat'), 'lng', _num(b, 'lng'));

  elsif p_action = 'addDrawing' then
    insert into drawings values (_t(b, 'id'), site, _t(b, 'name'), _num(b, 'kb')::int, _now(), u) on conflict do nothing;
    return jsonb_build_object('ok', true, 'id', _t(b, 'id'));

  elsif p_action = 'deleteDrawing' then
    delete from drawings where "ファイルID" = _t(b, 'id');
    get diagnostics n = row_count;
    if n = 0 then return jsonb_build_object('ok', false, 'error', '図面が見つかりません'); end if;
    return jsonb_build_object('ok', true);

  elsif p_action = 'saveAccount' then
    nm := _t(b, 'name');
    if nm = '' then return jsonb_build_object('ok', false, 'error', '名前を入れてください'); end if;
    e := _pw_problem(_t(b, 'newPw'), coalesce(nullif(_t(b, 'oldName'), ''), nm));
    if e <> '' then return jsonb_build_object('ok', false, 'error', e); end if;
    if exists (select 1 from accounts where "名前" = nm and "名前" <> coalesce(nullif(_t(b, 'oldName'), ''), nm)) then
      return jsonb_build_object('ok', false, 'error', '同じ名前の職人がいます');
    end if;
    if (b ->> 'isNew')::boolean is true and exists (select 1 from accounts where "名前" = nm) then
      return jsonb_build_object('ok', false, 'error', '同じ名前の職人がいます（変えるときは一覧の「保存」で）');
    end if;
    update accounts set "名前" = nm, "パスワード" = _t(b, 'newPw'), "担当名" = coalesce(nullif(_t(b, 'worker'), ''), nm)
      where "名前" = coalesce(nullif(_t(b, 'oldName'), ''), nm);
    get diagnostics n = row_count;
    if n = 0 then insert into accounts values (nm, _t(b, 'newPw'), coalesce(nullif(_t(b, 'worker'), ''), nm), _t(b, 'memo')); end if;
    return jsonb_build_object('ok', true);

  elsif p_action = 'deleteAccount' then
    delete from accounts where "名前" = _t(b, 'name');
    get diagnostics n = row_count;
    if n = 0 then return jsonb_build_object('ok', false, 'error', '見つかりません'); end if;
    return jsonb_build_object('ok', true);

  elsif p_action = 'setSharedPw' then
    v := _t(b, 'newPw');
    if length(v) < 6 then return jsonb_build_object('ok', false, 'error', 'パスワードは6文字以上にしてください'); end if;
    if v = case when _t(b, 'kind') = 'meeting' then _cfg('pw_normal') else _cfg('pw_meeting') end
       or exists (select 1 from accounts where "パスワード" = v) then
      return jsonb_build_object('ok', false, 'error', 'ほかのパスワードと同じです');
    end if;
    insert into cfg values (case when _t(b, 'kind') = 'meeting' then 'pw_meeting' else 'pw_normal' end, v)
      on conflict (k) do update set v = excluded.v;
    return jsonb_build_object('ok', true);
  end if;
  return jsonb_build_object('ok', false, 'error', 'unknown action');
end $$;

-- ---------- GAS から呼ぶ関数（gas_token で確かめる。トークンはGASが作ってbootstrapで登録） ----------
create or replace function _gas_ok(p_token text) returns boolean language sql stable as
$$ select coalesce(p_token, '') <> '' and p_token = _cfg('gas_token') $$;

-- 取り込みの前に、いまの住所・座標を渡す
create or replace function gas_sites(p_token text) returns jsonb language plpgsql stable security definer set search_path = public as $$
begin
  if not _gas_ok(p_token) then raise exception 'bad token'; end if;
  return (select coalesce(jsonb_object_agg(n, j), '{}') from (
    select "現場名" n, jsonb_build_object('住所', "住所", '緯度', "緯度", '経度', "経度") j from manual
    where not exists (select 1 from sites s where s."現場名" = manual."現場名")
    union all
    select "現場名", jsonb_build_object('住所', "住所", '緯度', "緯度", '経度', "経度") from sites) t);
end $$;

-- Excelの取り込み。rows は GAS版の物件シートと同じ見出しのオブジェクト（_prevLat＝GASが読んだときの緯度）
create or replace function gas_import(p_token text, p_rows jsonb, p_amounts jsonb, p_imported text, p_excel text) returns jsonb language plpgsql security definer set search_path = public as $$
declare cur jsonb;
begin
  if not _gas_ok(p_token) then raise exception 'bad token'; end if;
  if jsonb_array_length(coalesce(p_rows, '[]')) = 0 then raise exception '物件行が0件でした'; end if;
  perform pg_advisory_xact_lock(4711);
  -- 読み込みの間に「位置を直す」で変わった座標は、そちらを使う
  select coalesce(jsonb_object_agg("現場名", jsonb_build_object('住所', "住所", '緯度', "緯度", '経度', "経度")), '{}') into cur from sites where "緯度" is not null;
  delete from sites;
  insert into sites
  select _t(x, '現場名'), ord::int, _t(x, '元請'), _t(x, '職人'), _t(x, '消検月'),
    _t(x, '図面'), _t(x, 'ELV'), _t(x, '見積'), _t(x, '受注'), _t(x, '発注'), _t(x, '着工'), _t(x, '設置'), _t(x, '下検査'),
    _t(x, '注意点'),
    case when moved then cur -> _t(x, '現場名') ->> '住所' else _t(x, '住所') end,
    case when moved then (cur -> _t(x, '現場名') ->> '緯度')::float8 else _num(x, '緯度') end,
    case when moved then (cur -> _t(x, '現場名') ->> '経度')::float8 else _num(x, '経度') end,
    case when moved and _t(x, '要確認') not like '同名%' then '' else _t(x, '要確認') end,
    _now()
  from (select x, ord, (cur ? _t(x, '現場名') and (cur -> _t(x, '現場名') ->> '緯度') is distinct from nullif(_t(x, '_prevLat'), '')
                       and (_t(x, '_prevLat') = '' or (cur -> _t(x, '現場名') ->> '緯度')::float8 <> _num(x, '_prevLat'))) as moved
        from jsonb_array_elements(p_rows) with ordinality as t(x, ord)) z
  on conflict do nothing;
  delete from amounts;
  insert into amounts select _t(x, '現場名'), nullif(_t(x, '金額（千円）'), '')::numeric, _now() from jsonb_array_elements(coalesce(p_amounts, '[]')) x
    on conflict do nothing;
  -- Excel側で変わった項目のアプリの記録を消す（Excelに無い物件＝アプリで追加した物件は残す）
  delete from edits e using sites s where e."現場名" = s."現場名" and e."項目" <> '金額'
    and coalesce(to_jsonb(s) ->> e."項目", '') <> e."元の値";
  delete from edits e where e."項目" = '金額' and exists (select 1 from sites s where s."現場名" = e."現場名")
    and coalesce((select "金額（千円）"::text from amounts a where a."現場名" = e."現場名"), '') <> e."元の値";
  insert into cfg values ('imported_at', p_imported), ('excel_at', p_excel) on conflict (k) do update set v = excluded.v;
  return jsonb_build_object('ok', true, 'n', (select count(*) from sites));
end $$;

-- 図面の整理（消防検査から30日、かつ登録から30日）。消した図面のIDと、いま使っているファイルIDの全部を返す
create or replace function gas_files(p_token text) returns jsonb language plpgsql security definer set search_path = public as $$
declare gone jsonb;
begin
  if not _gas_ok(p_token) then raise exception 'bad token'; end if;
  with fire as (
    select n, max(d) d from (
      select "現場名" n, (to_date("消検月" || '-01', 'YYYY-MM-DD') + interval '1 month - 1 day')::timestamp d from sites where "消検月" ~ '^\d{4}-\d{2}'
      union all select "現場名", (to_date("消検月" || '-01', 'YYYY-MM-DD') + interval '1 month - 1 day')::timestamp from manual where "消検月" ~ '^\d{4}-\d{2}'
    ) t group by n
  ), due as (
    select coalesce(e."消検日"::date::timestamp, f.d) d, coalesce(e."現場名", f.n) n
    from fire f full join extra e on e."現場名" = f.n
  ), del as (
    delete from drawings dr using due where due.n = dr."現場名" and due.d is not null
      and due.d < _now() - interval '30 days' and coalesce(dr."登録日時", '1970-01-01') < _now() - interval '30 days'
    returning dr."ファイルID"
  ) select coalesce(jsonb_agg("ファイルID"), '[]') into gone from del;
  return jsonb_build_object('ok', true, 'trash', gone,
    'keep', (select coalesce(jsonb_agg(id), '[]') from (
      select "ファイルID" id from drawings
      union select x from visits, unnest(string_to_array("写真", E'\n')) x where x <> '') t));
end $$;

-- 最初の1回だけ：スプレッドシートの中身を丸ごと入れる（cfg が空のときだけ動く）
create or replace function bootstrap(p_dump jsonb) returns jsonb language plpgsql security definer set search_path = public as $$
declare d jsonb := p_dump;
begin
  if exists (select 1 from cfg) then raise exception 'already initialized'; end if;
  perform pg_advisory_xact_lock(4711);
  insert into cfg select key, value from jsonb_each_text(d -> 'cfg');
  insert into accounts select _t(x, '名前'), _t(x, 'パスワード'), _t(x, '担当名'), _t(x, 'メモ')
    from jsonb_array_elements(d -> 'accounts') x where _t(x, '名前') <> '' on conflict do nothing;
  insert into sites select _t(x, '現場名'), ord::int, _t(x, '元請'), _t(x, '職人'), _t(x, '消検月'),
    _t(x, '図面'), _t(x, 'ELV'), _t(x, '見積'), _t(x, '受注'), _t(x, '発注'), _t(x, '着工'), _t(x, '設置'), _t(x, '下検査'),
    _t(x, '注意点'), _t(x, '住所'), _num(x, '緯度'), _num(x, '経度'), _t(x, '要確認'), _ts(x, '取込日時')
    from jsonb_array_elements(d -> 'sites') with ordinality as t(x, ord) where _t(x, '現場名') <> '' on conflict do nothing;
  insert into manual select _t(x, '現場名'), _t(x, '元請'), _t(x, '職人'), _t(x, '消検月'), _t(x, '注意点'), _t(x, '住所'),
    _num(x, '緯度'), _num(x, '経度'), _ts(x, '登録日時'), _t(x, '登録者'), _ts(x, '更新日時'), _t(x, '更新者')
    from jsonb_array_elements(d -> 'manual') x where _t(x, '現場名') <> '' on conflict do nothing;
  insert into edits select _t(x, '現場名'), _t(x, '項目'), _t(x, '値'), _t(x, '元の値'), _ts(x, '日時'), _t(x, '更新者')
    from jsonb_array_elements(d -> 'edits') x where _t(x, '現場名') <> '' on conflict do nothing;
  insert into amounts select _t(x, '現場名'), nullif(_t(x, '金額（千円）'), '')::numeric, _ts(x, '取込日時')
    from jsonb_array_elements(d -> 'amounts') x where _t(x, '現場名') <> '' on conflict do nothing;
  insert into extra select _t(x, '現場名'), _ts(x, '消検日'), _ts(x, '建築確認検査日'), _t(x, '搬入'), _t(x, '工事進捗メモ'),
    _ts(x, '更新日時'), _t(x, '更新者'), _t(x, '正式名称')
    from jsonb_array_elements(d -> 'extra') x where _t(x, '現場名') <> '' on conflict do nothing;
  insert into visits ("日時", "記録者", "現場名", "建物進捗", "メモ", "写真", "緯度", "経度", "距離m")
    select _ts(x, '日時'), _t(x, '記録者'), _t(x, '現場名'), _t(x, '建物進捗'), _t(x, 'メモ'), _t(x, '写真'),
      _num(x, '緯度'), _num(x, '経度'), round(_num(x, '距離m'))::int
    from jsonb_array_elements(d -> 'visits') with ordinality as t(x, ord) where _t(x, '現場名') <> '' order by ord;
  insert into drawings select _t(x, 'ファイルID'), _t(x, '現場名'), _t(x, 'ファイル名'), _num(x, 'サイズKB')::int, _ts(x, '登録日時'), _t(x, '登録者')
    from jsonb_array_elements(d -> 'drawings') x where _t(x, 'ファイルID') <> '' on conflict do nothing;
  return jsonb_build_object('ok', true,
    'sites', (select count(*) from sites), 'manual', (select count(*) from manual), 'edits', (select count(*) from edits),
    'amounts', (select count(*) from amounts), 'extra', (select count(*) from extra), 'visits', (select count(*) from visits),
    'drawings', (select count(*) from drawings), 'accounts', (select count(*) from accounts));
end $$;

-- ---------- 変更の知らせ（リアルタイム） ----------
-- 表が変わったら公開チャンネル genba-map に「変わった」とだけ流す（中身は流さない）。画面はそれを受けて読み直す
create or replace function _notify() returns trigger language plpgsql security definer set search_path = public as $$
begin
  perform realtime.send(jsonb_build_object('t', TG_TABLE_NAME), 'changed', 'genba-map', false);
  return null;
exception when others then return null; -- 知らせに失敗しても保存は止めない
end $$;
do $$ declare t text; begin
  foreach t in array array['sites', 'manual', 'edits', 'amounts', 'extra', 'visits', 'drawings', 'accounts'] loop
    execute format('drop trigger if exists notify_%1$s on %1$I', t);
    execute format('create trigger notify_%1$s after insert or update or delete on %1$I for each statement execute function _notify()', t);
  end loop;
end $$;

-- ---------- 呼べる関数を絞る ----------
revoke execute on all functions in schema public from public, anon, authenticated;
grant execute on function users() to anon;
grant execute on function load(text) to anon;
grant execute on function whoami(text, text) to anon;
grant execute on function act(text, text, text, jsonb) to anon;
grant execute on function gas_sites(text) to anon;
grant execute on function gas_import(text, jsonb, jsonb, text, text) to anon;
grant execute on function gas_files(text) to anon;
grant execute on function bootstrap(jsonb) to anon;
alter default privileges in schema public revoke execute on functions from public, anon, authenticated;
