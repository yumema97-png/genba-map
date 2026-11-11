/**
 * 現場マップ GAS バックエンド
 * スプレッドシート「現場マップDB」に紐付けて使う（拡張機能 → Apps Script）
 *
 * シート構成（setup() で自動作成）
 *   物件      : Excelから週1で取り込む（手で触らない）
 *   物件追加  : アプリから入力する物件ごとの情報（消検日など）
 *   訪問記録  : 行くたびに1行追加
 *   設定      : パスワード、記録者一覧、Excelファイル名
 *
 * 初回手順
 *   1. setup() を実行（シート作成）
 *   2. 設定シートの パスワード / 記録者 / Excelファイル名 を書く
 *   3. サービス追加 → Drive API（Excel読み込みに使う）
 *   4. デプロイ → ウェブアプリ（実行者: 自分 / アクセス: 全員）→ URLを index.html の GAS_URL に貼る
 *   5. 週1更新は トリガー → importFromExcel を 週次 に設定
 */

const SHEETS = {
  sites: '物件',
  extra: '物件追加',
  visits: '訪問記録',
  config: '設定',
};
const SITE_COLS = ['現場名', '元請', '職人', '消検月', '図面', 'ELV', '見積', '受注', '発注', '着工', '設置', '下検査', '注意点', '住所', '緯度', '経度', '要確認', '取込日時'];
// 正式名称：現場名（Excelの略称）はそのまま鍵に使い、確定した物件名を別に持つ（2026-11-09 追加）
const EXTRA_COLS = ['現場名', '消検日', '建築確認検査日', '搬入', '工事進捗メモ', '更新日時', '更新者', '正式名称'];
const VISIT_COLS = ['日時', '記録者', '現場名', '建物進捗', 'メモ', '写真', '緯度', '経度', '距離m'];
// アプリで追加した物件。Excel取り込みは触らない。同じ現場名がExcelに載ったらExcel側を出す
const MANUAL_SHEET = '物件手入力';
// 図面〜下検査の○／△をアプリで切り替えた記録。「元の値」はそのときのExcelの値。
// Excelがその後も同じ値ならアプリの値を出す。Excel側が変わったらExcelを優先（新しいほうが勝つ）
const FLAG_SHEET = 'フラグ変更';
const FLAG_COLS = ['現場名', '項目', '値', '元の値', '日時', '更新者'];
const FLAG_KEYS = ['図面', 'ELV', '見積', '受注', '発注', '着工', '設置', '下検査'];
const FLAG_VALUES = ['', '〇', '／', '△'];
// 金額（千円）は物件シートに入れない（普段のパスワードでは返さない）。会議用パスワードのときだけ返す
const AMOUNT_SHEET = '金額';
const AMOUNT_COLS = ['現場名', '金額（千円）', '取込日時'];
const MANUAL_COLS = ['現場名', '元請', '職人', '消検月', '注意点', '住所', '緯度', '経度', '登録日時', '登録者', '更新日時', '更新者'];
const PHOTO_FOLDER = '現場写真';
// 図面PDF：ドライブ「仕事用/現場マップ/図面/<現場名>/」に置く。「リンクを知っている人は閲覧可」（2026-11-10〜。リンクはアプリの中でだけ出す）
// 消防検査から30日たったら自動でゴミ箱へ（ゴミ箱でさらに30日は戻せる）
const DRAW_SHEET = '図面';
const DRAW_COLS = ['現場名', 'ファイル名', 'ファイルID', 'サイズKB', '登録日時', '登録者'];
const DRAW_KEEP_DAYS = 30;

// ---------- 初期化 ----------
function setup() {
  const ss = SpreadsheetApp.getActive();
  // 新規スプレッドシートは米国時間のことがある。日時の文字列を日本時間で読ませる
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  const ensure = (name, cols) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  // 最初のシートが「物件」でなければ改名（CSV取り込みで作った場合）
  const first = ss.getSheets()[0];
  if (!ss.getSheetByName(SHEETS.sites) && first.getRange(1, 1).getValue() === '現場名') first.setName(SHEETS.sites);
  ensure(SHEETS.sites, SITE_COLS);
  ensure(SHEETS.extra, EXTRA_COLS);
  ensure(SHEETS.visits, VISIT_COLS);
  manualSheet();
  const cfg = ensure(SHEETS.config, ['キー', '値']);
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 4, 2).setValues([
      ['パスワード', 'genba'],
      ['記録者', '横山,高橋,小松'],
      ['Excelファイル名', 'R8年度_マンション系.xlsx'],
      ['現地判定m', '300'],
    ]);
  }
}

function cfg(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.config);
  const rows = sh.getDataRange().getValues();
  for (const r of rows) if (r[0] === key) return String(r[1]);
  return '';
}

// 普段のパスワード → 'normal'、会議用パスワード（設定シート、空欄なら無効）→ 'meeting'、
// 職人の個人パスワード（アカウントシート）→ 'worker'
function roleOf(pw) { return auth_(pw).role; }
function auth_(pw) {
  if (!pw) return { role: '' };
  const m = cfg('会議用パスワード');
  if (m && pw === m) return { role: 'meeting' };
  if (pw === cfg('パスワード')) return { role: 'normal' };
  const a = accounts_().find(x => x.pw && x.pw === pw);
  if (a) return { role: 'worker', name: a.name, worker: a.worker || a.name };
  return { role: '' };
}

// ---------- 職人のアカウント（2026-11-09〜） ----------
// アカウントシート：名前・パスワード・担当名（Excelの職人欄に出てくる名前）・メモ
// 職人で入ると、職人欄に担当名がある物件だけを返す。金額は返さない。書けるのは訪問記録（本人名）だけ
const ACCOUNT_SHEET = 'アカウント';
const ACCOUNT_COLS = ['名前', 'パスワード', '担当名', 'メモ'];
function accountSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ACCOUNT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ACCOUNT_SHEET);
    sh.getRange(1, 1, 1, ACCOUNT_COLS.length).setValues([ACCOUNT_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');
  }
  return sh;
}
function accounts_() {
  return readSheet(accountSheet_().getName()).map(r => ({ name: clean(r['名前']), pw: clean(r['パスワード']), worker: clean(r['担当名']) })).filter(a => a.name);
}
const splitW_ = v => String(v || '').split(/[、,，・\s]+/).filter(Boolean);
// 職人が見てよい現場名の集合
function workerSites_(worker) {
  return new Set(allSites().filter(s => splitW_(s['職人']).indexOf(worker) >= 0).map(s => s['現場名']));
}
// 職人ができる操作
const WORKER_ACTIONS = ['visit', 'deleteVisit', 'getDrawing'];
// 会議用（管理者）だけができる操作：職人とパスワードの管理
const ADMIN_ACTIONS = ['listAccounts', 'saveAccount', 'deleteAccount', 'setSharedPw'];

// ---------- 管理（会議用パスワードのときだけ） ----------
function listAccounts() {
  return { ok: true, accounts: accounts_().map(a => ({ name: a.name, pw: a.pw, worker: a.worker })), normalPw: cfg('パスワード'), meetingPw: cfg('会議用パスワード') };
}
// 使えるパスワードか（短すぎ・ほかの人と同じ はだめ）
function pwProblem_(pw, exceptName) {
  if (!pw) return '';
  if (pw.length < 6) return 'パスワードは6文字以上にしてください';
  if (pw === cfg('パスワード') || pw === cfg('会議用パスワード')) return '普段用・会議用と同じパスワードは使えません';
  if (accounts_().some(a => a.pw === pw && a.name !== exceptName)) return 'ほかの職人と同じパスワードです';
  return '';
}
// 職人を追加・変更。oldName があれば変更（名前の付け替えも可）。パスワードが空ならログインはできない（候補に出るだけ）
function saveAccount(b) {
  // 新しいパスワードは newPw で受け取る（pw はログインのパスワードの欄なので使わない）
  const name = clean(b.name), pw = clean(b.newPw), worker = clean(b.worker) || name, oldName = clean(b.oldName);
  if (!name) return { ok: false, error: '名前を入れてください' };
  const p = pwProblem_(pw, oldName || name); if (p) return { ok: false, error: p };
  const sh = accountSheet_(), v = sh.getDataRange().getValues();
  let r = 0;
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === (oldName || name)) r = i + 1;
    else if (v[i][0] === name) return { ok: false, error: '同じ名前の職人がいます' };
  }
  if (b.isNew && r) return { ok: false, error: '同じ名前の職人がいます（変えるときは一覧の「保存」で）' }; // 追加で既存を上書きしない
  const row = [txt(name), txt(pw), txt(worker), txt(r ? v[r - 1][3] : b.memo)];
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  return { ok: true };
}
function deleteAccount(b) {
  const sh = accountSheet_(), r = findRow(sh, clean(b.name));
  if (!r) return { ok: false, error: '見つかりません' };
  sh.deleteRow(r);
  return { ok: true };
}
// 普段用・会議用のパスワードを変える
function setSharedPw(b) {
  const pw = clean(b.newPw), key = b.kind === 'meeting' ? '会議用パスワード' : 'パスワード';
  if (pw.length < 6) return { ok: false, error: 'パスワードは6文字以上にしてください' };
  const other = b.kind === 'meeting' ? cfg('パスワード') : cfg('会議用パスワード');
  if (pw === other || accounts_().some(a => a.pw === pw)) return { ok: false, error: 'ほかのパスワードと同じです' };
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.config), v = sh.getDataRange().getValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === key) { sh.getRange(i + 1, 2).setValue(txt(pw)); return { ok: true }; }
  sh.appendRow([key, txt(pw)]);
  return { ok: true };
}

function AMOUNT_SHEET_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(AMOUNT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AMOUNT_SHEET);
    sh.getRange(1, 1, 1, AMOUNT_COLS.length).setValues([AMOUNT_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return AMOUNT_SHEET;
}

// ---------- Web API ----------
// 2026-11-11 Supabase 版に切り替えた。true の間は、スプレッドシート版の読み書きを止める
// （古い画面を開きっぱなしの人が書いても、シートにだけ入って見えなくなるのを防ぐ）。
// 毎朝の取り込みも Supabase にだけ入れる。シートは2026-11-11時点の控えとして残る
const SB_ONLY = true;
const SWITCHED_MSG = '新しい版に切り替えました。ページを開き直してください（ホーム画面のアイコンからでも同じです）';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (SB_ONLY) return json({ ok: false, error: SWITCHED_MSG });
  // 記録者の名前だけはパスワードなしで返す（ログイン画面で最初から選べるように。2026-11-09 Yoko 了承）
  if (p.mode === 'users') return json({ ok: true, users: cfg('記録者').split(',').map(s => s.trim()).filter(Boolean) });
  const a = auth_(p.pw), role = a.role;
  if (!role) return json({ ok: false, error: 'パスワードが違います' });
  let sites = allSites(), extra = readSheet(SHEETS.extra), visits = readSheet(SHEETS.visits), drawings = readSheet(drawSheet_().getName());
  if (role === 'worker') {
    // 担当の物件だけ。ほかの物件の情報は送らない
    const mine = new Set(sites.filter(s => splitW_(s['職人']).indexOf(a.worker) >= 0).map(s => s['現場名']));
    const only = r => mine.has(r['現場名']);
    sites = sites.filter(only); extra = extra.filter(only); visits = visits.filter(only); drawings = drawings.filter(only);
  }
  return json({
    ok: true,
    role,
    me: role === 'worker' ? a.name : undefined,
    // 職人の候補（名前だけ。パスワードは返さない）。物件の職人欄に無い新しい人もプルダウンに出す
    staff: role === 'worker' ? undefined : accounts_().map(x => x.worker || x.name),
    amounts: role === 'meeting' ? amountsWithEdits_() : undefined,
    sites, extra, visits, drawings,
    users: role === 'worker' ? [a.name] : cfg('記録者').split(',').map(s => s.trim()).filter(Boolean),
    radius: Number(cfg('現地判定m')) || 300,
    updated: new Date().toISOString(),
    // 画面上部に出す「Excel取り込み」日時と、そのときのExcelの更新日時
    importedAt: PropertiesService.getScriptProperties().getProperty('importedAt') || '',
    excelAt: PropertiesService.getScriptProperties().getProperty('excelAt') || '',
  });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad json' }); }
  // Supabase 版の画面からの写真・図面の保存（パスワードは Supabase に確かめる）
  if (body.action === 'sbPhotos' || body.action === 'sbDrawing') return json(sbUpload_(body));
  if (SB_ONLY) return json({ ok: false, error: SWITCHED_MSG });
  const who = auth_(body.pw);
  if (!who.role) return json({ ok: false, error: 'パスワードが違います' });
  // 金額を書けるのは会議用で入ったときだけ（普段用には金額を見せていないので、書かせもしない）
  if (who.role !== 'meeting' && body.fields) delete body.fields['金額'];
  // 職人・パスワードの管理は会議用だけ
  if (ADMIN_ACTIONS.indexOf(body.action) >= 0) {
    if (who.role !== 'meeting') return json({ ok: false, error: '会議用パスワードで入ったときだけ使えます' });
    if (body.action === 'listAccounts') return json(listAccounts());
  }
  if (who.role === 'worker') {
    // 職人：できる操作と物件を絞る。記録者は本人の名前に固定（なりすまし防止）
    if (WORKER_ACTIONS.indexOf(body.action) < 0) return json({ ok: false, error: 'この操作はできません（業務部に頼んでください）' });
    body.user = who.name;
    if (body.action === 'getDrawing') {
      const row = readSheet(drawSheet_().getName()).find(r => r['ファイルID'] === body.id);
      if (!row || !workerSites_(who.worker).has(row['現場名'])) return json({ ok: false, error: '図面が見つかりません' });
    } else if (!workerSites_(who.worker).has(body.site)) return json({ ok: false, error: '担当の物件ではありません' });
    if (body.action === 'deleteVisit' && body.by !== who.name) return json({ ok: false, error: '自分の記録だけ削除できます' });
  }
  // 図面を開く（読むだけなのでロック不要）
  if (body.action === 'getDrawing') return json(getDrawing(body));
  // 写真・図面のドライブ保存は時間がかかるので、ロックを取る前に済ませる（他の人の保存を待たせない）
  if (body.action === 'visit') body.photoIds = (body.photos || []).map((p, i) => savePhoto(body.site, p, i)).filter(Boolean);
  if (body.action === 'addDrawing') { try { body.file = saveDrawingFile_(body); } catch (err) { return json({ ok: false, error: '図面を保存できませんでした：' + err.message }); } }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json({ ok: false, error: '混み合っています。少し待ってもう一度押してください' });
  try {
    if (body.action === 'visit') return json(addVisit(body));
    if (body.action === 'extra') return json(saveExtra(body));
    if (body.action === 'deleteVisit') return json(deleteVisit(body));
    if (body.action === 'addSite') return json(addSite(body));
    if (body.action === 'updateSite') return json(updateSite(body));
    if (body.action === 'deleteSite') return json(deleteSite(body));
    if (body.action === 'setLocation') return json(setLocation(body));
    if (body.action === 'setFlag') return json(setFlag(body));
    if (body.action === 'editSite') return json(editSite(body));
    if (body.action === 'saveAccount') return json(saveAccount(body));
    if (body.action === 'deleteAccount') return json(deleteAccount(body));
    if (body.action === 'setSharedPw') return json(setSharedPw(body));
    if (body.action === 'addDrawing') return json(addDrawingRow_(body));
    if (body.action === 'deleteDrawing') return json(deleteDrawing(body));
    return json({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}

// ---------- 物件（アプリで追加・位置の修正） ----------
function manualSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(MANUAL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MANUAL_SHEET);
    sh.getRange(1, 1, 1, MANUAL_COLS.length).setValues([MANUAL_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('D:D').setNumberFormat('@'); // 消検月 "2026-11" を日付にしない
  }
  return sh;
}

// Excelの物件＋アプリで追加した物件（Excelに同名があるものは除く）
function allSites() {
  const xs = readSheet(SHEETS.sites);
  const names = new Set(xs.map(s => s['現場名']));
  manualSheet();
  const ms = readSheet(MANUAL_SHEET).filter(m => !names.has(m['現場名'])).map(m => ({ ...m, '手入力': 1, '要確認': m['緯度'] === '' ? '位置未設定' : '' }));
  const all = xs.concat(ms), by = {};
  all.forEach(s => (by[s['現場名']] = s));
  flagSheet();
  readSheet(FLAG_SHEET).forEach(f => {
    const s = by[f['現場名']], k = f['項目'];
    if (!s || EDIT_KEYS.indexOf(k) < 0) return; // 金額はここでは扱わない（会議用だけ amountsWithEdits_ で）
    if (String(s[k] || '') !== String(f['元の値'] || '')) return; // Excel側が変わった
    s['元_' + k] = String(s[k] || '');
    s[k] = String(f['値'] || '');
    s['アプリ_' + k] = 1;
  });
  return all;
}

// ---------- アプリでの編集（2026-11-09〜） ----------
// Excelの物件の 元請・職人・消検月・注意点・図面〜下検査・金額 をアプリで変えた記録はシート「フラグ変更」に置く
// （名前は昔のまま。中身は全項目の変更）。Excelがその後も同じ値ならアプリの値、Excel側が変わったらExcelを優先。
// アプリで追加した物件は、元請〜注意点は物件手入力シートを直接書き換え、図面〜下検査・金額はここ（元の値は空）
const EDIT_KEYS = ['元請', '職人', '消検月', '注意点'].concat(FLAG_KEYS);
// いまのExcel側の値（記録の「元の値」）。アプリで追加した物件は空
function baseValue_(site, key) {
  if (key === '金額') { const a = readSheet(AMOUNT_SHEET_()).find(r => r['現場名'] === site); return a ? String(a['金額（千円）']) : ''; }
  const s = readSheet(SHEETS.sites).find(r => r['現場名'] === site);
  return s ? String(s[key] || '') : '';
}
// 1項目の変更を記録（Excelと同じ値に戻したら記録を消す）
function setEdit_(site, key, value, base, user) {
  const sh = flagSheet(), v = sh.getDataRange().getValues();
  const same = String(value) === (FLAG_KEYS.indexOf(key) >= 0 ? normCircle(base) : String(base));
  const row = [txt(site), key, txt(value), txt(base), new Date(), txt(user)];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === site && v[i][1] === key) {
      if (same) sh.deleteRow(i + 1); else sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  if (!same) sh.appendRow(row);
}
// 値を確かめる。だめなら理由を返す
function checkEdit_(key, v) {
  if (FLAG_KEYS.indexOf(key) >= 0) return FLAG_VALUES.indexOf(v) >= 0 ? '' : key + 'の値が不正です';
  if (key === '消検月') return v === '' || /^\d{4}-\d{2}$/.test(v) ? '' : '消検月は 2026-12 の形で';
  if (key === '金額') return v === '' || /^\d+(\.\d+)?$/.test(v) ? '' : '金額は数字（千円）で';
  return EDIT_KEYS.indexOf(key) >= 0 ? '' : '変えられない項目です：' + key;
}
// 物件の項目を変える。fields = { 元請: 'CASE', 図面: '〇', 金額: '1340', … }（変えた項目だけ）
// 金額は会議用で入ったときだけ（doPost で外している）
function editSite(b) {
  const site = b.site, f = b.fields || {};
  for (const k in f) { const e = checkEdit_(k, clean(f[k])); if (e) return { ok: false, error: e }; }
  const inExcel = !!findRow(SpreadsheetApp.getActive().getSheetByName(SHEETS.sites), site);
  const mr = inExcel ? 0 : findRow(manualSheet(), site);
  if (!inExcel && !mr) return { ok: false, error: '物件が見つかりません' };
  for (const k in f) {
    const v = clean(f[k]);
    const col = ['元請', '職人', '消検月', '注意点'].indexOf(k);
    if (mr && col >= 0) { manualSheet().getRange(mr, col + 2).setValue(txt(v)); continue; } // 手入力の物件はシートを直接
    setEdit_(site, k, v, inExcel ? baseValue_(site, k) : '', b.user);
  }
  if (mr) manualSheet().getRange(mr, 11, 1, 2).setValues([[new Date(), txt(b.user)]]);
  return { ok: true };
}
// 会議用に返す金額：Excelの金額＋アプリで入れた金額（Excel側が変わっていなければアプリの値）
function amountsWithEdits_() {
  const am = {}; readSheet(AMOUNT_SHEET_()).forEach(r => (am[r['現場名']] = r['金額（千円）']));
  readSheet(FLAG_SHEET).forEach(f => {
    if (f['項目'] !== '金額') return;
    const n = f['現場名'], base = String(f['元の値'] || '');
    if (String(am[n] == null ? '' : am[n]) !== base) return; // Excel側が変わった
    am[n] = f['値'] === '' ? '' : Number(f['値']);
  });
  return Object.keys(am).map(n => ({ '現場名': n, '金額（千円）': am[n] }));
}

function flagSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FLAG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(FLAG_SHEET);
    sh.getRange(1, 1, 1, FLAG_COLS.length).setValues([FLAG_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 図面〜下検査を1つ切り替える（詳細のタップ）。元の値はサーバーでExcelから取る
function setFlag(b) {
  if (FLAG_KEYS.indexOf(b.key) < 0 || FLAG_VALUES.indexOf(b.value) < 0) return { ok: false, error: '項目か値が不正です' };
  setEdit_(b.site, b.key, b.value, baseValue_(b.site, b.key), b.user);
  return { ok: true };
}

function normCircle(v) { return v === '○' ? '〇' : String(v || ''); }

// 取り込み後、Excel側で変わった項目のアプリ記録を消す（古い記録が後で生き返らないように）
function pruneFlags_() {
  const sh = flagSheet(), v = sh.getDataRange().getValues();
  const xs = {}; readSheet(SHEETS.sites).forEach(s => (xs[s['現場名']] = s));
  const am = {}; readSheet(AMOUNT_SHEET_()).forEach(r => (am[r['現場名']] = r['金額（千円）']));
  for (let i = v.length - 1; i >= 1; i--) {
    const s = xs[v[i][0]]; if (!s) continue; // Excelに無い物件（アプリで追加）は消さない
    const now = v[i][1] === '金額' ? (am[v[i][0]] == null ? '' : am[v[i][0]]) : s[v[i][1]];
    if (String(now == null ? '' : now) !== String(v[i][3] || '')) sh.deleteRow(i + 1);
  }
}

function findRow(sh, name) {
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === name) return i + 1;
  return 0;
}

function clean(s) { return String(s == null ? '' : s).trim(); }

// 人が入力した文字をシートに書くとき用。先頭に ' を付けて「文字」として置く。
// こうしないと「=…」が数式として動き（他のシートの中身を読めてしまう）、「11/10」「1-2」は日付に化ける。
// 読み返すと ' は付かない（getValues は中身だけ返す）
function txt(s) {
  const v = clean(s);
  return v === '' ? '' : "'" + v;
}

// 住所だけで緯度経度が無ければ座標を取る
function locate_(addr, lat, lng) {
  if (lat !== '' && lat != null && lng !== '' && lng != null) return { lat: Number(lat), lng: Number(lng) };
  if (!addr) return { lat: '', lng: '' };
  const g = geocode(/^北海道/.test(addr) ? addr : '北海道' + addr);
  return g || { lat: '', lng: '' };
}

function addSite(b) {
  const name = norm(b.name);
  if (!name) return { ok: false, error: '現場名を入れてください' };
  if (findRow(SpreadsheetApp.getActive().getSheetByName(SHEETS.sites), name) || findRow(manualSheet(), name)) return { ok: false, error: '同じ現場名がすでにあります' };
  const addr = clean(b.addr) || nameToAddress(name);
  const p = locate_(addr, b.lat, b.lng);
  manualSheet().appendRow([txt(name), txt(b.moto), txt(b.worker), txt(b.ym), txt(b.note), txt(addr), p.lat, p.lng, new Date(), txt(b.user), '', '']);
  // 図面〜下検査・金額も一緒に入れられる（金額は会議用のときだけ doPost を通ってくる）
  const f = b.fields || {};
  for (const k in f) { const v = clean(f[k]); if (v !== '' && (FLAG_KEYS.indexOf(k) >= 0 || k === '金額') && !checkEdit_(k, v)) setEdit_(name, k, v, '', b.user); }
  return { ok: true, name, lat: p.lat, lng: p.lng, addr };
}

function updateSite(b) {
  const sh = manualSheet(), r = findRow(sh, b.site);
  if (!r) return { ok: false, error: 'アプリで追加した物件だけ編集できます（Excelの物件はExcelで直してください）' };
  const cur = sh.getRange(r, 1, 1, MANUAL_COLS.length).getValues()[0];
  const addr = clean(b.addr);
  const moved = addr !== String(cur[5]) || b.lat !== undefined;
  const p = moved ? locate_(addr, b.lat, b.lng) : { lat: cur[6], lng: cur[7] };
  sh.getRange(r, 2, 1, 7).setValues([[txt(b.moto), txt(b.worker), txt(b.ym), txt(b.note), txt(addr), p.lat, p.lng]]);
  sh.getRange(r, 11, 1, 2).setValues([[new Date(), txt(b.user)]]);
  return { ok: true, lat: p.lat, lng: p.lng };
}

// アプリで追加した物件を消す。訪問記録と物件追加の行は残す（同じ名前で登録し直せば戻る）
function deleteSite(b) {
  const sh = manualSheet(), r = findRow(sh, b.site);
  if (!r) return { ok: false, error: 'アプリで追加した物件だけ削除できます' };
  sh.deleteRow(r);
  return { ok: true };
}

// どの物件でも位置（住所・緯度経度）を直せる。Excelの物件は次の取り込みでもこの値が引き継がれる
function setLocation(b) {
  const addr = clean(b.addr);
  const p = locate_(addr, b.lat, b.lng);
  if (p.lat === '') return { ok: false, error: '位置が分かりませんでした。地図をタップして指定してください' };
  let sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.sites), r = findRow(sh, b.site);
  if (r) {
    const c = n => SITE_COLS.indexOf(n) + 1;
    sh.getRange(r, c('住所')).setValue(txt(addr));
    sh.getRange(r, c('緯度'), 1, 3).setValues([[p.lat, p.lng, '']]); // 緯度・経度・要確認
    return { ok: true, lat: p.lat, lng: p.lng };
  }
  sh = manualSheet(); r = findRow(sh, b.site);
  if (!r) return { ok: false, error: '物件が見つかりません' };
  sh.getRange(r, 6, 1, 3).setValues([[txt(addr), p.lat, p.lng]]);
  return { ok: true, lat: p.lat, lng: p.lng };
}

function addVisit(b) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.visits);
  const urls = b.photoIds || []; // doPost でロックの前に保存済み
  const row = [new Date(), txt(b.user), txt(b.site), txt(b.stage), txt(b.memo), urls.join('\n'), b.lat || '', b.lng || '', b.dist === undefined || b.dist === null ? '' : Math.round(b.dist)];
  sh.appendRow(row);
  return { ok: true, photos: urls };
}

// 訪問記録を1行消す。日時・現場名・記録者が一致する行。写真はドライブのゴミ箱へ（30日は戻せる）
function deleteVisit(b) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.visits);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const r = data[i];
    const at = r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss") : String(r[0]);
    if (at === b.at && r[2] === b.site && String(r[1]) === String(b.by || '')) {
      String(r[5] || '').split('\n').filter(Boolean).forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: '該当する記録が見つかりません（更新してからもう一度）' };
}

// 日付 "2026-12-01" と時刻 "10:00"（任意）を、シートが日時として読む文字列にする
function dateTime(d, t) {
  if (!d) return '';
  return d.replace(/-/g, '/') + (t ? ' ' + t : '');
}

// 送られてきた項目だけ書き換える（送られていない項目は今の値のまま）。
// 画面を開きっぱなしの人が保存しても、その間に別の人が入れた値を古い値で戻さないため
function saveExtra(b) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.extra);
  // 列が増えたとき（正式名称）は見出しを足す
  if (sh.getLastColumn() < EXTRA_COLS.length) sh.getRange(1, 1, 1, EXTRA_COLS.length).setValues([EXTRA_COLS]).setFontWeight('bold');
  const data = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), EXTRA_COLS.length).getValues();
  let r = 0;
  for (let i = 1; i < data.length; i++) if (data[i][0] === b.site) { r = i + 1; break; }
  // 今の行（無ければ空）。日時は Date のまま、文字は ' を付け直して書く（数式・日付化を防ぐ）
  const old = r ? data[r - 1] : [b.site, '', '', '', '', '', '', ''];
  const cur = [txt(old[0]), old[1], old[2], txt(old[3]), txt(old[4]), new Date(), txt(b.user), txt(old[7])];
  if (b.fireDate !== undefined) cur[1] = dateTime(b.fireDate, b.fireTime);
  if (b.bldgDate !== undefined) cur[2] = dateTime(b.bldgDate, b.bldgTime);
  if (b.delivery !== undefined) cur[3] = txt(b.delivery);
  if (b.progress !== undefined) cur[4] = txt(b.progress);
  if (b.alias !== undefined) cur[7] = txt(b.alias);
  if (r) sh.getRange(r, 1, 1, cur.length).setValues([cur]);
  else sh.appendRow(cur);
  return { ok: true };
}

// ---------- 置き場所（ドライブ直下を散らかさない） ----------
// 「仕事用/現場マップ」フォルダ。見つけたIDは覚えておく
function appFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('appFolderId');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  let found = null;
  const it = DriveApp.getFoldersByName('現場マップ');
  while (it.hasNext()) {
    const f = it.next(); if (f.isTrashed()) continue;
    const ps = f.getParents(); let underWork = false;
    while (ps.hasNext()) if (ps.next().getName() === '仕事用') underWork = true;
    if (underWork || !found) found = f;
    if (underWork) break;
  }
  if (!found) found = DriveApp.createFolder('現場マップ');
  props.setProperty('appFolderId', found.getId());
  return found;
}
function subFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ---------- 図面PDF ----------
function drawSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DRAW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DRAW_SHEET);
    sh.getRange(1, 1, 1, DRAW_COLS.length).setValues([DRAW_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
// ファイルの保存（ロックの外）。data は base64（"data:…;base64," 付きでも可）
function saveDrawingFile_(b) {
  const name = clean(b.name).replace(/[\\/:*?"<>|]/g, '_') || '図面.pdf';
  if (!/\.pdf$/i.test(name)) throw new Error('PDFだけ保存できます');
  const data = String(b.data || '').replace(/^data:[^,]*,/, '');
  const bytes = Utilities.base64Decode(data);
  if (bytes.length > 30 * 1024 * 1024) throw new Error('30MBまでです');
  const folder = subFolder_(subFolder_(appFolder_(), '図面'), clean(b.site).replace(/[\\/:*?"<>|]/g, '_'));
  const file = folder.createFile(Utilities.newBlob(bytes, 'application/pdf', name));
  // 2026-11-10 Yoko 決定：図面も写真と同じく「リンクを知っている人は閲覧可」にして、ドライブのビューアで直接開く（GASを通すと10秒前後かかるため）
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: file.getId(), name, kb: Math.round(bytes.length / 1024) };
}
function addDrawingRow_(b) {
  drawSheet_().appendRow([txt(b.site), txt(b.file.name), b.file.id, b.file.kb, new Date(), txt(b.user)]);
  return { ok: true, id: b.file.id };
}
// 図面を返す（base64）。シートに載っているIDだけ（ドライブの他のファイルは読ませない）
function getDrawing(b) {
  const v = drawSheet_().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][2] === b.id) {
      const blob = DriveApp.getFileById(b.id).getBlob();
      return { ok: true, name: v[i][1], data: Utilities.base64Encode(blob.getBytes()) };
    }
  }
  return { ok: false, error: '図面が見つかりません（削除済みかもしれません）' };
}
function deleteDrawing(b) {
  const sh = drawSheet_(), v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (v[i][2] === b.id) {
      try { DriveApp.getFileById(b.id).setTrashed(true); } catch (e) {}
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: '図面が見つかりません' };
}
// 一覧にある図面を全部「リンクを知っている人は閲覧可」にする（切り替え前に入れた図面用。何度実行してもよい）
function shareDrawings() {
  const v = drawSheet_().getDataRange().getValues(); let n = 0;
  for (let i = 1; i < v.length; i++) {
    try {
      const f = DriveApp.getFileById(v[i][2]);
      if (f.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); n++; }
    } catch (e) {}
  }
  return n;
}

// 消防検査から DRAW_KEEP_DAYS 日たった図面をゴミ箱へ。消検日（物件追加）、なければ消検月の月末で数える。
// 毎朝の importIfChanged から呼ぶ
function cleanupDrawings_() {
  const sh = drawSheet_(), v = sh.getDataRange().getValues();
  if (v.length < 2) return 0;
  const fire = {};
  readSheet(SHEETS.sites).concat(readSheet(MANUAL_SHEET)).forEach(s => {
    const ym = String(s['消検月'] || '');
    if (/^\d{4}-\d{2}/.test(ym)) fire[s['現場名']] = new Date(+ym.slice(0, 4), +ym.slice(5, 7), 0); // 月末
  });
  readSheet(SHEETS.extra).forEach(e => { if (e['消検日']) fire[e['現場名']] = new Date(String(e['消検日']).slice(0, 10).replace(/-/g, '/')); });
  const limit = new Date(); limit.setDate(limit.getDate() - DRAW_KEEP_DAYS);
  let n = 0;
  for (let i = v.length - 1; i >= 1; i--) {
    const d = fire[v[i][0]], added = v[i][4] instanceof Date ? v[i][4] : new Date(0);
    // 消検から30日、かつ登録から30日（消検後に入れた図面をすぐ消さない）
    if (d && d < limit && added < limit) {
      try {
        const f = DriveApp.getFileById(v[i][2]); const ps = f.getParents(); f.setTrashed(true);
        // 物件のフォルダが空になったらフォルダもゴミ箱へ
        while (ps.hasNext()) { const p = ps.next(); if (p.getName() !== '図面' && !p.getFiles().hasNext() && !p.getFolders().hasNext()) p.setTrashed(true); }
      } catch (e) {}
      sh.deleteRow(i + 1); n++;
    }
  }
  return n;
}

function savePhoto(site, dataUrl, idx) {
  try {
    const m = /^data:(image\/\w+);base64,(.+)$/.exec(dataUrl);
    if (!m) return '';
    const folder = subFolder_(appFolder_(), PHOTO_FOLDER); // 仕事用/現場マップ/現場写真
    const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
    const name = `${site}_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss')}_${idx + 1}.${ext}`;
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getId();
  } catch (err) {
    return '';
  }
}

function getFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function readSheet(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const head = v[0];
  return v.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {};
    head.forEach((h, i) => {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");
      o[h] = x;
    });
    return o;
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Excel 取り込み ----------
// 毎朝トリガーで呼ぶ。ドライブのExcelが前回取り込み後に更新されていたときだけ取り込む。
// Excelはパソコンのタスク「現場マップExcel複製」が火曜・土曜の夜にサーバーから複製している。
function importIfChanged() {
  if (!SB_ONLY) { try { cleanupDrawings_(); } catch (e) { console.error('図面の整理に失敗: ' + e.message); } } // 古い図面の整理（毎朝。切り替え後は Supabase 側の sbCleanup_ だけ）
  try { sbCleanup_(); } catch (e) { console.error('Supabase の図面・写真の整理に失敗: ' + e.message); }
  const xlsx = excelFile_();
  const updated = xlsx.getLastUpdated().getTime();
  const props = PropertiesService.getScriptProperties();
  if (Number(props.getProperty('excelUpdated') || 0) >= updated) return '変更なし';
  const n = importFromExcel(xlsx);
  props.setProperty('excelUpdated', String(updated));
  return n + '件取り込み';
}

// 取り込むExcelを1つに決める。同じ名前のファイルが複数あっても、ゴミ箱のものは除き、
// 「現場マップ」フォルダにあるものを優先、その中で更新日時がいちばん新しいものを使う
function excelFile_() {
  const fname = cfg('Excelファイル名');
  const it = DriveApp.getFilesByName(fname);
  let best = null, bestScore = -1;
  while (it.hasNext()) {
    const f = it.next();
    if (f.isTrashed()) continue;
    let inFolder = false;
    const ps = f.getParents();
    while (ps.hasNext()) if (ps.next().getName() === '現場マップ') inFolder = true;
    const score = (inFolder ? 1e15 : 0) + f.getLastUpdated().getTime();
    if (score > bestScore) { best = f; bestScore = score; }
  }
  if (!best) throw new Error('Excelが見つかりません: ' + fname);
  return best;
}

function importFromExcel(file) {
  const xlsx = file && file.getBlob ? file : excelFile_(); // エディタやトリガーから呼ぶと file は無い／イベントになる
  // xlsx → 一時スプレッドシートに変換（Drive API 必要）
  const tmp = Drive.Files.create(
    { name: '_tmp_import_' + Date.now(), mimeType: 'application/vnd.google-apps.spreadsheet' },
    xlsx.getBlob()
  );
  const ss = SpreadsheetApp.openById(tmp.id);
  const sitesSh = SpreadsheetApp.getActive().getSheetByName(SHEETS.sites);
  const old = {};
  // アプリで先に登録してあった物件がExcelに載ったら、その住所・座標を使う
  manualSheet();
  readSheet(MANUAL_SHEET).forEach(r => (old[r['現場名']] = { '住所': r['住所'], '緯度': r['緯度'], '経度': r['経度'] }));
  readSheet(SHEETS.sites).forEach(r => (old[r['現場名']] = r));

  const now = new Date();
  const C = n => SITE_COLS.indexOf(n);
  let rows, amts, sb = null;
  try {
    ({ rows, amts } = parseExcel_(ss, old, now));
    // Supabase 版にも入れる（併用の間）。住所・座標は Supabase 側のいまの値から
    try {
      const sbOld = sbRpc_('gas_sites', { p_token: sbToken_() });
      sb = { old: sbOld, ...parseExcel_(ss, sbOld, now) };
    } catch (e) { console.error('Supabase の取り込みの準備に失敗: ' + e.message); }
  } finally {
    DriveApp.getFileById(tmp.id).setTrashed(true);
  }
  if (rows.length === 0) throw new Error('物件行が0件でした');
  if (!SB_ONLY) writeImport_(rows, amts, now, xlsx, old);
  if (sb) {
    try { sbPushImport_(sb, now, xlsx); } catch (e) { console.error('Supabase への取り込みに失敗: ' + e.message); }
  }
  return rows.length;
}

// Excel（一時スプレッドシートに変換したもの）から物件行と金額を作る。old＝いまの住所・座標（現場名→{住所,緯度,経度}）
function parseExcel_(ss, old, now) {
  const rows = [], amts = [], seen = {};
  const fy = now.getMonth() < 3 ? now.getFullYear() - 1 : now.getFullYear(); // 今年度（4月始まり）
  {
    ss.getSheets().forEach(sh => {
      if (!/^R\d+年度_/.test(sh.getName())) return;
      const moto = sh.getName().replace(/^R\d+年度_/, '');
      const v = sh.getDataRange().getValues();
      if (v.length < 8) return; // 物件行が無いシート
      // 8行目以降: B=項番 C=現場名 D=職人 E=消検月 F..M=○ N=金額 O=注意点 P=消検年月(日付)
      const baseYear = Number((v[1] || [])[17]) || fy; // R2 = 集計基準年
      for (let i = 7; i < v.length; i++) {
        const r = v[i];
        let name = norm(r[2]);
        if (!name) continue;
        // 同じ現場名が別の元請シートにもあると、記録や消検日が混ざる。2つ目以降は「名前（元請）」にして分ける
        let dup = '';
        if (seen[name]) { name = `${name}（${moto}）`; dup = '同名あり（Excelで名前を分けてください）'; }
        seen[name] = 1;
        let ym = '';
        if (r[15] instanceof Date) ym = Utilities.formatDate(r[15], 'Asia/Tokyo', 'yyyy-MM');
        else if (r[4] !== '' && !isNaN(r[4])) {
          const m = Number(r[4]);
          ym = `${m >= 4 ? baseYear : baseYear + 1}-${String(m).padStart(2, '0')}`;
        }
        const prev = old[name] || {};
        let lat = prev['緯度'] || '', lng = prev['経度'] || '';
        // 地図で位置を決めた物件（住所は空）に、推定住所を勝手に入れない
        let addr = lat !== '' ? String(prev['住所'] || '') : (prev['住所'] || nameToAddress(name) || '');
        if ((lat === '' || lng === '') && addr) {
          const g = geocode(addr);
          if (g) { lat = g.lat; lng = g.lng; }
        }
        // 要確認は毎回作り直す（座標が取れたら ⚠ を消す）
        const chk = dup || (lat !== '' ? '' : addr ? '座標取得失敗' : '住所不明');
        amts.push([txt(name), typeof r[13] === 'number' ? r[13] : '', now]); // N列 金額（千円）
        rows.push([txt(name), txt(moto), txt(r[3]), ym, txt(r[5]), txt(r[6]), txt(r[7]), txt(r[8]), txt(r[9]), txt(r[10]), txt(r[11]), txt(r[12]), txt(r[14]), txt(addr), lat, lng, chk, now]);
      }
    });
  }
  return { rows, amts };
}

// スプレッドシート版への書き込み
// old＝Excelを読む前の住所・座標（読んでいる間に「位置を直す」で変わったものを見分ける）
function writeImport_(rows, amts, now, xlsx, old) {
  const C = n => SITE_COLS.indexOf(n);
  const sitesSh = SpreadsheetApp.getActive().getSheetByName(SHEETS.sites);
  // ここから書き込み。アプリからの保存（位置を直す等）とぶつからないようロックを取る
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    // 読み込みの間に「位置を直す」で変わった座標があれば、そちらを使う
    const nowSites = {}; readSheet(SHEETS.sites).forEach(s => (nowSites[s['現場名']] = s));
    rows.forEach(row => {
      const name = String(row[0]).replace(/^'/, ''), cur = nowSites[name], prev = old[name] || {};
      if (cur && cur['緯度'] !== '' && String(cur['緯度']) !== String(prev['緯度'] || '')) {
        row[C('住所')] = txt(cur['住所']); row[C('緯度')] = cur['緯度']; row[C('経度')] = cur['経度'];
        if (!String(row[C('要確認')]).startsWith('同名')) row[C('要確認')] = '';
      }
    });
    // 消検月 "2026-11" を日付に化けさせない（文字列のまま置く）
    sitesSh.getRange(2, C('消検月') + 1, rows.length, 1).setNumberFormat('@');
    sitesSh.getRange(1, 1, 1, SITE_COLS.length).setValues([SITE_COLS]).setFontWeight('bold');
    // 先に上書きしてから余った行だけ消す（途中で「0件」の瞬間を作らない）
    sitesSh.getRange(2, 1, rows.length, SITE_COLS.length).setValues(rows);
    const extraRows = sitesSh.getLastRow() - 1 - rows.length;
    if (extraRows > 0) sitesSh.getRange(rows.length + 2, 1, extraRows, sitesSh.getLastColumn()).clearContent();
    const amSh = SpreadsheetApp.getActive().getSheetByName(AMOUNT_SHEET_());
    amSh.getRange(2, 1, amts.length, AMOUNT_COLS.length).setValues(amts);
    const extraAm = amSh.getLastRow() - 1 - amts.length;
    if (extraAm > 0) amSh.getRange(amts.length + 2, 1, extraAm, AMOUNT_COLS.length).clearContent();
    pruneFlags_();
    const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");
    PropertiesService.getScriptProperties().setProperties({ importedAt: fmt(now), excelAt: fmt(xlsx.getLastUpdated()) });
  } finally {
    lock.releaseLock();
  }
  return rows.length;
}

const GEO_MEMO_ = {}; // 同じ住所を1回の実行で何度も引かない（取り込みでスプレッドシート版とSupabase版の2回作るため）
function geocode(addr) {
  if (addr in GEO_MEMO_) return GEO_MEMO_[addr];
  return (GEO_MEMO_[addr] = geocodeRaw_(addr));
}
function geocodeRaw_(addr) {
  try {
    const res = Maps.newGeocoder().setRegion('jp').setLanguage('ja').geocode(addr);
    if (res.status !== 'OK' || !res.results.length) return null;
    const loc = res.results[0].geometry.location;
    return { lat: loc.lat, lng: loc.lng };
  } catch (e) { return null; }
}

function norm(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/[！-～]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/　/g, ' ').trim();
}

// 現場名 → 住所（札幌の条丁目表記を読む。読めなければ ''）
const KU = { '美園': '豊平区', '平岸': '豊平区', '月寒西': '豊平区', '月寒東': '豊平区', '中の島': '豊平区', '西岡': '豊平区', '福住': '豊平区', '豊平': '豊平区',
  '東札幌': '白石区', '菊水': '白石区', '発寒': '西区', '二十四軒': '西区', '西野': '西区', '新琴似': '北区', '澄川': '南区', '宮の森': '中央区',
  '厚別中央': '厚別区', '宮の沢': '手稲区', '手稲本町': '手稲区', '本町': '東区' };
function nameToAddress(name) {
  let n = norm(name);
  if (/^(千歳市|恵庭市|北広島市|江別市)/.test(n)) return '北海道' + n.replace(/(MS|マンション|賃貸MS|の集合住宅|ホテル|PROJECT|計画).*$/, '').replace(/\s*[A-D]棟$/, '');
  n = n.replace(/^(GA-|PS-|LE-|NUTS-|CORE-|CORE )/, '');
  n = n.replace(/(MS|BL)?\s*(②|ⅲ|Ⅲ|\(A棟\)|\(B棟\)|[A-D]棟|東棟|西棟|賃貸MS|賃貸|の集合住宅|マンション|HOTEL|ホテル|学生寮|民泊|MS).*$/, '').trim();
  let m;
  if ((m = /^([SN])(\d+)([EW])(\d+)/.exec(n))) {
    const ns = m[1] === 'S' ? '南' : '北', ew = m[3] === 'E' ? '東' : '西';
    return `北海道札幌市${kuFor(ns, +m[2], ew)}${ns}${m[2]}条${ew}${m[4]}丁目`;
  }
  if ((m = /^([南北])\s*(\d+)\s*条?\s*([東西])\s*(\d+)/.exec(n))) {
    return `北海道札幌市${kuFor(m[1], +m[2], m[3])}${m[1]}${m[2]}条${m[3]}${m[4]}丁目`;
  }
  if ((m = /^(麻生町?)\s*(\d+)/.exec(n))) return `北海道札幌市北区麻生町${m[2]}丁目`;
  if ((m = /^(本通|本郷通|南郷通)り?\s*(\d+)丁目?\s*([南北])?/.exec(n))) return `北海道札幌市白石区${m[1]}${m[2]}丁目${m[3] || ''}`;
  if ((m = /^(栄通|西町北|厚別南|青葉町)\s*(\d+)/.exec(n))) {
    const ku = { '栄通': '白石区', '西町北': '西区', '厚別南': '厚別区', '青葉町': '厚別区' }[m[1]];
    return `北海道札幌市${ku}${m[1]}${m[2]}丁目`;
  }
  if ((m = /^白石中央\s*(\d+)-(\d+)/.exec(n))) return `北海道札幌市白石区中央${m[1]}条${m[2]}丁目`;
  if ((m = /^月寒中央通\s*(\d+)/.exec(n))) return `北海道札幌市豊平区月寒中央通${m[1]}丁目`;
  if ((m = /^([一-龥ぁ-んァ-ンの]+?)\s*(\d+)\s*(?:-|条)\s*(\d+)/.exec(n))) {
    return `北海道札幌市${KU[m[1]] || ''}${m[1]}${m[2]}条${m[3]}丁目`;
  }
  if (/^中央区/.test(n)) return '北海道札幌市' + n;
  return '';
}
function kuFor(ns, j, ew) {
  if (ns === '南') return '中央区';
  if (ew === '東') return j >= 6 ? '東区' : '中央区';
  return j >= 10 ? '北区' : '中央区';
}

// ======================================================================
// Supabase 版（2026-11-10〜）
// データは Supabase（プロジェクト genba-map）。写真・図面のドライブ保存と、毎朝のExcel取り込みだけGASでやる。
// GAS → Supabase はトークン（sbToken。GASが自分で作ってスクリプトのプロパティに置き、bootstrap で Supabase に登録）で確かめる。
// 下の URL と publishable キーは公開してよいもの（表は直接読めない。関数だけ呼べる）
// ======================================================================
const SB_URL = 'https://twoyzxitmyusbqlkjzfi.supabase.co';
const SB_KEY = 'sb_publishable_kohm5Neuh80PKUETukktug_5YqJQ1Ii';

function sbRpc_(fn, args) {
  const r = UrlFetchApp.fetch(SB_URL + '/rest/v1/rpc/' + fn, {
    method: 'post', contentType: 'application/json', headers: { apikey: SB_KEY },
    payload: JSON.stringify(args || {}), muteHttpExceptions: true,
  });
  const t = r.getContentText();
  if (r.getResponseCode() >= 300) throw new Error(fn + ': ' + r.getResponseCode() + ' ' + t.slice(0, 300));
  return JSON.parse(t);
}
function sbToken_() {
  const p = PropertiesService.getScriptProperties();
  let t = p.getProperty('sbToken');
  if (!t) { t = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('sbToken', t); }
  return t;
}

// スプレッドシートの中身を丸ごと Supabase に入れる。エディタから手で実行する。
// 1回目：Supabase が空のときだけ通る。2回目以降（切り替え直前の入れ直し）は Supabase の中身を全部消してから入れる
function migrateToSupabase() {
  // 切り替えた後に流すと、Supabase の最新の内容を古いシートで上書きしてしまう
  if (SB_ONLY) throw new Error('切り替え済み（SB_ONLY）。シートは古いので Supabase へは流さない');
  const props = PropertiesService.getScriptProperties();
  manualSheet(); flagSheet(); accountSheet_(); AMOUNT_SHEET_(); drawSheet_();
  const dump = {
    cfg: {
      pw_normal: cfg('パスワード'), pw_meeting: cfg('会議用パスワード'), users: cfg('記録者'), radius: cfg('現地判定m'),
      gas_token: sbToken_(), imported_at: props.getProperty('importedAt') || '', excel_at: props.getProperty('excelAt') || '',
    },
    accounts: readSheet(ACCOUNT_SHEET),
    sites: readSheet(SHEETS.sites),
    manual: readSheet(MANUAL_SHEET),
    edits: readSheet(FLAG_SHEET),
    amounts: readSheet(AMOUNT_SHEET),
    extra: readSheet(SHEETS.extra),
    visits: readSheet(SHEETS.visits),
    drawings: readSheet(DRAW_SHEET),
  };
  const r = sbRpc_('bootstrap', { p_dump: dump, p_token: sbToken_() });
  console.log(JSON.stringify(r));
  return r;
}

// Supabase 版の画面からの写真・図面の保存。誰か・その物件を触ってよいかは Supabase に聞く
function sbUpload_(b) {
  let who;
  try { who = sbRpc_('whoami', { p_pw: b.pw, p_site: b.site }); } catch (e) { return { ok: false, error: '確認できませんでした：' + e.message }; }
  if (!who.ok) return who;
  if (b.action === 'sbPhotos') {
    const ids = (b.photos || []).map((p, i) => savePhoto(b.site, p, i)).filter(Boolean);
    return { ok: true, ids };
  }
  if (who.role === 'worker') return { ok: false, error: 'この操作はできません（業務部に頼んでください）' };
  try { return { ok: true, file: saveDrawingFile_(b) }; } catch (e) { return { ok: false, error: '図面を保存できませんでした：' + e.message }; }
}

// 取り込んだ物件を Supabase へ。sb.old＝Supabase のいまの住所・座標（読んでいる間に位置が直されたか見分ける）
function sbPushImport_(sb, now, xlsx) {
  const strip = v => (typeof v === 'string' ? v.replace(/^'/, '') : v);
  const rows = sb.rows.map(row => {
    const o = {};
    SITE_COLS.forEach((c, i) => { o[c] = c === '取込日時' ? '' : strip(row[i]); });
    const prev = sb.old[o['現場名']] || {};
    o._prevLat = prev['緯度'] == null ? '' : prev['緯度'];
    return o;
  });
  const amts = sb.amts.map(a => ({ '現場名': strip(a[0]), '金額（千円）': a[1] }));
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");
  const r = sbRpc_('gas_import', { p_token: sbToken_(), p_rows: rows, p_amounts: amts, p_imported: fmt(now), p_excel: fmt(xlsx.getLastUpdated()) });
  console.log('Supabase 取り込み: ' + JSON.stringify(r));
}

// 毎朝：消防検査から30日たった図面をゴミ箱へ。訪問記録の削除などで使われなくなった写真・図面もゴミ箱へ（ゴミ箱で30日は戻せる）。
// 併用の間はスプレッドシート版で使っているファイルも残す。念のため、作って3日以内のファイルと、一度に30件を超えるときは触らない
function sbCleanup_() {
  const r = sbRpc_('gas_files', { p_token: sbToken_() });
  const keep = new Set(r.keep || []);
  if (!SB_ONLY) { // 併用の間はスプレッドシート版で使っているファイルも残す
    readSheet(SHEETS.visits).forEach(v => String(v['写真'] || '').split('\n').filter(Boolean).forEach(id => keep.add(id)));
    drawSheet_(); readSheet(DRAW_SHEET).forEach(d => keep.add(String(d['ファイルID'])));
  }
  (r.trash || []).forEach(id => { if (!keep.has(id)) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} } });
  if (keep.size === 0) return;
  const limit = new Date(Date.now() - 3 * 86400000), orphans = [];
  const scan = folder => {
    const fs = folder.getFiles();
    while (fs.hasNext()) { const f = fs.next(); if (!keep.has(f.getId()) && f.getDateCreated() < limit) orphans.push(f); }
  };
  const app = appFolder_();
  scan(subFolder_(app, PHOTO_FOLDER));
  const dws = subFolder_(app, '図面').getFolders();
  while (dws.hasNext()) scan(dws.next());
  if (orphans.length > 30) { console.error('使われていないファイルが' + orphans.length + '件。多すぎるので消さない（確かめてください）'); return; }
  orphans.forEach(f => f.setTrashed(true));
  if (orphans.length) console.log('使われていないファイルをゴミ箱へ: ' + orphans.length + '件');
}
