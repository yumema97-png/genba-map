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

// 普段のパスワード → 'normal'、会議用パスワード（設定シート、空欄なら無効）→ 'meeting'
function roleOf(pw) {
  if (!pw) return '';
  const m = cfg('会議用パスワード');
  if (m && pw === m) return 'meeting';
  return pw === cfg('パスワード') ? 'normal' : '';
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
function doGet(e) {
  const p = (e && e.parameter) || {};
  // 記録者の名前だけはパスワードなしで返す（ログイン画面で最初から選べるように。2026-11-09 Yoko 了承）
  if (p.mode === 'users') return json({ ok: true, users: cfg('記録者').split(',').map(s => s.trim()).filter(Boolean) });
  const role = roleOf(p.pw);
  if (!role) return json({ ok: false, error: 'パスワードが違います' });
  return json({
    ok: true,
    role,
    amounts: role === 'meeting' ? readSheet(AMOUNT_SHEET_()) : undefined,
    sites: allSites(),
    extra: readSheet(SHEETS.extra),
    visits: readSheet(SHEETS.visits),
    users: cfg('記録者').split(',').map(s => s.trim()).filter(Boolean),
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
  if (!roleOf(body.pw)) return json({ ok: false, error: 'パスワードが違います' });
  // 写真のドライブ保存は時間がかかるので、ロックを取る前に済ませる（他の人の保存を待たせない）
  if (body.action === 'visit') body.photoIds = (body.photos || []).map((p, i) => savePhoto(body.site, p, i)).filter(Boolean);
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
    if (!s || FLAG_KEYS.indexOf(k) < 0) return;
    if (String(s[k] || '') !== String(f['元の値'] || '')) return; // Excel側が変わった
    s['元_' + k] = String(s[k] || '');
    s[k] = String(f['値'] || '');
    s['アプリ_' + k] = 1;
  });
  return all;
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

// 1つの項目を切り替える。base は画面に出ていた元の値（Excelの値）
function setFlag(b) {
  if (FLAG_KEYS.indexOf(b.key) < 0 || FLAG_VALUES.indexOf(b.value) < 0) return { ok: false, error: '項目か値が不正です' };
  const sh = flagSheet(), v = sh.getDataRange().getValues();
  const base = String(b.base || '');
  const same = b.value === normCircle(base); // Excelの ○（U+25CB）は 〇 と同じ扱い
  const row = [txt(b.site), b.key, txt(b.value), txt(base), new Date(), txt(b.user)];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === b.site && v[i][1] === b.key) {
      if (same) sh.deleteRow(i + 1); // Excelと同じ値に戻したら記録は要らない
      else sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { ok: true };
    }
  }
  if (!same) sh.appendRow(row);
  return { ok: true };
}

function normCircle(v) { return v === '○' ? '〇' : String(v || ''); }

// 取り込み後、Excel側で変わった項目のアプリ記録を消す（古い記録が後で生き返らないように）
function pruneFlags_() {
  const sh = flagSheet(), v = sh.getDataRange().getValues();
  const xs = {}; readSheet(SHEETS.sites).forEach(s => (xs[s['現場名']] = s));
  for (let i = v.length - 1; i >= 1; i--) {
    const s = xs[v[i][0]];
    if (s && String(s[v[i][1]] || '') !== String(v[i][3] || '')) sh.deleteRow(i + 1);
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

function savePhoto(site, dataUrl, idx) {
  try {
    const m = /^data:(image\/\w+);base64,(.+)$/.exec(dataUrl);
    if (!m) return '';
    const folder = getFolder(PHOTO_FOLDER);
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

  const rows = [], amts = [], seen = {};
  const now = new Date();
  const fy = now.getMonth() < 3 ? now.getFullYear() - 1 : now.getFullYear(); // 今年度（4月始まり）
  const C = n => SITE_COLS.indexOf(n);
  try {
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
  } finally {
    DriveApp.getFileById(tmp.id).setTrashed(true);
  }
  if (rows.length === 0) throw new Error('物件行が0件でした');

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

function geocode(addr) {
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
