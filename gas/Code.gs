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
const EXTRA_COLS = ['現場名', '消検日', '建築確認検査日', '搬入', '工事進捗メモ', '更新日時', '更新者'];
const VISIT_COLS = ['日時', '記録者', '現場名', '建物進捗', 'メモ', '写真', '緯度', '経度', '距離m'];
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

// ---------- Web API ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.pw !== cfg('パスワード')) return json({ ok: false, error: 'パスワードが違います' });
  return json({
    ok: true,
    sites: readSheet(SHEETS.sites),
    extra: readSheet(SHEETS.extra),
    visits: readSheet(SHEETS.visits),
    users: cfg('記録者').split(',').map(s => s.trim()).filter(Boolean),
    radius: Number(cfg('現地判定m')) || 300,
    updated: new Date().toISOString(),
  });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad json' }); }
  if (body.pw !== cfg('パスワード')) return json({ ok: false, error: 'パスワードが違います' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (body.action === 'visit') return json(addVisit(body));
    if (body.action === 'extra') return json(saveExtra(body));
    if (body.action === 'deleteVisit') return json(deleteVisit(body));
    return json({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}

function addVisit(b) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.visits);
  const urls = (b.photos || []).map((p, i) => savePhoto(b.site, p, i)).filter(Boolean);
  const row = [new Date(), b.user || '', b.site, b.stage || '', b.memo || '', urls.join('\n'), b.lat || '', b.lng || '', b.dist === undefined ? '' : Math.round(b.dist)];
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

function saveExtra(b) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.extra);
  const data = sh.getDataRange().getValues();
  const row = [b.site, dateTime(b.fireDate, b.fireTime), dateTime(b.bldgDate, b.bldgTime), b.delivery || '', b.progress || '', new Date(), b.user || ''];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === b.site) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { ok: true };
    }
  }
  sh.appendRow(row);
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

// ---------- Excel 週次取り込み ----------
function importFromExcel() {
  const fname = cfg('Excelファイル名');
  const it = DriveApp.getFilesByName(fname);
  if (!it.hasNext()) throw new Error('Excelが見つかりません: ' + fname);
  const xlsx = it.next();
  // xlsx → 一時スプレッドシートに変換（Drive API 必要）
  const tmp = Drive.Files.create(
    { name: '_tmp_import_' + Date.now(), mimeType: 'application/vnd.google-apps.spreadsheet' },
    xlsx.getBlob()
  );
  const ss = SpreadsheetApp.openById(tmp.id);
  const sitesSh = SpreadsheetApp.getActive().getSheetByName(SHEETS.sites);
  const old = {};
  readSheet(SHEETS.sites).forEach(r => (old[r['現場名']] = r));

  const rows = [];
  const now = new Date();
  ss.getSheets().forEach(sh => {
    if (!/^R\d+年度_/.test(sh.getName())) return;
    const moto = sh.getName().replace(/^R\d+年度_/, '');
    const v = sh.getDataRange().getValues();
    // 8行目以降: B=項番 C=現場名 D=職人 E=消検月 F..M=○ N=金額 O=注意点 P=消検年月(日付)
    const baseYear = Number(v[1][17]) || now.getFullYear(); // R2 = 集計基準年
    for (let i = 7; i < v.length; i++) {
      const r = v[i];
      const name = norm(r[2]);
      if (!name) continue;
      let ym = '';
      if (r[15] instanceof Date) ym = Utilities.formatDate(r[15], 'Asia/Tokyo', 'yyyy-MM');
      else if (r[4] !== '' && !isNaN(r[4])) {
        const m = Number(r[4]);
        ym = `${m >= 4 ? baseYear : baseYear + 1}-${String(m).padStart(2, '0')}`;
      }
      const prev = old[name] || {};
      let addr = prev['住所'] || nameToAddress(name) || '';
      let lat = prev['緯度'] || '', lng = prev['経度'] || '', chk = prev['要確認'] || '';
      if ((lat === '' || lng === '') && addr) {
        const g = geocode(addr);
        if (g) { lat = g.lat; lng = g.lng; } else chk = '座標取得失敗';
      }
      if (!addr) chk = '住所不明';
      rows.push([name, moto, r[3] || '', ym, r[5] || '', r[6] || '', r[7] || '', r[8] || '', r[9] || '', r[10] || '', r[11] || '', r[12] || '', r[14] || '', addr, lat, lng, chk, now]);
    }
  });
  DriveApp.getFileById(tmp.id).setTrashed(true);
  if (rows.length === 0) throw new Error('物件行が0件でした');
  sitesSh.clearContents();
  sitesSh.getRange(1, 1, 1, SITE_COLS.length).setValues([SITE_COLS]).setFontWeight('bold');
  sitesSh.getRange(2, 1, rows.length, SITE_COLS.length).setValues(rows);
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
