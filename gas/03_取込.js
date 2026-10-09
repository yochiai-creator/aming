/**
 * 生産管理Excel（出荷予定 日程表変更A(26年10月7日).xlsm など）→ アーム台帳
 * 元フォルダの一番新しいExcelを、まだ取り込んでいなければ取り込む（Excelは毎回3年分の累積なので最新1本で足りる）。
 * 元ファイルは動かさない。台帳はExcel由来のマスタ。タップ記録は別シートなので取込で消えることはない。
 */
function importLatest(force) {
  const f = latestSourceFile_();
  if (!f) { console.log('取込対象なし'); return null; }
  const stamp = f.getId() + '@' + f.getLastUpdated().getTime();
  const props = PropertiesService.getScriptProperties();
  // トリガーから呼ばれると force にイベントが入るので true のときだけ強制
  if (force !== true && props.getProperty('LAST_IMPORTED') === stamp) {
    console.log('取込済み: %s', f.getName());
    return {file: f.getName(), skipped: true};
  }
  const res = importFile_(f);
  props.setProperty('LAST_IMPORTED', stamp);
  return res;
}

// ファイル名の (26年10月7日) で一番新しいもの。日付が読めなければ作成日時
function latestSourceFile_() {
  const it = DriveApp.getFolderById(cfg_('INBOX_FOLDER_ID')).getFiles();
  let best = null, bestKey = '';
  while (it.hasNext()) {
    const f = it.next();
    if (!CONFIG.SRC_FILE_RE.test(f.getName())) continue;
    const key = fileDateKey_(f.getName()) + Utilities.formatDate(f.getDateCreated(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
    if (key > bestKey) { best = f; bestKey = key; }
  }
  return best;
}
function fileDateKey_(name) {
  const m = /(\d{2})年(\d{1,2})月(\d{1,2})日/.exec(name);
  return m ? '20' + m[1] + m[2].padStart(2, '0') + m[3].padStart(2, '0') : '00000000';
}

function importFile_(file) {
  const started = new Date();
  const stamp = Utilities.formatDate(started, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  // xlsm → Googleスプレッドシート（一時ファイル）
  const tmp = Drive.Files.create(
    {name: '_tmp_' + file.getName(), mimeType: MimeType.GOOGLE_SHEETS, parents: [cfg_('DONE_FOLDER_ID')]},
    file.getBlob()
  );
  try {
    const sh = SpreadsheetApp.openById(tmp.id).getSheetByName(cfg_('SRC_SHEET'));
    if (!sh) throw new Error('シート「' + cfg_('SRC_SHEET') + '」がありません: ' + file.getName());
    const parsed = parseShipValues_(sh.getDataRange().getValues());
    const stat = mergeArms_(parsed.arms, stamp, file.getName());
    const dupRows = parsed.dups.map(d => [stamp, 'Excel内で図番+号機が重複', d.key, d.note]);
    if (dupRows.length) appendRows_(CONFIG.SHEETS.CHECK, dupRows);
    appendRows_(CONFIG.SHEETS.LOG, [[stamp, file.getName(), parsed.arms.length, stat.added, stat.updated, parsed.dups.length, parsed.skipped]]);
    console.log('取込: %s 件 (新規 %s / 更新 %s / 日付等の変更 %s / 重複 %s) %s', parsed.arms.length, stat.added, stat.updated, stat.changes, parsed.dups.length, file.getName());
    return {file: file.getName(), count: parsed.arms.length, added: stat.added, updated: stat.updated, changes: stat.changes, dups: parsed.dups.length};
  } finally {
    DriveApp.getFileById(tmp.id).setTrashed(true);
  }
}

/**
 * シートの2次元配列 → [{キー, 図番, 号機, ...}]（純粋関数）
 * 同じ図番+号機が複数行ある場合は出荷日が新しい方を採用し、dups に記録する。
 */
function parseShipValues_(values) {
  const hi = values.findIndex(r => String(r[0]).trim() === CONFIG.SRC_HEADER_MARK);
  if (hi < 0) throw new Error('ヘッダー行（' + CONFIG.SRC_HEADER_MARK + '）が見つかりません');
  const hdr = values[hi].map(h => String(h).replace(/\s+$/, ''));
  const col = {};
  Object.keys(CONFIG.SRC_COLS).forEach(src => {
    const i = hdr.indexOf(src);
    if (i < 0) throw new Error('列「' + src + '」が見つかりません');
    col[CONFIG.SRC_COLS[src]] = i;
  });
  const byKey = {}, dups = [];
  const exclude = cfgList_('EXCLUDE_KISHU');
  let skipped = 0;
  for (let r = hi + 1; r < values.length; r++) {
    const row = values[r];
    const z = String(row[col['図番']] || '').trim().toUpperCase();
    const g = row[col['号機']];
    const kishu = String(row[col['機種']] || '').trim();
    if (!CONFIG.ZUBAN_RE.test(z) || !/^\d+$/.test(String(g).trim()) || exclude.indexOf(kishu) >= 0) { if (z) skipped++; continue; }
    const arm = {'キー': armKey_(z, String(+g)), '図番': z, '号機': String(+g)};
    Object.keys(col).forEach(name => {
      if (name === '図番' || name === '号機') return;
      arm[name] = DATE_COLS.indexOf(name) >= 0 ? fmtDate_(row[col[name]]) : cleanCell_(row[col[name]]);
    });
    const prev = byKey[arm['キー']];
    if (prev) {
      dups.push({key: arm['キー'], note: '出荷日 ' + (prev['出荷日'] || '-') + ' / ' + (arm['出荷日'] || '-') + '（' + (r + 1) + '行目）'});
      if ((arm['出荷日'] || '') < (prev['出荷日'] || '')) continue;
    }
    byKey[arm['キー']] = arm;
  }
  return {arms: Object.keys(byKey).map(k => byKey[k]), dups: dups, skipped: skipped};
}

function cleanCell_(v) {
  if (v === null || v === undefined) return '';
  if (isDate_(v)) return fmtDate_(v);
  const s = String(v).trim();
  return /^#(N\/A|NUM!|VALUE!|REF!|DIV\/0!|NAME\?)$/.test(s) ? '' : s;
}
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
// 1900年以前（Excelの0日付）や文字列は空にする
function fmtDate_(v) {
  if (!isDate_(v) || isNaN(v) || v.getFullYear() < 2000) {
    const m = /^(20\d{2})[\/-](\d{1,2})[\/-](\d{1,2})/.exec(String(v || ''));
    return m ? m[1] + '/' + m[2].padStart(2, '0') + '/' + m[3].padStart(2, '0') : '';
  }
  return v.getFullYear() + '/' + String(v.getMonth() + 1).padStart(2, '0') + '/' + String(v.getDate()).padStart(2, '0');
}

// 台帳に上書きマージ（Excelから消えたアームは残す）
function mergeArms_(arms, at, fileName) {
  const sh = sheet_(CONFIG.SHEETS.ARMS);
  const cur = readTable_(sh);
  const byKey = {};
  cur.forEach(a => { byKey[a['キー']] = a; });
  let added = 0, updated = 0;
  const changes = [];
  arms.forEach(a => {
    const old = byKey[a['キー']];
    if (!old) added++;
    else if (ARM_COLS.some(c => c !== '更新日時' && String(old[c] || '') !== String(a[c] || ''))) {
      updated++;
      diffArm_(old, a).forEach(d => changes.push([at, a['キー'], d[0], d[1], d[2], fileName || '']));
    }
    else return;
    a['更新日時'] = at;
    byKey[a['キー']] = a;
  });
  const rows = Object.keys(byKey).sort().map(k => ARM_COLS.map(c => byKey[k][c] === undefined ? '' : byKey[k][c]));
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, ARM_COLS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, ARM_COLS.length).setNumberFormat('@').setValues(rows);
  ARMS_MEMO = null;
  if (changes.length) appendRows_(CONFIG.SHEETS.CHANGES, changes);
  return {added: added, updated: updated, changes: changes.length};
}

// 追跡する列のうち値が変わったもの [[列, 前, 後]]（純粋関数）
function diffArm_(old, a) {
  return TRACK_CHANGE_COLS.filter(c => String(old[c] || '') !== String(a[c] || '')).map(c => [c, String(old[c] || ''), String(a[c] || '')]);
}

function readTable_(sh) {
  // シートは全部テキスト書式なので getValues で十分（getDisplayValues より速い）。日付が混ざっても文字にそろえる
  const v = sh.getDataRange().getValues().map(r => r.map(x => isDate_(x) ? Utilities.formatDate(x, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : String(x)));
  const hdr = v.shift();
  return v.filter(r => r[0] !== '').map(r => { const o = {}; hdr.forEach((h, i) => { o[h] = r[i]; }); return o; });
}
function appendRows_(name, rows) {
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

if (typeof module !== 'undefined') module.exports = {parseShipValues_, fmtDate_, cleanCell_, fileDateKey_, diffArm_};
