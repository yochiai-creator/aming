/**
 * Webアプリ本体と、画面から呼ぶAPI（google.script.run）
 */
function doGet(e) {
  const fn = e && e.parameter && e.parameter.fn;
  if (fn) return runFn_(fn, e.parameter.arg);
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('アーム トレーサビリティ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// CLI（gas-run.sh）から ?fn=名前 で実行できる関数。ここに無いものは動かさない
const CLI_FUNCTIONS = {
  healthCheck: () => healthCheck(),
  importLatest: () => importLatest(),
  importForce: () => importLatest(true),
  apiDays: () => apiDays(),
  apiUnchecked: () => apiUnchecked().length,
  apiSearch: q => apiSearch(q)
};
function runFn_(fn, arg) {
  let body;
  if (!Object.prototype.hasOwnProperty.call(CLI_FUNCTIONS, fn)) body = {ok: false, error: 'Unknown function: ' + fn, allowed: Object.keys(CLI_FUNCTIONS)};
  else {
    try { body = {ok: true, function: fn, result: CLI_FUNCTIONS[fn](arg) || null}; }
    catch (err) { body = {ok: false, function: fn, error: String(err && err.message || err), stack: err && err.stack}; }
  }
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

/** 設定とデータの状態確認 */
function healthCheck() {
  const ss = ss_();
  const counts = {};
  Object.keys(CONFIG.SHEETS).forEach(k => { const sh = ss.getSheetByName(CONFIG.SHEETS[k]); counts[CONFIG.SHEETS[k]] = sh ? sh.getLastRow() - 1 : 'なし'; });
  const latest = latestSourceFile_();
  return {ss: ss.getName(), counts: counts, latest: latest ? latest.getName() : null, lastImported: PropertiesService.getScriptProperties().getProperty('LAST_IMPORTED'), ai: CONFIG.AI_PROVIDER, apiKey: !!PropertiesService.getScriptProperties().getProperty(aiKeyName_()), photoFolder: DriveApp.getFolderById(cfg_('PHOTO_FOLDER_ID')).getName()};
}

/** 毎朝のトリガーを入れる（何度実行しても同じ2本だけになる） */
function setupTriggers() {
  const plan = [['importLatest', 6], ['notifyUnchecked', 7]];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (plan.some(p => p[0] === t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  plan.forEach(p => ScriptApp.newTrigger(p[0]).timeBased().everyDays(1).atHour(p[1]).inTimezone('Asia/Tokyo').create());
  ScriptApp.getProjectTriggers().forEach(t => console.log('トリガー: %s', t.getHandlerFunction()));
}

/** 初回だけ実行：台帳スプレッドシートとフォルダを作る */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!cfg_('SS_ID')) {
    const ss = SpreadsheetApp.create('アーム トレーサビリティ台帳');
    props.setProperty('SS_ID', ss.getId());
  }
  const ss = ss_();
  const heads = {};
  heads[CONFIG.SHEETS.ARMS] = ARM_COLS;
  heads[CONFIG.SHEETS.TAPS] = TAP_COLS;
  heads[CONFIG.SHEETS.HOLES] = ['図番', 'ねじ穴数', '更新日時'];
  heads[CONFIG.SHEETS.PEOPLE] = ['名前'];
  heads[CONFIG.SHEETS.CHECK] = ['日時', '種別', 'キー', '内容'];
  heads[CONFIG.SHEETS.LOG] = ['日時', 'ファイル名', '件数', '新規', '更新', '重複', '対象外'];
  Object.keys(heads).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, heads[name].length).setValues([heads[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.getRange(1, 1, sh.getMaxRows(), heads[name].length).setNumberFormat('@');
    }
  });
  const first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);

  const root = DriveApp.getFileById(ss.getId()).getParents().next();
  [['INBOX_FOLDER_ID', 'アーム出荷明細_取込'], ['DONE_FOLDER_ID', 'アーム出荷明細_取込済'], ['PHOTO_FOLDER_ID', 'アーム刻印写真']].forEach(p => {
    if (!cfg_(p[0])) props.setProperty(p[0], root.createFolder(p[1]).getId());
  });
  console.log('台帳: %s', ss.getUrl());
  ['INBOX_FOLDER_ID', 'DONE_FOLDER_ID', 'PHOTO_FOLDER_ID'].forEach(k => console.log('%s: https://drive.google.com/drive/folders/%s', k, cfg_(k)));
  if (!props.getProperty(aiKeyName_())) console.log('※ スクリプトプロパティに %s を登録してください（AI読取に必要）', aiKeyName_());
}

/* ---------------- データ読み出し ---------------- */
let ARMS_MEMO = null;
function loadArms_() {
  if (!ARMS_MEMO) ARMS_MEMO = readTable_(sheet_(CONFIG.SHEETS.ARMS));
  return ARMS_MEMO;
}
function loadTaps_() { return readTable_(sheet_(CONFIG.SHEETS.TAPS)); }
// キー → 最新の有効なタップ記録
function latestTaps_(taps) {
  const m = {};
  taps.forEach(t => { if (t['取消'] !== '1') m[t['キー']] = t; }); // 追記順なので後勝ち
  return m;
}
function loadHoles_() {
  const m = {};
  readTable_(sheet_(CONFIG.SHEETS.HOLES)).forEach(r => { m[r['図番']] = +r['ねじ穴数'] || 0; });
  return m;
}
function today_() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'); }

function armView_(a, last) {
  return {
    key: a['キー'], z: a['図番'], g: a['号機'], kishu: a['機種'], spec: a['仕様'], kenki: a['建機号機'],
    to: a['出荷先'], ship: a['出荷日'], deliv: a['納入日'],
    status: last ? last['結果'] : '', by: last ? last['確認者'] : '', at: last ? last['日時'] : ''
  };
}

/* ---------------- API ---------------- */
function apiBoot() {
  return {
    today: today_(),
    me: Session.getActiveUser().getEmail(),
    people: readTable_(sheet_(CONFIG.SHEETS.PEOPLE)).map(r => r['名前']),
    days: apiDays(),
    hasKey: !!PropertiesService.getScriptProperties().getProperty(aiKeyName_())
  };
}

/** 図番・号機・建機号機・注文番号で検索（スペース区切りはAND） */
function apiSearch(q) {
  const words = String(q || '').toUpperCase().split(/[\s　]+/).filter(Boolean);
  if (!words.length) return [];
  const last = latestTaps_(loadTaps_());
  const hits = loadArms_().filter(a => {
    const hay = [a['図番'], a['建機号機'], a['注文番号'], a['製缶図番']].join(' ').toUpperCase();
    return words.every(w => /^\d{1,4}$/.test(w) ? (a['号機'] === String(+w) || hay.indexOf(w) >= 0) : hay.indexOf(w) >= 0);
  });
  hits.sort((x, y) => (y['出荷日'] || '').localeCompare(x['出荷日'] || ''));
  return hits.slice(0, 40).map(a => armView_(a, last[a['キー']]));
}

/** 1本の全履歴 */
function apiArm(key) {
  const a = loadArms_().find(r => r['キー'] === key);
  if (!a) return null;
  const taps = loadTaps_().filter(t => t['キー'] === key);
  const live = taps.filter(t => t['取消'] !== '1');
  const view = armView_(a, live[live.length - 1]);
  view.steps = [['着工', a['着工']], ['検査完了', a['検査完了日']], ['塗装完了', a['塗装完了日']],
    ['塗装後修正', a['塗装後修正完了日']], ['出荷', a['出荷日']], ['納入', a['納入日']]];
  view.order = a['注文番号']; view.base = a['製缶図番']; view.updated = a['更新日時'];
  view.taps = taps.map(t => ({id: t['記録ID'], at: t['日時'], result: t['結果'], holes: t['ねじ穴数'], treated: t['処置数'],
    by: t['確認者'], note: t['備考'], photo: t['写真ID'] ? 'https://drive.google.com/file/d/' + t['写真ID'] + '/view' : '', canceled: t['取消'] === '1'})).reverse();
  view.holes = loadHoles_()[a['図番']] || '';
  return view;
}

/** タップチェックを記録（追記のみ。訂正は取消→再記録） */
function apiSaveTap(rec) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const a = loadArms_().find(r => r['キー'] === rec.key);
    if (!a) throw new Error('台帳にないアームです: ' + rec.key);
    if (rec.result !== '良' && rec.result !== '否') throw new Error('結果は 良 か 否');
    if (!rec.person) throw new Error('確認者を選んでください');
    const holes = +rec.holes || 0, treated = +rec.treated || 0;
    if (holes && treated > holes) throw new Error('処置数がねじ穴数より多いです');
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const id = Utilities.getUuid().slice(0, 8);
    appendRows_(CONFIG.SHEETS.TAPS, [[id, now, a['キー'], a['図番'], a['号機'], rec.result, holes || '', treated, rec.person, rec.note || '', rec.photoId || '', '']]);
    if (holes) saveHoles_(a['図番'], holes, now);
  } finally { lock.releaseLock(); }
  return apiArm(rec.key);
}

function apiCancelTap(id, key) {
  const sh = sheet_(CONFIG.SHEETS.TAPS);
  const ids = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
  const i = ids.findIndex(r => String(r[0]) === String(id));
  if (i < 0) throw new Error('記録が見つかりません');
  sh.getRange(i + 2, TAP_COLS.indexOf('取消') + 1).setValue('1');
  return apiArm(key);
}

function saveHoles_(z, n, now) {
  const sh = sheet_(CONFIG.SHEETS.HOLES);
  const zs = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = zs.indexOf(z);
  if (i < 0) sh.appendRow([z, n, now]);
  else sh.getRange(i + 2, 2, 1, 2).setValues([[n, now]]);
}

/** 出荷日ごとの本数と未チェック数（7日前〜21日後） */
function apiDays() {
  const today = today_(), from = shiftDate_(today, -7), to = shiftDate_(today, 21);
  const last = latestTaps_(loadTaps_());
  const m = {};
  loadArms_().forEach(a => {
    const d = a['出荷日'];
    if (!d || d < from || d > to) return;
    const s = m[d] || (m[d] = {date: d, total: 0, todo: 0, ng: 0});
    s.total++;
    const t = last[a['キー']];
    if (!t) { if (d >= CONFIG.TRACK_FROM) s.todo++; } else if (t['結果'] === '否') s.ng++;
  });
  return Object.keys(m).sort().map(k => m[k]);
}

/** ある出荷日のアーム一覧 */
function apiDay(date) {
  const last = latestTaps_(loadTaps_());
  return loadArms_().filter(a => a['出荷日'] === date)
    .map(a => armView_(a, last[a['キー']]))
    .sort((x, y) => (x.to + x.z + x.g.padStart(5, '0')).localeCompare(y.to + y.z + y.g.padStart(5, '0')));
}

/** 今日〜N日後に出荷するのに、良の記録がないアーム（運用開始日より前の出荷は除く） */
function apiUnchecked(daysAhead) {
  const today = today_(), to = shiftDate_(today, daysAhead || CONFIG.ALERT_DAYS_AHEAD);
  return uncheckedBetween_(today, to);
}
function uncheckedBetween_(from, to) {
  if (from < CONFIG.TRACK_FROM) from = CONFIG.TRACK_FROM;
  const last = latestTaps_(loadTaps_());
  return loadArms_().filter(a => a['出荷日'] >= from && a['出荷日'] <= to && !(last[a['キー']] && last[a['キー']]['結果'] === '良'))
    .map(a => armView_(a, last[a['キー']]))
    .sort((x, y) => (x.ship + x.z).localeCompare(y.ship + y.z));
}
/** 今日より後で一番近い出荷日（台帳の出荷日から。土日祝や休みは自然に飛ぶ） */
function nextShipDate_() {
  const today = today_();
  let next = '';
  loadArms_().forEach(a => { const d = a['出荷日']; if (d > today && (!next || d < next)) next = d; });
  return next;
}

function apiSavePeople(names) {
  const sh = sheet_(CONFIG.SHEETS.PEOPLE);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
  const rows = names.map(n => String(n).trim()).filter(Boolean).map(n => [n]);
  if (rows.length) sh.getRange(2, 1, rows.length, 1).setValues(rows);
  return rows.map(r => r[0]);
}

/** 画面から「今すぐ取込」 */
function apiImportNow() { return importLatest(); }

/** CSV（タップ記録＋台帳の主要項目）。出荷日の範囲指定 */
function apiCsv(from, to) {
  const arms = {};
  loadArms_().forEach(a => { arms[a['キー']] = a; });
  const head = ['出荷日', '出荷先', '図番', '号機', '建機号機', '結果', 'ねじ穴数', '処置数', '確認者', '記録日時', '備考'];
  const rows = loadTaps_().filter(t => t['取消'] !== '1').map(t => {
    const a = arms[t['キー']] || {};
    return [a['出荷日'] || '', a['出荷先'] || '', t['図番'], t['号機'], a['建機号機'] || '', t['結果'], t['ねじ穴数'], t['処置数'], t['確認者'], t['日時'], t['備考']];
  }).filter(r => (!from || r[0] >= from) && (!to || r[0] <= to));
  const esc = v => /[",\n]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v);
  return '﻿' + [head].concat(rows).map(r => r.map(esc).join(',')).join('\r\n');
}

