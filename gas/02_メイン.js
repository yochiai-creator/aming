/**
 * Webアプリ本体と、画面から呼ぶAPI（google.script.run）
 */
function doGet(e) {
  const fn = e && e.parameter && e.parameter.fn;
  if (fn) return runFn_(fn, e.parameter.arg);
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('aming')
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
  return {ss: ss.getName(), counts: counts, latest: latest ? latest.getName() : null, lastImported: PropertiesService.getScriptProperties().getProperty('LAST_IMPORTED'), ai: cfg_('AI_PROVIDER'), apiKey: !!PropertiesService.getScriptProperties().getProperty(aiKeyName_()), photoFolder: DriveApp.getFolderById(cfg_('PHOTO_FOLDER_ID')).getName()};
}

/** 毎朝のトリガーを入れる（何度実行しても同じ2本だけになる） */
function setupTriggers() {
  const plan = [['importLatest', cfgNum_('IMPORT_HOUR') || 6], ['notifyUnchecked', cfgNum_('NOTIFY_HOUR') || 7]];
  // 撮るだけモードの取りこぼしを5分おきに読む
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'processQueue') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processQueue').timeBased().everyMinutes(5).create();
  console.log('トリガー: processQueue 5分おき');
  ScriptApp.getProjectTriggers().forEach(t => {
    if (plan.some(p => p[0] === t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  plan.forEach(p => ScriptApp.newTrigger(p[0]).timeBased().everyDays(1).atHour(p[1]).inTimezone('Asia/Tokyo').create());
  plan.forEach(p => console.log('トリガー: %s 毎朝%s時台', p[0], p[1]));
}

/** 初回だけ実行：台帳スプレッドシートとフォルダを作る */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!cfg_('SS_ID')) {
    const ss = SpreadsheetApp.create('aming 台帳');
    props.setProperty('SS_ID', ss.getId());
  }
  const ss = ss_();
  const heads = {};
  heads[CONFIG.SHEETS.ARMS] = ARM_COLS;
  heads[CONFIG.SHEETS.TAPS] = TAP_COLS;
  heads[CONFIG.SHEETS.HOLES] = ['図番', 'ねじ穴数', '更新日時'];
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
// 台帳はキャッシュ（gzip＋分割、6時間）から読む。無ければシートから読んでキャッシュに置く。取込時に clearArmsCache_()
function loadArms_() {
  if (ARMS_MEMO) return ARMS_MEMO;
  ARMS_MEMO = armsFromCache_();
  if (!ARMS_MEMO) { ARMS_MEMO = readTable_(sheet_(CONFIG.SHEETS.ARMS)); armsToCache_(ARMS_MEMO); }
  return ARMS_MEMO;
}
const ARMS_CACHE_CHUNK = 90000;
function armsFromCache_() {
  try {
    const cache = CacheService.getScriptCache(), n = +cache.get('arms_n');
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push('arms_' + i);
    const parts = cache.getAll(keys);
    if (keys.some(k => !parts[k])) return null;
    const json = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(keys.map(k => parts[k]).join('')), 'application/x-gzip')).getDataAsString();
    return JSON.parse(json).map(r => { const o = {}; ARM_COLS.forEach((c, i) => { o[c] = r[i]; }); return o; });
  } catch (e) { console.warn('台帳キャッシュ読込失敗: ' + e); return null; }
}
function armsToCache_(arms) {
  try {
    const rows = arms.map(a => ARM_COLS.map(c => a[c] || ''));
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(rows), 'application/json')).getBytes());
    const put = {};
    let n = 0;
    for (let i = 0; i < b64.length; i += ARMS_CACHE_CHUNK) put['arms_' + n++] = b64.slice(i, i + ARMS_CACHE_CHUNK);
    put.arms_n = String(n);
    CacheService.getScriptCache().putAll(put, 21600);
  } catch (e) { console.warn('台帳キャッシュ保存失敗: ' + e); }
}
function clearArmsCache_() { ARMS_MEMO = null; try { CacheService.getScriptCache().remove('arms_n'); } catch (e) {} }
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

// キー → 現場で記録済みの工程（取消を除く）
function liveSteps_() {
  const m = {};
  readTable_(sheet_(CONFIG.SHEETS.STEPS)).forEach(r => { if (r['取消'] !== '1') (m[r['キー']] = m[r['キー']] || {})[r['工程']] = r['日時']; });
  return m;
}
// 今の段階：未 → 塗装完了 → 塗装後チェック済（タップ良）→ 出荷済。否は「要処置」
function stage_(last, steps) {
  steps = steps || {};
  if (steps['出荷']) return '出荷済';
  if (last && last['結果'] === '否') return '要処置';
  if (last && last['結果'] === '良') return '塗装後チェック済';
  if (steps['塗装後修正完了'] || steps['塗装完了']) return '塗装完了';
  return '未';
}
function armView_(a, last, steps) {
  return {
    key: a['キー'], z: a['図番'], g: a['号機'], kishu: a['機種'], spec: a['仕様'], kenki: a['建機号機'],
    to: a['出荷先'], ship: a['出荷日'], deliv: a['納入日'],
    paintPlan: a['塗装完了日'] || '', fixPlan: a['塗装後修正完了日'] || '', // 生産管理Excelの予定日
    status: last ? last['結果'] : '', by: last ? last['確認者'] : '', at: last ? last['日時'] : '',
    stage: stage_(last, steps),
    // 一覧の絞込み用：塗装済＝塗装完了以降の記録あり、塗装後修正済＝塗装後修正完了・タップ良・出荷のどれか
    painted: !!(steps && (steps['塗装完了'] || steps['塗装後修正完了'] || steps['出荷'])) || !!last,
    fixed: !!(steps && (steps['塗装後修正完了'] || steps['出荷'])) || !!(last && last['結果'] === '良')
  };
}

/* ---------------- API ---------------- */
function apiBoot() {
  return {
    today: today_(),
    me: Session.getActiveUser().getEmail(),
    people: cfgList_('PEOPLE'),
    days: apiDays(),
    hasKey: !!PropertiesService.getScriptProperties().getProperty(aiKeyName_())
  };
}

/** 図番・号機・建機号機・注文番号で検索（スペース区切りはAND） */
function apiSearch(q) {
  const words = String(q || '').toUpperCase().split(/[\s　]+/).filter(Boolean);
  if (!words.length) return [];
  const last = latestTaps_(loadTaps_()), st = liveSteps_();
  const hits = loadArms_().filter(a => {
    const hay = [a['図番'], a['建機号機'], a['注文番号'], a['製缶図番']].join(' ').toUpperCase();
    return words.every(w => /^\d{1,4}$/.test(w) ? (a['号機'] === String(+w) || hay.indexOf(w) >= 0) : hay.indexOf(w) >= 0);
  });
  hits.sort((x, y) => (y['出荷日'] || '').localeCompare(x['出荷日'] || ''));
  return hits.slice(0, 40).map(a => armView_(a, last[a['キー']], st[a['キー']]));
}

/** 1本の全履歴 */
function apiArm(key) {
  const a = loadArms_().find(r => r['キー'] === key);
  if (!a) return null;
  const taps = loadTaps_().filter(t => t['キー'] === key);
  const live = taps.filter(t => t['取消'] !== '1');
  const view = armView_(a, live[live.length - 1], liveSteps_()[key]);
  const changes = readTable_(sheet_(CONFIG.SHEETS.CHANGES)).filter(c => c['キー'] === key);
  const recs = readTable_(sheet_(CONFIG.SHEETS.STEPS)).filter(r => r['キー'] === key);
  const fieldStep = {};
  STEPS.forEach(s => { fieldStep[s[1]] = s[0]; });
  // 工程ごと：Excelの日付・その変更履歴・現場の記録
  view.steps = [['着工', '着工'], ['検査完了', '検査完了日'], ['塗装完了', '塗装完了日'],
    ['塗装後修正', '塗装後修正完了日'], ['出荷', '出荷日'], ['納入', '納入日']].map(p => ({
      name: p[0], date: a[p[1]], step: fieldStep[p[1]] || '',
      changes: changes.filter(c => c['項目'] === p[1]).map(c => ({at: c['日時'], from: c['前'], to: c['後']})),
      recs: recs.filter(r => r['工程'] === fieldStep[p[1]]).map(r => ({id: r['記録ID'], at: r['日時'], by: r['確認者'], note: r['備考'], canceled: r['取消'] === '1'}))
    }));
  view.otherChanges = changes.filter(c => DATE_COLS.indexOf(c['項目']) < 0).map(c => ({at: c['日時'], col: c['項目'], from: c['前'], to: c['後']}));
  view.order = a['注文番号']; view.base = a['製缶図番']; view.updated = a['更新日時'];
  view.taps = taps.map(t => ({id: t['記録ID'], at: t['日時'], result: t['結果'], holes: t['ねじ穴数'], treated: t['処置数'],
    by: t['確認者'], note: t['備考'], photo: t['写真ID'] ? 'https://drive.google.com/file/d/' + t['写真ID'] + '/view' : '', canceled: t['取消'] === '1'})).reverse();
  view.holes = loadHoles_()[a['図番']] || '';
  view.photos = readTable_(sheet_(CONFIG.SHEETS.PHOTOS)).filter(r => r['キー'] === key)
    .map(r => ({id: r['写真ID'], at: r['日時'], name: r['ファイル名'], purpose: r['目的'] || ''})).reverse();
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

/** 工程（塗装完了／塗装後修正完了／出荷）を現場で記録（追記のみ） */
function apiSaveStep(rec) {
  if (!STEPS.some(s => s[0] === rec.step)) throw new Error('工程が不明です: ' + rec.step);
  if (!rec.person) throw new Error('確認者を選んでください');
  const a = loadArms_().find(r => r['キー'] === rec.key);
  if (!a) throw new Error('台帳にないアームです: ' + rec.key);
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  appendRows_(CONFIG.SHEETS.STEPS, [[Utilities.getUuid().slice(0, 8), now, a['キー'], a['図番'], a['号機'], rec.step, rec.person, rec.note || '', rec.photoId || '', '']]);
  return apiArm(rec.key);
}
function apiCancelStep(id, key) {
  cancelRow_(CONFIG.SHEETS.STEPS, STEP_COLS, id);
  return apiArm(key);
}
function cancelRow_(name, cols, id) {
  const sh = sheet_(name);
  const ids = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getValues();
  const i = ids.findIndex(r => String(r[0]) === String(id));
  if (i < 0) throw new Error('記録が見つかりません');
  sh.getRange(i + 2, cols.indexOf('取消') + 1).setValue('1');
}

function apiCancelTap(id, key) {
  cancelRow_(CONFIG.SHEETS.TAPS, TAP_COLS, id);
  return apiArm(key);
}

function saveHoles_(z, n, now) {
  const sh = sheet_(CONFIG.SHEETS.HOLES);
  const zs = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = zs.indexOf(z);
  if (i < 0) sh.appendRow([z, n, now]);
  else sh.getRange(i + 2, 2, 1, 2).setValues([[n, now]]);
}

/** 出荷日ごとの本数と未チェック数（今日〜21日後。過ぎた出荷日は出さない） */
function apiDays() {
  const today = today_(), from = today, to = shiftDate_(today, 21);
  const last = latestTaps_(loadTaps_()), st = liveSteps_();
  const m = {};
  loadArms_().forEach(a => {
    const d = a['出荷日'];
    if (!d || d < from || d > to) return;
    const s = m[d] || (m[d] = {date: d, total: 0, todo: 0, ng: 0});
    s.total++;
    const t = last[a['キー']];
    if (!t) { if (d >= cfg_('TRACK_FROM')) s.todo++; } else if (t['結果'] === '否') s.ng++;
  });
  return Object.keys(m).sort().map(k => m[k]);
}

/** ある出荷日のアーム一覧 */
function apiDay(date) {
  const last = latestTaps_(loadTaps_()), st = liveSteps_();
  return loadArms_().filter(a => a['出荷日'] === date)
    .map(a => armView_(a, last[a['キー']], st[a['キー']]))
    .sort((x, y) => (x.to + x.z + x.g.padStart(5, '0')).localeCompare(y.to + y.z + y.g.padStart(5, '0')));
}

/** 一覧画面用にまとめて1回で返す（出荷日の帯・今日〜21日後の全アーム・出荷前未チェック）。画面側でキャッシュして日付切替は通信なし */
function apiList() {
  const today = today_(), to = shiftDate_(today, 21), from = cfg_('TRACK_FROM'), alertTo = shiftDate_(today, cfgNum_('ALERT_DAYS_AHEAD'));
  const last = latestTaps_(loadTaps_()), st = liveSteps_();
  const rows = loadArms_().filter(a => a['出荷日'] >= today && a['出荷日'] <= to)
    .map(a => armView_(a, last[a['キー']], st[a['キー']]))
    .sort((x, y) => (x.ship + x.to + x.z + x.g.padStart(5, '0')).localeCompare(y.ship + y.to + y.z + y.g.padStart(5, '0')));
  const m = {};
  rows.forEach(a => {
    const s = m[a.ship] || (m[a.ship] = {date: a.ship, total: 0, todo: 0, ng: 0});
    s.total++;
    if (!a.status) { if (a.ship >= from) s.todo++; } else if (a.status === '否') s.ng++;
  });
  return {days: Object.keys(m).sort().map(k => m[k]), rows: rows,
    unchecked: rows.filter(a => a.ship >= from && a.ship <= alertTo && a.status !== '良')};
}

/** 今日〜N日後に出荷するのに、良の記録がないアーム（運用開始日より前の出荷は除く） */
function apiUnchecked(daysAhead) {
  const today = today_(), to = shiftDate_(today, daysAhead || cfgNum_('ALERT_DAYS_AHEAD'));
  return uncheckedBetween_(today, to);
}
function uncheckedBetween_(from, to) {
  if (from < cfg_('TRACK_FROM')) from = cfg_('TRACK_FROM');
  const last = latestTaps_(loadTaps_()), st = liveSteps_();
  return loadArms_().filter(a => a['出荷日'] >= from && a['出荷日'] <= to && !(last[a['キー']] && last[a['キー']]['結果'] === '良'))
    .map(a => armView_(a, last[a['キー']], st[a['キー']]))
    .sort((x, y) => (x.ship + x.z).localeCompare(y.ship + y.z));
}
/** 今日より後で一番近い出荷日（台帳の出荷日から。土日祝や休みは自然に飛ぶ） */
function nextShipDate_() {
  const today = today_();
  let next = '';
  loadArms_().forEach(a => { const d = a['出荷日']; if (d > today && (!next || d < next)) next = d; });
  return next;
}


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

