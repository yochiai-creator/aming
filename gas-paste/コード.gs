/* ===== 01_設定.js ===== */
/**
 * アーム トレーサビリティ（野田組）設定
 * ID類は空なら setup() が作ってスクリプトプロパティに保存する。直書きしたい場合はここに入れる。
 */
const CONFIG = {
  SS_ID: '13Fj89c-17Ec0YsUUuKdIykSkZV2442SrA35UNoMHUd8',        // 台帳スプレッドシート（ドライブ「アームトレーサビリティ」）
  INBOX_FOLDER_ID: '1NS4WoClO0xlGWSxvFimFcqFGUT0jOKQL',  // '001_アーム出荷明細（元Excelがたまるフォルダ。読むだけで動かさない）
  DONE_FOLDER_ID: '14XGF3BnMr2-fdt4wiqFMejk7UD5RxSuW',   // 変換用の一時ファイル置き場（取込後すぐゴミ箱へ）
  PHOTO_FOLDER_ID: '1g82k8Y4gCAlBVXAWbeSTkvlHUvcNJ1_0',  // アーム刻印写真

  SRC_FILE_RE: /\.xls[xm]$/i,      // 対象ファイル
  SRC_SHEET: '出荷明細',           // Excel側のシート名
  SRC_HEADER_MARK: '注文番号（写し）', // ヘッダー行を探す目印（A列）

  // Excel列名 → 台帳列名。列位置ではなく見出しで探すので、列が増減しても動く
  SRC_COLS: {
    '図番': '図番', '号機': '号機', '機器　機種': '機種', '仕様': '仕様',
    '建機号機': '建機号機', '出荷先': '出荷先', '着工': '着工',
    '検査完了日': '検査完了日', '塗装完了日': '塗装完了日', '塗装後修正完了日': '塗装後修正完了日',
    '最新出荷日': '出荷日', '納入日': '納入日', '注文番号': '注文番号', '使用製缶ベース情報': '製缶図番'
  },
  ZUBAN_RE: /^[A-Z]{2}\d{2}B\d{5}[FP]\d(G\d)?$/,
  EXCLUDE_KISHU: ['ブームブラケット', 'オプションブロック取り付け図'], // アーム以外

  SHEETS: {
    ARMS: 'アーム台帳', TAPS: 'タップ記録', HOLES: '穴数マスタ',
    PEOPLE: '確認者', CHECK: '要確認', LOG: '取込ログ'
  },

  // AI読取：'gemini'（スクリプトプロパティ GEMINI_API_KEY）か 'claude'（ANTHROPIC_API_KEY）
  AI_PROVIDER: 'gemini',
  GEMINI_MODEL: 'gemini-3.8-flash',
  CLAUDE_MODEL: 'claude-opus-5-5',
  CLAUDE_EFFORT: 'medium',

  // 未チェック通知（notifyUnchecked）
  ALERT_TO: '',          // カンマ区切り。空なら実行者に送る
  ALERT_DAYS_AHEAD: 3,   // 画面の「出荷前で未チェック」：今日〜N日後に出荷するアーム
  TRACK_FROM: '2026/10/08' // 運用開始日。これより前に出荷したアームは未チェック扱いにしない
};

const ARM_COLS = ['キー', '図番', '号機', '機種', '仕様', '建機号機', '出荷先', '着工', '検査完了日',
  '塗装完了日', '塗装後修正完了日', '出荷日', '納入日', '注文番号', '製缶図番', '更新日時'];
const TAP_COLS = ['記録ID', '日時', 'キー', '図番', '号機', '結果', 'ねじ穴数', '処置数', '確認者', '備考', '写真ID', '取消'];
const DATE_COLS = ['着工', '検査完了日', '塗装完了日', '塗装後修正完了日', '出荷日', '納入日'];

function cfg_(key) {
  return CONFIG[key] || PropertiesService.getScriptProperties().getProperty(key) || '';
}
function ss_() {
  const id = cfg_('SS_ID');
  if (!id) throw new Error('台帳がありません。先に setup() を実行してください');
  return SpreadsheetApp.openById(id);
}
function aiKeyName_() { return CONFIG.AI_PROVIDER === 'claude' ? 'ANTHROPIC_API_KEY' : 'GEMINI_API_KEY'; }
function sheet_(name) { return ss_().getSheetByName(name); }
function armKey_(z, g) { return z + '_' + g; }

/* ===== 02_メイン.js ===== */
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


/* ===== 03_AI読取.js ===== */
/**
 * 刻印のAI読取（Gemini / Claude）と台帳照合
 * 1回目：自由読取 → 台帳から上位5候補 → 2回目：写真と候補を見比べて選択（1回目で確定なら省略）
 */
const P1 = [
  '1枚目は刻印機（Telesis）のドットフォント見本です。上段 0123456789、中段 ?@ABCDEFGHIJKLM、下段 NOPQRSTUVWXYZ。',
  '2枚目が建機アームに打刻された刻印を切り抜いた写真、3枚目がその凹凸を強調した画像です。刻印は見本と同じフォントで、点の打刻で文字ができています。',
  '刻印は2行あります。',
  '1行目: 図番。形式は 英大文字2文字 + 数字2桁 + "B" + 数字5桁 + "F" + 数字1桁、末尾に "G"+数字1桁 が付くことがある（例: LS12B10010F1, YY12B00902F1G2）。',
  '2行目: ◇マーク、製作年月4桁(YYMM)、"-"またはすき間、号機(1〜4桁の数字)、検査記号(例 UM)。',
  'このフォントの見分け方: 0は斜線入り（Oは斜線なし）。2は角ばって斜めの線。9は丸い頭とまっすぐ下に伸びる右の縦線。5は上が平ら。B・Dは左側がまっすぐな縦線（8・0と区別）。1は上に小さな旗。',
  '文字ごとに見本と見比べてください。自信のない文字は ? にしてください。推測で埋めないこと。'
].join('\n');
const P1_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['zuban', 'seizo', 'goki', 'kensa', 'note'],
  properties: {zuban: {type: 'string'}, seizo: {type: 'string'}, goki: {type: 'string'}, kensa: {type: 'string'}, note: {type: 'string'}}
};
const P2_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['choice', 'confidence', 'reason'],
  properties: {choice: {type: 'string', enum: ['A', 'B', 'C', 'D', 'E', 'none']}, confidence: {type: 'string', enum: ['高', '中', '低']}, reason: {type: 'string'}}
};
function p2Prompt_(top) {
  return '1枚目は刻印機のドットフォント見本（上段 0-9、中段 ?@A-M、下段 N-Z）、2枚目・3枚目が同じ刻印の写真（元画像・凹凸強調）です。\n' +
    'この刻印は次の候補のどれと一致しますか。1文字ずつ見本の字形と写真を見比べて判断してください。\n' +
    '特に号機（2行目の "-" の後の数字）と図番の数字部分を丁寧に確認すること。\n' +
    top.map((c, i) => 'ABCDE'[i] + ': 図番 ' + c.arm.z + '　号機 ' + c.arm.g).join('\n') +
    '\nどれとも一致しなければ choice は "none"。';
}

/**
 * 画面から呼ぶ。orig/enh は base64(JPEG)。last は直前に記録したアーム {z,g}
 * 戻り値 {read:{zuban,goki,seizo,kensa,fixed}, cands:[armView+cost+seq], pick:{index,confidence,reason}|null, sure, photoId}
 */
function apiRead(orig, enh, last) {
  const photoId = savePhoto_(orig);
  const imgs = [{data: FONT_REF_PNG, mime: 'image/png'}, orig, enh];
  const r1 = callAI_(P1, imgs, P1_SCHEMA);
  const fz = fitZuban_(r1.zuban), g = fitGoki_(r1.goki);

  const taps = latestTaps_(loadTaps_());
  const arms = loadArms_().map(a => ({z: a['図番'], g: a['号機'], ship: a['出荷日'], done: !!taps[a['キー']], src: a}));
  const top = rankArms_(fz.z, g, arms, {today: today_(), last: last}).slice(0, 5);
  const sure = isSure_(top);
  let pick = sure ? {index: 0, confidence: '高', reason: '1回目で一致'} : null;
  if (!sure && top.length) {
    try {
      const r2 = callAI_(p2Prompt_(top), imgs, P2_SCHEMA);
      const i = 'ABCDE'.indexOf(r2.choice);
      pick = i >= 0 && i < top.length ? {index: i, confidence: r2.confidence, reason: r2.reason} : {index: -1, confidence: r2.confidence, reason: r2.reason};
    } catch (e) { console.warn('2回目失敗: ' + e); }
  }
  return {
    read: {zuban: fz.z, fixed: fz.fixed, goki: g, seizo: r1.seizo, kensa: r1.kensa, note: r1.note},
    cands: top.map(c => { const v = armView_(c.arm.src, taps[c.arm.src['キー']]); v.cost = Math.round(c.cost * 100) / 100; v.seq = c.seq; return v; }),
    pick: pick, sure: sure, photoId: photoId
  };
}

// 画像は base64文字列（JPEG）か {data, mime}
function img_(x) { return typeof x === 'string' ? {data: x, mime: 'image/jpeg'} : x; }
function callAI_(prompt, imgsB64, schema) {
  return CONFIG.AI_PROVIDER === 'claude' ? callClaude_(prompt, imgsB64, schema) : callGemini_(prompt, imgsB64, schema);
}

function callGemini_(prompt, imgsB64, schema) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY が未設定です');
  const parts = imgsB64.map(img_).map(i => ({inline_data: {mime_type: i.mime, data: i.data}}));
  parts.push({text: prompt});
  const body = {
    contents: [{role: 'user', parts: parts}],
    generationConfig: {responseMimeType: 'application/json', responseSchema: geminiSchema_(schema)}
  };
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.GEMINI_MODEL + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {'x-goog-api-key': key},
    payload: JSON.stringify(body)
  });
  const code = res.getResponseCode(), json = JSON.parse(res.getContentText() || '{}');
  if (code === 429 || code === 503) throw new Error('混み合っています。少し待って撮り直してください');
  if (code !== 200) throw new Error('AI読取エラー ' + code + ': ' + (json.error && json.error.message || ''));
  const cand = (json.candidates || [])[0];
  if (!cand || !cand.content) throw new Error('AIが読取を返しませんでした（' + (cand && cand.finishReason || (json.promptFeedback && json.promptFeedback.blockReason) || '不明') + '）');
  return JSON.parse(cand.content.parts.map(p => p.text || '').join(''));
}
// JSON Schema → Gemini の responseSchema（OpenAPI形式。additionalProperties は使えない）
function geminiSchema_(s) {
  const o = {type: String(s.type).toUpperCase()};
  if (s.enum) o.enum = s.enum;
  if (s.required) o.required = s.required;
  if (s.properties) {
    o.properties = {};
    Object.keys(s.properties).forEach(k => { o.properties[k] = geminiSchema_(s.properties[k]); });
    o.propertyOrdering = Object.keys(s.properties);
  }
  return o;
}

function callClaude_(prompt, imgsB64, schema) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY が未設定です');
  const content = imgsB64.map(img_).map(i => ({type: 'image', source: {type: 'base64', media_type: i.mime, data: i.data}}));
  content.push({type: 'text', text: prompt});
  const body = {
    model: CONFIG.CLAUDE_MODEL,
    max_tokens: 16000,
    output_config: {effort: CONFIG.CLAUDE_EFFORT, format: {type: 'json_schema', schema: schema}},
    fallbacks: 'default', // 安全判定で断られたとき別モデルで再実行（サーバー側）
    messages: [{role: 'user', content: content}]
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {'x-api-key': key, 'anthropic-version': '2023-06-01', 'anthropic-beta': 'server-side-fallback-2026-07-01'},
    payload: JSON.stringify(body)
  });
  const code = res.getResponseCode(), json = JSON.parse(res.getContentText() || '{}');
  if (code === 429 || code === 529) throw new Error('混み合っています。少し待って撮り直してください');
  if (code !== 200) throw new Error('AI読取エラー ' + code + ': ' + (json.error && json.error.message || ''));
  if (json.stop_reason === 'refusal') throw new Error('AIが読取を断りました');
  const text = (json.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  return JSON.parse(text);
}

function savePhoto_(b64) {
  const id = cfg_('PHOTO_FOLDER_ID');
  if (!id) return '';
  const name = 'kokuin_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '.jpg';
  return DriveApp.getFolderById(id).createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name)).getId();
}

/* ===== 03_フォント見本.js ===== */
/**
 * 刻印機（Telesis TMC470）11x16 ドットフォントの見本（数字0-9・英大文字）。AI読取の1枚目に添える
 * 元: data/telesis_font_11x16.png（取扱説明書 5.3 から切り出し）
 */
const FONT_REF_PNG = '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';

/* ===== 03_取込.js ===== */
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
    const sh = SpreadsheetApp.openById(tmp.id).getSheetByName(CONFIG.SRC_SHEET);
    if (!sh) throw new Error('シート「' + CONFIG.SRC_SHEET + '」がありません: ' + file.getName());
    const parsed = parseShipValues_(sh.getDataRange().getValues());
    const stat = mergeArms_(parsed.arms, stamp);
    const dupRows = parsed.dups.map(d => [stamp, 'Excel内で図番+号機が重複', d.key, d.note]);
    if (dupRows.length) appendRows_(CONFIG.SHEETS.CHECK, dupRows);
    appendRows_(CONFIG.SHEETS.LOG, [[stamp, file.getName(), parsed.arms.length, stat.added, stat.updated, parsed.dups.length, parsed.skipped]]);
    console.log('取込: %s 件 (新規 %s / 更新 %s / 重複 %s) %s', parsed.arms.length, stat.added, stat.updated, parsed.dups.length, file.getName());
    return {file: file.getName(), count: parsed.arms.length, added: stat.added, updated: stat.updated, dups: parsed.dups.length};
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
  let skipped = 0;
  for (let r = hi + 1; r < values.length; r++) {
    const row = values[r];
    const z = String(row[col['図番']] || '').trim().toUpperCase();
    const g = row[col['号機']];
    const kishu = String(row[col['機種']] || '').trim();
    if (!CONFIG.ZUBAN_RE.test(z) || !/^\d+$/.test(String(g).trim()) || CONFIG.EXCLUDE_KISHU.indexOf(kishu) >= 0) { if (z) skipped++; continue; }
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
function mergeArms_(arms, at) {
  const sh = sheet_(CONFIG.SHEETS.ARMS);
  const cur = readTable_(sh);
  const byKey = {};
  cur.forEach(a => { byKey[a['キー']] = a; });
  let added = 0, updated = 0;
  arms.forEach(a => {
    const old = byKey[a['キー']];
    if (!old) added++;
    else if (ARM_COLS.some(c => c !== '更新日時' && String(old[c] || '') !== String(a[c] || ''))) updated++;
    else return;
    a['更新日時'] = at;
    byKey[a['キー']] = a;
  });
  const rows = Object.keys(byKey).sort().map(k => ARM_COLS.map(c => byKey[k][c] === undefined ? '' : byKey[k][c]));
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, ARM_COLS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, ARM_COLS.length).setNumberFormat('@').setValues(rows);
  ARMS_MEMO = null;
  return {added: added, updated: updated};
}

function readTable_(sh) {
  const v = sh.getDataRange().getDisplayValues();
  const hdr = v.shift();
  return v.filter(r => r[0] !== '').map(r => { const o = {}; hdr.forEach((h, i) => { o[h] = r[i]; }); return o; });
}
function appendRows_(name, rows) {
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}


/* ===== 03_照合.js ===== */
/**
 * 刻印読取結果と台帳の照合（純粋関数。node でもテストできる）
 */
const SIM_GROUPS = [['0','O','D','Q'],['1','I','L','7','T'],['2','Z'],['5','S','6'],['8','B','3'],['9','7','4'],['6','G','5'],['F','E','P'],['N','M'],['U','V']];
const SIM_MAP = {};
SIM_GROUPS.forEach(g => g.forEach(a => g.forEach(b => { if (a !== b) SIM_MAP[a + b] = 1; })));

function subCost_(a, b) { if (a === b) return 0; if (a === '?' || b === '?') return 0.25; return SIM_MAP[a + b] ? 0.35 : 1; }
// 打刻で紛らわしい文字ペアを安くした編集距離
function wdist_(a, b) {
  const m = a.length, n = b.length, d = [];
  for (let i = 0; i <= m; i++) d.push([i]);
  for (let j = 1; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
    d[i][j] = Math.min(d[i-1][j] + 1, d[i][j-1] + 1, d[i-1][j-1] + subCost_(a[i-1], b[j-1]));
  return d[m][n];
}
function normKokuin_(s) { return String(s || '').toUpperCase().replace(/[ØΦθ∅]/g, '0').replace(/[^A-Z0-9?]/g, ''); }

// 桁ごとの型で補正：英2・数2・B・数5・F(P)・数1（+G数1）
const TO_DIGIT = {O:'0',D:'0',Q:'0',I:'1',L:'1',T:'1',Z:'2',S:'5',B:'8',G:'6',A:'4'};
const TO_ALPHA = {'0':'O','1':'I','2':'Z','5':'S','8':'B','6':'G','4':'A'};
function fitZuban_(raw) {
  const z = normKokuin_(raw);
  const tpl = z.length === 14 ? 'AANNBNNNNNFNGN' : z.length === 12 ? 'AANNBNNNNNFN' : null;
  if (!tpl) return {z: z, fixed: []};
  let out = '';
  const fixed = [];
  for (let i = 0; i < z.length; i++) {
    const c = z[i], t = tpl[i];
    let n = c;
    if (c === '?') n = c;
    else if (t === 'N' && !/\d/.test(c)) n = TO_DIGIT[c] || c;
    else if (t === 'A' && /\d/.test(c)) n = TO_ALPHA[c] || c;
    else if (t === 'B' || t === 'G') n = t;
    else if (t === 'F') n = (c === 'P' ? 'P' : 'F');
    if (n !== c) fixed.push(i);
    out += n;
  }
  return {z: out, fixed: fixed};
}
function fitGoki_(raw) { return normKokuin_(raw).split('').map(c => /\d|\?/.test(c) ? c : (TO_DIGIT[c] || '')).join(''); }

/**
 * arms: [{z, g, ship(yyyy/MM/dd), done(bool)}], opts: {today, last:{z,g}}
 * 戻り値: cost 昇順の [{arm, cost, seq}]（図番+号機の重複なし）
 */
function rankArms_(z, g, arms, opts) {
  opts = opts || {};
  const last = opts.last, today = opts.today || '';
  const out = arms.map(a => {
    let cost = (z ? wdist_(z, a.z) : 3) + (g ? wdist_(g, String(a.g)) * 1.5 : 2);
    let seq = false;
    // 直前に記録したのと同じ図番の続き番号なら優先
    if (last && last.z === a.z && /^\d+$/.test(String(a.g)) && /^\d+$/.test(String(last.g))) {
      const dg = +a.g - +last.g;
      if (dg >= 1 && dg <= 5) { cost = Math.max(0, cost - 0.3); seq = true; }
    }
    if (a.done) cost += 0.5;                              // チェック済み
    if (today && a.ship && a.ship < shiftDate_(today, -30)) cost += 0.3; // 出荷から1か月以上前
    return {arm: a, cost: cost, seq: seq};
  });
  out.sort((x, y) => x.cost - y.cost);
  const seen = {}, uniq = [];
  out.forEach(s => { const k = s.arm.z + '#' + s.arm.g; if (!seen[k]) { seen[k] = 1; uniq.push(s); } });
  return uniq;
}
// 1位が完全一致で、2位以下がはっきり離れていれば確認（2回目の読取）を省く
function isSure_(top) { return !!top.length && top[0].cost === 0 && (!top[1] || top[1].cost >= 2); }

function shiftDate_(ymd, days) {
  const p = ymd.split('/').map(Number), d = new Date(p[0], p[1] - 1, p[2] + days);
  return d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + String(d.getDate()).padStart(2, '0');
}


/* ===== 03_通知.js ===== */
/**
 * 出荷前の未チェック通知（時間主導トリガーで毎朝実行する想定。トリガーは手動で設定）
 * 次の出荷日（翌営業日）に出荷するアームのうち、ねじ穴タップ「良」の記録がないものだけを送る
 */
function notifyUnchecked() {
  const day = nextShipDate_();
  if (!day) { console.log('次の出荷日なし'); return 0; }
  const list = uncheckedBetween_(day, day);
  if (!list.length) { console.log('%s 出荷分は全部チェック済み', day); return 0; }
  const to = CONFIG.ALERT_TO || Session.getEffectiveUser().getEmail();
  const lines = list.map(a => a.z + ' ' + a.g + '号機  ' + (a.to || '') + (a.status === '否' ? '  ★否（処置待ち）' : ''));
  const url = ScriptApp.getService().getUrl() || '';
  MailApp.sendEmail(to, '【アーム】' + day.slice(5) + ' 出荷分 タップ未チェック ' + list.length + '本',
    day + ' に出荷するアームで、ねじ穴タップ「良」の記録がないもの：\n\n' + lines.join('\n') + (url ? '\n\n記録はこちら: ' + url : ''));
  console.log('通知: %s 出荷分 %s 本 → %s', day, list.length, to);
  return list.length;
}

/* ===== 04_テスト.js ===== */
/**
 * GASエディタから実行して確認する関数
 */
// 元フォルダの最新Excelを取り込む（取込済みでも強制）→ ログに件数が出ればOK
function testImport() {
  console.log('対象: %s', (latestSourceFile_() || {getName: () => 'なし'}).getName());
  console.log(JSON.stringify(importLatest(true)));
}
// 台帳を読めるか・検索できるか
function testSearch() {
  console.log('台帳 %s 本', loadArms_().length);
  console.log(JSON.stringify(apiSearch('YN12B20269 1152'), null, 1));
}
// 出荷日ごとの集計と、出荷前未チェック
function testDays() {
  console.log(JSON.stringify(apiDays()));
  console.log('未チェック %s 本', apiUnchecked().length);
}
// AI読取：刻印写真フォルダの最新写真で1回目だけ試す
function testReadLatestPhoto() {
  const it = DriveApp.getFolderById(cfg_('PHOTO_FOLDER_ID')).getFiles();
  let f = null;
  while (it.hasNext()) { const x = it.next(); if (!f || x.getDateCreated() > f.getDateCreated()) f = x; }
  if (!f) { console.log('写真がありません'); return; }
  const b = Utilities.base64Encode(f.getBlob().getBytes());
  console.log(JSON.stringify(callAI_(P1, [{data: FONT_REF_PNG, mime: 'image/png'}, b, b], P1_SCHEMA)));
}
// AIキーが通るか（画像なしで短く1回呼ぶ）
function testAiKey() {
  const r = callAI_('「OK」とだけ答えて。', [], {type: 'object', required: ['answer'], properties: {answer: {type: 'string'}}});
  console.log('%s %s → %s', CONFIG.AI_PROVIDER, CONFIG.AI_PROVIDER === 'claude' ? CONFIG.CLAUDE_MODEL : CONFIG.GEMINI_MODEL, JSON.stringify(r));
}
// 照合ロジック
function testMatch() {
  console.log(JSON.stringify(fitZuban_('YN12BZ0Z69F1')));  // → YN12B20269F1
  console.log(fitGoki_('11S2'));                           // → 1152
}

