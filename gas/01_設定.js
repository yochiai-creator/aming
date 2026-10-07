/**
 * aming（アーム トレーサビリティ・野田組）設定
 * 運用で変える値は台帳スプレッドシートの「設定」シートが優先（SETTINGS 参照）。ここはその初期値。
 * SS_ID だけはコードで持つ（設定シートの場所なので）。APIキーはスクリプトプロパティ。
 */
const CONFIG = {
  SS_ID: '13Fj89c-17Ec0YsUUuKdIykSkZV2442SrA35UNoMHUd8',        // aming 台帳（ドライブ「aming」フォルダ）
  INBOX_FOLDER_ID: '1NS4WoClO0xlGWSxvFimFcqFGUT0jOKQL',  // '001_アーム出荷明細（元Excelがたまるフォルダ。読むだけで動かさない）
  DONE_FOLDER_ID: '14XGF3BnMr2-fdt4wiqFMejk7UD5RxSuW',   // aming 作業用（変換の一時ファイル。取込後すぐゴミ箱へ）
  PHOTO_FOLDER_ID: '1g82k8Y4gCAlBVXAWbeSTkvlHUvcNJ1_0',  // aming 刻印写真

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
    SETTINGS: '設定',
    ARMS: 'アーム台帳', TAPS: 'タップ記録', HOLES: '穴数マスタ',
    CHECK: '要確認', LOG: '取込ログ',
    STEPS: '工程記録', CHANGES: '変更履歴', PHOTOS: '刻印写真'
  },

  // AI読取：'gemini'（スクリプトプロパティ GEMINI_API_KEY）か 'claude'（ANTHROPIC_API_KEY）
  AI_PROVIDER: 'gemini',
  GEMINI_MODEL: 'gemini-3.8-flash',
  CLAUDE_MODEL: 'claude-opus-5-5',
  CLAUDE_EFFORT: 'medium',

  // 未チェック通知（notifyUnchecked）
  ALERT_TO: '',          // カンマ区切り。空なら実行者に送る
  ALERT_DAYS_AHEAD: 3,   // 画面の「出荷前で未チェック」：今日〜N日後に出荷するアーム
  IMPORT_HOUR: 6,        // 毎朝の取込トリガーの時（setupTriggers）
  NOTIFY_HOUR: 7,        // 毎朝の通知トリガーの時（setupTriggers）
  PEOPLE: '',            // 確認者（カンマ区切り）
  TRACK_FROM: '2026/10/08' // 運用開始日。これより前に出荷したアームは未チェック扱いにしない
};

const ARM_COLS = ['キー', '図番', '号機', '機種', '仕様', '建機号機', '出荷先', '着工', '検査完了日',
  '塗装完了日', '塗装後修正完了日', '出荷日', '納入日', '注文番号', '製缶図番', '更新日時'];
const TAP_COLS = ['記録ID', '日時', 'キー', '図番', '号機', '結果', 'ねじ穴数', '処置数', '確認者', '備考', '写真ID', '取消'];
const DATE_COLS = ['着工', '検査完了日', '塗装完了日', '塗装後修正完了日', '出荷日', '納入日'];
const STEP_COLS = ['記録ID', '日時', 'キー', '図番', '号機', '工程', '確認者', '備考', '写真ID', '取消'];
const CHANGE_COLS = ['日時', 'キー', '項目', '前', '後', '元ファイル'];
const PHOTO_COLS = ['日時', 'キー', '図番', '号機', '写真ID', 'ファイル名', 'AI読取'];
// 現場で記録する工程（名前 → 台帳の対応する日付列）
const STEPS = [['塗装完了', '塗装完了日'], ['塗装後修正完了', '塗装後修正完了日'], ['出荷', '出荷日']];
// 取込で値が変わったら変更履歴に残す列
const TRACK_CHANGE_COLS = DATE_COLS.concat(['出荷先', '建機号機']);

// 「設定」シートに出す項目：[キー, 項目名, 説明]。値はシートが優先、空ならCONFIG
const SETTINGS = [
  ['AI_PROVIDER', 'AI読取', 'gemini か claude'],
  ['GEMINI_MODEL', 'Geminiモデル', '例 gemini-3.8-flash'],
  ['CLAUDE_MODEL', 'Claudeモデル', 'AI読取が claude のとき'],
  ['TRACK_FROM', '運用開始日', 'これより前に出荷したアームは未チェック扱いにしない（yyyy/MM/dd）'],
  ['ALERT_TO', '通知メール宛先', 'カンマ区切りで複数可。空なら自分'],
  ['ALERT_DAYS_AHEAD', '未チェック表示日数', '一覧の「出荷前で未チェック」に今日から何日後まで出すか'],
  ['IMPORT_HOUR', '取込の時刻', '毎朝何時台にExcelを取り込むか（変えたら setupTriggers を再実行）'],
  ['NOTIFY_HOUR', '通知の時刻', '毎朝何時台に未チェック通知を送るか（変えたら setupTriggers を再実行）'],
  ['EXCLUDE_KISHU', '除外する機種', 'アーム以外として取り込まない機種。カンマ区切り'],
  ['SRC_SHEET', 'Excelのシート名', '生産管理Excelで読むシート'],
  ['INBOX_FOLDER_ID', '取込元フォルダID', '生産管理Excelがたまるフォルダ'],
  ['PHOTO_FOLDER_ID', '刻印写真フォルダID', '撮った刻印写真の保存先'],
  ['DONE_FOLDER_ID', '作業用フォルダID', '取込時の一時ファイル置き場'],
  // 確認者は一番下。B列に1行1人で下に足していく（A列は空でOK）
  ['PEOPLE', '確認者', '1行に1人。下の行のB列に続けて入れる']
];
let SETTINGS_MEMO = null;
function sheetSettings_() {
  if (SETTINGS_MEMO) return SETTINGS_MEMO;
  SETTINGS_MEMO = {};
  if (typeof SpreadsheetApp === 'undefined') return SETTINGS_MEMO; // node テスト
  const sh = ss_().getSheetByName(CONFIG.SHEETS.SETTINGS) || makeSheet_(CONFIG.SHEETS.SETTINGS);
  const byLabel = parseSettingRows_(sh.getDataRange().getDisplayValues().slice(1));
  SETTINGS.forEach(d => { if (byLabel[d[1]]) SETTINGS_MEMO[d[0]] = byLabel[d[1]]; });
  return SETTINGS_MEMO;
}
// [[項目, 値, 説明], ...] → {項目: 値}。A列が空の行は直前の項目の続き（確認者を1行1人で書ける）（純粋関数）
function parseSettingRows_(rows) {
  const out = {};
  let label = '';
  rows.forEach(r => {
    const a = String(r[0] || '').trim(), b = String(r[1] || '').trim();
    if (a) label = a;
    if (!label || !b) return;
    out[label] = out[label] ? out[label] + ',' + b : b;
  });
  return out;
}
// 設定値：設定シート → CONFIG → スクリプトプロパティ の順
function cfg_(key) {
  if (key !== 'SS_ID') { const v = sheetSettings_()[key]; if (v) return v; }
  const c = CONFIG[key];
  if (c !== undefined && c !== '') return Array.isArray(c) ? c.join(',') : c;
  return (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties().getProperty(key)) || '';
}
function cfgNum_(key) { return Number(cfg_(key)) || 0; }
function cfgList_(key) { return String(cfg_(key)).split(/[,、，\n]/).map(x => x.trim()).filter(Boolean); }
function ss_() {
  const id = cfg_('SS_ID');
  if (!id) throw new Error('台帳がありません。先に setup() を実行してください');
  return SpreadsheetApp.openById(id);
}
function aiKeyName_() { return cfg_('AI_PROVIDER') === 'claude' ? 'ANTHROPIC_API_KEY' : 'GEMINI_API_KEY'; }
function sheet_(name) { return ss_().getSheetByName(name) || makeSheet_(name); }
// 後から増えたシートは初回アクセス時に作る
function makeSheet_(name) {
  const cols = {};
  cols[CONFIG.SHEETS.STEPS] = STEP_COLS;
  cols[CONFIG.SHEETS.CHANGES] = CHANGE_COLS;
  cols[CONFIG.SHEETS.PHOTOS] = PHOTO_COLS;
  if (name === CONFIG.SHEETS.SETTINGS) {
    const sh = ss_().insertSheet(name, 0);
    const rows = [['項目', '値', '説明']].concat(SETTINGS.map(d => [d[1], String(Array.isArray(CONFIG[d[0]]) ? CONFIG[d[0]].join(',') : CONFIG[d[0]]), d[2]]));
    sh.getRange(1, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  if (!cols[name]) return null;
  const sh = ss_().insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), cols[name].length).setNumberFormat('@');
  sh.getRange(1, 1, 1, cols[name].length).setValues([cols[name]]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function armKey_(z, g) { return z + '_' + g; }
