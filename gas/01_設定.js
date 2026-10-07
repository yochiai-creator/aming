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
  ALERT_DAYS_AHEAD: 3    // 今日〜N日後に出荷するアームを対象
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
