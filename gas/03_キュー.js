/**
 * 撮るだけモード：写真を撮ったら「読取キュー」に積み、裏で
 *   ① 写真全体から刻印の場所（枠）を見つけて読む（Gemini）
 *   ② 台帳と照合
 *   ③ 画面側が枠で切り出した画像を受け取り、確定ならアームにひも付け（確定しなければ切り出し画像で読み直し）
 * 画面を閉じても、5分おきのトリガー（processQueue）が残りを読む。切り出しは次に画面を開いたとき。
 */
const P_FULL = [
  '1枚目は刻印機（Telesis）のドットフォント見本です。上段 0123456789、中段 ?@ABCDEFGHIJKLM、下段 NOPQRSTUVWXYZ。',
  '2枚目は建機アームを撮った写真です。どこかに点の打刻で2行の刻印があります（白い塗装面で薄いことが多い）。',
  'まず刻印の2行がある範囲を見つけて box に [ymin, xmin, ymax, xmax]（画像全体を0〜1000とした整数）で返してください。見つからなければ found=false。',
  '1行目: 図番。形式は 英大文字2文字 + 数字2桁 + "B" + 数字5桁 + "F" + 数字1桁、末尾に "G"+数字1桁 が付くことがある（例: LS12B10010F1, YY12B00902F1G2）。',
  '2行目: ◇マーク、製作年月4桁(YYMM)、"-"またはすき間、号機(1〜4桁の数字)、検査記号(例 UM)。',
  'このフォントの見分け方: 0は斜線入り（Oは斜線なし）。2は角ばって斜めの線。9は丸い頭とまっすぐ下に伸びる右の縦線。5は上が平ら。B・Dは左側がまっすぐな縦線（8・0と区別）。1は上に小さな旗。',
  '自信のない文字は ? にしてください。推測で埋めないこと。'
].join('\n');
const P_FULL_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['found', 'box', 'zuban', 'seizo', 'goki', 'kensa', 'note'],
  properties: {found: {type: 'boolean'}, box: {type: 'array', items: {type: 'integer'}},
    zuban: {type: 'string'}, seizo: {type: 'string'}, goki: {type: 'string'}, kensa: {type: 'string'}, note: {type: 'string'}}
};

function queueSheet_() { return sheet_(CONFIG.SHEETS.QUEUE); }
function queueRows_() {
  const sh = queueSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, QUEUE_COLS.length).getValues()
    .map((r, i) => { const o = {_row: i + 2}; QUEUE_COLS.forEach((c, j) => { o[c] = isDate_(r[j]) ? Utilities.formatDate(r[j], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') : String(r[j]); }); return o; });
}
function queueFind_(id) { const r = queueRows_().find(x => x['ID'] === id); if (!r) throw new Error('キューにありません: ' + id); return r; }
function queueSet_(row, patch) {
  const sh = queueSheet_();
  Object.keys(patch).forEach(k => { row[k] = patch[k]; sh.getRange(row._row, QUEUE_COLS.indexOf(k) + 1).setValue(patch[k]); });
  return row;
}
function queueView_(r) {
  let cands = [];
  try { cands = r['候補'] ? JSON.parse(r['候補']) : []; } catch (e) {}
  let box = null;
  try { box = r['枠'] ? JSON.parse(r['枠']) : null; } catch (e) {}
  return {id: r['ID'], at: r['日時'], purpose: r['目的'], state: r['状態'], fullId: r['元写真ID'], cropId: r['切り出しID'],
    box: box, read: r['AI読取'], key: r['キー'], cands: cands, error: r['エラー'], seen: r['確認'] === '済'};
}

/** 撮った写真（全体・縮小済み）を積む。すぐ返る */
function apiQueueAdd(b64, purpose) {
  const fullId = savePhoto_(b64, 'raw');
  const id = Utilities.getUuid().slice(0, 8);
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  appendRows_(CONFIG.SHEETS.QUEUE, [[id, now, purpose || '', '待ち', fullId, '', '', '', '', '', '', '']]);
  return queueView_(queueFind_(id));
}

/** 1件読む：写真全体から枠＋読み → 照合 */
function apiQueueProcess(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let row;
  try {
    row = queueFind_(id);
    if (row['状態'] !== '待ち' && row['状態'] !== '失敗') return queueView_(row);
    queueSet_(row, {'状態': '処理中'});
  } finally { lock.releaseLock(); }
  READ_DEADLINE = Date.now() + 45000; READ_LOG = [];
  const t0 = Date.now();
  try {
    const b64 = Utilities.base64Encode(DriveApp.getFileById(row['元写真ID']).getBlob().getBytes());
    const r1 = callAI_(P_FULL, [{data: FONT_REF_PNG, mime: 'image/png'}, b64], P_FULL_SCHEMA);
    const box = r1.found && r1.box && r1.box.length === 4 ? r1.box : null;
    const m = matchRead_(r1, null);
    const read = m.fz.z + ' / ' + (r1.seizo || '?') + '-' + (m.g || '?') + ' ' + (r1.kensa || '');
    queueSet_(row, {'枠': box ? JSON.stringify(box) : '', 'AI読取': read.trim(), '候補': JSON.stringify(m.cands),
      'キー': m.sure && m.cands[0] ? m.cands[0].key : '', '状態': m.sure ? '確定' : (box ? '候補' : '刻印なし'), 'エラー': ''});
    rlog_('キュー ' + id + ' → ' + row['状態'] + ' ' + read);
  } catch (e) {
    queueSet_(row, {'状態': '失敗', 'エラー': String(e.message || e).slice(0, 300)});
    rlog_('キュー ' + id + ' 失敗 ' + e);
  }
  writeReadLog_(t0, row);
  return queueView_(row);
}

/** 画面が枠で切り出した画像を受け取る。確定ならひも付け、まだなら切り出し画像（＋強調）で読み直し */
function apiQueueCrop(id, cropB64, enhB64) {
  let row = queueFind_(id);
  if (!row['切り出しID']) row = queueSet_(row, {'切り出しID': savePhoto_(cropB64, 'kokuin')});
  if (row['状態'] === '候補' && enhB64) {
    READ_DEADLINE = Date.now() + 40000; READ_LOG = [];
    const t0 = Date.now();
    try {
      const r1 = callAI_(P1, [{data: FONT_REF_PNG, mime: 'image/png'}, cropB64, enhB64], P1_SCHEMA);
      const m = matchRead_(r1, null);
      const read = m.fz.z + ' / ' + (r1.seizo || '?') + '-' + (m.g || '?') + ' ' + (r1.kensa || '');
      queueSet_(row, {'AI読取': read.trim(), '候補': JSON.stringify(m.cands), 'キー': m.sure && m.cands[0] ? m.cands[0].key : '', '状態': m.sure ? '確定' : '候補'});
      rlog_('キュー ' + id + ' 切り出しで読み直し → ' + row['状態'] + ' ' + read);
    } catch (e) { rlog_('キュー ' + id + ' 読み直し失敗 ' + e); }
    writeReadLog_(t0, row);
  }
  if (row['状態'] === '確定' && row['キー']) apiLinkPhoto(row['切り出しID'], row['キー'], row['AI読取'], row['目的']);
  return queueView_(row);
}

/** 候補から人が選んだ */
function apiQueuePick(id, key) {
  const row = queueSet_(queueFind_(id), {'キー': key, '状態': '確定'});
  const photo = row['切り出しID'] || row['元写真ID'];
  if (photo) apiLinkPhoto(photo, key, row['AI読取'], row['目的']);
  return queueView_(row);
}
/** 一覧から消す（確認済み） */
function apiQueueDone(id) { return queueView_(queueSet_(queueFind_(id), {'確認': '済'})); }

/** 未確認のもの（新しい順・最大40件） */
function apiQueueList() {
  return queueRows_().filter(r => r['確認'] !== '済').slice(-40).reverse().map(queueView_);
}

/** トリガー用：画面を閉じても残った「待ち」を読む（2分以上前のもの） */
function processQueue() {
  const limit = Date.now() + 4 * 60 * 1000, old = Utilities.formatDate(new Date(Date.now() - 120000), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  queueRows_().filter(r => r['状態'] === '待ち' && r['日時'] < old).forEach(r => { if (Date.now() < limit) apiQueueProcess(r['ID']); });
}

function writeReadLog_(t0, row) {
  try {
    appendRows_(CONFIG.SHEETS.READLOG, [[Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'), ((Date.now() - t0) / 1000).toFixed(1),
      'キュー:' + row['状態'] + (row['エラー'] ? ' ' + row['エラー'].slice(0, 120) : ''), row['AI読取'] || '', row['キー'] || '', READ_LOG.join('\n').slice(0, 3000)]]);
  } catch (e) { console.warn(e); }
}
