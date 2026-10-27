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
  '2行目: ◇マーク、製作年月4桁(YYMM)、"-"またはすき間、号機(1〜4桁の数字。前に英字1文字が付くことがある 例 2610-K2022 → 号機は 2022。英字は goki に入れず note に書く)、検査記号(例 UM, U)。',
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
// 読取キューの掃除：「済」になって日数（既定30日）たった行は消す（一覧を開くたびに全行読むので、たまると遅くなる）
// rows は [[ID, 日時, …], …]（見出し除く）。残す行だけ返す（純粋関数）
function queueKeepRows_(rows, cutoff) {
  const iT = QUEUE_COLS.indexOf('日時'), iC = QUEUE_COLS.indexOf('確認');
  return rows.filter(r => {
    if (String(r[iC]) !== '済') return true;
    const t = r[iT] instanceof Date ? r[iT] : new Date(String(r[iT]).replace(/-/g, '/'));
    return isNaN(t) || t >= cutoff;
  });
}
function cleanupQueue_(days) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = queueSheet_(), n = sh.getLastRow() - 1;
    if (n < 1) return 0;
    const w = Math.max(QUEUE_COLS.length, sh.getLastColumn());
    const rows = sh.getRange(2, 1, n, w).getValues();
    const keep = queueKeepRows_(rows, new Date(Date.now() - (days || 30) * 864e5));
    if (keep.length === rows.length) return 0;
    sh.getRange(2, 1, n, w).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, w).setValues(keep);
    console.log('読取キュー掃除: %s 行削除（残り %s 行）', rows.length - keep.length, keep.length);
    return rows.length - keep.length;
  } finally { lock.releaseLock(); }
}
function queueFind_(id) { const r = queueRows_().find(x => x['ID'] === id); if (!r) throw new Error('キューにありません: ' + id); return r; }
function queueSet_(row, patch) {
  const sh = queueSheet_();
  Object.keys(patch).forEach(k => { row[k] = patch[k]; sh.getRange(row._row, QUEUE_COLS.indexOf(k) + 1).setValue(patch[k]); });
  return row;
}
// st：工程記録（liveSteps_）。同じ号機に同じ工程の記録がもうあれば already に日時を入れる（照合のダブり・撮り直しに気づけるように）
function queueView_(r, st) {
  let cands = [];
  try { cands = r['候補'] ? JSON.parse(r['候補']) : []; } catch (e) {}
  let box = null;
  try { box = r['枠'] ? JSON.parse(r['枠']) : null; } catch (e) {}
  return {id: r['ID'], at: r['日時'], shotAt: r['撮影日時'] || '', purpose: r['目的'], state: r['状態'], fullId: r['元写真ID'], cropId: r['切り出しID'],
    box: box, read: r['AI読取'], key: r['キー'], cands: cands, error: r['エラー'], seen: r['確認'] === '済', rec: r['記録'] || '',
    already: alreadyStep_(r, st)};
}

function alreadyStep_(r, st) {
  const step = r['目的'], key = r['キー'];
  if (!key || r['記録'] || (step !== '塗装完了' && step !== '出荷')) return '';
  st = st || liveSteps_();
  return (st[key] && st[key][step]) || '';
}

/** 撮った写真（全体・縮小済み）を積む。すぐ返る */
// 読取キューの見出し：列を足したら（記録・鮮明度・撮影日時…）足りない分だけ書き足す
function queueHeader_() {
  const sh = queueSheet_(), n = QUEUE_COLS.length, cur = sh.getRange(1, 1, 1, n).getDisplayValues()[0];
  if (cur.join('|') === QUEUE_COLS.join('|')) return;
  if (cur.every((h, i) => !h || h === QUEUE_COLS[i])) sh.getRange(1, 1, 1, n).setValues([QUEUE_COLS]).setFontWeight('bold');
}
// shotAt：写真の撮影日時（画面が写真のEXIFから読む。'yyyy/MM/dd HH:mm'）。アルバムからまとめて読んだ時に、記録の日時をこれにする
function apiQueueAdd(b64, purpose, shotAt) {
  const fullId = savePhoto_(b64, 'raw');
  const id = Utilities.getUuid().slice(0, 8);
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  queueHeader_();
  const row = QUEUE_COLS.map(c => ({'ID': id, '日時': now, '目的': purpose || '', '状態': '待ち', '元写真ID': fullId, '撮影日時': /^\d{4}\/\d\d\/\d\d \d\d:\d\d$/.test(shotAt || '') ? shotAt : ''})[c] || '');
  appendRows_(CONFIG.SHEETS.QUEUE, [row]);
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
    const r1 = callAI_(P_FULL + learnHint_(learned_()), [{data: FONT_REF_PNG, mime: 'image/png'}, b64], P_FULL_SCHEMA);
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
function apiQueueCrop(id, cropB64, enhB64, sharp) {
  let row = queueFind_(id);
  if (!row['切り出しID']) { queueHeader_(); row = queueSet_(row, {'切り出しID': savePhoto_(cropB64, 'kokuin'), '切り出し版': CROP_VER}); }
  if (sharp !== undefined && sharp !== null) { queueHeader_(); row = queueSet_(row, {'鮮明度': sharp}); }
  if (row['状態'] === '候補' && enhB64) {
    READ_DEADLINE = Date.now() + 40000; READ_LOG = [];
    const t0 = Date.now();
    try {
      const r1 = callAI_(P1 + learnHint_(learned_()), [{data: FONT_REF_PNG, mime: 'image/png'}, cropB64, enhB64], P1_SCHEMA);
      const m = matchRead_(r1, null);
      const read = m.fz.z + ' / ' + (r1.seizo || '?') + '-' + (m.g || '?') + ' ' + (r1.kensa || '');
      queueSet_(row, {'AI読取': read.trim(), '候補': JSON.stringify(m.cands), 'キー': m.sure && m.cands[0] ? m.cands[0].key : '', '状態': m.sure ? '確定' : '候補'});
      rlog_('キュー ' + id + ' 切り出しで読み直し → ' + row['状態'] + ' ' + read);
    } catch (e) { rlog_('キュー ' + id + ' 読み直し失敗 ' + e); }
    writeReadLog_(t0, row);
  }
  if (row['状態'] === '確定' && row['キー']) apiLinkPhoto(row['切り出しID'], row['キー'], row['AI読取'], row['目的'], row['鮮明度'], row['撮影日時']);
  return queueView_(row);
}

/** 候補から人が選んだ */
function apiQueuePick(id, key) {
  const row = queueSet_(queueFind_(id), {'キー': key, '状態': '確定'});
  const photo = row['切り出しID'] || row['元写真ID'];
  if (photo) apiLinkPhoto(photo, key, row['AI読取'], row['目的'], row['切り出しID'] ? row['鮮明度'] : '', row['撮影日時']);
  return queueView_(row);
}
/** 特定済みの写真から、目的の工程（塗装完了・出荷）をまとめて記録。詳細を開かずに一覧から押せる */
function apiQueueRecord(ids, person) {
  if (!person) throw new Error('読取画面の「確認者」で自分の名前を選んでください');
  // 2台が同時に「まとめて記録」しても二重に記録しないように、読むところから鍵をかける
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return queueRecord_(ids, person); } finally { lock.releaseLock(); }
}
function queueRecord_(ids, person) {
  queueHeader_();
  const rows = queueRows_(), arms = {};
  loadArms_().forEach(a => { arms[a['キー']] = a; });
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), add = [], done = [];
  ids.forEach(id => {
    const r = rows.find(x => x['ID'] === id), step = r && r['目的'];
    if (!r || r['状態'] !== '確定' || r['記録'] || (step !== '塗装完了' && step !== '出荷')) return;
    const a = arms[r['キー']];
    if (!a) return;
    const recId = Utilities.getUuid().slice(0, 8);
    // 撮影日時がわかる写真（アルバムから読んだ分）は、その日時で記録する
    const at = r['撮影日時'] || now;
    add.push([recId, at, a['キー'], a['図番'], a['号機'], step, person, r['撮影日時'] ? '撮影日時で記録（' + now + ' 登録）' : '', r['切り出しID'] || r['元写真ID'] || '', '']);
    done.push([r, recId]);
  });
  if (add.length) appendRows_(CONFIG.SHEETS.STEPS, add);
  return done.map(d => queueView_(queueSet_(d[0], {'記録': d[1]})));
}
/** 一覧から記録した工程を取り消す */
function apiQueueUnrecord(id) {
  const row = queueFind_(id);
  if (row['記録']) cancelRow_(CONFIG.SHEETS.STEPS, STEP_COLS, row['記録']);
  return queueView_(queueSet_(row, {'記録': ''}));
}
/** 一覧から消す（確認済み） */
function apiQueueDone(id) { return queueView_(queueSet_(queueFind_(id), {'確認': '済'})); }

/** 未確認のもの（新しい順・最大40件） */
function apiQueueList() {
  const st = liveSteps_();
  return queueRows_().filter(r => r['確認'] !== '済').slice(-40).reverse().map(r => queueView_(r, st));
}

/** トリガー用：画面を閉じても残った「待ち」を読む（2分以上前のもの） */
function processQueue() {
  const limit = Date.now() + 4 * 60 * 1000, old = Utilities.formatDate(new Date(Date.now() - 120000), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  queueRows_().filter(r => r['状態'] === '待ち' && r['日時'] < old).forEach(r => { if (Date.now() < limit) apiQueueProcess(r['ID']); });
  // 作り直した写真を元に戻す続き（画面を開かなくても5分おきに進める）
  if (limit - Date.now() > 30000) apiRevertBadRecrops(limit - Date.now());
}

function writeReadLog_(t0, row) {
  try {
    appendRows_(CONFIG.SHEETS.READLOG, [[Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'), ((Date.now() - t0) / 1000).toFixed(1),
      'キュー:' + row['状態'] + (row['エラー'] ? ' ' + row['エラー'].slice(0, 120) : ''), row['AI読取'] || '', row['キー'] || '', READ_LOG.join('\n').slice(0, 3000)]]);
  } catch (e) { console.warn(e); }
}

/* ===== 切り出し直し（10/27 ユーザー：切り出しがちゃんと出来てない写真がある）=====
 * 撮るだけの写真は 元写真＋AIの枠 が残っているので、画面の新しい切り出しでやり直せる。
 * ファイルの中身だけ差し替える（写真ID・ファイル名・刻印写真シートのひも付けはそのまま） */
/** 切り出し直せる写真：[{cropId, fullId, box}]。cropId を渡すとその1枚、無ければ直近 days 日 */
// oldOnly：今の版（CROP_VER）より前のやり方で切り出したものだけ（画面が開いた時に裏で少しずつ作り直す）
function apiRecropList(cropId, days, oldOnly) {
  const from = Utilities.formatDate(new Date(Date.now() - (days || 14) * 86400000), 'Asia/Tokyo', 'yyyy/MM/dd');
  const rows = queueRows_();
  const out = rows.filter(r => r['切り出しID'] && r['元写真ID'] && r['枠'] && (cropId ? r['切り出しID'] === cropId : r['日時'] >= from) && (!oldOnly || (+r['切り出し版'] || 0) < CROP_VER))
    .map(r => ({cropId: r['切り出しID'], fullId: r['元写真ID'], box: JSON.parse(r['枠'])}));
  // その場で撮った写真（キューに無い＝黄色い枠の中だけ保存）は作り直さない（10/27 やってみたら欠けたので中止。元に戻す apiRevertBadRecrops）
  return out;
}
function photoRows_() {
  const sh = sheet_(CONFIG.SHEETS.PHOTOS), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, PHOTO_COLS.length).getDisplayValues().map((r, i) => { const o = {_row: i + 2}; PHOTO_COLS.forEach((c, j) => { o[c] = r[j]; }); return o; });
}
/** 切り出し画像の中身を差し替え＋鮮明度を書き直す */
// whole：その場で撮った写真（版は刻印写真シートに書く）。b64 が空なら作り直さず「済」にするだけ。元の画像はドライブの版の履歴に残す
function apiRecropSave(cropId, b64, sharp, whole) {
  if (b64) Drive.Files.update({}, cropId, Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg'), {supportsAllDrives: true, keepRevisionForever: true});
  const sh = sheet_(CONFIG.SHEETS.PHOTOS);
  if (whole) {
    if (sh.getRange(1, PHOTO_COLS.length).getValue() !== '切り出し版') sh.getRange(1, 1, 1, PHOTO_COLS.length).setValues([PHOTO_COLS]).setFontWeight('bold');
    photoRows_().filter(p => p['写真ID'] === cropId).forEach(p => sh.getRange(p._row, PHOTO_COLS.indexOf('切り出し版') + 1).setValue(WHOLE_VER));
    if (!b64) return true;
  }
  queueHeader_();
  const q = queueRows_().find(r => r['切り出しID'] === cropId);
  if (q) queueSet_(q, {'切り出し版': CROP_VER});
  if (sharp !== undefined && sharp !== null) {
    if (q) queueSet_(q, {'鮮明度': sharp});
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      const ids = sh.getRange(2, PHOTO_COLS.indexOf('写真ID') + 1, n, 1).getDisplayValues().map(r => r[0]), i = ids.indexOf(cropId);
      if (i >= 0) sh.getRange(i + 2, PHOTO_COLS.indexOf('鮮明度') + 1).setValue(sharp);
    }
  }
  return true;
}
/** その場で撮った写真（黄色い枠の中だけ）の刻印の範囲をAIに探させる（撮るだけと同じ仕上がりにするため。10/27 ユーザー：LQ12B10421F1_102 の仕上がりに）。無ければ null */
function apiPhotoBox(photoId) {
  READ_DEADLINE = Date.now() + 40000; READ_LOG = [];
  const b64 = Utilities.base64Encode(origBytes_(photoId));
  const r = callAI_(P_FULL, [{data: FONT_REF_PNG, mime: 'image/png'}, b64], P_FULL_SCHEMA);
  return r.found && r.box && r.box.length === 4 ? r.box : null;
}

/** 写真の一番最初の版（切り出し直す前の元画像）。版が1つならそのまま */
function origBytes_(id) {
  try {
    const revs = (Drive.Revisions.list(id, {fields: 'revisions(id,modifiedTime)'}).revisions || []);
    if (revs.length > 1) {
      revs.sort((a, b) => String(a.modifiedTime).localeCompare(String(b.modifiedTime)));
      const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/revisions/' + revs[0].id + '?alt=media&supportsAllDrives=true',
        {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true});
      if (res.getResponseCode() === 200) return res.getBlob().getBytes();
    }
  } catch (e) { console.warn('元の版が読めない ' + id + ' ' + e); }
  return DriveApp.getFileById(id).getBlob().getBytes();
}
/** 切り出し直す前の元画像に戻す（10/27 ユーザー：作り直しがダメな写真）。もう自動で作り直さない */
function apiRecropRevert(id) {
  Drive.Files.update({}, id, Utilities.newBlob(origBytes_(id), 'image/jpeg'), {supportsAllDrives: true});
  apiRecropSave(id, '', null, true);
  return apiPhoto(id);
}
/** 画面用：元画像（data URL） */
function apiPhotoOrig(id) { return 'data:image/jpeg;base64,' + Utilities.base64Encode(origBytes_(id)); }

// 10/27 その場で撮った写真（黄色い枠の中だけ）の作り直しは、もともと枠いっぱいの写真をさらに切るので
// 1行目が欠けたり一部だけ拡大されたりした（LC12B10494F1_63・YY12B00902F1G2_2072 など）。作り直した分は全部元の画像に戻し、以後はやらない
function apiRevertBadRecrops(ms) {
  const budget = typeof ms === 'number' ? ms : 240000;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('RECROP_REVERTED_WHOLE') === '1') return 0;
  const c = CacheService.getScriptCache();
  if (c.get('reverting')) return 0;
  c.put('reverting', '1', 600);
  let n = 0;
  const t0 = Date.now();
  try {
    const inQ = {};
    queueRows_().forEach(r => { inQ[r['切り出しID']] = 1; inQ[r['元写真ID']] = 1; });
    const done = JSON.parse(props.getProperty('RECROP_REVERTED_IDS2') || '[]');
    const ids = photoRows_().filter(p => p['写真ID'] && !inQ[p['写真ID']] && p['切り出し版']).map(p => p['写真ID']);
    for (const id of ids) {
      if (done.indexOf(id) >= 0) continue;
      if (Date.now() - t0 > budget) return n;   // 時間切れ前にやめる（次に開いた時に続き）
      try {
        const revs = Drive.Revisions.list(id, {fields: 'revisions(id)'}).revisions || [];
        if (revs.length > 1) { Drive.Files.update({}, id, Utilities.newBlob(origBytes_(id), 'image/jpeg'), {supportsAllDrives: true}); n++; }
      } catch (e) { console.warn('戻せない ' + id + ' ' + e); }
      done.push(id); props.setProperty('RECROP_REVERTED_IDS2', JSON.stringify(done));
    }
    props.setProperty('RECROP_REVERTED_WHOLE', '1');
  } finally { c.remove('reverting'); }
  return n;
}
