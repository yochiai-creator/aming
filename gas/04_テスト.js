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
