/**
 * GASエディタから実行して確認する関数
 */
// 取込フォルダの最新Excelを取り込む → ログに件数が出ればOK
function testImport() {
  console.log(JSON.stringify(importLatest()));
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
  console.log(JSON.stringify(callClaude_(P1, [b], P1_SCHEMA)));
}
// 照合ロジック
function testMatch() {
  console.log(JSON.stringify(fitZuban_('YN12BZ0Z69F1')));  // → YN12B20269F1
  console.log(fitGoki_('11S2'));                           // → 1152
}
