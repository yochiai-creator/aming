// GASの純粋関数（照合・Excel解析）と画面を node で確認する
const fs = require("fs"), path = require("path"), vm = require("vm"), assert = require("assert");
const G = path.join(__dirname, "..", "gas");
const ctx = vm.createContext({console});
["01_設定.js", "03_照合.js", "03_取込.js"].forEach(f => vm.runInContext(fs.readFileSync(path.join(G, f), "utf8").replace(/^if \(typeof module.*$/m, ""), ctx, {filename: f}));
const run = s => vm.runInContext(s, ctx);
const errors = [];
function t(name, fn) { try { fn(); console.log("ok  ", name); } catch (e) { errors.push(name); console.log("NG  ", name, e.message); } }

t("fitZuban 桁型補正", () => {
  assert.strictEqual(run("fitZuban_('YN12BZ0Z69F1').z"), "YN12B20269F1");
  assert.strictEqual(run("fitZuban_('YY12B0O902F1G2').z"), "YY12B00902F1G2");
  assert.strictEqual(run("fitGoki_('11S2')"), "1152");
});
t("rankArms 完全一致が1位・確定", () => {
  const r = run(`rankArms_('YN12B20269F1', '1152', [
    {z:'YN12B20269F1', g:'1151', ship:'2026/10/05'}, {z:'YN12B20269F1', g:'1152', ship:'2026/10/07'},
    {z:'YN12B20271F1', g:'1152', ship:'2026/10/07'}, {z:'LS12B10010F1', g:'39', ship:'2026/10/07'}], {today:'2026/10/07'})`);
  assert.strictEqual(r[0].arm.g, "1152"); assert.strictEqual(r[0].arm.z, "YN12B20269F1");
  assert.strictEqual(run("isSure_")(r), false); // 1151 が近いので2回目へ回す
});
t("isExact 完全一致だけ確定", () => {
  assert.ok(run("isExact_")("YY12B00902F1G2", "2020", {z: "YY12B00902F1G2", g: "2020"}));
  assert.ok(run("isExact_")("YY12B00902F1G2", "0020", {z: "YY12B00902F1G2", g: "20"}));
  assert.ok(!run("isExact_")("YY12B00902F1G2", "2724", {z: "YY12B00902F1G2", g: "2024"}));
  assert.ok(!run("isExact_")("Y?12B00273F1", "247", {z: "YB12B00273F1", g: "247"}));
});
t("自動学習：読み間違いを集めて照合に効かせる", () => {
  assert.strictEqual(JSON.stringify(run("parseRead_('YY12B00902F1G2 / 2512-2724 U')")), '{"z":"YY12B00902F1G2","g":"2724"}');
  assert.deepStrictEqual(Array.from(run("diffPairs_('2724', '2024')")), ["7>0"]);
  const st = run(`learnStats_([
    {read:'YY12B00902F1G2 / 2512-2724 U', z:'YY12B00902F1G2', g:'2024'},
    {read:'YY12B00902F1G2 / 2610-2722 U', z:'YY12B00902F1G2', g:'2022'},
    {read:'YY12B10167F1 / 2610-1331 UM', z:'YY12B10167F1', g:'1331'},
    {read:'Y?12B00273F1 / 2510-247 UM', z:'YB12B00273F1', g:'247'}])`);
  assert.strictEqual(st.pairs["7>0"], 2); assert.strictEqual(st.total, 4); assert.strictEqual(st.miss, 2);
  assert.ok(run("learnHint_")(st).includes("「0」を「7」"));
  // 学習あり：2724 の読みで 2024 が 2724 寄りの別号機より上に来る
  const arms = `[{z:'YY12B00902F1G2', g:'2024'}, {z:'YY12B00902F1G2', g:'2124'}]`;
  assert.strictEqual(run(`rankArms_('YY12B00902F1G2', '2724', ${arms}, {learned: {'7>0': 2}})`)[0].arm.g, "2024");
  assert.ok(!run("isSure_")(run(`rankArms_('YY12B00902F1G2', '2724', ${arms}, {learned: {'7>0': 2}})`))); // 学習だけで自動確定はしない
});
t("rankArms 続き番号優先", () => {
  const r = run(`rankArms_('LS12B10010F1', '4?', [{z:'LS12B10010F1', g:'40'}, {z:'LS12B10010F1', g:'47'}], {last:{z:'LS12B10010F1', g:'39'}})`);
  assert.strictEqual(r[0].arm.g, "40"); assert.ok(r[0].seq);
});
t("parseShipValues 見出しで列を探す・重複・対象外", () => {
  const hdr = ["注文番号（写し）", "合計", "過不足", "払出完了日", "検査完了日", "機器　機種", "建機号機", "納入週", "図番", "x", "y", "号機", "仕様",
    "着工", "塗装完了日", "塗装後修正完了日", "最新出荷日", "納入日", "使用製缶ベース情報", "出荷先", "注文番号"];
  const row = (z, g, ship, extra) => { const r = hdr.map(() => ""); Object.assign(r, {8: z, 11: g, 16: ship, 5: "13ton", 6: "YY09066073", 4: new Date(2026, 8, 25), 13: new Date(1899, 11, 30), 19: "正和(13ton)", 20: "#N/A"}, extra || {}); return r; };
  const vals = [["", "タイトル"], hdr,
    row("YY12B00902F1G2", 2007, new Date(2026, 9, 5)),
    row("YY12B00902F1G2", 2007, new Date(2026, 9, 9)),   // 重複 → 新しい方
    row("LS02B02470P1", 1, new Date(2026, 9, 5), {5: "ブームブラケット"}), // アーム以外
    row("YB12B00209F1", "仮", "")];
  const p = run("parseShipValues_")(vals);
  assert.strictEqual(p.arms.length, 1);
  const a = p.arms[0];
  assert.strictEqual(a["キー"], "YY12B00902F1G2_2007");
  assert.strictEqual(a["出荷日"], "2026/10/09");
  assert.strictEqual(a["検査完了日"], "2026/09/25");
  assert.strictEqual(a["着工"], "");       // 1899年 → 空
  assert.strictEqual(a["注文番号"], "");   // #N/A → 空
  assert.strictEqual(a["出荷先"], "正和(13ton)");
  assert.strictEqual(p.dups.length, 1);
  assert.strictEqual(p.skipped, 2);
});

t("diffArm 日付と出荷先の変更だけ拾う", () => {
  const d = run("diffArm_")({"塗装完了日": "2026/09/30", "出荷先": "正和", "仕様": "A"}, {"塗装完了日": "2026/10/02", "出荷先": "正和", "仕様": "B"});
  assert.deepStrictEqual(JSON.parse(JSON.stringify(d)), [["塗装完了日", "2026/09/30", "2026/10/02"]]);
});
t("parseSettingRows 確認者を1行1人で", () => {
  const o = run("parseSettingRows_")([["AI読取", "gemini", ""], ["確認者", "A.ましゃたか", "説明"], ["", "田中", ""], ["", "", ""], ["", "佐藤", ""]]);
  assert.strictEqual(o["確認者"], "A.ましゃたか,田中,佐藤"); assert.strictEqual(o["AI読取"], "gemini");
});
t("fileDateKey ファイル名の日付", () => {
  assert.strictEqual(run("fileDateKey_")("出荷予定　日程表変更A(26年10月7日).xlsm"), "20261007");
  assert.ok(run("fileDateKey_")("出荷予定　日程表変更A(26年10月7日).xlsm") > run("fileDateKey_")("出荷予定　日程表変更A(26年9月24日).xlsm"));
  assert.strictEqual(run("fileDateKey_")("x.xlsm"), "00000000");
});

// 実データ（リポジトリに入れない）で解析だけ確認：test/fixtures/ship_values.json があれば
const fx = path.join(__dirname, "fixtures", "ship_values.json");
if (fs.existsSync(fx)) t("実データ解析", () => {
  const vals = JSON.parse(fs.readFileSync(fx, "utf8")).map(r => r.map(v => v && v.$d ? new Date(v.$d) : v));
  const p = run("parseShipValues_")(vals);
  console.log("     arms", p.arms.length, "dups", p.dups.length, "skipped", p.skipped, JSON.stringify(p.arms.find(a => a["図番"] === "YN12B20269F1")));
  assert.ok(p.arms.length > 5000);
});

/* ---------- 画面（google.script.run をモック） ---------- */
const {JSDOM} = require("jsdom");
const arm = {key: "YN12B20269F1_1152", z: "YN12B20269F1", g: "1152", kishu: "SK200　B165", spec: "HD10型", kenki: "YN13-12345", to: "正和", ship: "2026/10/07", deliv: "2026/10/08", status: "", by: "", at: ""};
const calls = [];
const api = {
  apiBoot: () => ({today: "2026/10/07", me: "", people: ["落合", "田中"], hasKey: true, days: [{date: "2026/10/07", total: 3, todo: 2, ng: 0}, {date: "2026/10/08", total: 5, todo: 5, ng: 0}]}),
  apiDays: () => [{date: "2026/10/07", total: 3, todo: 1, ng: 0}],
  apiDay: () => [arm], apiUnchecked: () => [arm], apiSearch: () => [arm],
  apiArm: () => Object.assign({}, arm, {steps: [
    {name: "着工", date: "2026/09/24", step: "", changes: [], recs: []},
    {name: "塗装完了", date: "2026/10/02", step: "塗装完了", changes: [{at: "2026/10/07 06:00", from: "2026/09/30", to: "2026/10/02"}], recs: []},
    {name: "出荷", date: "2026/10/07", step: "出荷", changes: [], recs: []}], otherChanges: [], photos: [{id: "p1", at: "2026/10/07 10:00", name: "x.jpg"}], order: "A5", base: "YY", updated: "", taps: [], holes: 33}),
  apiQueueList: () => [{id: "q1", state: "確定", key: arm.key, purpose: "塗装完了", cands: [arm], read: "YN12B20269F1 / 2609-1152 UM"}, {id: "q2", state: "候補", purpose: "出荷", cands: [arm], read: "YN12B2026?F1 / 1152"}],
  apiQueueDone: id => ({id, seen: true}),
  apiQueuePick: (id, key) => ({id, state: "確定", key, cands: [arm]}),
  apiDeletePhoto: () => Object.assign(api.apiArm(), {photos: []}),
  apiPhoto: () => "data:image/jpeg;base64,AAAA",
  apiSaveStep: rec => { const v = api.apiArm(); v.steps[1].recs.push({id: "s1", at: "2026/10/07 10:00", by: rec.person, note: rec.note, canceled: false}); return v; },
  apiSaveTap: rec => Object.assign(api.apiArm(), {status: rec.result, taps: [{id: "x", at: "2026/10/07 10:00", result: rec.result, holes: rec.holes, treated: rec.treated, by: rec.person, note: "", photo: "", canceled: false}]})
};
const html = fs.readFileSync(path.join(G, "index.html"), "utf8").replace(/<link[^>]+>/g, "");
const dom = new JSDOM(html, {runScripts: "outside-only", pretendToBeVisual: true});
const w = dom.window;
w.scrollTo = () => {}; w.HTMLElement.prototype.scrollIntoView = () => {};
w.google = {script: {get run() {
  const h = {};
  const p = new Proxy({}, {get: (_, k) => k === "withSuccessHandler" ? f => (h.ok = f, p) : k === "withFailureHandler" ? f => (h.ng = f, p) : (...a) => { calls.push(k); setTimeout(() => { try { h.ok(api[k](...a)); } catch (e) { h.ng(e); } }, 0); }});
  return p;
}}};
w.eval(html.match(/<script>([\s\S]*)<\/script>/)[1]);
const wait = ms => new Promise(r => setTimeout(r, ms));
const click = sel => w.document.querySelector(sel).dispatchEvent(new w.MouseEvent("click", {bubbles: true}));
process.on('unhandledRejection', e => { errors.push('unhandled'); console.log('NG   unhandled', e); });
(async () => {
  await wait(30);
  t("撮るだけ：キューが出る", () => { assert.strictEqual(w.document.querySelectorAll("#qList .qi").length, 2); assert.ok(w.document.querySelector("#qList").textContent.includes("特定")); });
  click('[data-act="qopen"][data-id="q1"]'); await wait(40);
  t("撮るだけ：確定→記録を開く（目的バナー）", () => { assert.ok(calls.includes("apiArm")); assert.ok(w.document.querySelector(".goal").textContent.includes("塗装完了")); });
  click('[data-act="nav"][data-v="scan"]'); await wait(10);
  click('[data-act="qpick"][data-id="q2"]'); await wait(40);
  t("撮るだけ：候補を選ぶ", () => assert.ok(calls.includes("apiQueuePick")));
  click('[data-act="nav"][data-v="scan"]'); await wait(10);
  click('[data-act="purpose"][data-v="出荷"]');
  t("撮影の目的：選べる", () => assert.ok(w.document.querySelector('[data-act="purpose"][data-v="出荷"]').classList.contains("on")));
  t("起動：未チェックバッジ", () => assert.strictEqual(w.document.getElementById("todoBadge").textContent, "7"));
  click('[data-act="nav"][data-v="list"]'); await wait(30);
  t("一覧：出荷日と未チェック", () => {
    assert.ok(w.document.querySelector(".days .on").textContent.includes("10/7"));
    assert.ok(w.document.getElementById("uncheckedSec").textContent.includes("1本"));
  });
  click('#dayList [data-act="arm"]'); await wait(30);
  t("詳細：ねじ穴数の自動入力", () => assert.strictEqual(w.document.getElementById("fHoles").value, "33"));
  click('[data-act="allok"]'); await wait(30);
  t("全数OK：確認者なしは止める", () => { assert.ok(w.document.getElementById("formMsg").textContent.includes("確認者")); assert.ok(!calls.includes("apiSaveTap")); });
  click('[data-act="person"][data-n="落合"]'); click('[data-act="allok"]'); await wait(30);
  t("全数OK：記録される", () => {
    assert.ok(calls.includes("apiSaveTap"));
    assert.ok(w.document.querySelector(".hist").textContent.includes("処置 33"));
  });
  t("刻印写真：詳細に表示", () => assert.strictEqual(w.document.querySelector(".photos img").getAttribute("src"), "data:image/jpeg;base64,AAAA"));
  w.confirm = () => true; click('[data-act="delPhoto"]'); await wait(30);
  t("刻印写真：削除", () => { assert.ok(calls.includes("apiDeletePhoto")); assert.ok(!w.document.querySelector(".photos")); });
  t("工程ログ：日程変更の表示", () => assert.ok(w.document.querySelector(".steps .chg").textContent.includes("2026/09/30")));
  w.prompt = () => "メモ";
  click('[data-act="nav"][data-v="settings"]'); await wait(10); click('[data-act="me"][data-n="田中"]'); await wait(10);
  click('[data-act="nav"][data-v="list"]'); await wait(30); click('#dayList [data-act="arm"]'); await wait(30);
  click('[data-act="step"][data-s="塗装完了"]'); await wait(30);
  t("工程ログ：塗装完了を記録", () => { assert.ok(calls.includes("apiSaveStep")); assert.ok(w.document.querySelector(".steps .rec").textContent.includes("田中")); });
  console.log("errors", JSON.stringify(errors));
  process.exitCode = errors.length ? 1 : 0;
})();
