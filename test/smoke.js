const {JSDOM} = require("jsdom");
const fs = require("fs");
const html = fs.readFileSync("./dist/arm-kensa-demo.html", "utf8");
const errs = [];
const dom = new JSDOM(html, {runScripts: "dangerously", pretendToBeVisual: true, beforeParse(w) { w.addEventListener("error", e => errs.push(e.message)); w.console.error = (...a) => errs.push(a.join(" ")); }});
const w = dom.window, d = w.document; w.scrollTo = () => {};
const click = sel => { const el = typeof sel === "string" ? d.querySelector(sel) : sel; if (!el) throw new Error("missing " + sel); el.dispatchEvent(new w.MouseEvent("click", {bubbles: true})); };
setTimeout(() => {
  try {
    click('#nav [data-v="list"]');
    console.log("list:", d.querySelector("#view h2").textContent, d.querySelectorAll("#view .arm").length);
    click('#view .arm[data-g="396"]');
    console.log("arm head:", d.querySelector("#view .plate").textContent.trim().replace(/\s+/g," "), "|", d.querySelector(".tabs").textContent);
    // mark item
    const b = d.querySelector('#view [data-act="set"][data-p="mfg.B.items.0.v"][data-v="否"]'); click(b);
    console.log("after 否:", d.querySelector(".tabs").textContent, d.querySelectorAll("#view .item.ng").length);
    click('#view [data-act="armTab"][data-v="paint"]');
    console.log("paint film inputs:", d.querySelectorAll("#view .films input").length);
    click('#view [data-act="sign"][data-p="paint.touchup"]');
    console.log("modal:", d.querySelector("#modalBody").textContent.slice(0,40));
    click('#modalBody [data-act="close"]');
    click('#view [data-act="nav"][data-v="list"]');
    click('#view [data-act="nav"][data-v="day"]');
    console.log("day rows:", d.querySelectorAll("#view .sheet tbody tr").length, d.querySelectorAll("#view .signs .sign").length);
    // open unscanned arm and sign without me
    click('#view .sheet [data-act="openArm"][data-g="1152"]');
    click('#view [data-act="allok"][data-sec="C"]');
    console.log("allok:", d.querySelector(".tabs").textContent);
    click('#view [data-act="sign"][data-p="mfg.C.sign"]');
    console.log("picker names:", d.querySelectorAll('#modalBody [data-act="pickDo"]').length);
    click('#modalBody [data-act="pickDo"]');
    console.log("signed:", d.querySelector('#view [data-p="mfg.C.sign"]').textContent.replace(/\s+/g," "));
    // input change
    const inp = d.querySelector('#view input[data-p="head.kanGoki"]'); inp.value = "1152"; inp.dispatchEvent(new w.Event("change", {bubbles: true}));
    console.log("kanGoki:", d.querySelector('#view input[data-p="head.kanGoki"]').value);
    click('#nav [data-v="settings"]');
    console.log("settings names:", d.querySelectorAll("#view .nm").length);
    // manual search → open
    click('#nav [data-v="scan"]');
    const q = d.getElementById("q"); q.value = "20269 1153"; q.dispatchEvent(new w.Event("input"));
    click('#qres .cand'); console.log("opened via search:", d.querySelector("#view .plate").textContent.trim().replace(/\s+/g," "));
  } catch (e) { console.log("FAIL", e.stack); }
  console.log("errors:", errs);
  process.exit(0);
}, 300);
