const {JSDOM} = require("jsdom"); const fs = require("fs");
const errs = [];
const dom = new JSDOM(fs.readFileSync("./dist/arm-kensa-demo.html","utf8"), {runScripts:"dangerously", pretendToBeVisual:true, beforeParse(w){ w.scrollTo=()=>{}; w.addEventListener("error", e=>errs.push(e.message)); }});
const w = dom.window, d = w.document, click = s => { const el = d.querySelector(s); if(!el) throw new Error("missing "+s); el.dispatchEvent(new w.MouseEvent("click",{bubbles:true})); };
setTimeout(()=>{ try{
 click('#nav [data-v="list"]'); console.log(d.querySelector("#view h2").textContent, d.querySelector("#view .arm").textContent.replace(/\s+/g," "));
 click('#view .arm[data-g="39"]'); console.log("arm:", !!d.querySelector("#view .films"), !!d.querySelector(".tabs"));
 click('#view [data-act="set"][data-p="paint.sabi"][data-v="有"]'); click('#view [data-act="nav"][data-v="list"]'); click('#view [data-act="nav"][data-v="day"]');
 console.log("day rows:", d.querySelectorAll("#view .sheet tbody tr").length);
 const q=d.getElementById("q"); click('#nav [data-v="scan"]'); q.value="20269 1152"; q.dispatchEvent(new w.Event("input")); click('#qres .cand'); console.log("search open:", d.querySelector("#view .plate").textContent.replace(/\s+/g," "));
}catch(e){console.log("FAIL",e.message)} console.log("errors",errs); process.exit(0);},300);
