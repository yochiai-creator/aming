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

// AIが読んだ図番・号機が台帳のアームと1文字も違わず一致（? なし）なら確定扱い
function isExact_(z, g, arm) { return !!arm && !!z && !!g && !/\?/.test(z + g) && z === arm.z && String(+g) === String(arm.g); }

function shiftDate_(ymd, days) {
  const p = ymd.split('/').map(Number), d = new Date(p[0], p[1] - 1, p[2] + days);
  return d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + String(d.getDate()).padStart(2, '0');
}

if (typeof module !== 'undefined') module.exports = {wdist_, fitZuban_, fitGoki_, rankArms_, isSure_, isExact_, shiftDate_, normKokuin_};
