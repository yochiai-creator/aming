/**
 * 刻印読取結果と台帳の照合（純粋関数。node でもテストできる）
 */
const SIM_GROUPS = [['0','O','D','Q'],['1','I','L','7','T'],['2','Z'],['5','S','6'],['8','B','3'],['9','7','4'],['6','G','5'],['F','E','P'],['N','M'],['U','V']];
const SIM_MAP = {};
SIM_GROUPS.forEach(g => g.forEach(a => g.forEach(b => { if (a !== b) SIM_MAP[a + b] = 1; })));

// L = 学習した読み間違い {'AIの字>正しい字': 回数}。2回以上あった間違いはほぼ同じ字として扱う
function subCost_(a, b, L) {
  if (a === b) return 0; if (a === '?' || b === '?') return 0.25;
  const n = L ? L[a + '>' + b] || 0 : 0, c = SIM_MAP[a + b] ? 0.35 : 1;
  return n >= 2 ? Math.min(c, 0.15) : n === 1 ? Math.min(c, 0.5) : c;
}
// 打刻で紛らわしい文字ペアを安くした編集距離（a=AIの読み, b=台帳）
function wdist_(a, b, L) {
  const m = a.length, n = b.length, d = [];
  for (let i = 0; i <= m; i++) d.push([i]);
  for (let j = 1; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
    d[i][j] = Math.min(d[i-1][j] + 1, d[i][j-1] + 1, d[i-1][j-1] + subCost_(a[i-1], b[j-1], L));
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
    let cost = (z ? wdist_(z, a.z, opts.learned) : 3) + (g ? wdist_(g, String(a.g), opts.learned) * 1.5 : 2);
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

/* ===== 自動学習：確定した刻印写真（AIの読み と 正解）から読み間違いを集める ===== */
// 「YY12B00902F1G2 / 2512-2724 U」→ {z, g}
function parseRead_(text) {
  const p = String(text || '').split(' / ');
  const m = (p[1] || '').match(/-\s*([0-9A-Z?]+)/);
  const z = normKokuin_(p[0]), g = m ? fitGoki_(m[1]) : '';
  return {z: /^\?*$/.test(z) ? '' : z, g: /^\?*$/.test(g) ? '' : g}; // 全部 ? は読めてないので使わない
}
// 同じ長さの部分で、AIの字→正しい字 の置き換えを拾う（編集距離の逆たどり）
function diffPairs_(a, b) {
  const m = a.length, n = b.length, d = [];
  for (let i = 0; i <= m; i++) { d.push([i]); }
  for (let j = 1; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
    d[i][j] = Math.min(d[i-1][j] + 1, d[i][j-1] + 1, d[i-1][j-1] + (a[i-1] === b[j-1] ? 0 : 1));
  const out = [];
  let i = m, j = n;
  while (i > 0 && j > 0) {
    const c = a[i-1] === b[j-1] ? 0 : 1;
    if (d[i][j] === d[i-1][j-1] + c) { if (c && a[i-1] !== '?') out.push(a[i-1] + '>' + b[j-1]); i--; j--; }
    else if (d[i][j] === d[i-1][j] + 1) i--; else j--;
  }
  return out;
}
// rows: [{read:'AI読取', z:'図番', g:'号機'}] → {pairs:{'7>0':n}, examples:[{ai, ok}], total, miss}
function learnStats_(rows) {
  const pairs = {}, examples = [];
  let total = 0, miss = 0;
  rows.forEach(r => {
    const ai = parseRead_(r.read);
    if (!ai.z && !ai.g) return;
    total++;
    const ps = (ai.z ? diffPairs_(ai.z, r.z) : []).concat(ai.g ? diffPairs_(ai.g, String(r.g)) : []);
    if (!ps.length && (!ai.z || ai.z.length === r.z.length) && (!ai.g || ai.g.length === String(r.g).length)) return; // ? だけなら間違いに数えない
    miss++;
    ps.forEach(k => { pairs[k] = (pairs[k] || 0) + 1; });
    examples.push({ai: ai.z + ' ' + ai.g, ok: r.z + ' ' + r.g});
  });
  return {pairs: pairs, examples: examples.slice(-8), total: total, miss: miss};
}
// AIに渡すヒント文（2回以上の間違いと最近の例）
function learnHint_(st) {
  if (!st || !st.miss) return '';
  const top = Object.keys(st.pairs).filter(k => st.pairs[k] >= 2).sort((x, y) => st.pairs[y] - st.pairs[x]).slice(0, 8);
  const lines = ['これまでの実績で、あなたは次の読み間違いをしています。同じ間違いに注意してください。'];
  if (top.length) lines.push('よくある間違い: ' + top.map(k => '「' + k.split('>')[1] + '」を「' + k.split('>')[0] + '」と読んだ(' + st.pairs[k] + '回)').join('、'));
  st.examples.slice(-5).forEach(e => lines.push('例: 読み ' + e.ai + ' → 正解 ' + e.ok));
  return '\n' + lines.join('\n');
}

if (typeof module !== 'undefined') module.exports = {parseRead_, diffPairs_, learnStats_, learnHint_, wdist_, fitZuban_, fitGoki_, rankArms_, isSure_, isExact_, shiftDate_, normKokuin_};
