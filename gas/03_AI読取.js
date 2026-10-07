/**
 * 刻印のAI読取（Gemini / Claude）と台帳照合
 * 1回目：自由読取 → 台帳から上位5候補 → 2回目：写真と候補を見比べて選択（1回目で確定なら省略）
 */
const P1 = [
  '建機アームに打刻された刻印を切り抜いた写真です。1枚目が元画像、2枚目が凹凸を強調した画像です。',
  '刻印は2行あります。',
  '1行目: 図番。形式は 英大文字2文字 + 数字2桁 + "B" + 数字5桁 + "F" + 数字1桁、末尾に "G"+数字1桁 が付くことがある（例: LS12B10010F1, YY12B00902F1G2）。',
  '2行目: ◇マーク、製作年月4桁(YYMM)、"-"、号機(1〜4桁の数字)、検査記号(例 UM)。',
  'スラッシュ付きの0は数字の0です。自信のない文字は ? にしてください。推測で埋めないこと。'
].join('\n');
const P1_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['zuban', 'seizo', 'goki', 'kensa', 'note'],
  properties: {zuban: {type: 'string'}, seizo: {type: 'string'}, goki: {type: 'string'}, kensa: {type: 'string'}, note: {type: 'string'}}
};
const P2_SCHEMA = {
  type: 'object', additionalProperties: false,
  required: ['choice', 'confidence', 'reason'],
  properties: {choice: {type: 'string', enum: ['A', 'B', 'C', 'D', 'E', 'none']}, confidence: {type: 'string', enum: ['高', '中', '低']}, reason: {type: 'string'}}
};
function p2Prompt_(top) {
  return '同じ刻印の写真です。この刻印は次の候補のどれと一致しますか。1文字ずつ写真と見比べて判断してください。\n' +
    '特に号機（2行目の "-" の後の数字）と図番の数字部分を丁寧に確認すること。\n' +
    top.map((c, i) => 'ABCDE'[i] + ': 図番 ' + c.arm.z + '　号機 ' + c.arm.g).join('\n') +
    '\nどれとも一致しなければ choice は "none"。';
}

/**
 * 画面から呼ぶ。orig/enh は base64(JPEG)。last は直前に記録したアーム {z,g}
 * 戻り値 {read:{zuban,goki,seizo,kensa,fixed}, cands:[armView+cost+seq], pick:{index,confidence,reason}|null, sure, photoId}
 */
function apiRead(orig, enh, last) {
  const photoId = savePhoto_(orig);
  const imgs = [orig, enh];
  const r1 = callAI_(P1, imgs, P1_SCHEMA);
  const fz = fitZuban_(r1.zuban), g = fitGoki_(r1.goki);

  const taps = latestTaps_(loadTaps_());
  const arms = loadArms_().map(a => ({z: a['図番'], g: a['号機'], ship: a['出荷日'], done: !!taps[a['キー']], src: a}));
  const top = rankArms_(fz.z, g, arms, {today: today_(), last: last}).slice(0, 5);
  const sure = isSure_(top);
  let pick = sure ? {index: 0, confidence: '高', reason: '1回目で一致'} : null;
  if (!sure && top.length) {
    try {
      const r2 = callAI_(p2Prompt_(top), imgs, P2_SCHEMA);
      const i = 'ABCDE'.indexOf(r2.choice);
      pick = i >= 0 && i < top.length ? {index: i, confidence: r2.confidence, reason: r2.reason} : {index: -1, confidence: r2.confidence, reason: r2.reason};
    } catch (e) { console.warn('2回目失敗: ' + e); }
  }
  return {
    read: {zuban: fz.z, fixed: fz.fixed, goki: g, seizo: r1.seizo, kensa: r1.kensa, note: r1.note},
    cands: top.map(c => { const v = armView_(c.arm.src, taps[c.arm.src['キー']]); v.cost = Math.round(c.cost * 100) / 100; v.seq = c.seq; return v; }),
    pick: pick, sure: sure, photoId: photoId
  };
}

function callAI_(prompt, imgsB64, schema) {
  return CONFIG.AI_PROVIDER === 'claude' ? callClaude_(prompt, imgsB64, schema) : callGemini_(prompt, imgsB64, schema);
}

function callGemini_(prompt, imgsB64, schema) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY が未設定です');
  const parts = imgsB64.map(b => ({inline_data: {mime_type: 'image/jpeg', data: b}}));
  parts.push({text: prompt});
  const body = {
    contents: [{role: 'user', parts: parts}],
    generationConfig: {responseMimeType: 'application/json', responseSchema: geminiSchema_(schema)}
  };
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.GEMINI_MODEL + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {'x-goog-api-key': key},
    payload: JSON.stringify(body)
  });
  const code = res.getResponseCode(), json = JSON.parse(res.getContentText() || '{}');
  if (code === 429 || code === 503) throw new Error('混み合っています。少し待って撮り直してください');
  if (code !== 200) throw new Error('AI読取エラー ' + code + ': ' + (json.error && json.error.message || ''));
  const cand = (json.candidates || [])[0];
  if (!cand || !cand.content) throw new Error('AIが読取を返しませんでした（' + (cand && cand.finishReason || (json.promptFeedback && json.promptFeedback.blockReason) || '不明') + '）');
  return JSON.parse(cand.content.parts.map(p => p.text || '').join(''));
}
// JSON Schema → Gemini の responseSchema（OpenAPI形式。additionalProperties は使えない）
function geminiSchema_(s) {
  const o = {type: String(s.type).toUpperCase()};
  if (s.enum) o.enum = s.enum;
  if (s.required) o.required = s.required;
  if (s.properties) {
    o.properties = {};
    Object.keys(s.properties).forEach(k => { o.properties[k] = geminiSchema_(s.properties[k]); });
    o.propertyOrdering = Object.keys(s.properties);
  }
  return o;
}

function callClaude_(prompt, imgsB64, schema) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY が未設定です');
  const content = imgsB64.map(b => ({type: 'image', source: {type: 'base64', media_type: 'image/jpeg', data: b}}));
  content.push({type: 'text', text: prompt});
  const body = {
    model: CONFIG.CLAUDE_MODEL,
    max_tokens: 16000,
    output_config: {effort: CONFIG.CLAUDE_EFFORT, format: {type: 'json_schema', schema: schema}},
    fallbacks: 'default', // 安全判定で断られたとき別モデルで再実行（サーバー側）
    messages: [{role: 'user', content: content}]
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {'x-api-key': key, 'anthropic-version': '2023-06-01', 'anthropic-beta': 'server-side-fallback-2026-07-01'},
    payload: JSON.stringify(body)
  });
  const code = res.getResponseCode(), json = JSON.parse(res.getContentText() || '{}');
  if (code === 429 || code === 529) throw new Error('混み合っています。少し待って撮り直してください');
  if (code !== 200) throw new Error('AI読取エラー ' + code + ': ' + (json.error && json.error.message || ''));
  if (json.stop_reason === 'refusal') throw new Error('AIが読取を断りました');
  const text = (json.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  return JSON.parse(text);
}

function savePhoto_(b64) {
  const id = cfg_('PHOTO_FOLDER_ID');
  if (!id) return '';
  const name = 'kokuin_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '.jpg';
  return DriveApp.getFolderById(id).createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name)).getId();
}
