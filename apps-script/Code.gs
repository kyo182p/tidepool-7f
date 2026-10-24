/**
 * 對話整理 — 後端(Google Apps Script,綁定一份 Google 試算表)
 *
 * 網頁(GitHub Pages)把「代碼 + 動作」用 POST 送到這裡;
 * 這裡驗證代碼、用存在指令碼屬性的 ANTHROPIC_API_KEY 呼叫 Claude,
 * 再把逐字稿、每日摘要、單字寫進試算表。
 *
 * 注意:UrlFetchApp 單次呼叫約 60 秒就會逾時,所以
 *   - 轉錄一次只送 1 張小圖(網頁端切好)
 *   - 摘要和單字分成兩個動作,各自一次呼叫
 */

/* ===== 可調整的常數 ===== */
const MODELS = {
  transcribe: 'claude-opus-5-5',   // 看截圖轉逐字稿,精準最重要
  summary:    'claude-opus-5-5',   // 每日摘要(中英對照)
  vocab:      'claude-opus-5-5',   // 英文單字擷取
  attach:     'claude-opus-5-5',   // 附件(圖片、PDF)內容擷取
  match:      'claude-opus-5-5',   // 附件配對到訊息
  events:     'claude-opus-5-5',   // 每日事件抽取(期限、決策、突發、人員、船機)
  report:     'claude-opus-5-5',   // 期間報告
  link:       'claude-opus-5-5',   // 信件 ↔ 對話 找相關訊息
  ask:        'claude-opus-5-5',   // 排班工具的「問 AI」
};
// 每百萬 token 美元 [輸入, 輸出],只用來估算「用量」工作表
const PRICE = { 'claude-opus-5-5': [5, 25], 'claude-sonnet-5-5': [3, 15], 'claude-haiku-4-5-20251001': [1, 5] };
// 實測 Opus 約每秒 100 token;UrlFetchApp 約 60 秒逾時,所以單次輸出抓 4500 以內
const MAX_TOKENS = { transcribe: 3500, summary: 4500, vocab: 3500, attach: 4000, match: 1000, events: 4000, report: 4500, link: 1500, ask: 2500 };
const LINK_WINDOW = 7;     // 信件前後幾天內的對話拿來比對
const LINK_CANDS = 45;     // 每封信最多送幾則候選對話給 AI
const ATT_FOLDER = '對話整理附件';   // 雲端硬碟資料夾名稱
const MAX_IMAGES = 2;
const TZ = 'Asia/Taipei';

const SHEETS = {
  users:  { name: '使用者',   head: ['代碼', '名稱', '啟用', '備註', '最後使用'] },
  groups: { name: '群組',     head: ['群組ID', '名稱', '說明', '建立時間', '啟用', '訊息數', '最早日期', '最新日期', '類型'] },
  msgs:   { name: '訊息',     head: ['訊息ID', '群組ID', '日期', '時間', '發言人', '內容', '引用', '附件', '語言', '首次上傳', '出現的上傳', '建立時間'] },
  sums:   { name: '摘要',     head: ['群組ID', '日期', '更新時間', '模型', '標題', '重點', 'JSON'] },
  vocab:  { name: '單字',     head: ['單字', '中文', '說明', '例句', '群組ID', '日期', '訊息ID', '上傳ID', '建立時間'] },
  runs:   { name: '上傳紀錄', head: ['上傳ID', '時間', '使用者', '群組ID', '檔名', '截圖日期', '片數', '訊息數', '新增訊息', '最早日期', '最新日期'] },
  atts:   { name: '附件',     head: ['附件ID', '群組ID', '訊息ID', '日期', '檔名', '類型', '檔案ID', '雲端連結', '標題', '內容', '上傳ID', '建立時間'] },
  events: { name: '事件',     head: ['事件ID', '群組ID', '來源日期', '類型', '事件日期', '日期說明', '標題', '內容', '相關人', '對象', '主題', '狀態', '嚴重度', '訊息ID', '更新時間'] },
  reports:{ name: '報告',     head: ['報告ID', '建立時間', '使用者', '起', '迄', '群組', '標題', 'JSON1', 'JSON2', 'JSON3', 'JSON4', 'JSON5', 'JSON6'] },
  vmarks: { name: '單字標記', head: ['單字', '標記', '練習次數', '答對', '答錯', '最後練習', '最近結果', '更新時間', '使用者'] },
  mails:  { name: '信件',     head: ['訊息ID', '群組ID', '日期', '時間', '寄件人', '收件人', '主旨', '來源', '上傳ID', '建立時間'] },
  links:  { name: '相關對應', head: ['對應ID', '信件訊息ID', '對話訊息ID', '理由', '把握', '建立方式', '狀態', '建立時間'] },
  usage:  { name: '用量',     head: ['時間', '使用者', '動作', '模型', '圖片數', '輸入token', '輸出token', '估計美元'] },
};

/* 先幫 Tina 建好的對話群(來自她 Teams 側欄的工作群組,可在試算表或網頁改名) */
const SEED_GROUPS = [
  'Yunlin AOC & Jenny',
  'Yunlin Site Management Group',
  'Yunlin & Training',
  'YUNLIN_IOC & AOC',
  'Yunlin TSWA Muehlhan',
  'Yunlin RMPL',
  "Yunlin Heli Tech's Group 雲林直升機技師",
  'YUN Heli Plan Alignment',
  'YL Heli Ops *external*',
  'Heli Team Chat_中文頻道',
  'Yunlin Offshore PTW Status (New)',
  'HHO Daily Planning – Winter Campaign 2026/27',
  'Winter & SGRE',
  'Yunlin PO approve with KWIYOMI 倉管群',
  'HHO – AOC Coordination | Winter 2026/27',
  'Yunlin team 2 lead chat',
];

/* ===== 初始化(可重複執行,只補缺的) ===== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const k in SHEETS) {
    const def = SHEETS[k];
    let sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    const cur = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    if (cur.join('|') !== def.head.join('|')) sh.getRange(1, 1, 1, def.head.length).setValues([def.head]);
    sh.setFrozenRows(1);
    // 全部存成純文字,避免 2026-10-20、10:31、001 被自動轉型
    sh.getRange(1, 1, sh.getMaxRows(), def.head.length).setNumberFormat('@');
  }
  const users = sheet('users');
  if (users.getLastRow() < 2) {
    users.getRange(2, 1, 2, 4).setValues([
      [newCode('ADM'), '管理者', 'TRUE', 'setup 自動產生'],
      [newCode('AOC'), 'Wu, Tina', 'TRUE', 'setup 自動產生'],
    ]);
  }
  const groups = sheet('groups');
  if (groups.getLastRow() < 2) {
    const now = stamp();
    const rows = SEED_GROUPS.map((n, i) => ['g' + String(i + 1).padStart(3, '0'), n, '', now, 'TRUE', '0', '', '', 'chat']);
    groups.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  if (!groupRows().some(g => g.kind === 'mail')) groupSave({ group: { name: '信件', note: '從 Outlook 貼上的信件(原信或 Copilot 整理)', kind: 'mail' } });
  attFolder();   // 順便建立雲端硬碟資料夾(第一次會要求雲端硬碟權限)
  Logger.log('完成。代碼請看「使用者」工作表。');
}

/* ===== 進入點 ===== */
function doGet() { return out({ ok: true, service: 'chat-digest' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, code: 'bad_request' }); }
  try {
    const user = findUser(req.code);
    if (!user) return out({ ok: false, code: 'bad_code' });
    switch (req.action) {
      case 'ping':         touchUser(user); return out({ ok: true, name: user.name });
      case 'bootstrap':    return out(bootstrap());
      case 'groupSave':    return out(groupSave(req));
      case 'transcribe':   return out(transcribe(user, req));
      case 'saveRun':      return out(saveRun(user, req));
      case 'summarizeDay': return out(summarizeDay(user, req));
      case 'vocabDay':     return out(vocabDay(user, req));
      case 'groupData':    return out(groupData(req));
      case 'dateData':     return out(dateData(req));
      case 'runData':      return out(runData(req));
      case 'search':       return out(search(req));
      case 'deleteMsg':    return out(deleteMsg(req));
      case 'readAttachment':   return out(readAttachment(user, req));
      case 'matchAttachments': return out(matchAttachments(user, req));
      case 'attGet':       return out(attGet(req));
      case 'eventsDay':    return out(eventsDay(user, req));
      case 'periodReport': return out(periodReport(user, req));
      case 'reportGet':    return out(reportGet(req));
      case 'vocabMark':    return out(vocabMark(user, req));
      case 'vocabPractice': return out(vocabPractice(user, req));
      case 'saveMails':    return out(saveMails(user, req));
      case 'linkTargets':  return out(linkTargets(req));
      case 'linkMails':    return out(linkMails(user, req));
      case 'linkRemove':   return out(linkRemove(req));
      case 'ask':          return out(ask(user, req));
    }
    return out({ ok: false, code: 'bad_action' });
  } catch (err) {
    if (err && err.code) return out({ ok: false, code: err.code, message: String(err.message || '') });
    return out({ ok: false, code: 'server_error', message: String((err && err.message) || err) });
  }
}

function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ===== 試算表小工具 ===== */
function sheet(k) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS[k].name);
  if (!sh) throw { code: 'not_setup', message: '請先在 Apps Script 執行一次 setup()' };
  return sh;
}
function readAll(k) {
  const sh = sheet(k), n = sh.getLastRow() - 1, w = SHEETS[k].head.length;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, w).getValues().map(r => r.map(v => (v === null || v === undefined) ? '' : String(v)));
}
function readOpt(k) { try { return readAll(k); } catch (e) { return []; } }   // 新工作表還沒 setup 時當成空的
function appendRows(k, rows) {
  if (!rows.length) return;
  const sh = sheet(k), w = SHEETS[k].head.length;
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, w).setNumberFormat('@').setValues(rows.map(r => {
    const o = r.slice(0, w).map(v => v === null || v === undefined ? '' : String(v));
    while (o.length < w) o.push('');
    return o;
  }));
}
function stamp() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function newCode(prefix) {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += abc[Math.floor(Math.random() * abc.length)];
  return prefix + '-' + s;
}
function md5hex(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function norm(s) { return String(s || '').toLowerCase().replace(/[\s\p{P}\p{S}]/gu, ''); }
function clip(s, n) { s = String(s === null || s === undefined ? '' : s); return s.length > n ? s.slice(0, n) : s; }
function validDate(s) { s = String(s || '').trim(); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; }
function validTime(s) {
  s = String(s || '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  return m ? String(+m[1]).padStart(2, '0') + ':' + m[2] : '';
}
function isTrue(v) { return String(v).toUpperCase() === 'TRUE'; }

/* ===== 使用者 ===== */
function findUser(code) {
  code = String(code || '').trim();
  if (!code) return null;
  const rows = readAll('users');
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === code && isTrue(rows[i][2])) return { code, name: rows[i][1] || code, row: i + 2 };
  }
  return null;
}
function touchUser(u) { try { sheet('users').getRange(u.row, 5).setValue(stamp()); } catch (e) {} }

/* ===== 群組 ===== */
function groupRows() {
  return readAll('groups').map((r, i) => ({ id: r[0], name: r[1], note: r[2], created: r[3], enabled: isTrue(r[4]),
    count: +r[5] || 0, first: r[6], last: r[7], kind: r[8] === 'mail' ? 'mail' : 'chat', row: i + 2 }));
}
function findGroup(id) { return groupRows().find(g => g.id === id) || null; }

function groupSave(req) {
  const g = req.group || {};
  const name = clip(String(g.name || '').trim(), 120);
  if (!name) throw { code: 'bad_request', message: '群組名稱不能空白' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const all = groupRows();
    if (g.id) {
      const hit = all.find(x => x.id === g.id);
      if (!hit) throw { code: 'bad_group' };
      sheet('groups').getRange(hit.row, 2, 1, 4).setValues([[name, clip(g.note || '', 500), hit.created, g.enabled === false ? 'FALSE' : 'TRUE']]);
      return { ok: true, group: Object.assign({}, hit, { name, note: g.note || '', enabled: g.enabled !== false }) };
    }
    const dup = all.find(x => x.name.toLowerCase() === name.toLowerCase());
    if (dup) return { ok: true, group: dup, existed: true };
    let max = 0;
    all.forEach(x => { const n = parseInt(String(x.id).replace(/\D/g, ''), 10); if (n > max) max = n; });
    const id = 'g' + String(max + 1).padStart(3, '0');
    const kind = g.kind === 'mail' ? 'mail' : 'chat';
    appendRows('groups', [[id, name, clip(g.note || '', 500), stamp(), 'TRUE', '0', '', '', kind]]);
    return { ok: true, group: { id, name, note: g.note || '', enabled: true, count: 0, first: '', last: '', kind } };
  } finally { lock.releaseLock(); }
}

/* ===== Claude 呼叫 ===== */
function callClaude(user, action, model, content, maxTokens, system, nImages) {
  const payload = { model, max_tokens: maxTokens, messages: [{ role: 'user', content }] };
  if (system) payload.system = system;
  const body = postAnthropic(payload);
  const text = (body.content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  logUsage(user, action, model, nImages || 0, body.usage || {});
  return { text, truncated: body.stop_reason === 'max_tokens' };
}
function postAnthropic(payload) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw { code: 'no_key', message: '指令碼屬性沒有 ANTHROPIC_API_KEY' };
  let status = 0, body = {};
  for (let attempt = 0; attempt < 2; attempt++) {
    let res;
    try {
      res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json',
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify(payload), muteHttpExceptions: true,
      });
    } catch (err) {
      const m = String((err && err.message) || err);
      if (/timeout|timed out|逾時|超時/i.test(m)) throw { code: 'timeout', message: m };
      throw { code: 'network', message: m };
    }
    status = res.getResponseCode();
    body = {}; try { body = JSON.parse(res.getContentText()); } catch (e) {}
    if ((status === 429 || status === 529 || status === 503) && attempt === 0) { Utilities.sleep(4000); continue; }
    break;
  }
  if (status !== 200) {
    const msg = (body.error && body.error.message) || ('HTTP ' + status);
    if (/credit balance/i.test(msg)) throw { code: 'no_credit', message: msg };
    if (status === 429 || status === 529) throw { code: 'rate_limited', message: msg };
    throw { code: 'api_error', message: msg };
  }
  return body;
}

function logUsage(user, action, model, n, u) {
  try {
    const p = PRICE[model] || [0, 0];
    const inTok = (u.input_tokens || 0) + (u.cache_creation_input_tokens || 0) + (u.cache_read_input_tokens || 0);
    const outTok = u.output_tokens || 0;
    const usd = (inTok * p[0] + outTok * p[1]) / 1e6;
    appendRows('usage', [[stamp(), user.name, action, model, n, inTok, outTok, usd.toFixed(4)]]);
  } catch (e) {}
}

/* 寬鬆解析:去掉 ``` 圍欄;被截斷時退回最後一個完整物件再補上括號 */
function parseLoose(text) {
  let s = String(text || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  const a = s.indexOf('{'); if (a < 0) return null;
  s = s.slice(a);
  try { return JSON.parse(s); } catch (e) {}
  const b = s.lastIndexOf('}');
  if (b > 0) { try { return JSON.parse(s.slice(0, b + 1)); } catch (e) {} }
  // 修補:由後往前找 '}',切掉後面再補閉合
  let tries = 0;
  for (let i = s.length - 1; i > 0 && tries < 80; i--) {
    if (s[i] !== '}') continue;
    tries++;
    const cut = s.slice(0, i + 1);
    const close = closers(cut);
    if (close === null) continue;
    try { return JSON.parse(cut + close); } catch (e) {}
  }
  return null;
}
function closers(s) {
  const st = []; let inStr = false, esc = false;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (inStr) { if (esc) esc = false; else if (c === '\\') esc = true; else if (c === '"') inStr = false; continue; }
    if (c === '"') inStr = true;
    else if (c === '{' || c === '[') st.push(c);
    else if (c === '}' || c === ']') st.pop();
  }
  if (inStr) return null;
  return st.reverse().map(c => c === '{' ? '}' : ']').join('');
}

/* ===== 1. 轉錄 ===== */
const TRANSCRIBE_SYSTEM = [
  '你是逐字稿轉錄員。輸入是 Microsoft Teams 聊天視窗的截圖片段(長截圖切成的一段)。',
  '你的工作是「一字不漏地照抄」畫面上的每一則訊息,轉成 JSON。不要翻譯、不要摘要、不要修正錯字、不要補寫看不到的內容。',
  '',
  '規則:',
  '1. 發言區塊的標頭是「姓名 星期/日期 時間」(例如「Lee, Ealeen  Tuesday 8:20 PM」)。標頭下方連續的每個氣泡都是同一人、同一時間的「各自一則訊息」,每個氣泡輸出一筆。',
  '2. 片段最上方若是沒有標頭的氣泡(標頭在上一段),sender 與 time 填 null,其他照抄。',
  '3. 靠右對齊、沒有姓名的氣泡是截圖者本人發的,sender 填使用者提供的「截圖者姓名」。',
  '4. 置中的日期分隔線(Tuesday、Yesterday、Today、October 14 等)代表其後訊息的日期。依使用者給的對照表換算成 YYYY-MM-DD。看不出日期就填 null,不要猜。',
  '5. 時間換算成 24 小時制 HH:MM(8:20 PM → 20:20)。',
  '6. 有框線、內含「姓名 + 完整日期時間 + 文字」的是「引用」,放進該訊息的 quote,不要另外當成一則訊息。',
  '7. 訊息裡的圖片、截圖、表格、檔案卡,在 attachments 用繁體中文簡述,並抄出看得清楚的關鍵文字(人名、日期、數字、標題),例如「[截圖] Muehlhan 最新名單:Team 1 Grzegorz Selwanowicz…」「[檔案] TW Legal MEWP.pdf」。看不清楚就寫「[截圖] 內容不清楚」。',
  '8. 忽略介面文字:Translate、Edited、Never translate…、表情反應與數字、Last read、輸入框、「… are out of office」提示、頭像。',
  '9. 訊息被片段上緣或下緣切到時,照抄看得到的部分,edge 填 "top" 或 "bottom";完整的填 null。',
  '10. 保留原文換行(用 \\n)、表情符號、@提及的名字、英文大小寫。中英夾雜照原樣。',
  '11. lang:純中文 "zh"、純英文 "en"、中英夾雜 "mixed"。',
  '12. 如果畫面上方看得到聊天室名稱(對話標題),填在 chat_title;沒有就 null。',
  '',
  '只輸出一個 JSON 物件,不要任何說明文字,格式:',
  '{"chat_title":null,"messages":[{"sender":"Lee, Ealeen","when":"Tuesday 8:20 PM","date":"2026-10-13","time":"20:20","text":"…","quote":null,"attachments":[],"lang":"zh","edge":null}]}',
  'quote 的格式:{"sender":"Chiang, Jenny","when":"10/15/2026 10:38 AM","text":"…"}',
].join('\n');

function transcribe(user, req) {
  const images = (Array.isArray(req.images) ? req.images : []).slice(0, MAX_IMAGES);
  if (!images.length) throw { code: 'bad_request', message: '沒有圖片' };
  const content = images.map(b64 => ({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: String(b64) } }));
  const ctx = [
    '截圖者姓名(靠右氣泡的發言人):' + (clip(req.owner, 80) || 'Wu, Tina'),
    '截圖日期:' + (validDate(req.captureDate) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')),
    '日期對照表(Teams 顯示 → 實際日期):',
    clip(req.dateMap || '', 2000),
    '只寫月/日沒寫年份的,用截圖日期的年份;若因此晚於截圖日期,改用前一年。',
    images.length > 1 ? '共 ' + images.length + ' 張,依序是同一段對話由上而下的連續片段。' : '',
    '請轉錄這個片段。',
  ].filter(Boolean).join('\n');
  content.push({ type: 'text', text: ctx });
  const r = callClaude(user, 'transcribe', MODELS.transcribe, content, MAX_TOKENS.transcribe, TRANSCRIBE_SYSTEM, images.length);
  const j = parseLoose(r.text);
  if (!j || !Array.isArray(j.messages)) throw { code: 'bad_ai', message: 'AI 回傳的格式解析失敗' };
  const msgs = j.messages.map(m => ({
    sender: m.sender ? clip(m.sender, 120) : null,
    when: m.when ? clip(m.when, 60) : null,
    date: validDate(m.date) || null,
    time: validTime(m.time) || null,
    text: clip(m.text || '', 20000),
    quote: m.quote && (m.quote.text || m.quote.sender) ? { sender: clip(m.quote.sender || '', 120), when: clip(m.quote.when || '', 60), text: clip(m.quote.text || '', 4000) } : null,
    attachments: Array.isArray(m.attachments) ? m.attachments.map(a => clip(a, 2000)).filter(Boolean) : [],
    lang: ['zh', 'en', 'mixed'].indexOf(m.lang) >= 0 ? m.lang : 'zh',
    edge: m.edge === 'top' || m.edge === 'bottom' ? m.edge : null,
  }));
  return { ok: true, chat_title: j.chat_title ? clip(j.chat_title, 200) : null, messages: msgs, truncated: r.truncated };
}

/* ===== 2. 存檔(與既有訊息去重合併) ===== */
function msgId(gid, o) {
  return 'm' + md5hex([gid, o.sender, o.date, o.time, norm(o.text + (o.attach || '')).slice(0, 60)].join('|')).slice(0, 14);
}
function slotKey(o) { return [o.sender, o.date, o.time].join('|'); }

function saveRun(user, req) {
  const g = findGroup(req.groupId);
  if (!g) throw { code: 'bad_group' };
  const incoming = Array.isArray(req.messages) ? req.messages : [];
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const runs = readAll('runs');
    let maxRun = 0;
    runs.forEach(r => { const n = parseInt(String(r[0]).replace(/\D/g, ''), 10); if (n > maxRun) maxRun = n; });
    const runId = 'r' + String(maxRun + 1).padStart(4, '0');
    const now = stamp();

    const data = readAll('msgs');
    const byId = {}, bySlot = {};
    const index = o => {
      byId[o.id] = o;
      const k = slotKey(o); (bySlot[k] = bySlot[k] || []).push(o);
    };
    data.forEach((r, i) => {
      if (r[1] !== g.id) return;
      index({ id: r[0], date: r[2], time: r[3], sender: r[4], text: r[5], attach: r[7], runs: r[10], arr: r });
    });

    const appends = [];
    let newCount = 0, changed = false;
    const dates = {};
    const resIds = [];
    for (let idx = 0; idx < incoming.length; idx++) {
      const m = incoming[idx] || {};
      const quote = m.quote && (m.quote.text || m.quote.sender)
        ? clip((m.quote.sender || '') + (m.quote.when ? ' ' + m.quote.when : '') + ':' + (m.quote.text || ''), 4000) : '';
      const attach = Array.isArray(m.attachments) ? clip(m.attachments.join('\n'), 6000) : clip(m.attach || '', 6000);
      const o = {
        dupSeq: 0,
        sender: clip(String(m.sender || '').trim(), 120) || '(未知)',
        date: validDate(m.date), time: validTime(m.time),
        text: clip(m.text || '', 20000), quote, attach,
        lang: ['zh', 'en', 'mixed'].indexOf(m.lang) >= 0 ? m.lang : 'zh',
      };
      if (!o.text && !o.attach) continue;
      if (o.date) dates[o.date] = true;
      const id = msgId(g.id, o);
      let hit = byId[id];
      if (!hit) {
        const n = norm(o.text);
        hit = (bySlot[slotKey(o)] || []).find(c => {
          const cn = norm(c.text);
          return n.length >= 4 && cn.length >= 4 && (cn.indexOf(n) >= 0 || n.indexOf(cn) >= 0);
        });
      }
      if (hit && hit.fresh && norm(o.text) === norm(hit.text)) {
        // 同一次上傳裡兩則一模一樣的短訊息(例如連續兩個「好」)= 兩則不同的訊息
        hit = null;
        o.dupSeq = (o.dupSeq || 0) + 1;
      }
      if (hit) {
        const rs = String(hit.runs || '').split(',').filter(Boolean);
        if (rs.indexOf(runId) < 0) { rs.push(runId); hit.runs = rs.join(','); hit.arr[10] = hit.runs; changed = true; }
        if (norm(o.text).length > norm(hit.text).length) { hit.text = o.text; hit.arr[5] = o.text; changed = true; }
        if (!hit.arr[7] && o.attach) { hit.arr[7] = o.attach; changed = true; }
        if (!hit.arr[6] && o.quote) { hit.arr[6] = o.quote; changed = true; }
        resIds[idx] = hit.id;
        continue;
      }
      let nid = id;
      while (byId[nid]) nid = id + '-' + (++o.dupSeq || 1);
      const arr = [nid, g.id, o.date, o.time, o.sender, o.text, o.quote, o.attach, o.lang, runId, runId, now];
      appends.push(arr);
      index({ id: nid, date: o.date, time: o.time, sender: o.sender, text: o.text, attach: o.attach, runs: runId, arr, fresh: true });
      resIds[idx] = nid;
      newCount++;
    }
    if (changed && data.length) {
      const sh = sheet('msgs');
      sh.getRange(2, 6, data.length, 3).setValues(data.map(r => [r[5], r[6], r[7]]));
      sh.getRange(2, 11, data.length, 1).setValues(data.map(r => [r[10]]));
    }
    appendRows('msgs', appends);

    // 附件:對到這次存的訊息(或只掛在某一天)
    const attIn = Array.isArray(req.attachments) ? req.attachments : [];
    let attAdded = 0;
    if (attIn.length) {
      const old = readAll('atts');
      let maxA = 0;
      old.forEach(r => { const n = parseInt(String(r[0]).replace(/\D/g, ''), 10); if (n > maxA) maxA = n; });
      const seen = {};
      old.forEach(r => { if (r[1] === g.id) seen[r[6]] = true; });
      const rows = [];
      attIn.forEach(a => {
        if (!a || !a.fileId || seen[a.fileId]) return;
        const mi = (a.msg === null || a.msg === undefined || a.msg === '') ? -1 : +a.msg;
        const mid = mi >= 0 ? (resIds[mi] || '') : '';
        const date = validDate(mi >= 0 && incoming[mi] ? incoming[mi].date : a.date) || validDate(a.date);
        if (date) dates[date] = true;
        seen[a.fileId] = true;
        rows.push(['a' + String(++maxA).padStart(5, '0'), g.id, mid, date, clip(a.name || '', 200), clip(a.mime || '', 80),
          clip(a.fileId, 100), clip(a.url || '', 300), clip(a.title || '', 300), clip(a.content || '', 45000), runId, now]);
      });
      appendRows('atts', rows);
      attAdded = rows.length;
    }

    // 群組統計
    let cnt = 0, first = '', last = '';
    data.concat(appends).forEach(r => {
      if (r[1] !== g.id) return;
      cnt++;
      if (r[2]) { if (!first || r[2] < first) first = r[2]; if (!last || r[2] > last) last = r[2]; }
    });
    sheet('groups').getRange(g.row, 6, 1, 3).setValues([[String(cnt), first, last]]);

    const ds = Object.keys(dates).sort();
    appendRows('runs', [[runId, now, user.name, g.id, clip(req.fileName || '', 300), validDate(req.captureDate),
      String(req.segments || ''), String(incoming.length), String(newCount), ds[0] || '', ds[ds.length - 1] || '']]);
    return { ok: true, runId, newCount, total: incoming.length, dates: ds, attAdded };
  } finally { lock.releaseLock(); }
}

/* 取某群組某天的訊息(依時間排序,同時間照寫入順序) */
function dayMessages(gid, date) {
  const list = [];
  readAll('msgs').forEach((r, i) => { if (r[1] === gid && r[2] === date) list.push(rowToMsg(r, i)); });
  list.sort((a, b) => (a.time && b.time && a.time !== b.time) ? (a.time < b.time ? -1 : 1) : a.i - b.i);
  return list;
}
function rowToMsg(r, i) {
  return { id: r[0], group: r[1], date: r[2], time: r[3], sender: r[4], text: r[5], quote: r[6], attach: r[7], lang: r[8], firstRun: r[9], runs: r[10], i };
}
function dayAtts(gid, date) {
  return readAll('atts').filter(r => r[1] === gid && r[3] === date).map(attOut);
}
function attOut(r) {
  return { id: r[0], group: r[1], msg: r[2], date: r[3], name: r[4], mime: r[5], fileId: r[6], url: r[7], title: r[8], content: r[9], run: r[10], created: r[11] };
}
function linesFor(list, atts) {
  atts = atts || [];
  const ids = {}; list.forEach(m => ids[m.id] = true);
  let out = list.map((m, k) => {
    let s = '[m' + (k + 1) + '] ' + (m.time || '--:--') + ' ' + m.sender + ':' + m.text;
    if (m.quote) s += '\n    (引用 ' + m.quote + ')';
    if (m.attach) s += '\n    (附件 ' + m.attach.replace(/\n/g, ' / ') + ')';
    atts.filter(a => a.msg === m.id).forEach(a => { s += '\n    (附件原檔「' + a.title + '」內容:\n' + clip(a.content, 4000) + '\n    )'; });
    return s;
  }).join('\n');
  const loose = atts.filter(a => !ids[a.msg]);
  if (loose.length) out += '\n\n當天另外附上的檔案:\n' + loose.map(a => '[附件]「' + a.title + '」\n' + clip(a.content, 4000)).join('\n');
  return out;
}
function mapRefs(refs, list) {
  return (Array.isArray(refs) ? refs : [refs]).map(x => {
    const n = parseInt(String(x || '').replace(/\D/g, ''), 10);
    return n >= 1 && n <= list.length ? list[n - 1].id : null;
  }).filter(Boolean);
}

/* ===== 3. 每日摘要(英文對話給中英對照) ===== */
const SUMMARY_SYSTEM = [
  '你是風電運維現場協調(AOC)的助理,要把一天的 Teams 對話整理成摘要給 AOC 同仁看。',
  '只根據對話內容,不要臆測;人名、日期、數字照原文。',
  '使用繁體中文(台灣用語)。英文專有名詞、公司名、人名保留英文。',
  '若當天對話有相當比例是英文(或主要內容由英文訊息傳達),bilingual 設為 true,每一條都要同時給 zh 與 en(en 為自然的英文);否則 bilingual=false、en 留空字串。',
  '每一條都要附 refs,列出依據的訊息編號(例如 ["m3","m7"])。',
  '精簡:points 最多 6 條、todos 最多 8 條、decisions 最多 5 條、questions 最多 5 條;每條中文 60 字以內、英文 40 字以內,同一件事只寫一次(寫在 points 的不要在 decisions 重複)。',
  'refs 每條最多 3 個,挑最關鍵的訊息。',
  'todos 是有人要去做的事:who 填負責人(對話中看得出來才填,否則空字串),due 填期限(YYYY-MM-DD 或原文說法,沒有就空字串)。',
  'decisions 是已經拍板的結論;questions 是還沒結論、需要追問的事。',
  '只輸出 JSON:',
  '{"bilingual":false,"title":{"zh":"","en":""},"points":[{"zh":"","en":"","refs":[]}],"todos":[{"who":"","zh":"","en":"","due":"","refs":[]}],"decisions":[{"zh":"","en":"","refs":[]}],"questions":[{"zh":"","en":"","refs":[]}]}',
].join('\n');

function summarizeDay(user, req) {
  const g = findGroup(req.groupId); if (!g) throw { code: 'bad_group' };
  const date = validDate(req.date); if (!date) throw { code: 'bad_request', message: '日期格式錯誤' };
  const list = dayMessages(g.id, date), atts = dayAtts(g.id, date);
  if (!list.length && !atts.length) throw { code: 'no_messages', message: '這一天沒有訊息' };
  const prompt = '對話群:' + g.name + mailNote(g) + '\n日期:' + date + '\n共 ' + list.length + ' 則訊息:\n\n' + clip(linesFor(list, atts), 120000);
  const r = callClaude(user, 'summary', MODELS.summary, [{ type: 'text', text: prompt }], MAX_TOKENS.summary, SUMMARY_SYSTEM, 0);
  const j = parseLoose(r.text);
  if (!j) throw { code: 'bad_ai', message: 'AI 回傳的摘要解析失敗' };
  const fix = (arr, extra) => (Array.isArray(arr) ? arr : []).map(x => {
    const o = { zh: clip(x.zh || '', 1000), en: clip(x.en || '', 1000), refs: mapRefs(x.refs, list) };
    if (extra) { o.who = clip(x.who || '', 120); o.due = clip(x.due || '', 60); }
    return o;
  }).filter(o => o.zh || o.en);
  const sum = {
    bilingual: !!j.bilingual,
    title: { zh: clip((j.title && j.title.zh) || '', 200), en: clip((j.title && j.title.en) || '', 200) },
    points: fix(j.points), todos: fix(j.todos, true), decisions: fix(j.decisions), questions: fix(j.questions),
    msgCount: list.length, truncated: r.truncated,
  };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet('sums');
    const rows = readAll('sums');
    const json = clip(JSON.stringify(sum), 49000);
    const row = [g.id, date, stamp(), MODELS.summary, sum.title.zh, sum.points.map(p => '・' + p.zh).join('\n'), json];
    const idx = rows.findIndex(x => x[0] === g.id && x[1] === date);
    if (idx >= 0) sh.getRange(idx + 2, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    else appendRows('sums', [row]);
  } finally { lock.releaseLock(); }
  return { ok: true, summary: sum, updated: stamp() };
}

/* ===== 4. 英文單字(給 Tina 練習) ===== */
const VOCAB_SYSTEM = [
  '你是職場英文老師,學生是在台灣風電運維公司做現場協調(AOC)的同仁。',
  '從一天的 Teams 對話中,挑出值得學的英文單字或片語(包含中文訊息裡夾雜的英文,例如 onboarding、charge back、induction、rotation),最多 10 個。',
  '優先:工作上常用、對話裡真的出現、學生可能不熟的。不要挑人名、公司名、地名、太簡單的字(ok、yes、team)。',
  '每個給:term(對話中的原形或片語)、pos(詞性,英文縮寫如 n./v./phr.)、zh(簡短中文意思)、explain(繁體中文 50 字以內,說明在這個情境的用法)、example(對話中出現的原句,照抄;句子太長就只抄含這個字的那一句)、example_zh(原句的中文翻譯)、ref(出處訊息編號,如 "m12")。',
  '只輸出 JSON:{"vocab":[{"term":"","pos":"","zh":"","explain":"","example":"","example_zh":"","ref":""}]}',
  '如果當天完全沒有英文,輸出 {"vocab":[]}。',
].join('\n');

function vocabDay(user, req) {
  const g = findGroup(req.groupId); if (!g) throw { code: 'bad_group' };
  const date = validDate(req.date); if (!date) throw { code: 'bad_request', message: '日期格式錯誤' };
  const list = dayMessages(g.id, date), atts = dayAtts(g.id, date);
  let items = [];
  if (list.some(m => /[A-Za-z]{3,}/.test(m.text + m.quote)) || atts.some(a => /[A-Za-z]{3,}/.test(a.content))) {
    const skip = readAll('vmarks').filter(r => r[1] === 'skip').map(r => r[0]).slice(-200);
    const prompt = '對話群:' + g.name + mailNote(g) + '\n日期:' + date +
      (skip.length ? '\n學生已經標示不需要學、不要再挑的字:' + skip.join('、') : '') + '\n\n' + clip(linesFor(list, atts), 120000);
    const r = callClaude(user, 'vocab', MODELS.vocab, [{ type: 'text', text: prompt }], MAX_TOKENS.vocab, VOCAB_SYSTEM, 0);
    const j = parseLoose(r.text);
    if (!j) throw { code: 'bad_ai', message: 'AI 回傳的單字解析失敗' };
    items = (Array.isArray(j.vocab) ? j.vocab : []).filter(v => v && v.term).slice(0, 15).map(v => {
      const ids = mapRefs(v.ref, list);
      const m = ids.length ? list.find(x => x.id === ids[0]) : null;
      return {
        term: clip(String(v.term).trim(), 120), pos: clip(v.pos || '', 20), zh: clip(v.zh || '', 200),
        explain: clip(v.explain || '', 600), example: clip(v.example || '', 600), example_zh: clip(v.example_zh || '', 600),
        msg: ids[0] || '', run: m ? m.firstRun : '',
      };
    });
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet('vocab');
    const rows = readAll('vocab');
    for (let i = rows.length - 1; i >= 0; i--) if (rows[i][4] === g.id && rows[i][5] === date) sh.deleteRow(i + 2);
    const now = stamp();
    appendRows('vocab', items.map(v => [v.term, (v.pos ? v.pos + ' ' : '') + v.zh, v.explain,
      v.example + (v.example_zh ? '\n' + v.example_zh : ''), g.id, date, v.msg, v.run, now]));
  } finally { lock.releaseLock(); }
  return { ok: true, vocab: items.map(v => vocabOut([v.term, (v.pos ? v.pos + ' ' : '') + v.zh, v.explain,
    v.example + (v.example_zh ? '\n' + v.example_zh : ''), g.id, date, v.msg, v.run, ''])) };
}
function vocabOut(r) {
  const ex = String(r[3] || '').split('\n');
  return { term: r[0], zh: r[1], explain: r[2], example: ex[0] || '', example_zh: ex.slice(1).join('\n'), group: r[4], date: r[5], msg: r[6], run: r[7], created: r[8] };
}

/* ===== 5. 讀取 ===== */
function bootstrap() {
  const groups = groupRows().map(g => ({ id: g.id, name: g.name, note: g.note, enabled: g.enabled, count: g.count, first: g.first, last: g.last, kind: g.kind }));
  const runs = readAll('runs').slice(-400).reverse().map(r => ({ id: r[0], time: r[1], user: r[2], group: r[3], file: r[4],
    capture: r[5], segments: r[6], total: r[7], added: r[8], from: r[9], to: r[10] }));
  const sums = readAll('sums').map(r => ({ group: r[0], date: r[1], updated: r[2], title: r[4] }));
  const vocab = readAll('vocab').map(vocabOut);
  const evDays = {};
  readAll('events').forEach(r => { evDays[r[1] + '|' + r[2]] = (evDays[r[1] + '|' + r[2]] || 0) + 1; });
  const msgDays = {};
  readAll('msgs').forEach(r => { if (r[2]) msgDays[r[1] + '|' + r[2]] = (msgDays[r[1] + '|' + r[2]] || 0) + 1; });
  const reports = readAll('reports').reverse().map(r => ({ id: r[0], created: r[1], user: r[2], from: r[3], to: r[4], groups: r[5], title: r[6] }));
  const vmarks = readAll('vmarks').map(vmarkOut);
  return { ok: true, groups, runs, sums, vocab, evDays, msgDays, reports, vmarks };
}
function groupData(req) {
  const gid = String(req.groupId || '');
  const msgs = [], rows = readAll('msgs');
  rows.forEach((r, i) => { if (r[1] === gid) msgs.push(rowToMsg(r, i)); });
  const sums = readAll('sums').filter(r => r[0] === gid).map(sumOut);
  const atts = readAll('atts').filter(r => r[1] === gid).map(attOut);
  return Object.assign({ ok: true, messages: msgs, summaries: sums, atts }, linkInfo(msgs, rows));
}
function dateData(req) {
  const d = validDate(req.date);
  const msgs = [], rows = readAll('msgs');
  rows.forEach((r, i) => { if (r[2] === d) msgs.push(rowToMsg(r, i)); });
  const sums = readAll('sums').filter(r => r[1] === d).map(sumOut);
  const atts = readAll('atts').filter(r => r[3] === d).map(attOut);
  return Object.assign({ ok: true, messages: msgs, summaries: sums, atts }, linkInfo(msgs, rows));
}
function sumOut(r) {
  let s = null; try { s = JSON.parse(r[6]); } catch (e) {}
  return { group: r[0], date: r[1], updated: r[2], model: r[3], summary: s };
}
function runData(req) {
  const id = String(req.runId || '');
  const msgs = [], rows = readAll('msgs');
  rows.forEach((r, i) => { if (String(r[10]).split(',').indexOf(id) >= 0) msgs.push(rowToMsg(r, i)); });
  const atts = readAll('atts').filter(r => r[10] === id).map(attOut);
  return Object.assign({ ok: true, messages: msgs, atts }, linkInfo(msgs, rows));
}
function search(req) {
  const q = String(req.q || '').trim().toLowerCase();
  if (!q) return { ok: true, messages: [] };
  const terms = q.split(/\s+/).filter(Boolean);
  const out = [], rows = readAll('msgs');
  rows.forEach((r, i) => {
    if (out.length >= 300) return;
    const hay = (r[4] + ' ' + r[5] + ' ' + r[6] + ' ' + r[7]).toLowerCase();
    if (terms.every(t => hay.indexOf(t) >= 0)) out.push(rowToMsg(r, i));
  });
  const atts = readAll('atts').filter(r => terms.every(t => (r[4] + ' ' + r[8] + ' ' + r[9]).toLowerCase().indexOf(t) >= 0)).slice(0, 100).map(attOut);
  return Object.assign({ ok: true, messages: out, atts }, linkInfo(out, rows));
}
function deleteMsg(req) {
  const id = String(req.id || '');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const rows = readAll('msgs');
    const i = rows.findIndex(r => r[0] === id);
    if (i < 0) return { ok: true, deleted: 0 };
    sheet('msgs').deleteRow(i + 2);
    const mr = readOpt('mails'); for (let k = mr.length - 1; k >= 0; k--) if (mr[k][0] === id) sheet('mails').deleteRow(k + 2);
    const lr = readOpt('links'); for (let k = lr.length - 1; k >= 0; k--) if (lr[k][1] === id || lr[k][2] === id) sheet('links').deleteRow(k + 2);
    const gid = rows[i][1];
    const g = findGroup(gid);
    if (g) sheet('groups').getRange(g.row, 6).setValue(String(Math.max(0, g.count - 1)));
    return { ok: true, deleted: 1 };
  } finally { lock.releaseLock(); }
}

/* ===== 6. 附件(圖片、PDF) ===== */
function attFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ATT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(ATT_FOLDER);
  props.setProperty('ATT_FOLDER_ID', f.getId());
  return f;
}

const ATTACH_SYSTEM = [
  '你是資料擷取員。輸入是 Microsoft Teams 對話裡分享的附件(截圖、照片或 PDF),來自台灣風電運維現場協調(AOC)的工作對話。',
  '把附件裡的資訊完整擷取成文字,讓之後不看原圖也知道內容:',
  '- 表格:轉成 Markdown 表格,保留所有欄位與列,人名、日期、數字、代碼一個都不能漏。',
  '- 名單、清單:逐行列出。',
  '- 信件、訊息截圖:寄件者、收件者、日期、主旨,以及內文(重要段落照抄)。',
  '- 文件(PDF、證書、報告):標題、對象、日期、有效期限、重點內容。',
  '- 照片:描述看到的東西與可辨識的文字。',
  '照原文語言抄寫,不要翻譯;看不清楚的地方寫「(看不清楚)」,不要猜。',
  'title 用繁體中文一句話描述這個附件(例如「Muehlhan 葉片組 10/22 最新名單」)。',
  '只輸出 JSON:{"title":"","content":""}(content 內可以用 \\n 與 Markdown)',
].join('\n');

function readAttachment(user, req) {
  const name = clip(req.name || '附件', 200), mime = String(req.mime || '');
  const isPdf = mime === 'application/pdf';
  if (!isPdf && !/^image\//.test(mime)) throw { code: 'bad_request', message: '只支援圖片和 PDF' };
  if (!req.original) throw { code: 'bad_request', message: '沒有檔案內容' };
  let content;
  if (isPdf) content = [{ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: String(req.original) } }];
  else {
    const imgs = (Array.isArray(req.images) ? req.images : []).slice(0, 4);
    if (!imgs.length) throw { code: 'bad_request', message: '沒有圖片' };
    content = imgs.map(b => ({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: String(b) } }));
  }
  const n = isPdf ? 0 : content.length;
  content.push({ type: 'text', text: '檔名:' + name + (n > 1 ? '\n這張圖很長,已切成 ' + n + ' 段,依序由上而下,相鄰段落有重疊。' : '') + '\n請擷取內容。' });
  const r = callClaude(user, 'attach', MODELS.attach, content, MAX_TOKENS.attach, ATTACH_SYSTEM, n);
  const j = parseLoose(r.text) || {};
  // AI 成功才存原檔,避免失敗重試時重複存
  const file = attFolder().createFile(Utilities.newBlob(Utilities.base64Decode(String(req.original)), mime,
    Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss') + '_' + name));
  return { ok: true, fileId: file.getId(), url: file.getUrl(), title: clip(j.title || name, 300),
    content: clip(j.content || r.text || '', 45000), truncated: r.truncated };
}

const MATCH_SYSTEM = [
  '你要把附件配對到它在 Teams 對話裡出現的那一則訊息。',
  '訊息裡的「附件描述」是從對話縮圖看到的內容(可能不完整);附件內容是原檔擷取的完整內容。',
  '依內容(人名、日期、標題、表格欄位)、發言脈絡判斷。不確定就填 null,不要硬配。',
  '一則訊息可以有多個附件。',
  '只輸出 JSON:{"pairs":[{"att":0,"msg":12}]}',
].join('\n');

function matchAttachments(user, req) {
  const cands = (Array.isArray(req.cands) ? req.cands : []).slice(0, 250);
  const atts = (Array.isArray(req.atts) ? req.atts : []).slice(0, 30);
  if (!cands.length || !atts.length) return { ok: true, pairs: [] };
  const lines = ['訊息:'];
  cands.forEach(c => lines.push('#' + c.i + ' ' + (c.date || '') + ' ' + (c.time || '') + ' ' + clip(c.sender, 60) + ':' +
    clip(c.text, 200).replace(/\n/g, ' ') + (c.attach ? '  【附件描述】' + clip(c.attach, 400).replace(/\n/g, ' ') : '')));
  lines.push('', '附件:');
  atts.forEach(a => lines.push('附件' + a.j + ' 檔名 ' + clip(a.name, 100) + '「' + clip(a.title, 200) + '」:' + clip(a.content, 800).replace(/\n/g, ' ')));
  const r = callClaude(user, 'match', MODELS.match, [{ type: 'text', text: lines.join('\n') }], MAX_TOKENS.match, MATCH_SYSTEM, 0);
  const j = parseLoose(r.text) || {};
  const ok = {}; cands.forEach(c => ok[c.i] = true);
  const pairs = (Array.isArray(j.pairs) ? j.pairs : []).map(p => ({
    att: parseInt(p.att, 10), msg: p.msg === null || p.msg === undefined || !ok[parseInt(p.msg, 10)] ? null : parseInt(p.msg, 10),
  })).filter(p => p.att >= 0);
  return { ok: true, pairs };
}

function attGet(req) {
  const row = readAll('atts').find(r => r[0] === String(req.id || ''));
  if (!row) throw { code: 'bad_request', message: '找不到附件' };
  const blob = DriveApp.getFileById(row[6]).getBlob();
  return { ok: true, mime: blob.getContentType() || row[5], name: row[4], data: Utilities.base64Encode(blob.getBytes()) };
}

/* ===== 7. 每日事件(給期間報告用的結構化資料) ===== */
const EVENT_TYPES = ['deadline', 'decision', 'incident', 'personnel', 'asset'];
const EVENTS_SYSTEM = [
  '你是台灣離岸風場運維現場協調(AOC)的資訊整理員。從一天的 Teams 對話裡,抽出之後需要追蹤或回顧的「事件」。',
  '類型 type:',
  '- deadline:有日期要完成或要發生的事(報到、出發、開工、到期、截止、交件、訪客、會議)',
  '- decision:已經拍板的決定',
  '- incident:突發、異常、事故、安全問題、臨時變更、延誤、客訴',
  '- personnel:人員異動(調派、延長、離開、報到、換人、請假、資格證照問題)',
  '- asset:船舶、直升機、風機、設備、工具的狀態變化',
  '每個事件欄位:',
  'type;title(繁體中文 25 字內);detail(繁體中文 60 字內,寫清楚誰、什麼、何時);',
  'event_date(事件發生或截止的日期 YYYY-MM-DD,依對話日期換算「明天」「週四」等;不知道留空);date_text(原文對日期的說法);',
  'who(相關人名陣列,照對話寫法);entities(風機編號、船名、公司、文件等陣列);',
  'topic(主題名稱:優先沿用下面列出的既有主題,同一件事一定要用同一個名字;新主題用簡短名詞片語,例如「Muehlhan TSWA 到場」「ActSafe 檢驗與訓練」);',
  'status(open 尚未完成/done 已完成/cancelled 取消/info 僅告知);severity(incident 才填 high/medium/low,其他留空);refs(依據的訊息編號,最多 3 個)。',
  '最多 15 個,同一件事只列一次,閒聊與寒暄不列。只輸出 JSON:{"events":[{"type":"","title":"","detail":"","event_date":"","date_text":"","who":[],"entities":[],"topic":"","status":"","severity":"","refs":[]}]}',
].join('\n');

function knownTopics(limit) {
  const seen = {}, out = [];
  const rows = readAll('events');
  for (let i = rows.length - 1; i >= 0 && out.length < (limit || 80); i--) {
    const t = rows[i][10]; if (t && !seen[t]) { seen[t] = true; out.push(t); }
  }
  return out;
}
function eventOut(r) {
  return { id: r[0], group: r[1], src: r[2], type: r[3], date: r[4], dateText: r[5], title: r[6], detail: r[7],
    who: r[8], entities: r[9], topic: r[10], status: r[11], severity: r[12], msgs: String(r[13] || '').split(',').filter(Boolean), updated: r[14] };
}
function eventsDay(user, req) {
  const g = findGroup(req.groupId); if (!g) throw { code: 'bad_group' };
  const date = validDate(req.date); if (!date) throw { code: 'bad_request', message: '日期格式錯誤' };
  const list = dayMessages(g.id, date), atts = dayAtts(g.id, date);
  if (!list.length && !atts.length) throw { code: 'no_messages', message: '這一天沒有訊息' };
  const topics = knownTopics(80);
  const prompt = '對話群:' + g.name + mailNote(g) + '\n對話日期:' + date + '(星期' + '日一二三四五六'[new Date(date + 'T00:00:00Z').getUTCDay()] + ')\n' +
    '既有主題:' + (topics.length ? topics.join('、') : '(還沒有)') + '\n\n' + clip(linesFor(list, atts), 120000);
  const r = callClaude(user, 'events', MODELS.events, [{ type: 'text', text: prompt }], MAX_TOKENS.events, EVENTS_SYSTEM, 0);
  const j = parseLoose(r.text);
  if (!j) throw { code: 'bad_ai', message: 'AI 回傳的事件解析失敗' };
  const arr = v => (Array.isArray(v) ? v : (v ? [v] : [])).map(x => clip(String(x), 80)).filter(Boolean).join('、');
  const items = (Array.isArray(j.events) ? j.events : []).filter(e => e && e.title && EVENT_TYPES.indexOf(e.type) >= 0).slice(0, 20);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet('events');
    const rows = readAll('events');
    let max = 0;
    rows.forEach(x => { const n = parseInt(String(x[0]).replace(/\D/g, ''), 10); if (n > max) max = n; });
    for (let i = rows.length - 1; i >= 0; i--) if (rows[i][1] === g.id && rows[i][2] === date) sh.deleteRow(i + 2);
    const now = stamp();
    const out = items.map(e => ['e' + String(++max).padStart(5, '0'), g.id, date, e.type, validDate(e.event_date), clip(e.date_text || '', 80),
      clip(e.title, 120), clip(e.detail || '', 400), arr(e.who), arr(e.entities), clip(String(e.topic || '').trim(), 80),
      ['open', 'done', 'cancelled', 'info'].indexOf(e.status) >= 0 ? e.status : 'info',
      e.type === 'incident' && ['high', 'medium', 'low'].indexOf(e.severity) >= 0 ? e.severity : '',
      mapRefs(e.refs, list).slice(0, 5).join(','), now]);
    appendRows('events', out);
    return { ok: true, events: out.map(eventOut), truncated: r.truncated };
  } finally { lock.releaseLock(); }
}

/* ===== 8. 期間報告 ===== */
const REPORT_INTRO = [
  '你是台灣離岸風場運維現場協調(AOC)的資深助理,要幫 AOC 同仁掌握一段期間內「整個專案」的走向。',
  '輸入是這段期間各對話群抽出的事件(e 開頭編號)、每日摘要標題與待追問(q 開頭編號)。資訊散在不同群組;群組也可能是電子郵件(標「Copilot 整理」的是二手摘要,日期數字以對話或原信為準)。',
  '只根據輸入內容,不要臆測。文字用繁體中文(台灣用語),人名、船名、風機編號照原文。',
];
// 第一次呼叫:整體走向+跨群組議題
const REPORT_SYSTEM_A = REPORT_INTRO.concat([
  '1. overview:3 到 5 句、250 字以內,說明整體走向、最重要的變化、需要特別注意的事。',
  '2. threads:把同一件事(可能跨群組、跨多天)串成議題,最多 6 個,挑最重要的。每個議題:topic、status(一句話,40 字內)、',
  '   priority(high/medium/low)、timeline(依日期排序,最多 5 步,每步 {date(MM/DD), text(40 字內), events:[最多 3 個編號]})、next(下一步,40 字內,沒有就空字串)。',
  '只輸出 JSON:{"overview":"","threads":[{"topic":"","status":"","priority":"","timeline":[{"date":"","text":"","events":[]}],"next":""}]}',
]).join('\n');
// 第二次呼叫:合併重複、狀態更新、重點、待追問答案
const REPORT_SYSTEM_B = REPORT_INTRO.concat([
  '1. merge:不同群組或不同天講的「其實是同一件事」的事件編號分組,例如 [["e3","e17"]]。',
  '2. updates:後來的事件顯示某事件已經完成、取消或改期,列 {event, status(done/cancelled/changed), note(30 字內,改期要寫新日期)}。',
  '3. highlights:最需要 AOC 注意的事件編號,最多 8 個。',
  '4. resolved:待追問(q 開頭)裡,後來的事件或摘要已經有答案的,列 {q, note(30 字內說明答案)}。',
  '只輸出 JSON:{"merge":[],"updates":[],"highlights":[],"resolved":[]}',
]).join('\n');

function periodReport(user, req) {
  const from = validDate(req.from), to = validDate(req.to);
  if (!from || !to || from > to) throw { code: 'bad_request', message: '日期區間錯誤' };
  const all = groupRows();
  const gids = Array.isArray(req.groups) && req.groups.length ? req.groups.map(String) : all.map(g => g.id);
  const gname = id => (all.find(g => g.id === id) || {}).name || id;
  const inR = d => d && d >= from && d <= to;
  // 涵蓋範圍
  const cov = {};
  gids.forEach(id => cov[id] = { group: id, name: gname(id), msgDays: {}, sumDays: {}, evDays: {} });
  readAll('msgs').forEach(r => { if (cov[r[1]] && inR(r[2])) cov[r[1]].msgDays[r[2]] = (cov[r[1]].msgDays[r[2]] || 0) + 1; });
  const sums = readAll('sums').filter(r => cov[r[0]] && inR(r[1])).map(sumOut);
  sums.forEach(s => cov[s.group].sumDays[s.date] = true);
  const cop = {}; readOpt('mails').forEach(r => { if (r[7] === 'copilot') cop[r[0]] = true; });
  const evAll = readAll('events').map(eventOut).filter(e => cov[e.group]).map(e => (e.msgs.length && e.msgs.every(m => cop[m]) ? Object.assign(e, { copilot: true }) : e));
  evAll.forEach(e => { if (inR(e.src)) cov[e.group].evDays[e.src] = true; });
  // 期間內提到的事件,加上期間前提到、但日期落在期間內的期限
  const evs = evAll.filter(e => inR(e.src) || (e.src < from && inR(e.date)));
  const coverage = gids.map(id => {
    const c = cov[id];
    const days = Object.keys(c.msgDays).sort();
    return { group: id, name: c.name, days: days.map(d => ({ date: d, msgs: c.msgDays[d], summary: !!c.sumDays[d], events: !!c.evDays[d] })) };
  });
  const missing = [];
  coverage.forEach(c => c.days.forEach(d => { if (!d.events) missing.push({ group: c.group, date: d.date }); }));
  if (!evs.length) throw { code: 'no_events', message: '這段期間沒有事件資料' + (missing.length ? '(有 ' + missing.length + ' 天還沒抽事件)' : '') };

  sums.sort((a, b) => a.date < b.date ? -1 : 1);
  const questions = [];
  sums.forEach(s => (s.summary && s.summary.questions || []).forEach(q => questions.push({ id: 'q' + (questions.length + 1),
    group: s.group, groupName: gname(s.group), date: s.date, zh: q.zh, refs: q.refs || [] })));
  const sumLines = sums.map(s => '[' + s.date + '][' + gname(s.group) + '] ' + ((s.summary && s.summary.title && s.summary.title.zh) || ''));
  const qLines = questions.map(q => q.id + ' [' + q.date + '][' + q.groupName + '] ' + q.zh);
  const evLines = evs.map(e => e.id + ' [' + e.src + '][' + gname(e.group) + (e.copilot ? '/Copilot 整理' : '') + '][' + e.type + (e.severity ? '/' + e.severity : '') + '][' + (e.topic || '-') + ']' +
    (e.date ? '[事件日 ' + e.date + ']' : '') + '[' + e.status + '] ' + e.title + ':' + e.detail + (e.who ? '(' + e.who + ')' : ''));
  const withData = coverage.filter(c => c.days.length).map(c => c.name);
  const prompt = '期間:' + from + ' ~ ' + to + '\n有資料的對話群:' + withData.join('、') + '\n\n事件(' + evs.length + ' 個):\n' +
    clip(evLines.join('\n'), 90000) + '\n\n每日摘要標題:\n' + clip(sumLines.join('\n'), 20000) +
    (qLines.length ? '\n\n待追問:\n' + clip(qLines.join('\n'), 20000) : '');
  const ra = callClaude(user, 'report', MODELS.report, [{ type: 'text', text: prompt }], MAX_TOKENS.report, REPORT_SYSTEM_A, 0);
  const rb = callClaude(user, 'report', MODELS.report, [{ type: 'text', text: prompt }], MAX_TOKENS.report, REPORT_SYSTEM_B, 0);
  const j = Object.assign({}, parseLoose(rb.text) || {}, parseLoose(ra.text) || {});
  const r = { truncated: ra.truncated || rb.truncated };
  const okId = {}; evs.forEach(e => okId[e.id] = true);
  const ids = v => (Array.isArray(v) ? v : []).map(String).filter(x => okId[x]);
  const rep = {
    from, to, groups: gids, groupNames: gids.map(gname), created: stamp(), model: MODELS.report, truncated: r.truncated,
    overview: clip(j.overview || '', 3000),
    threads: (Array.isArray(j.threads) ? j.threads : []).slice(0, 10).map(t => ({
      topic: clip(t.topic || '', 80), status: clip(t.status || '', 300), next: clip(t.next || '', 300),
      priority: ['high', 'medium', 'low'].indexOf(t.priority) >= 0 ? t.priority : 'medium', events: ids(t.events),
      timeline: (Array.isArray(t.timeline) ? t.timeline : []).slice(0, 8).map(s => ({ date: clip(s.date || '', 20), text: clip(s.text || '', 300), events: ids(s.events).slice(0, 3) })),
    })),
    merge: (Array.isArray(j.merge) ? j.merge : []).map(ids).filter(m => m.length > 1),
    updates: (Array.isArray(j.updates) ? j.updates : []).filter(u => u && okId[u.event]).map(u => ({ event: u.event,
      status: ['done', 'cancelled', 'changed'].indexOf(u.status) >= 0 ? u.status : 'changed', note: clip(u.note || '', 200) })),
    highlights: ids(j.highlights),
    events: evs.map(e => Object.assign({}, e, { groupName: gname(e.group) })),
    questions: (() => {
      const res = {}; (Array.isArray(j.resolved) ? j.resolved : []).forEach(x => { if (x && x.q) res[String(x.q)] = clip(x.note || '已有答案', 300); });
      return questions.map(q => Object.assign({}, q, res[q.id] ? { resolved: res[q.id] } : {}));
    })(),
    coverage, missing,
  };
  const json = JSON.stringify(rep);
  if (json.length > 6 * 45000) throw { code: 'too_big', message: '期間太長,請縮短日期區間' };
  const parts = []; for (let i = 0; i < 6; i++) parts.push(json.slice(i * 45000, (i + 1) * 45000));
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let id;
  try {
    let max = 0;
    readAll('reports').forEach(x => { const n = parseInt(String(x[0]).replace(/\D/g, ''), 10); if (n > max) max = n; });
    id = 'p' + String(max + 1).padStart(4, '0');
    const title = from + ' ~ ' + to + '(' + (Array.isArray(req.groups) && req.groups.length ? gids.length + ' 個對話群' : '全部對話群') + ')';
    appendRows('reports', [[id, rep.created, user.name, from, to, gids.join(','), title].concat(parts)]);
  } finally { lock.releaseLock(); }
  rep.id = id;
  return { ok: true, report: rep };
}
function reportGet(req) {
  const row = readAll('reports').find(r => r[0] === String(req.id || ''));
  if (!row) throw { code: 'bad_request', message: '找不到報告' };
  const rep = JSON.parse(row.slice(7, 13).join(''));
  rep.id = row[0];
  return { ok: true, report: rep };
}

/* ===== 9. 單字標記與練習紀錄(Tina 的學習狀況) ===== */
function vmarkOut(r) {
  return { term: r[0], mark: r[1], tries: +r[2] || 0, ok: +r[3] || 0, ng: +r[4] || 0, last: r[5], recent: r[6] || '', updated: r[7], user: r[8] };
}
function vmarkUpsert(user, terms, fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet('vmarks'), rows = readAll('vmarks'), idx = {};
    rows.forEach((r, i) => idx[r[0]] = i);
    const now = stamp(), add = [], out = [];
    terms.forEach(t => {
      const key = clip(String(t || '').trim().toLowerCase(), 120); if (!key) return;
      let r = idx[key] !== undefined ? rows[idx[key]] : null;
      const isNew = !r;
      if (isNew) r = [key, '', '0', '0', '0', '', '', '', ''];
      fn(r, t);
      r[7] = now; r[8] = user.name;
      if (isNew) { add.push(r); idx[key] = -1; } else sh.getRange(idx[key] + 2, 1, 1, r.length).setNumberFormat('@').setValues([r.map(String)]);
      out.push(vmarkOut(r));
    });
    appendRows('vmarks', add);
    return out;
  } finally { lock.releaseLock(); }
}
function vocabMark(user, req) {
  const mark = ['important', 'skip', ''].indexOf(req.mark) >= 0 ? req.mark : '';
  const terms = (Array.isArray(req.terms) ? req.terms : [req.term]).filter(Boolean).slice(0, 500);
  return { ok: true, marks: vmarkUpsert(user, terms, r => { r[1] = mark; }) };
}
function vocabPractice(user, req) {
  const res = (Array.isArray(req.results) ? req.results : []).filter(x => x && x.term).slice(0, 100);
  const by = {}; res.forEach(x => { const k = String(x.term).trim().toLowerCase(); (by[k] = by[k] || []).push(!!x.ok); });
  return { ok: true, marks: vmarkUpsert(user, Object.keys(by), (r, k) => {
    const list = by[k];
    r[2] = String((+r[2] || 0) + list.length);
    r[3] = String((+r[3] || 0) + list.filter(Boolean).length);
    r[4] = String((+r[4] || 0) + list.filter(x => !x).length);
    r[5] = stamp();
    r[6] = (String(r[6] || '') + list.map(x => x ? '1' : '0').join('')).slice(-10);
  }) };
}

/* ===== 10. 信件(從 Outlook 貼上:原信或 Copilot 整理) ===== */
function mailNote(g) {
  return g && g.kind === 'mail' ? '(這個群組是電子郵件,不是 Teams 對話;每則是一封信,開頭寫主旨;標【Copilot 整理】的是 Copilot 整理過的重點,不是原文)' : '';
}
function nextRunId() {
  let maxRun = 0;
  readAll('runs').forEach(r => { const n = parseInt(String(r[0]).replace(/\D/g, ''), 10); if (n > maxRun) maxRun = n; });
  return 'r' + String(maxRun + 1).padStart(4, '0');
}
function guessLang(t) {
  const cjk = (String(t).match(/[一-鿿]/g) || []).length, lat = (String(t).match(/[A-Za-z]/g) || []).length;
  return cjk < 10 ? (lat ? 'en' : 'zh') : (lat > cjk ? 'mixed' : 'zh');
}
function mailText(m) {
  return (m.source === 'copilot' ? '【Copilot 整理】\n' : '') + '主旨:' + m.subject + (m.to ? '\n收件人:' + m.to : '') + '\n\n' + m.body;
}
function saveMails(user, req) {
  const g = findGroup(req.groupId);
  if (!g) throw { code: 'bad_group' };
  if (g.kind !== 'mail') throw { code: 'bad_request', message: '這個群組不是信件群組' };
  const incoming = (Array.isArray(req.mails) ? req.mails : []).slice(0, 300);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const runId = nextRunId(), now = stamp();
    const data = readAll('msgs'), mrows = readAll('mails');
    const msgIdx = {}; data.forEach((r, i) => msgIdx[r[0]] = i);
    const metas = mrows.filter(r => r[1] === g.id).map(r => ({ id: r[0], date: r[2], time: r[3], s: norm(r[4]), subj: norm(r[6]), source: r[7], row: r }));
    const appends = [], mailAppends = [], ids = [], fresh = [];
    let newCount = 0, upgraded = 0, msgChanged = false, metaChanged = false;
    const dates = {};
    incoming.forEach((x, idx) => {
      x = x || {};
      const m = {
        date: validDate(x.date), time: validTime(x.time),
        sender: clip(String(x.sender || '').trim(), 120), to: clip(String(x.to || '').trim(), 400),
        subject: clip(String(x.subject || '').trim(), 300), body: clip(String(x.body || '').trim(), 18000),
        source: x.source === 'copilot' ? 'copilot' : 'original',
      };
      if (!m.date || !m.sender || (!m.subject && !m.body)) { ids[idx] = ''; return; }
      dates[m.date] = true;
      const s = norm(m.sender), subj = norm(m.subject);
      const hit = metas.find(o => o.date === m.date && o.s === s && o.subj === subj && (!o.time || !m.time || o.time === m.time));
      const text = mailText(m), lang = guessLang(m.subject + ' ' + m.body);
      if (hit) {
        ids[idx] = hit.id;
        const di = msgIdx[hit.id];
        const arr = di !== undefined ? data[di] : (appends.find(a => a[0] === hit.id) || null);
        if (arr) {
          const rs = String(arr[10] || '').split(',').filter(Boolean);
          if (rs.indexOf(runId) < 0) { rs.push(runId); arr[10] = rs.join(','); msgChanged = true; }
          // 原信取代 Copilot 整理的版本
          if (hit.source === 'copilot' && m.source === 'original') {
            arr[5] = text; arr[8] = lang; if (m.time) arr[3] = m.time; msgChanged = true;
            hit.row[3] = m.time || hit.row[3]; hit.row[5] = m.to || hit.row[5]; hit.row[7] = 'original'; hit.row[8] = runId; hit.source = 'original';
            metaChanged = true; upgraded++; fresh.push(hit.id);
          }
        }
        return;
      }
      let id = 'x' + md5hex([g.id, m.date, s, subj, m.time].join('|')).slice(0, 14), k = 0;
      while (msgIdx[id] !== undefined || appends.some(a => a[0] === id)) id = id.replace(/-\d+$/, '') + '-' + (++k);
      appends.push([id, g.id, m.date, m.time, m.sender, text, '', '', lang, runId, runId, now]);
      const row = [id, g.id, m.date, m.time, m.sender, m.to, m.subject, m.source, runId, now];
      mailAppends.push(row);
      metas.push({ id, date: m.date, time: m.time, s, subj, source: m.source, row });
      ids[idx] = id; fresh.push(id); newCount++;
    });
    if (msgChanged && data.length) sheet('msgs').getRange(2, 1, data.length, 12).setNumberFormat('@').setValues(data.map(r => r.slice(0, 12)));
    if (metaChanged && mrows.length) sheet('mails').getRange(2, 1, mrows.length, 10).setNumberFormat('@').setValues(mrows);
    appendRows('msgs', appends);
    appendRows('mails', mailAppends);
    let cnt = 0, first = '', last = '';
    data.concat(appends).forEach(r => {
      if (r[1] !== g.id) return;
      cnt++;
      if (r[2]) { if (!first || r[2] < first) first = r[2]; if (!last || r[2] > last) last = r[2]; }
    });
    sheet('groups').getRange(g.row, 6, 1, 3).setValues([[String(cnt), first, last]]);
    const ds = Object.keys(dates).sort();
    appendRows('runs', [[runId, now, user.name, g.id, clip(req.fileName || '貼上信件', 300), '', '',
      String(incoming.length), String(newCount), ds[0] || '', ds[ds.length - 1] || '']]);
    return { ok: true, runId, newCount, upgraded, total: incoming.length, dates: ds, ids, fresh: fresh.filter((v, i, a) => a.indexOf(v) === i) };
  } finally { lock.releaseLock(); }
}

/* 信件 ↔ 對話:哪幾則對話在講同一件事 */
function mailMeta() {
  const m = {}; readOpt('mails').forEach(r => m[r[0]] = { id: r[0], group: r[1], date: r[2], time: r[3], sender: r[4], to: r[5], subject: r[6], source: r[7] });
  return m;
}
function brief(r) { return r ? { id: r[0], group: r[1], date: r[2], time: r[3], sender: r[4], text: clip(r[5], 240) } : null; }
function linkInfo(msgs, rows) {
  const want = {}; msgs.forEach(m => want[m.id] = true);
  const lr = readOpt('links').filter(r => r[6] === 'ok' && (want[r[1]] || want[r[2]]));
  const meta = mailMeta(), mails = {};
  msgs.forEach(m => { if (meta[m.id]) mails[m.id] = meta[m.id]; });
  if (!lr.length) return { links: [], mails };
  const byId = {}; rows.forEach(r => byId[r[0]] = r);
  const links = lr.filter(r => byId[r[1]] && byId[r[2]]).map(r => ({ id: r[0], mail: r[1], msg: r[2], reason: r[3], conf: r[4], by: r[5],
    mailBrief: Object.assign(brief(byId[r[1]]), { subject: (meta[r[1]] || {}).subject || '', source: (meta[r[1]] || {}).source || '' }),
    msgBrief: brief(byId[r[2]]) }));
  return { links, mails };
}
/* 對話存檔後:找出日期在前後 LINK_WINDOW 天內的信件,讓網頁逐封重新比對 */
function linkTargets(req) {
  const ds = (Array.isArray(req.dates) ? req.dates : []).map(validDate).filter(Boolean);
  if (!ds.length) return { ok: true, ids: [] };
  const lo = shiftDate(ds.slice().sort()[0], -LINK_WINDOW), hi = shiftDate(ds.slice().sort().pop(), LINK_WINDOW);
  const ids = readAll('mails').filter(r => r[2] >= lo && r[2] <= hi).map(r => r[0]);
  return { ok: true, ids };
}
function shiftDate(iso, n) { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
const STOP = ('the and for with this that from have has had will would can could should please thanks thank you your our are was were been be not but all any ' +
  'about into also there their them they its it\'s just let know need needs per via re fw fwd dear hi hello regards best kind team today tomorrow yesterday ' +
  'yunlin sgre siemens gamesa energy wind farm').split(' ');
function tokens(text) {
  const t = String(text || '').toLowerCase(), out = {};
  (t.match(/[a-z][a-z0-9_\-]{2,}|\d+[\/.:-]\d+(?:[\/.:-]\d+)?|[a-z]*\d{2,}[a-z0-9]*/g) || []).forEach(w => { if (STOP.indexOf(w) < 0) out[w] = 1; });
  const cjk = t.replace(/[^一-鿿]+/g, ' ').split(' ');
  cjk.forEach(seg => { for (let i = 0; i + 1 < seg.length; i++) out[seg.substr(i, 2)] = 1; });
  return Object.keys(out);
}
const LINK_SYSTEM = [
  '你要找出 Teams 對話裡,跟一封電子郵件講「同一件具體事情」的訊息。',
  '同一件事 = 同一批人、同一個日期或期限、同一艘船、同一台風機、同一份文件、同一個要求或問題。只是同一個大主題(例如都在講人力)不算。',
  '例如:信裡確認 10/24 報到的名單,對話裡有人問「Muehlhan 哪天到」→ 相關;信講 ActSafe 檢驗報告,對話講吊掛證照 → 不相關。',
  '每個配對寫 reason(繁體中文 40 字內,說明兩邊講的是哪件事)與 conf(high 很確定 / medium 大概是)。不確定就不要列。最多 8 則。',
  '只輸出 JSON:{"links":[{"c":3,"reason":"","conf":"high"}]}',
].join('\n');
function linkMails(user, req) {
  const ids = (Array.isArray(req.ids) ? req.ids : [req.id]).filter(Boolean).slice(0, 3);
  if (!ids.length) return { ok: true, added: [] };
  const rows = readAll('msgs'), meta = mailMeta(), groups = groupRows();
  const kind = {}; groups.forEach(g => kind[g.id] = g.kind);
  const gname = id => (groups.find(g => g.id === id) || {}).name || id;
  const existing = readAll('links');
  const added = [], skipped = [];
  ids.forEach(mid => {
    const mr = rows.find(r => r[0] === mid), mm = meta[mid];
    if (!mr || !mm) { skipped.push(mid); return; }
    const lo = shiftDate(mm.date, -LINK_WINDOW), hi = shiftDate(mm.date, LINK_WINDOW);
    const done = {}; existing.forEach(r => { if (r[1] === mid) done[r[2]] = true; });
    const pool = rows.filter(r => kind[r[1]] !== 'mail' && r[2] >= lo && r[2] <= hi && !done[r[0]]);
    if (!pool.length) return;
    // 先用關鍵字重疊挑候選,太常見的字權重低
    // 只比內文(人名到處都是,拿來比會誤配)
    const df = {}; const toks = pool.map(r => { const t = tokens(r[5] + ' ' + r[6] + ' ' + r[7]); t.forEach(w => df[w] = (df[w] || 0) + 1); return t; });
    const mailBody = String(mr[5]).split('\n').filter(l => !/^(收件人:|【Copilot 整理】)/.test(l)).join('\n');
    const mt = {}; tokens(mailBody).forEach(w => mt[w] = 1);
    const scored = pool.map((r, i) => {
      let sc = 0; toks[i].forEach(w => { if (mt[w]) sc += 1 / Math.log(2 + df[w]); });
      const gap = Math.abs((new Date(r[2]) - new Date(mm.date)) / 864e5);
      return { r, sc: sc * (1 - gap / (LINK_WINDOW * 3)) };
    }).filter(x => x.sc > 0.35).sort((a, b) => b.sc - a.sc).slice(0, LINK_CANDS);
    if (!scored.length) return;
    scored.sort((a, b) => (a.r[2] + a.r[3]).localeCompare(b.r[2] + b.r[3]));
    const lines = ['信件:', '日期 ' + mm.date + ' ' + (mm.time || '') + '|寄件人 ' + mm.sender + (mm.to ? '|收件人 ' + clip(mm.to, 200) : '') +
      (mm.source === 'copilot' ? '|(Copilot 整理的重點,日期可能有誤差)' : ''), clip(mr[5], 3500), '', '候選的 Teams 訊息:'];
    scored.forEach((x, k) => lines.push('c' + (k + 1) + ' [' + gname(x.r[1]) + '][' + x.r[2] + ' ' + (x.r[3] || '') + '] ' + clip(x.r[4], 40) + ':' +
      clip(x.r[5], 220).replace(/\n/g, ' ') + (x.r[7] ? ' (附件 ' + clip(x.r[7], 120).replace(/\n/g, ' ') + ')' : '')));
    const res = callClaude(user, 'link', MODELS.link, [{ type: 'text', text: lines.join('\n') }], MAX_TOKENS.link, LINK_SYSTEM, 0);
    const j = parseLoose(res.text) || {};
    (Array.isArray(j.links) ? j.links : []).slice(0, 10).forEach(l => {
      const k = parseInt(String(l && l.c).replace(/\D/g, ''), 10);
      if (!(k >= 1 && k <= scored.length)) return;
      const conf = l.conf === 'high' ? 'high' : l.conf === 'medium' ? 'medium' : '';
      if (!conf) return;
      const cid = scored[k - 1].r[0];
      if (added.some(a => a[1] === mid && a[2] === cid)) return;
      added.push(['', mid, cid, clip(l.reason || '', 200), conf, 'ai', 'ok', '']);
    });
  });
  if (added.length) {
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const cur = readAll('links'), have = {};
      let max = 0;
      cur.forEach(r => { have[r[1] + '|' + r[2]] = true; const n = parseInt(String(r[0]).replace(/\D/g, ''), 10); if (n > max) max = n; });
      const now = stamp(), rowsOut = [];
      added.forEach(a => { if (have[a[1] + '|' + a[2]]) return; have[a[1] + '|' + a[2]] = true; a[0] = 'k' + String(++max).padStart(5, '0'); a[7] = now; rowsOut.push(a); });
      appendRows('links', rowsOut);
      return { ok: true, added: rowsOut.map(a => ({ id: a[0], mail: a[1], msg: a[2], reason: a[3], conf: a[4] })), skipped };
    } finally { lock.releaseLock(); }
  }
  return { ok: true, added: [], skipped };
}
/* 刪掉配錯的(保留紀錄,之後重新比對不會再配回來) */
function linkRemove(req) {
  const id = String(req.id || '');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const rows = readAll('links'), i = rows.findIndex(r => r[0] === id);
    if (i < 0) return { ok: true, removed: 0 };
    sheet('links').getRange(i + 2, 7).setValue('removed');
    return { ok: true, removed: 1 };
  } finally { lock.releaseLock(); }
}

/* ===== 11. 排班工具的「問 AI」 =====
   網頁(雲林排班工具)把對話送來;AI 需要資料時回 tool_use,
   由網頁在使用者電腦上查 Excel、算好再送回來(資料不存在這裡,也不寫進試算表)。 */
const ASK_SYSTEM = [
  '你是台灣雲林離岸風場(Siemens Gamesa)現場協調 AOC 的排班助理,回答 AOC 同仁(主要是 Tina)關於排班的問題。',
  '資料來源是 AOC 電腦上已載入的兩個 Excel:人力總表(Resource Planning 的 Allowance 分頁)與每日出海計畫(Hand-Over 的 1.Daily Ops)。用工具查,不要憑印象。',
  '',
  '【一定要遵守】',
  '- 人數、名單、天數一律用工具查到的結果,不要自己心算或猜;查不到就說查不到。',
  '- 人名用總表上的寫法(例如「Chen, YiAn (Kenneth)」「Jacky Hong」);列名單用 1- 2- 3- 編號。',
  '- 用繁體中文(台灣用語),簡潔;人名、船名、風機編號保留原文。先給結論,再給依據。',
  '- 你只能回答與建議,不能改 Excel;要改的地方請明確告訴 AOC 改哪一天、哪一隊、哪個人。',
  '- 使用者附的檔案(圖片、PDF、Excel 轉成的表格文字)也可以拿來比對。',
  '',
  '【總表代碼】',
  '- 等級碼 TS/TSF/MT/MTF/CT/BT/HV/TSWA/W = 當天可派工(計畫區)。OF = 已出海、ON = 留岸上(已發生的紀錄,OF 有出海津貼)。',
  '- L = 請假、T = 訓練、fly = 交通日、空白 = 不在這一輪;FMS/HL/F2 = 當天不是為雲林工作。隱藏列 = 不排。',
  '- 總表以「今天」為界:今天以前多是 OF/ON 紀錄,今天起是等級碼計畫。全員 ON、沒有 OF = 天氣取消日,不比人數。',
  '- w 是等級之一,Yunlin Office 群組的人不出海;Gordon Huang 永不派工。',
  '',
  '【排班規則(AOC 確認過)】',
  '- 技師兩隊七休七,週二換班;分包商同步週二輪換。連續上班最多 14 天。',
  '- 隊長只能 TS/TSF/MT/MTF;CT、BT 只能當隊員(BT 不一定要 CT 陪)。每隊至少 3 人;每艘船總人數不超過 12。',
  '- 年保 AS 隊理想 6 人,4~5 人算正常。ST(Senior Tech)原則不出海,缺人時偶爾會派(提醒即可)。',
  '- 夜班目標 12 人(湊不齊就算了);夜班人選 = 總表姓名欄綠色的分包商 + Jacky Hong、Ken Chao。',
  '- 有歐洲 HV 技師的隊,隊長要 TSF/MTF,還要搭一位 CT 或 BT。TSWA(Muehlhan 葉片組)要成雙、整組同進同出。',
  '- OT- 前綴 = 輪休週加班;TT- = Global Travel Tech。班表姓名有刪除線 = 取消。',
  '- IPS 等分包可以調人,但預約後要取消須出海前三天通知。',
  '- 「班表已排」與「總表可出海」不一致時,用 day_overview 的 diff 說明差額是誰。',
].join('\n');
const ASK_TOOLS = [
  { name: 'day_overview', description: '某一天的出海計畫與檢核:每隊名單(含等級、取消、對不到總表的名字、備註)、人數摘要、差額是誰、檢核發現的問題。也會說畫面上有沒有這一天的草稿。',
    input_schema: { type: 'object', properties: { date: { type: 'string', description: 'YYYY-MM-DD' } }, required: ['date'] } },
  { name: 'person_schedule', description: '查某人在一段期間每天的總表代碼與被排到的隊伍(含夜班),並算出連續上班天數。名字可以只打一部分(英文名、姓)。',
    input_schema: { type: 'object', properties: { name: { type: 'string' }, from: { type: 'string', description: 'YYYY-MM-DD' }, to: { type: 'string', description: 'YYYY-MM-DD,最多 62 天' } }, required: ['name', 'from', 'to'] } },
  { name: 'available_people', description: '某一天總表可派工的人(可依等級篩選、只看還沒被排到的),附等級、群組、是否已排、被排在哪隊。找補人時用。',
    input_schema: { type: 'object', properties: { date: { type: 'string' }, levels: { type: 'array', items: { type: 'string' }, description: '例如 ["MT","MTF"],不填=全部' }, only_unscheduled: { type: 'boolean' } }, required: ['date'] } },
  { name: 'find_people', description: '用部分名字、等級或群組(分包商)找總表上的人。',
    input_schema: { type: 'object', properties: { query: { type: 'string', description: '名字片段、等級(如 CT)或群組(如 IPS)' } }, required: ['query'] } },
  { name: 'date_range_counts', description: '一段期間每天的人數摘要(班表已排、總表可出海、差額、夜班人數、要處理的問題數),看趨勢或找問題日用。',
    input_schema: { type: 'object', properties: { from: { type: 'string' }, to: { type: 'string', description: '最多 31 天' } }, required: ['from', 'to'] } },
];
function ask(user, req) {
  const msgs = Array.isArray(req.messages) ? req.messages : [];
  if (!msgs.length || msgs[0].role !== 'user' || msgs[msgs.length - 1].role !== 'user')
    throw { code: 'bad_request', message: '對話格式錯誤' };
  for (let i = 1; i < msgs.length; i++) if (msgs[i].role === msgs[i - 1].role) throw { code: 'bad_request', message: '對話格式錯誤' };
  const size = JSON.stringify(msgs).length;
  if (size > 24000000) throw { code: 'too_big', message: '對話太長或附件太大,請按「新對話」' };
  let nImg = 0;
  msgs.forEach(m => (Array.isArray(m.content) ? m.content : []).forEach(b => { if (b && (b.type === 'image' || b.type === 'document')) nImg++; }));
  const payload = {
    model: MODELS.ask, max_tokens: MAX_TOKENS.ask, tools: ASK_TOOLS, messages: msgs,
    system: [{ type: 'text', text: ASK_SYSTEM, cache_control: { type: 'ephemeral' } },
             { type: 'text', text: '目前狀態:\n' + clip(req.context || '', 4000) }],
  };
  const body = postAnthropic(payload);
  logUsage(user, 'ask', MODELS.ask, nImg, body.usage || {});
  return { ok: true, content: body.content || [], stop_reason: body.stop_reason || '' };
}
