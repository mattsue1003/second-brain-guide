// 用假的 Google 服務（試算表、UrlFetch、Gmail…）在本機跑教學裡的 Apps Script 程式碼。
// 驗證的是程式邏輯；真正的 Gemini 與 LINE 連線沒辦法在這裡測。
// 執行：node test.js
const fs = require('fs'), vm = require('vm'), path = require('path');
const GAS = path.join(__dirname, 'src', 'gas');

let pass = 0, fail = 0;
function ok(cond, name) { if (cond) pass++; else { fail++; console.log('  ✗ ' + name); } }

function makeWorld() {
  const sheets = {}, order = [];
  const props = {}, cache = {}, sent = { line: [], mail: [], triggers: [], files: [], gemini: [] };
  function Sheet(name) {
    this.name = name; this.rows = [];
  }
  Sheet.prototype = {
    getLastRow() { return this.rows.length; },
    appendRow(r) { this.rows.push(r.slice()); },
    getRange(r, c, nr, nc) {
      const sh = this; nr = nr || 1; nc = nc || 1;
      const api = {
        getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push((sh.rows[r - 1 + i] || [])[c - 1 + j] ?? ''); out.push(row); } return out; },
        setValues(v) { v.forEach((row, i) => { sh.rows[r - 1 + i] = sh.rows[r - 1 + i] || []; row.forEach((x, j) => { sh.rows[r - 1 + i][c - 1 + j] = x; }); }); return api; },
        setValue(v) { return api.setValues([[v]]); },
        setFontWeight() { return api; }
      };
      return api;
    },
    setFrozenRows() {}, setColumnWidth() {}
  };
  const ss = {
    getSheetByName: n => sheets[n] || null,
    insertSheet(n) { sheets[n] = new Sheet(n); order.push(n); return sheets[n]; }
  };
  // 假的向量：依關鍵字出現與否產生，讓「意思相近」可以被測到
  const KEYS = ['手', '失智', '糖尿病', '血糖', '運動', '團體', '音樂', '飲食'];
  function fakeVec(t) { return KEYS.map(k => (t.includes(k) ? 1 : 0)).concat([0.01]); }
  function resp(code, body, blob) {
    return { getResponseCode: () => code, getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)), getBlob: () => blob };
  }
  function gen(text) { return resp(200, { candidates: [{ content: { parts: [{ text }] } }] }); }
  const world = { sheets, props, cache, sent, geminiDown: false };
  const UrlFetchApp = {
    fetch(url, opt) {
      opt = opt || {};
      const body = opt.payload ? JSON.parse(opt.payload) : {};
      if (url.includes(':generateContent')) {
        sent.gemini.push({ url, body });
        if (world.geminiDown) return resp(503, { error: { code: 503, message: 'high demand' } });
        if ((opt.headers || {})['x-goog-api-key'] !== 'TESTKEY') return resp(400, { error: { code: 400, message: 'API key not valid' } });
        const parts = body.contents[0].parts, text = parts.map(p => p.text || '').join('');
        if (parts.some(p => p.inline_data)) return gen(parts.find(p => p.inline_data).inline_data.mime_type.startsWith('audio') ? '今天帶長輩做音樂律動，大家很投入' : '手指操步驟：握拳、張開、對指');
        if (text.includes('只輸出一個 JSON')) {
          const content = text.split('內容：\n')[1] || '';
          const cat = content.includes('手') ? '教案點子' : content.includes('糖尿病') ? '研究文獻' : '不存在的類別';
          return gen('```json\n' + JSON.stringify({ title: '標題-' + content.slice(0, 6), summary: '摘要：' + content.slice(0, 20), tags: ['標籤甲', content.includes('手') ? '手功能' : '其他標'], category: cat }) + '\n```');
        }
        if (text.includes('我的問題：')) return gen('根據筆記，重點是手指操 [#1]。');
        if (text.includes('每週回顧')) return gen('【這週在想什麼】手功能。\n【值得延伸的三件事】1. 做教案 [#1]\n【可以發表的題目】手指操');
        if (text.includes('任務：')) return gen('初稿：' + text.split('任務：')[1].slice(0, 10));
        return gen('第二大腦準備好了');
      }
      if (url.includes(':embedContent')) {
        if (!url.includes('gemini-embedding-2')) return resp(404, { error: { code: 404, message: 'no model' } });
        ok(body.output_dimensionality === 768, 'embed 要帶 output_dimensionality');
        return resp(200, { embedding: { values: fakeVec(body.content.parts[0].text) } });
      }
      if (url.startsWith('https://api.line.me/')) { sent.line.push({ url, body, auth: opt.headers.Authorization }); return resp(200, {}); }
      if (url.startsWith('https://api-data.line.me/')) {
        const blob = { getBytes: () => [1, 2, 3], getContentType: () => 'image/jpeg', setName(n) { this.name = n; return this; } };
        return resp(200, '', blob);
      }
      if (url.startsWith('https://example.com/')) return resp(200, '<html><head><title>糖尿病 飲食指南</title><style>x{}</style></head><body><script>bad()</script><p>糖尿病患者的飲食與血糖控制 &amp; 運動</p></body></html>');
      return resp(404, '');
    }
  };
  const g = {
    console, JSON, Math, Date, String, Number, Array, Object, RegExp, Error,
    Logger: { log() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty(k, v) { props[k] = v; }, deleteProperty(k) { delete props[k]; } }) },
    CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put(k, v) { cache[k] = v; } }) },
    Utilities: {
      formatDate(d, tz, f) { const p = n => String(n).padStart(2, '0'); const t = new Date(d.getTime() + 8 * 3600e3);
        return f.replace('yyyy', t.getUTCFullYear()).replace('MM', p(t.getUTCMonth() + 1)).replace('dd', p(t.getUTCDate())).replace('HH', p(t.getUTCHours())).replace('mm', p(t.getUTCMinutes())).replace('ss', p(t.getUTCSeconds())).replace('M', t.getUTCMonth() + 1).replace('d', t.getUTCDate()); },
      base64Encode: b => Buffer.from(b).toString('base64')
    },
    HtmlService: {
      createHtmlOutput: s => ({ kind: 'html', content: s }),
      createHtmlOutputFromFile: f => { const o = { kind: 'file', file: f, setTitle(t) { o.title = t; return o; }, addMetaTag(k, v) { o.meta = [k, v]; return o; } }; return o; }
    },
    DriveApp: {
      getFoldersByName: () => ({ hasNext: () => false }),
      createFolder: () => ({ createFile(b) { sent.files.push(b.name); return { getUrl: () => 'https://drive.google.com/file/' + b.name }; } })
    },
    MailApp: { sendEmail(m) { sent.mail.push(m); } },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'me@example.com' }) },
    ScriptApp: {
      WeekDay: { SUNDAY: 'SUN' },
      getProjectTriggers: () => sent.triggers.slice(),
      deleteTrigger(t) { sent.triggers.splice(sent.triggers.indexOf(t), 1); },
      newTrigger(fn) { const t = { fn, getHandlerFunction: () => fn }; const b = { timeBased: () => b, onWeekDay(d) { t.day = d; return b; }, atHour(h) { t.hour = h; return b; }, create() { sent.triggers.push(t); } }; return b; }
    },
    UrlFetchApp
  };
  world.g = g;
  return world;
}

const CONFIG = `
var GEMINI_API_KEY = 'TESTKEY';
var LINE_TOKEN = 'LINETOKEN';
var WEB_PASSWORD = 'pw-1234';
var MODEL = 'gemini-3.5-flash-lite';
var BACKUP_MODELS = ['gemini-3.5-flash-lite', 'gemini-3.6-flash', 'gemini-3.8-flash'];
var CATEGORIES = ['教案點子', '研究文獻', '活動觀察', '其他'];
var SAVE_FILES = true;
var WEEKLY_LINE = true;
`;

// 依階段載入檔案：模擬學員做到第 N 階段時，專案裡有哪些檔案
function load(world, files) {
  const ctx = vm.createContext(Object.assign({}, world.g));
  vm.runInContext(CONFIG, ctx);
  files.forEach(f => vm.runInContext(fs.readFileSync(path.join(GAS, f + '.gs'), 'utf8'), ctx, { filename: f + '.gs' }));
  return ctx;
}
let evN = 0;
function lineEvent(ctx, message, uid) {
  const ev = { type: 'message', webhookEventId: 'ev' + (++evN), replyToken: 'rt' + evN, source: { type: 'user', userId: uid || 'U-owner' }, message };
  return ctx.doPost({ postData: { contents: JSON.stringify({ events: [ev] }) } });
}
const lastReply = w => { const r = w.sent.line.filter(x => x.url.endsWith('/reply')).pop(); return r ? r.body.messages[0].text : ''; };

const w = makeWorld();

console.log('階段 1：LINE 隨手記');
let c = load(w, ['Capture']);
c.setup();
ok(w.sheets['筆記'] && w.sheets['筆記'].rows[0][0] === '編號', 'setup 建立「筆記」工作表與標題列');
let out = lineEvent(c, { type: 'text', id: 'm1', text: '手指操可以搭配懷舊歌曲，長輩比較願意做' });
ok(out.kind === 'html', 'doPost 回傳 HtmlService');
let rows = w.sheets['筆記'].rows;
ok(rows.length === 2 && rows[1][0] === 1 && rows[1][2] === '文字', '文字訊息存成第 1 則筆記');
ok(rows[1][6] === '教案點子' && rows[1][5] === '標籤甲、手功能', 'AI 整理出類別與標籤（能讀懂包在 ```json 裡的回覆）');
ok(rows[1][1] instanceof Date, '時間欄存的是日期');
ok(lastReply(w).includes('已存入 #1') && lastReply(w).includes('已把你設為'), '回覆「已存入 #1」並告知綁定主人');
ok(w.sent.line[0].auth === 'Bearer LINETOKEN', 'LINE 請求帶權杖');
ok(w.sent.line.some(x => x.url.endsWith('/chat/loading/start')), '有送出輸入中動畫');
ok(w.props.OWNER === 'U-owner', '第一個傳訊息的人成為主人');

lineEvent(c, { type: 'text', id: 'm2', text: '偷看一下' }, 'U-stranger');
ok(w.sheets['筆記'].rows.length === 2 && lastReply(w).includes('私人的筆記機器人'), '其他人的訊息不會被記錄');

const dup = { type: 'message', webhookEventId: 'same', replyToken: 'x', source: { userId: 'U-owner' }, message: { type: 'text', text: '重送測試 手' } };
c.doPost({ postData: { contents: JSON.stringify({ events: [dup] }) } });
c.doPost({ postData: { contents: JSON.stringify({ events: [dup] }) } });
ok(w.sheets['筆記'].rows.length === 3, 'LINE 重送同一則訊息只存一次');

lineEvent(c, { type: 'text', id: 'm3', text: '這篇不錯 https://example.com/dm' });
rows = w.sheets['筆記'].rows;
ok(rows[3][2] === '連結' && rows[3][8] === 'https://example.com/dm', '網址存成「連結」類型並保留網址');
ok(rows[3][7].includes('糖尿病患者的飲食') && !rows[3][7].includes('bad()') && rows[3][7].includes('我的備註：這篇不錯'), '抓到網頁內文、去掉 script、保留我的備註');
ok(rows[3][6] === '研究文獻', '連結也有分類');

lineEvent(c, { type: 'image', id: 'm4' });
rows = w.sheets['筆記'].rows;
ok(rows[4][2] === '照片' && rows[4][7].includes('手指操步驟'), '照片：文字被抄錄成原文');
ok(rows[4][8].startsWith('https://drive.google.com/file/') && w.sent.files[0].endsWith('.jpg'), '照片另存到雲端硬碟並留下連結');
lineEvent(c, { type: 'audio', id: 'm5' });
rows = w.sheets['筆記'].rows;
ok(rows[5][2] === '錄音' && rows[5][7].includes('音樂律動'), '語音：轉成逐字稿');
const audioCall = w.sent.gemini.filter(x => x.body.contents[0].parts.some(p => p.inline_data)).pop();
ok(audioCall.body.contents[0].parts[1].inline_data.mime_type === 'audio/m4a', '語音用 audio/m4a 送給 Gemini');
ok(rows[5][6] === '其他', 'AI 給了清單以外的類別時，改成「其他」');

lineEvent(c, { type: 'sticker', id: 'm6' });
ok(lastReply(w).includes('只收文字'), '貼圖等其他類型會說明不支援');
lineEvent(c, { type: 'text', id: 'm7', text: '？我記過什麼' });
ok(lastReply(w).includes('階段 3'), '還沒做階段 3 時，發問會提示');
lineEvent(c, { type: 'text', id: 'm8', text: '找 手指' });
ok(lastReply(w).includes('階段 2'), '還沒做階段 2 時，搜尋會提示');
ok(w.sheets['筆記'].rows.length === 6, '指令不會被存成筆記');

w.geminiDown = true;
lineEvent(c, { type: 'text', id: 'm9', text: 'AI 掛掉時也要存得進去 手' });
w.geminiDown = false;
rows = w.sheets['筆記'].rows;
ok(rows.length === 7 && rows[6][3] === 'AI 掛掉時也要存得進去 手', 'Gemini 忙線時筆記照樣存入，用原文開頭當標題');
ok(w.sent.gemini.slice(-3).map(x => x.url.match(/models\/([^:]+)/)[1]).join() === 'gemini-3.5-flash-lite,gemini-3.6-flash,gemini-3.8-flash', '忙線時依序換備用模型');
ok(w.sent.gemini.slice(-1)[0].body.generationConfig.thinkingConfig.thinkingLevel === 'low' && !w.sent.gemini.slice(-3)[0].body.generationConfig, '非 lite 模型調低思考，lite 不帶');

console.log('階段 2：知識總站');
c = load(w, ['Capture', 'Web']);
ok(c.doGet().file === 'Index' && c.doGet().meta[0] === 'viewport', 'doGet 回傳 Index 網頁，含手機 viewport');
ok(c.api('wrong', 'home').error === '密碼不對', '密碼錯誤被擋');
let h = c.api('pw-1234', 'home');
ok(h.count === 6 && h.notes[0].id === 6 && h.notes.length === 6, '首頁：總數與最近筆記（新到舊）');
ok(h.categories.find(x => x.name === '教案點子').count === 3 && h.tags[0].name === '標籤甲', '首頁：類別數量與常用標籤');
ok(h.canAsk === false && h.canCompose === false && h.kinds.length === 0, '後面階段還沒做時，功能標示為未開啟');
ok(h.notes[0].text === undefined, '清單不傳完整原文');
let s = c.api('pw-1234', 'search', { q: '糖尿病' });
ok(s.notes.length === 1 && s.notes[0].id === 3, '關鍵字搜尋找到原文裡的字');
s = c.api('pw-1234', 'search', { q: '', category: '教案點子' });
ok(s.notes.length === 3 && s.notes[0].id === 4, '只選類別：列出該類別，新到舊');
s = c.api('pw-1234', 'search', { q: '', tag: '其他標' });
ok(s.notes.length === 2, '只選標籤');
ok(c.api('pw-1234', 'search', { q: '完全沒有的字' }).notes.length === 0, '找不到時回空清單');
ok(c.api('pw-1234', 'note', { id: 3 }).note.text.includes('糖尿病'), '可以讀單則完整原文');
let a = c.api('pw-1234', 'add', { text: '網頁新增：血糖紀錄表的設計' });
ok(a.note.id === 7 && /^\d{4}\/\d\d\/\d\d$/.test(a.note.date), '網頁新增筆記');
ok(c.api('pw-1234', 'ask', { q: 'x' }).error.includes('階段 3') && c.api('pw-1234', 'compose', {}).error.includes('階段 5'), '未完成的功能回傳提示');
lineEvent(c, { type: 'text', id: 'm10', text: '找 糖尿病' });
ok(lastReply(w).includes('找到 1 則') && lastReply(w).includes('#3'), 'LINE「找 關鍵字」');
for (let i = 0; i < 5; i++) c.api('bad', 'home');
ok(c.api('pw-1234', 'home').error.includes('太多次'), '密碼錯 5 次後鎖住，連正確密碼也擋');
delete w.cache['pw-fails'];
w.g.SpreadsheetApp.getActiveSpreadsheet().insertSheet('匯入');
w.sheets['匯入'].rows = [['舊筆記：團體活動的座位安排'], ['已經匯過的', '已匯入 #1'], [''], ['https://example.com/old']];
c.importSheet();
ok(w.sheets['筆記'].rows.length === 10 && w.sheets['匯入'].rows[0][1] === '已匯入 #8' && w.sheets['匯入'].rows[3][1] === '已匯入 #9', '匯入：只處理還沒匯過的列，並標記');
c.importSheet();
ok(w.sheets['筆記'].rows.length === 10, '匯入再跑一次不會重複');

console.log('階段 3：問自己的 AI');
c = load(w, ['Capture', 'Web', 'Ask']);
c.buildIndex();
ok(w.sheets['向量'].rows.length === 10, 'buildIndex 幫 9 則舊筆記都建立索引');
c.buildIndex();
ok(w.sheets['向量'].rows.length === 10, 'buildIndex 再跑一次不會重複');
lineEvent(c, { type: 'text', id: 'm11', text: '運動後要量血糖' });
ok(w.sheets['向量'].rows.length === 11 && w.sheets['向量'].rows[10][0] === 10, '新筆記存入時自動建索引');
s = c.api('pw-1234', 'search', { q: '血糖 飲食' });
ok(s.notes.length >= 3 && s.notes.every(n => n.score > 0), '搜尋加上語意分數');
let r = c.api('pw-1234', 'related', { id: 10 });
ok(r.notes.length > 0 && r.notes[0].id !== 10 && [3, 7, 9].includes(r.notes[0].id), '相關筆記：血糖的筆記找到其他血糖／糖尿病筆記');
a = c.api('pw-1234', 'ask', { q: '手部活動有哪些' });
ok(a.text.includes('[#1]') && a.sources.length === 1 && a.sources[0].id === 1, '問答：回答帶出處，來源只列回答裡用到的筆記');
const askCall = w.sent.gemini.pop();
ok(askCall.body.system_instruction.parts[0].text.includes('只能根據') && askCall.body.contents[0].parts[0].text.includes('[#1]'), '問答：把筆記內容和「只能根據筆記」的規則送給 Gemini');
lineEvent(c, { type: 'text', id: 'm12', text: '?手部活動有哪些' });
ok(lastReply(w).includes('來源：#1'), 'LINE「？問題」會回答並列來源');
ok(w.sheets['筆記'].rows.length === 11, '發問不會被存成筆記');

console.log('階段 4：每週回顧信');
c = load(w, ['Capture', 'Web', 'Ask', 'Weekly']);
w.sheets['筆記'].rows[1][1] = new Date(Date.now() - 30 * 86400e3);   // 把 #1 變成一個月前的舊筆記
w.sheets['筆記'].rows[3][1] = new Date(Date.now() - 20 * 86400e3);
c.weeklyReview();
let mail = w.sent.mail.pop();
ok(mail.to === 'me@example.com' && mail.subject.includes('本週 8 則'), '回顧信寄給自己，主旨有本週則數');
ok(mail.htmlBody.includes('<h3') && mail.htmlBody.includes('這週在想什麼') && mail.htmlBody.includes('翻出來的舊筆記') && mail.htmlBody.includes('#1 '), '信裡有回顧段落與翻出來的舊筆記');
ok(w.sent.line.filter(x => x.url.endsWith('/push')).pop().body.to === 'U-owner', '同時用 LINE 推播給主人');
c.installWeekly(); c.installWeekly();
ok(w.sent.triggers.length === 1 && w.sent.triggers[0].day === 'SUN' && w.sent.triggers[0].hour === 20, 'installWeekly 執行兩次也只有一個週日 20 點的觸發條件');
c.removeWeekly();
ok(w.sent.triggers.length === 0, 'removeWeekly 取消');
const w2 = makeWorld(); const c2 = load(w2, ['Capture', 'Weekly']); c2.setup(); c2.weeklyReview();
ok(w2.sent.mail[0].subject.includes('本週 0 則') && w2.sent.mail[0].htmlBody.includes('這週沒有新筆記'), '沒有筆記、沒做階段 2–3 時回顧信也不會出錯');

console.log('階段 5：產出');
c = load(w, ['Capture', 'Web', 'Ask', 'Weekly', 'Compose']);
h = c.api('pw-1234', 'home');
ok(h.canAsk && h.canCompose && h.kinds.join() === 'FB 貼文,簡報大綱,教案初稿,衛教單張', '全部完成後首頁標示功能開啟，列出成品類型');
ok(c.api('pw-1234', 'compose', { ids: [], kind: 'FB 貼文' }).error.includes('勾選'), '沒勾筆記時提示');
let d = c.api('pw-1234', 'compose', { ids: ['1', '4'], kind: '教案初稿', extra: '對象是志工' });
ok(d.text.startsWith('初稿：') && d.kind === '教案初稿' && d.used.length === 2, '產出初稿');
const compCall = w.sent.gemini.pop().body.contents[0].parts[0].text;
ok(compCall.includes('60 分鐘') && compCall.includes('對象是志工') && compCall.includes('[#4]') && compCall.includes('待補'), '產出：帶入範本、補充要求、筆記內容與「不要自己編」');
ok(w.sheets['產出'].rows.length === 2 && w.sheets['產出'].rows[1][2] === '#1 #4', '初稿存到「產出」工作表');

console.log('金鑰錯誤');
const w3 = makeWorld(); const c3 = load(w3, ['Capture']);
vm.runInContext("GEMINI_API_KEY = 'WRONG'", c3);
let msg = ''; try { c3.askGemini('hi'); } catch (e) { msg = e.message; }
ok(msg.includes('金鑰好像不對') && w3.sent.gemini.length === 1, '金鑰錯誤：說明原因，且不會白白換模型重試');

console.log(`\n通過 ${pass}，失敗 ${fail}`);
process.exit(fail ? 1 : 0);
