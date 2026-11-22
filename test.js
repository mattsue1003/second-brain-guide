// 用假的 Google 服務（試算表、UrlFetch、Gmail…）在本機跑教學裡的 Apps Script 程式碼。
// 驗證的是程式邏輯；真正的 Gemini 與 LINE 連線沒辦法在這裡測。
// 執行：node test.js
const fs = require('fs'), vm = require('vm'), path = require('path');
const GAS = path.join(__dirname, 'src', 'gas');

let pass = 0, fail = 0;
function ok(cond, name) { if (cond) pass++; else { fail++; console.log('  ✗ ' + name); } }

function makeWorld() {
  const sheets = {}, order = [];
  const props = {}, cache = {}, sent = { line: [], mail: [], triggers: [], files: [], gemini: [], events: [] };
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
        const sys = body.system_instruction ? body.system_instruction.parts[0].text : '';
        if (text.includes('換成繁體字')) {   // 假的簡轉繁：只換測試用到的幾個字
          const map = { '这': '這', '个': '個', '动': '動', '对': '對', '长': '長', '说': '說', '话': '話', '时': '時', '间': '間', '记': '記', '录': '錄', '请': '請', '会': '會', '议': '議', '医': '醫', '学': '學', '练': '練', '习': '習' };
          return gen(world.convertBroken ? '' : text.split('\n\n').slice(1).join('\n\n').replace(/./g, ch => map[ch] || ch));
        }
        if (world.simplifiedOnce) { const t = world.simplifiedOnce; world.simplifiedOnce = null; return gen(t); }
        if (sys.includes('私人秘書')) return gen(world.plan ? world.plan(text) : '{"action":"note"}');
        if (parts.some(p => p.inline_data)) return gen(world.audioText || (parts.find(p => p.inline_data).inline_data.mime_type.startsWith('audio') ? '今天帶長輩做音樂律動，大家很投入' : '手指操步驟：握拳、張開、對指'));
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
    Session: { getEffectiveUser: () => ({ getEmail: () => 'me@example.com' }), getScriptTimeZone: () => 'Asia/Taipei' },
    CalendarApp: { getDefaultCalendar: () => ({
      createEvent(title, start, end, opt) { sent.events.push({ title, start, end, opt }); },
      createAllDayEvent(title, start, opt) { sent.events.push({ title, start, allDay: true, opt }); },
      getEvents: (from, to) => sent.events.filter(e => e.start >= from && e.start < to).map(e => ({ getStartTime: () => e.start, isAllDayEvent: () => !!e.allDay, getTitle: () => e.title }))
    }) },
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

console.log('階段 6：AI 秘書');
const ALL = ['Capture', 'Web', 'Ask', 'Weekly', 'Compose', 'Secretary'];
// 合併成同一個專案時，不同檔案不能有同名的函式或全域變數
const seen = {}, clash = [];
ALL.forEach(f => (fs.readFileSync(path.join(GAS, f + '.gs'), 'utf8').match(/^(?:function|var)\s+[\w$]+/gm) || []).forEach(d => {
  const n = d.split(/\s+/)[1]; if (seen[n] && seen[n] !== f) clash.push(n + '（' + seen[n] + '、' + f + '）'); seen[n] = f; }));
ok(clash.length === 0, '六個檔案之間沒有同名的函式或變數：' + clash.join('、'));
ok(Object.keys(seen).filter(n => n === 'doPost' || n === 'doGet').length === 2, '整個專案只有一個 doPost 和一個 doGet');
ok(!/GEMINI_API_KEY\s*=|LINE_TOKEN\s*=|ALLOWED_USERS|SpreadsheetApp\.create|openById/.test(fs.readFileSync(path.join(GAS, 'Secretary.gs'), 'utf8')), '秘書檔不另外放金鑰、不另開試算表、不用允許名單');

c = load(w, ALL);
const p2 = n => String(n).padStart(2, '0');
const dstr = off => { const t = new Date(); t.setDate(t.getDate() + off); return t.getFullYear() + '-' + p2(t.getMonth() + 1) + '-' + p2(t.getDate()); };
const md = off => { const t = new Date(); t.setDate(t.getDate() + off); return (t.getMonth() + 1) + '/' + t.getDate(); };
const notesN = () => w.sheets['筆記'].rows.length, gemN = () => w.sent.gemini.length;
w.plan = text => {
  if (text.includes('開會')) return JSON.stringify({ action: 'add_event', title: '跟王主任開會', start: dstr(1) + 'T15:00', location: '三樓會議室' });
  if (text.includes('保養')) return '```json\n' + JSON.stringify({ action: 'add_event', title: '車子保養', start: dstr(3) }) + '\n```';
  if (text.includes('有什麼事')) return JSON.stringify({ action: 'list_events', from: dstr(1), to: dstr(1) });
  if (text.includes('交報告')) return JSON.stringify({ action: 'add_todo', content: '交報告', due: dstr(4) });
  if (text.includes('報告交了')) return JSON.stringify({ action: 'done_todo', keyword: '報告' });
  if (text.includes('午餐')) return JSON.stringify({ action: 'add_expense', type: '支出', category: '餐飲', item: '午餐', amount: 120, date: '' });
  if (text.includes('薪水')) return JSON.stringify({ action: 'add_expense', type: '收入', category: '薪資', item: '薪水', amount: '45,000 元' });
  if (text.includes('亂回')) return '我不知道';
  return '{"action":"note"}';
};
let n0 = notesN(), g0 = gemN();
c.testSecretary();
ok(w.sheets['待辦'].rows[0][1] === '事項' && w.sheets['記帳'].rows[0][5] === '金額', 'testSecretary 在同一個試算表建立「待辦」「記帳」工作表');

lineEvent(c, { type: 'text', id: 's1', text: '明天下午三點跟王主任開會，地點在三樓會議室' });
let evs = w.sent.events;
ok(evs.length === 1 && evs[0].title === '跟王主任開會' && evs[0].start.getHours() === 15 && evs[0].end - evs[0].start === 3600e3 && evs[0].opt.location === '三樓會議室', '行程：寫進行事曆，沒說結束時間就當一小時');
ok(lastReply(w).includes('已記到行事曆') && lastReply(w).includes(md(1)) && lastReply(w).includes('15:00–16:00') && lastReply(w).includes('地點：三樓會議室'), '行程：回覆實際寫入的日期、時間、地點');
ok(notesN() === n0, '行程不會被存成筆記');
lineEvent(c, { type: 'text', id: 's2', text: '大後天車子要保養' });
ok(evs.length === 2 && evs[1].allDay && lastReply(w).includes('（整天）'), '只有日期的行程存成整天活動（AI 回覆包在 ```json 裡也讀得懂）');
lineEvent(c, { type: 'text', id: 's3', text: '我明天有什麼事' });
ok(lastReply(w).includes('15:00　跟王主任開會') && !lastReply(w).includes('車子保養'), '查行事曆：只列出那一天的行程');
g0 = gemN();
lineEvent(c, { type: 'text', id: 's4', text: '行程' });
ok(lastReply(w).includes('跟王主任開會') && lastReply(w).includes('整天　車子保養') && gemN() === g0, '傳「行程」列出未來七天，不用問 AI');

g0 = gemN();
lineEvent(c, { type: 'text', id: 's5', text: '待辦：打電話給王主任' });
ok(w.sheets['待辦'].rows.length === 2 && w.sheets['待辦'].rows[1][1] === '打電話給王主任' && w.sheets['待辦'].rows[1][3] === '未完成' && gemN() === g0, '「待辦：…」直接加入，不用問 AI');
lineEvent(c, { type: 'text', id: 's6', text: '記得週五前交報告' });
ok(w.sheets['待辦'].rows[2][2] instanceof Date && lastReply(w).includes('期限 ' + md(4)) && lastReply(w).includes('還有 2 件'), 'AI 判斷成待辦，帶期限');
lineEvent(c, { type: 'text', id: 's7', text: '待辦' });
ok(lastReply(w).includes('1. 打電話給王主任') && lastReply(w).includes('2. 交報告（期限'), '傳「待辦」看清單');
lineEvent(c, { type: 'text', id: 's8', text: '完成 1' });
ok(w.sheets['待辦'].rows[1][3] === '完成' && w.sheets['待辦'].rows[1][4] instanceof Date && lastReply(w).includes('已完成：打電話給王主任') && lastReply(w).includes('還有 1 件'), '「完成 1」把第一件標成完成');
lineEvent(c, { type: 'text', id: 's9', text: '報告交了' });
ok(w.sheets['待辦'].rows[2][3] === '完成', '用關鍵字完成待辦');
lineEvent(c, { type: 'text', id: 's10', text: '完成 5' });
ok(lastReply(w).includes('目前沒有待辦事項'), '沒有待辦時的回覆');

lineEvent(c, { type: 'text', id: 's11', text: '午餐 120' });
let ex = w.sheets['記帳'].rows;
ok(ex.length === 2 && ex[1][2] === '支出' && ex[1][3] === '餐飲' && ex[1][5] === 120 && lastReply(w).includes('支出累計 120 元'), '記帳：AI 分類後寫進「記帳」工作表並回報累計');
lineEvent(c, { type: 'text', id: 's12', text: '領薪水 45000' });
ok(ex[2][2] === '收入' && ex[2][5] === 45000 && lastReply(w).includes('45,000 元'), '收入：金額字串裡的逗號和「元」會被去掉');
g0 = gemN();
lineEvent(c, { type: 'text', id: 's13', text: '查帳' });
ok(lastReply(w).includes('（2 筆）') && lastReply(w).includes('支出 120 元') && lastReply(w).includes('收入 45,000 元') && lastReply(w).includes('餐飲 120 元') && gemN() === g0, '「查帳」回本月統計，不用問 AI');
w.geminiDown = true;
lineEvent(c, { type: 'text', id: 's14', text: '記帳 加油 1250' });
ok(ex.length === 4 && ex[3][4] === '加油' && ex[3][5] === 1250 && ex[3][3] === '未分類', 'AI 忙線時「記帳 項目 金額」還是記得進去');
n0 = notesN();
lineEvent(c, { type: 'text', id: 's15', text: 'AI 忙線時一般句子照樣存成筆記 手' });
ok(notesN() === n0 + 1 && lastReply(w).includes('已存入'), 'AI 忙線時，一般句子照樣存成筆記');
lineEvent(c, { type: 'text', id: 's16', text: '行程：後天開會' });
ok(lastReply(w).includes('AI 暫時無法處理') && notesN() === n0 + 1, 'AI 忙線時，「行程：」開頭的句子回報錯誤、不會誤存成筆記');
w.geminiDown = false;

n0 = notesN(); g0 = gemN();
lineEvent(c, { type: 'text', id: 's17', text: '手部精細動作可以用夾豆子練習' });
ok(notesN() === n0 + 1 && lastReply(w).includes('已存入 #') && gemN() === g0 + 2, '一般句子：AI 判斷不是秘書的事，照舊整理後存成筆記');
lineEvent(c, { type: 'text', id: 's18', text: 'AI 亂回的時候 手' });
ok(notesN() === n0 + 2, 'AI 沒照格式回時當成筆記');
g0 = gemN();
lineEvent(c, { type: 'text', id: 's19', text: '筆記：下次開會要帶手指操講義' });
ok(notesN() === n0 + 3 && w.sheets['筆記'].rows[notesN() - 1][7] === '下次開會要帶手指操講義' && gemN() === g0 + 1 && evs.length === 2, '「筆記：」開頭一定存成筆記（就算內容像行程），而且不多問一次 AI');
g0 = gemN();
lineEvent(c, { type: 'text', id: 's20', text: '開會資料 https://example.com/meeting' });
ok(w.sheets['筆記'].rows[notesN() - 1][2] === '連結' && gemN() === g0 + 1 && evs.length === 2, '有網址的訊息照舊存成連結筆記');
g0 = gemN(); n0 = notesN();
lineEvent(c, { type: 'text', id: 's21', text: '找 糖尿病' });
lineEvent(c, { type: 'text', id: 's22', text: '？開會要注意什麼' });
ok(notesN() === n0 && evs.length === 2 && lastReply(w).includes('根據筆記'), '「找」和「？」維持原本的功能，不會被秘書攔走');

lineEvent(c, { type: 'file', id: 's23', fileName: '講義.pdf' });
rows = w.sheets['筆記'].rows;
ok(rows[rows.length - 1][2] === '檔案' && rows[rows.length - 1][3] === '講義.pdf' && rows[rows.length - 1][8] === 'https://drive.google.com/file/講義.pdf' && lastReply(w).includes('講義.pdf'), '檔案：存到雲端硬碟，並在「筆記」留一則');
lineEvent(c, { type: 'video', id: 's24' });
ok(/^影片_\d{8}_\d{6}\.mp4$/.test(w.sent.files[w.sent.files.length - 1]), '影片：沒有檔名時自動命名');
lineEvent(c, { type: 'sticker', id: 's25' });
ok(lastReply(w).includes('只收文字'), '貼圖還是不收');

n0 = notesN(); let t0 = w.sheets['待辦'].rows.length;
lineEvent(c, { type: 'text', id: 's26', text: '待辦：偷加一筆' }, 'U-stranger');
lineEvent(c, { type: 'text', id: 's27', text: '明天開會' }, 'U-stranger');
ok(w.sheets['待辦'].rows.length === t0 && evs.length === 2 && lastReply(w).includes('私人的筆記機器人'), '不是主人的人不能用秘書功能');

vm.runInContext('SECRETARY_AUTO = false', c);
g0 = gemN(); n0 = notesN();
lineEvent(c, { type: 'text', id: 's28', text: '後天下午開會討論教案 手' });
ok(evs.length === 2 && notesN() === n0 + 1 && gemN() === g0 + 1, 'SECRETARY_AUTO 關閉：沒有指定開頭的句子直接存成筆記，只呼叫 AI 一次');
lineEvent(c, { type: 'text', id: 's29', text: '行程：後天下午開會' });
ok(evs.length === 3, 'SECRETARY_AUTO 關閉：「行程：」開頭還是會記到行事曆');
lineEvent(c, { type: 'text', id: 's30', text: '記帳 午餐 90' });
ok(w.sheets['記帳'].rows.length === 5 && w.sheets['記帳'].rows[4][5] === 120, 'SECRETARY_AUTO 關閉：「記帳 …」開頭還是會交給 AI 分類');

const w4 = makeWorld(); const c4 = load(w4, ['Capture']); c4.setup();
lineEvent(c4, { type: 'file', id: 'f1', fileName: 'a.pdf' });
ok(lastReply(w4).includes('只收文字') && w4.sent.files.length === 0, '還沒做階段 6 時，檔案訊息照舊說明不支援');
lineEvent(c4, { type: 'text', id: 'f2', text: '筆記：只做階段 1 也能用 手' });
ok(w4.sheets['筆記'].rows[1][7] === '只做階段 1 也能用 手' && !w4.sheets['待辦'], '還沒做階段 6 時，文字照舊存成筆記');

console.log('繁體中文');
const SIMP = c.SIMPLIFIED;
ok(w.sent.gemini.length > 40 && w.sent.gemini.every(x => x.body.system_instruction && x.body.system_instruction.parts[0].text.includes('一律使用繁體中文（台灣用語）') && x.body.system_instruction.parts[0].text.includes('不可以出現簡體字')), '六個階段每一次呼叫 Gemini 都帶「一律使用繁體中文」的規則');
ALL.concat(['Index']).forEach(f => { const src = fs.readFileSync(path.join(GAS, f + (f === 'Index' ? '.html' : '.gs')), 'utf8'); ok(!SIMP.test(src), f + ' 的程式碼和提示詞裡沒有簡體字'); });
ok(c.hasSimplified('这个活动对长辈很好') && !c.hasSimplified('這個活動對長輩很好，在台灣的社區據點秘書群組') , '簡體字偵測：簡體抓得到，繁體不會誤判');
const w5 = makeWorld(); const c5 = load(w5, ALL); c5.setup();
w5.audioText = '这个活动对长者很好';
lineEvent(c5, { type: 'audio', id: 't1' });
ok(w5.sheets['筆記'].rows[1][7] === '這個活動對長者很好' && !SIMP.test(w5.sheets['筆記'].rows[1].join('')), '語音逐字稿是簡體時，自動轉成繁體才存');
w5.audioText = null;
w5.simplifiedOnce = '```json\n{"title":"说话的时间","summary":"记录会议","tags":["医学"],"category":"其他"}\n```';
lineEvent(c5, { type: 'text', id: 't2', text: '筆記：開會心得' });
ok(w5.sheets['筆記'].rows[2][3] === '說話的時間' && w5.sheets['筆記'].rows[2][4] === '記錄會議' && w5.sheets['筆記'].rows[2][5] === '醫學' && !SIMP.test(lastReply(w5)), 'AI 整理的標題、摘要、標籤是簡體時轉成繁體（JSON 結構不會壞），LINE 回覆也沒有簡體');
let g5 = w5.sent.gemini.length;
lineEvent(c5, { type: 'text', id: 't3', text: '筆記：全繁體的句子 手' });
ok(w5.sent.gemini.length === g5 + 1, '沒有簡體字時不會多呼叫一次 AI');
w5.g.UrlFetchApp.fetch = (orig => (url, opt) => url.startsWith('https://cn.example/') ? { getResponseCode: () => 200, getContentText: () => '<title>练习</title><p>请记录时间</p>' } : orig(url, opt))(w5.g.UrlFetchApp.fetch);
const c6 = load(w5, ALL);
lineEvent(c6, { type: 'text', id: 't4', text: 'https://cn.example/a' });
ok(w5.sheets['筆記'].rows[4][7].includes('網頁標題：練習') && w5.sheets['筆記'].rows[4][7].includes('請記錄時間') && !SIMP.test(w5.sheets['筆記'].rows[4][7]), '簡體網站的內容轉成繁體才存進原文');
w5.sheets['筆記'].rows.push([5, new Date(), '文字', '旧的说话', '摘要', '', '其他', '这是旧笔记', '']);
c6.fixSimplified();
ok(w5.sheets['筆記'].rows[5][3] === '旧的說話' && w5.sheets['筆記'].rows[5][7] === '這是旧笔記', 'fixSimplified 把已經存進去的舊筆記轉成繁體');
w5.convertBroken = true; w5.audioText = '这个';
lineEvent(c6, { type: 'audio', id: 't5' });
ok(w5.sheets['筆記'].rows.length === 7 && lastReply(w5).includes('已存入'), '轉繁體失敗時筆記照樣存入，不會出錯');

console.log('金鑰錯誤');
const w3 = makeWorld(); const c3 = load(w3, ['Capture']);
vm.runInContext("GEMINI_API_KEY = 'WRONG'", c3);
let msg = ''; try { c3.askGemini('hi'); } catch (e) { msg = e.message; }
ok(msg.includes('金鑰好像不對') && w3.sent.gemini.length === 1, '金鑰錯誤：說明原因，且不會白白換模型重試');
ok(c3.explainError({ code: 402, message: 'prepayment credits are depleted' }).includes('儲值金用完了'), '402：說明是預付儲值金用完');

console.log(`\n通過 ${pass}，失敗 ${fail}`);
process.exit(fail ? 1 : 0);
