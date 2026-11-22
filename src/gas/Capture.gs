// ===== 第二大腦｜階段 1：核心功能＋LINE 隨手記 =====
// 這個檔案不用改。金鑰與個人設定都放在「設定.gs」。

var SHEET_NOTES = '筆記';
var HEADERS = ['編號', '時間', '類型', '標題', '摘要', '標籤', '類別', '原文', '連結'];

// ---- 第一次使用：建立「筆記」工作表，並測試 Gemini 金鑰 ----
function setup() {
  notesSheet();
  Logger.log('「筆記」工作表準備好了');
  Logger.log('Gemini 測試：' + askGemini('請用一句話說「第二大腦準備好了」'));
}

function book() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function notesSheet() {
  var ss = book();
  var sh = ss.getSheetByName(SHEET_NOTES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NOTES, 0);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(4, 220);
    sh.setColumnWidth(5, 380);
    sh.setColumnWidth(8, 300);
  }
  return sh;
}

// ---- 存一則筆記，回傳編號 ----
function saveNote(n) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var id;
  try {
    var sh = notesSheet();
    id = sh.getLastRow();   // 第 1 列是標題，所以第一則筆記的編號是 1
    sh.appendRow([id, new Date(), n.type || '文字', n.title || '', n.summary || '',
      (n.tags || []).join('、'), n.category || '其他', String(n.text || '').slice(0, 45000), n.link || '']);
  } finally {
    lock.releaseLock();
  }
  n.id = id;
  // 階段 3 完成後，新筆記會自動建立「語意索引」
  if (typeof indexNote === 'function') {
    try { indexNote(n); } catch (err) { Logger.log('建立索引失敗：' + err.message); }
  }
  return id;
}

// ---- 讀出全部筆記 ----
function readNotes() {
  var sh = notesSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues().filter(function (r) {
    return r[0] !== '';
  }).map(function (r) {
    var t = r[1] instanceof Date ? r[1] : new Date(r[1]);
    return {
      id: Number(r[0]), ms: t.getTime(), date: Utilities.formatDate(t, 'Asia/Taipei', 'yyyy/MM/dd'),
      type: String(r[2]), title: String(r[3]), summary: String(r[4]),
      tags: String(r[5]).split(/[、,，\s]+/).filter(String),
      category: String(r[6]), text: String(r[7]), link: String(r[8])
    };
  });
}

// ---- 請 Gemini 幫一段內容下標題、摘要、標籤、類別 ----
function digest(text, kind) {
  var prompt = '下面是一則我要存進個人筆記庫的' + (kind || '文字') + '。請只輸出一個 JSON 物件，不要有其他文字，格式：\n' +
    '{"title":"15 字內的標題","summary":"60 字內的重點摘要","tags":["2 到 4 個標籤"],"category":"類別"}\n' +
    'category 只能從這些選一個：' + CATEGORIES.join('、') + '\n' +
    '全部用繁體中文（台灣用語）。標籤用名詞，不要加 # 號。\n\n內容：\n' + String(text).slice(0, 8000);
  var out = {};
  try {
    var raw = askGemini(prompt, '你是筆記整理助手，只輸出 JSON。');
    out = JSON.parse(raw.slice(raw.indexOf('{'), raw.lastIndexOf('}') + 1));
  } catch (err) {
    Logger.log('整理失敗，改用原文開頭當標題：' + err.message);
  }
  var plain = String(text).replace(/\s+/g, ' ').trim();
  return {
    title: String(out.title || plain.slice(0, 20)),
    summary: String(out.summary || plain.slice(0, 60)),
    tags: Array.isArray(out.tags) ? out.tags.map(String).slice(0, 5) : [],
    category: CATEGORIES.indexOf(out.category) >= 0 ? out.category : '其他'
  };
}

// ---- 收一段文字（或網址），整理後存起來 ----
function captureText(text) {
  text = String(text).trim();
  var url = (text.match(/https?:\/\/[^\s]+/) || [])[0];
  var note;
  if (url) {
    var page = fetchPage(url);
    var memo = text.replace(url, '').trim();
    var body = (memo ? '我的備註：' + memo + '\n' : '') + (page.title ? '網頁標題：' + page.title + '\n' : '') + page.text;
    note = digest(body || url, '網頁');
    note.type = '連結';
    note.text = body || url;
    note.link = url;
  } else {
    note = digest(text, '文字');
    note.type = '文字';
    note.text = text;
  }
  saveNote(note);
  return note;
}

// ---- 抓網頁內容，去掉 HTML 標籤 ----
function fetchPage(url) {
  try {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return { title: '', text: '' };
    var html = res.getContentText();
    var title = (html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || '';
    var text = html.replace(/<script[\s\S]*?<\/script>/gi, ' ').replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/\s+/g, ' ').trim();
    return { title: title.replace(/\s+/g, ' ').trim(), text: text.slice(0, 8000) };
  } catch (err) {
    return { title: '', text: '' };
  }
}

// ---- 收一張照片或一段錄音：先請 Gemini 轉成文字，再整理 ----
function captureMedia(blob, kind) {
  var bytes = blob.getBytes();
  if (bytes.length > 15 * 1024 * 1024) throw new Error('檔案超過 15MB，太大了');
  var mime = kind === '錄音' ? 'audio/m4a' : (blob.getContentType() || 'image/jpeg');
  var ask = kind === '錄音'
    ? '請把這段錄音逐字轉成繁體中文文字，只輸出逐字稿。'
    : '請完整抄錄這張圖片裡的文字；如果沒有文字，就用三句話描述圖片內容。只輸出結果，用繁體中文。';
  var text = askGemini([{ text: ask }, { inline_data: { mime_type: mime, data: Utilities.base64Encode(bytes) } }]);
  var note = digest(text, kind);
  note.type = kind;
  note.text = text;
  if (SAVE_FILES) {
    try {
      var name = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss') + (kind === '錄音' ? '.m4a' : '.jpg');
      note.link = attachFolder().createFile(blob.setName(name)).getUrl();
    } catch (err) {
      Logger.log('附件沒有存成功：' + err.message);
    }
  }
  saveNote(note);
  return note;
}

// ---- 雲端硬碟的「第二大腦附件」資料夾，沒有就建立 ----
function attachFolder() {
  var it = DriveApp.getFoldersByName('第二大腦附件');
  return it.hasNext() ? it.next() : DriveApp.createFolder('第二大腦附件');
}

// ================= LINE =================

// LINE 傳訊息進來時會執行這裡。最後回傳 HtmlService，LINE 才會收到「成功」。
function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  (body.events || []).forEach(function (ev) {
    if (ev.type !== 'message') return;
    if (isDuplicate(ev.webhookEventId)) return;
    var answer;
    try {
      answer = handleEvent(ev);
    } catch (err) {
      answer = '程式出錯了：' + err.message;
    }
    if (answer) lineReply(ev.replyToken, answer);
  });
  return HtmlService.createHtmlOutput('OK');
}

function handleEvent(ev) {
  var uid = ev.source && ev.source.userId;
  var props = PropertiesService.getScriptProperties();
  var owner = props.getProperty('OWNER');
  var hello = '';
  if (!owner) {
    props.setProperty('OWNER', uid);   // 第一個傳訊息的人成為主人
    hello = '（已把你設為這個第二大腦的主人，其他人傳訊息不會被記錄）\n\n';
  } else if (owner !== uid) {
    return '這是私人的筆記機器人，不接受其他人的訊息。';
  }
  lineLoading(uid);
  var m = ev.message;
  if (m.type === 'text') return hello + handleText(m.text);
  if (m.type === 'image') return hello + savedMsg(captureMedia(lineContent(m.id), '照片'));
  if (m.type === 'audio') return hello + savedMsg(captureMedia(lineContent(m.id), '錄音'));
  // 階段 6 完成後，影片和檔案也會存到雲端硬碟
  if ((m.type === 'video' || m.type === 'file') && typeof saveAttachment === 'function') return hello + saveAttachment(m);
  return hello + '目前只收文字、網址、照片和語音訊息。';
}

// 文字訊息：開頭是「？」＝發問；開頭是「找」＝搜尋；其他＝存成筆記
// 階段 6 完成後，行程、待辦、記帳的句子會先交給「秘書」處理
function handleText(text) {
  text = String(text).trim();
  var q = text.match(/^[?？]\s*([\s\S]+)/);
  if (q) {
    if (typeof answerQuestion !== 'function') return '問答功能要完成「階段 3」才會開啟。';
    var a = answerQuestion(q[1]);
    return a.text + (a.sources.length ? '\n\n來源：' + a.sources.map(function (s) { return '#' + s.id + ' ' + s.title; }).join('；') : '');
  }
  var f = text.match(/^找\s+([\s\S]+)/);
  if (f) {
    if (typeof searchNotes !== 'function') return '搜尋功能要完成「階段 2」才會開啟。';
    var hits = searchNotes(f[1], {}).slice(0, 5);
    if (!hits.length) return '找不到和「' + f[1] + '」有關的筆記。';
    return '找到 ' + hits.length + ' 則：\n' + hits.map(function (n) {
      return '#' + n.id + '【' + n.title + '】' + n.date + '\n' + n.summary;
    }).join('\n\n');
  }
  if (typeof secretaryText === 'function') {
    var done = secretaryText(text);   // 回傳空字串代表「這是筆記」
    if (done) return done;
  }
  return savedMsg(captureText(text.replace(/^(?:筆記|記一下)\s*[:：\s]\s*/, '') || text));
}

function savedMsg(n) {
  return '已存入 #' + n.id + '｜' + n.category + '\n【' + n.title + '】\n' + n.summary +
    (n.tags.length ? '\n標籤：' + n.tags.join('、') : '');
}

function lineFetch(url, payload) {
  return UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
}
function lineReply(token, text) {
  lineFetch('https://api.line.me/v2/bot/message/reply', { replyToken: token, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] });
}
// 主動傳訊息給主人（每週回顧會用到；免費方案每月 200 則）
function linePush(text) {
  var owner = PropertiesService.getScriptProperties().getProperty('OWNER');
  if (owner) lineFetch('https://api.line.me/v2/bot/message/push', { to: owner, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] });
}
function lineLoading(uid) {
  if (uid) lineFetch('https://api.line.me/v2/bot/chat/loading/start', { chatId: uid, loadingSeconds: 40 });
}
function lineContent(messageId) {
  var res = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/message/' + messageId + '/content', {
    headers: { Authorization: 'Bearer ' + LINE_TOKEN }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('沒辦法從 LINE 下載檔案（' + res.getResponseCode() + '）');
  return res.getBlob();
}
function isDuplicate(id) {
  if (!id) return false;
  var cache = CacheService.getScriptCache();
  if (cache.get(id)) return true;
  cache.put(id, '1', 600);
  return false;
}
// 想換主人（例如換了 LINE 帳號）：執行這個，再傳一則訊息
function resetOwner() {
  PropertiesService.getScriptProperties().deleteProperty('OWNER');
  Logger.log('已清除主人，下一個傳訊息的人會成為新主人');
}

// ================= Gemini =================

// input 可以是一句話，也可以是 parts 陣列（含圖片或錄音）
function askGemini(input, system) {
  var parts = typeof input === 'string' ? [{ text: input }] : input;
  var hasMedia = parts.some(function (p) { return p.inline_data; });
  var models = [MODEL].concat(BACKUP_MODELS.filter(function (m) { return m !== MODEL; }));
  var start = Date.now(), data = {};
  for (var i = 0; i < models.length; i++) {
    data = callGemini(models[i], parts, system, true);
    if (data.error && data.error.code === 400 && /think/i.test(data.error.message)) data = callGemini(models[i], parts, system, false);
    if (!data.error) break;
    var retry = [404, 429, 500, 502, 503].indexOf(data.error.code) >= 0 || (hasMedia && data.error.code === 400);
    if (!retry) break;                        // 金鑰錯誤等問題，換模型也沒用
    if (Date.now() - start > 40000) break;    // 等太久就不再換
  }
  var answer = '';
  try {
    answer = data.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('').trim();
  } catch (err) {}
  if (answer) return answer;
  throw new Error(explainError(data.error));
}

function callGemini(model, parts, system, fast) {
  var body = { contents: [{ role: 'user', parts: parts }] };
  if (system) body.system_instruction = { parts: [{ text: system }] };
  // lite 模型幾乎不思考；其他模型預設會先想一陣子，調低才回得快
  if (fast && model.indexOf('lite') < 0) body.generationConfig = { thinkingConfig: { thinkingLevel: 'low' } };
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-goog-api-key': GEMINI_API_KEY },
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  try {
    return JSON.parse(res.getContentText());
  } catch (err) {
    return { error: { code: res.getResponseCode(), message: 'Gemini 沒有回傳正確的資料' } };
  }
}

function explainError(e) {
  if (!e) return 'AI 這次沒有給出答案，請再試一次。';
  var tips = {
    400: '金鑰好像不對，或這個模型不接受這種內容。',
    402: '這把金鑰的專案開了付費（預付制），但儲值金用完了。請到 AI Studio 建立一個沒有綁帳單的新專案、換一把新金鑰，或是去儲值。',
    403: '這把金鑰沒有權限，請到 AI Studio 重新建立。',
    404: '模型名稱已停用，請到「設定.gs」把 MODEL 換成新的名稱。',
    429: '免費額度用完了，或問得太快，請等一下再試。',
    503: 'Google 的 AI 現在太多人使用，請過一兩分鐘再試。'
  };
  return 'AI 暫時無法處理：' + (tips[e.code] || '') + '（' + e.code + ' ' + e.message + '）';
}
