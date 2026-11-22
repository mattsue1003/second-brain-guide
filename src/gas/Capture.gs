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
    page.title = toTraditional(page.title);   // 簡體網站的內容先轉成繁體再存
    page.text = toTraditional(page.text);
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
    ? '請把這段錄音逐字轉成文字，只輸出逐字稿。逐字稿一律用繁體中文（台灣用語）書寫，不可以出現簡體字。'
    : '請完整抄錄這張圖片裡的文字；如果沒有文字，就用三句話描述圖片內容。只輸出結果，一律用繁體中文（台灣用語），圖片裡是簡體字也要轉成繁體字。';
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
// 送出前再檢查一次：不管文字是 AI 寫的、舊筆記裡的，還是從語音轉出來的，LINE 上看到的一律是繁體
function lineReply(token, text) {
  text = toTraditional(text);
  lineFetch('https://api.line.me/v2/bot/message/reply', { replyToken: token, messages: [{ type: 'text', text: String(text).slice(0, 4900) }] });
}
// 主動傳訊息給主人（每週回顧會用到；免費方案每月 200 則）
function linePush(text) {
  var owner = PropertiesService.getScriptProperties().getProperty('OWNER');
  if (owner) lineFetch('https://api.line.me/v2/bot/message/push', { to: owner, messages: [{ type: 'text', text: toTraditional(text).slice(0, 4900) }] });
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

// 每一次呼叫 Gemini 都會帶上這條規則，所以六個階段的回覆、摘要、逐字稿都是繁體中文
var TW_RULE = '一律使用繁體中文（台灣用語）。就算輸入的文字、圖片或錄音是簡體中文，輸出也要轉成繁體中文，不可以出現簡體字。';
// 只有簡體中文才會用到的常用字，約 1200 個（用編碼寫，這份程式碼本身才不會出現簡體字）。這一行很長，不用看也不用改
var SIMPLIFIED = /[\u4e13\u4e1a-\u4e1d\u4e22\u4e24\u4e25\u4e27\u4e2a\u4e34\u4e3a\u4e3d\u4e3e\u4e49\u4e4c\u4e50\u4e54\u4e60\u4e61\u4e66\u4e70\u4e71\u4e89\u4e8f\u4e9a\u4ea7\u4ea9\u4eb2\u4ebf\u4ec5\u4ece\u4ed1\u4ed3\u4eea\u4eec\u4f17\u4f1a\u4f1e-\u4f20\u4f24\u4f26\u4f2a\u4fa0\u4fa3\u4fa5-\u4fa9\u4fe9\u4fed\u503a\u503e\u507f\u50a8\u5151\u5170\u5173\u5174\u5179\u517b\u517d\u5185\u5188\u518c\u5199\u519b\u519c\u51af\u51b2\u51b3\u51b5\u51bb\u51c0\u51c9\u51cf\u51d1\u51db\u51e4\u51ed\u51ef\u51fb\u51ff\u5218-\u521b\u5220\u522b\u5239\u523d\u5242\u5250\u5251\u5265\u5267\u529d\u529e\u52a1\u52a8\u52b1-\u52b3\u52bf\u52cb\u5300\u533a\u533b\u534e\u534f\u5355\u5356\u5362\u5364\u5367\u536b\u5374\u5385\u5386\u5389\u538b\u538c\u5395\u53a2\u53a6\u53a8\u53a9\u53bf\u53c1\u53c2\u53cc\u53d1\u53d8\u53d9\u53e0\u53f6\u53f7\u53f9\u5413\u5415\u5417\u542f\u5434\u5450\u5455\u5458\u545b\u545c\u548f\u5499\u54cd\u54d1\u54d7\u54df\u5524\u556e\u5578\u55b7\u5618\u5631\u56a3\u56e2\u56ed\u56f1\u56f4\u56fd\u56fe\u5706\u573a\u5757\u575a\u575b\u575d-\u5760\u5784\u5792\u57a6\u57ab\u5811\u5815\u5899\u58ee\u58f0\u58f3\u58f6\u5904\u5907\u591f\u5934\u5939\u593a\u594b\u5956\u5965\u5986-\u5988\u5a04\u5a07\u5a31\u5a74\u5a76\u5b59\u5b66\u5b6a\u5b9d\u5b9e\u5ba0\u5ba1\u5baa\u5bab\u5bbd\u5bbe\u5bdd\u5bf9\u5bfb\u5bfc\u5bff\u5c06\u5c14\u5c18\u5c1d\u5c27\u5c3d\u5c42\u5c49\u5c4a\u5c5e\u5c61\u5c7f\u5c81\u5c82\u5c97\u5c9b\u5cbf\u5ce1\u5ce6\u5d2d\u5de9\u5e01\u5e05\u5e08\u5e10\u5e1c\u5e26\u5e27\u5e2e\u5e42\u5e7f\u5e86\u5e90\u5e93\u5e94\u5e99\u5e9e\u5e9f\u5f00\u5f03\u5f20\u5f25\u5f2f\u5f39\u5f3a\u5f52\u5f53\u5f55\u5f66\u5f7b\u5f84\u5fc6\u5fe7\u6001\u6002\u603b\u604b\u6052\u6073\u6076\u607c\u60a6\u60ac\u60af\u60e7-\u60e9\u60eb\u60ed-\u60ef\u6124\u6151\u61d2\u620f\u6218\u6237\u6267\u6269\u626b\u626c\u629a\u629b\u62a0-\u62a2\u62a4\u62a5\u62c5\u62df\u62e2\u62e3\u62e5-\u62e9\u631a\u631b\u631d-\u6321\u6323-\u6325\u635e\u635f\u6361-\u6363\u63b3\u63b7\u63b8\u63ba\u63fd\u6400-\u6402\u6405\u643a\u6444\u6446-\u6448\u644a\u6491\u64b5\u64de\u6512\u654c\u655b\u6570\u658b\u65a9\u65ad\u65e0\u65e7\u65f6\u65f7\u663c\u663e\u664b\u6653\u6655\u6682\u672f\u6740\u6742\u6743\u6761\u6765\u6768\u67a2\u67a3\u67aa\u67ab\u67e0\u6805\u6807\u6808\u680b\u680f\u6811\u6837\u6863\u6865\u6868\u6869\u68a6\u68c0\u692d\u697c\u69db\u6a2a\u6a31\u6a71\u6b22\u6b27\u6b7c\u6b8b\u6bb4\u6bc1\u6bd5\u6bd9\u6be1\u6c22\u6c47\u6c49\u6c64\u6c79\u6c9f\u6ca1\u6ca4-\u6ca7\u6caa\u6cea\u6cfb-\u6cfd\u6d45-\u6d47\u6d4a\u6d4b\u6d4e\u6d51\u6d53\u6d9b\u6d9d\u6d9f\u6da1\u6da3\u6da4\u6da6-\u6da9\u6e0a\u6e0d\u6e10\u6e14\u6e17\u6e29\u6e7e\u6e7f\u6e83\u6e85\u6eda\u6ede\u6ee1\u6ee4-\u6ee6\u6ee8\u6ee9\u6f4d\u6f5c\u6f9c\u6fd2\u706d\u706f\u7075\u707e\u707f\u7089\u70b9\u70bc\u70bd\u70c1-\u70c3\u70db\u70df\u70e6\u70e7\u70e9\u70eb-\u70ed\u7115\u7231\u7237\u7275\u727a\u728a\u72b6\u72b9\u72c8\u72de\u72ec-\u72ee\u72f0\u72f1\u730e\u732a\u732b\u732e\u736d\u739b\u73af\u73b0\u73d0\u7410\u743c\u7476\u7535\u753b\u7545\u7574\u7597\u759f\u75a1\u75ae\u75af\u75c8\u75c9\u75ea\u762a\u762b\u7663\u7691\u76b1\u76cf-\u76d1\u76d6-\u76d8\u7741\u7792\u77a9\u77eb\u77fe\u77ff\u7801\u7816\u781a\u783e\u7840\u7855\u7877\u788d\u78b1\u793c\u7977\u7978\u7984\u79c3\u79c6\u79ef\u79f0\u79fd\u7a0e\u7a33\u7a77\u7a83\u7a8d\u7a91\u7a9c\u7a9d\u7aa5\u7ad6\u7ade\u7b0b\u7b14\u7b3a\u7b3c\u7b5b\u7b79\u7b7e\u7b80\u7ba9\u7bd3\u7bee\u7c7b\u7ca4\u7caa\u7cae\u7d27\u7ea0\u7ea2\u7ea4\u7ea6\u7ea7\u7eaa-\u7eac\u7eaf\u7eb1-\u7eb3\u7eb5-\u7eba\u7ebd\u7ebf\u7ec3-\u7ec8\u7eca\u7ecd-\u7ecf\u7ed1-\u7ed3\u7ed5\u7ed8-\u7eda\u7edc-\u7edf\u7ee2\u7ee3\u7ee5-\u7ee7\u7ee9\u7eea\u7eed\u7ef0\u7ef3-\u7ef5\u7ef7\u7ef8\u7efc\u7efd\u7eff\u7f00\u7f04-\u7f06\u7f09\u7f0e\u7f13-\u7f16\u7f18\u7f1a\u7f1d\u7f20\u7f28\u7f29\u7f2e\u7f34\u7f57\u7f5a\u7f62\u7fd8\u8038\u803b\u8042\u804b\u804c\u8054\u806a\u8083\u80a0\u80a4\u80be-\u80c1\u80c6\u80f6\u8109\u810f-\u8111\u8113\u811a\u8131\u8138\u817b\u817e\u8206\u8230\u8231\u8270\u8273\u827a\u8282\u829c\u82a6\u82c7\u82cd\u82cf\u830e\u8346\u835a\u8361\u8363\u8364\u8367\u836b\u836f\u83b1\u83b2\u83b7\u83b9\u841d\u8424\u8425\u8427\u8428\u8471\u848b\u84dd\u84df\u8537\u853c\u8574\u864f\u8651\u865a\u867d\u867e\u8680-\u8682\u86ca\u86ee\u86f0\u8715\u8717\u8747\u8749\u8845\u8854\u8865\u886c\u8884\u889c\u88ad\u88c5\u88e4\u89c1\u89c2\u89c4-\u89c6\u89c8\u89c9\u8a89\u8a8a\u8ba1-\u8ba5\u8ba8\u8ba9\u8bab\u8bad-\u8bb0\u8bb2\u8bb3\u8bb6\u8bb8-\u8bba\u8bbc-\u8bc1\u8bc4-\u8bc6\u8bc8-\u8bca\u8bcc\u8bcd\u8bd1\u8bd5\u8bd7\u8bda\u8bdb\u8bdd\u8bde\u8be1-\u8be3\u8be5-\u8be7\u8beb-\u8bed\u8bef\u8bf1\u8bf2\u8bf4\u8bf5\u8bf7\u8bf8\u8bfa\u8bfb\u8bfd\u8bfe\u8c01\u8c03\u8c05\u8c06\u8c08\u8c0a\u8c0b\u8c0d\u8c0e\u8c10\u8c13\u8c17\u8c1a\u8c1c\u8c22-\u8c24\u8c26\u8c28\u8c29\u8c2c\u8c2d\u8c30\u8c31\u8c34\u8d1d-\u8d1f\u8d21-\u8d31\u8d34\u8d35\u8d37-\u8d3a\u8d3c\u8d3e\u8d3f\u8d41-\u8d44\u8d4a-\u8d4c\u8d4e-\u8d50\u8d54\u8d56\u8d58\u8d5a\u8d5b\u8d5e\u8d60-\u8d63\u8d75\u8d8b\u8dc3\u8df5\u8e0c\u8e2a\u8e7f\u8eaf\u8f66-\u8f69\u8f6c\u8f6e-\u8f70\u8f74\u8f7b\u8f7d\u8f7f\u8f83\u8f85\u8f86\u8f88-\u8f8a\u8f90\u8f91\u8f93\u8f95-\u8f97\u8f99\u8f9e\u8fa9\u8fab\u8fb9\u8fbd\u8fbe\u8fc1\u8fc7\u8fc8\u8fd0\u8fd8\u8fd9\u8fdb-\u8fdf\u8ff9\u9009\u900a\u9012\u903b\u9057\u9065\u9093\u90ae\u90b9\u90bb\u90d1\u90e7\u90f8\u915d\u9171\u917f\u91ca\u9274\u9488\u9489\u948e\u9492\u9493\u9499\u949d-\u94a2\u94a5-\u94a9\u94ae\u94b1\u94b3\u94b5\u94bb\u94be\u94c0-\u94c3\u94c5\u94c6\u94dc\u94dd\u94e1\u94e3\u94ec\u94ed\u94f0-\u94f2\u94f6\u94f8\u94fa\u94fe\u9500\u9501\u9504\u9505\u9508\u950b\u950c\u9510\u9511\u9517\u9519\u951a\u9521\u9523-\u9526\u9528\u952d-\u9530\u9539\u953b\u9540\u9541\u9547\u954a\u954d\u9550\u9551\u955c\u9563\u956d\u9570\u9576\u957f\u95e8\u95ea\u95ed-\u95f0\u95f2\u95f4\u95f7-\u95fb\u95fd\u9600-\u9602\u9605\u9609\u960e\u9610\u9611\u9614\u961f\u9633-\u9636\u9645-\u9648\u9655\u9668\u9669\u968f\u9690\u96b6\u96be\u96cf\u96fe\u9759\u97e6\u97e7\u97e9\u97f5\u9875-\u9877\u9879-\u987b\u987d-\u987f\u9881\u9882\u9884-\u9888\u988a\u9890\u9891\u9893\u9896-\u9898\u989c\u989d\u98a0\u98a4\u98a7\u98ce\u98d8\u98de\u9965\u996d-\u9972\u9975\u9976\u997a\u997c\u997f\u9981\u9985\u9986\u9988\u998b\u998f\u9992\u9a6c-\u9a71\u9a73\u9a74\u9a76\u9a79\u9a7b\u9a7c\u9a7e\u9a82\u9a84\u9a86\u9a87\u9a8b\u9a8c\u9a8f\u9a91\u9a97\u9a9a\u9aa1\u9aa4\u9c7c\u9c81\u9c8d\u9c9c\u9ca4\u9cb8\u9cc3\u9cd6\u9cde\u9e1f\u9e21\u9e23\u9e25\u9e26\u9e2d\u9e2f\u9e33\u9e35\u9e3d\u9e3f\u9e43\u9e45\u9e4a\u9e4f\u9e64\u9e70\u9ea6\u9ec4\u9f50\u9f7f\u9f84\u9f8b\u9f99\u9f9a\u9f9f]/;

function hasSimplified(text) {
  return SIMPLIFIED.test(String(text));
}

// 文字裡有簡體字時，請 Gemini 轉成繁體；沒有簡體字就原樣回傳，不會多呼叫一次
function toTraditional(text) {
  text = String(text);
  if (!hasSimplified(text)) return text;
  try {
    return askGemini('把下面文字裡的簡體字全部換成繁體字（台灣用語）。其他內容、標點、換行和格式一個字都不要改，只輸出轉換後的文字：\n\n' + text.slice(0, 12000), '', true) || text;
  } catch (err) {
    Logger.log('轉成繁體失敗，先保留原文：' + err.message);
    return text;
  }
}

// 把「筆記」工作表裡已經存成簡體的舊筆記轉成繁體。一次最多跑 4 分半鐘，沒跑完再執行一次
function fixSimplified() {
  var sh = notesSheet(), last = sh.getLastRow(), start = Date.now(), done = 0;
  for (var r = 2; r <= last; r++) {
    if (Date.now() - start > 270000) { Logger.log('時間快到了，請再執行一次 fixSimplified 接著做'); break; }
    var changed = false;
    [4, 5, 6, 8].forEach(function (c) {   // 標題、摘要、標籤、原文
      var cell = sh.getRange(r, c), v = String(cell.getValues()[0][0]);
      if (!hasSimplified(v)) return;
      var t = toTraditional(v);
      if (t !== v) { cell.setValue(t); changed = true; }
    });
    if (changed) done++;
  }
  Logger.log('這次把 ' + done + ' 則筆記轉成繁體');
}

// input 可以是一句話，也可以是 parts 陣列（含圖片或錄音）。keep＝true 代表不用再檢查簡體字（轉換本身會用到）
function askGemini(input, system, keep) {
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
  if (answer) return keep ? answer : toTraditional(answer);
  throw new Error(explainError(data.error));
}

function callGemini(model, parts, system, fast) {
  var body = { contents: [{ role: 'user', parts: parts }] };
  body.system_instruction = { parts: [{ text: (system ? system + '\n' : '') + TW_RULE }] };
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
