// ===== 第二大腦｜階段 3：問自己的 AI（語意索引＋問答）=====

var SHEET_VEC = '向量';
var EMBED_MODEL = 'gemini-embedding-2';   // 把文字變成一串數字（向量）的模型
var EMBED_DIM = 768;

function vecSheet() {
  var ss = book();
  var sh = ss.getSheetByName(SHEET_VEC);
  if (!sh) {
    sh = ss.insertSheet(SHEET_VEC);
    sh.getRange(1, 1, 1, 2).setValues([['編號', '向量（請勿手動修改）']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 把一段文字變成向量：意思相近的文字，向量也會相近
function embed(text) {
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + EMBED_MODEL + ':embedContent', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-goog-api-key': GEMINI_API_KEY },
    payload: JSON.stringify({ content: { parts: [{ text: String(text).slice(0, 6000) }] }, output_dimensionality: EMBED_DIM }),
    muteHttpExceptions: true
  });
  var data = {};
  try { data = JSON.parse(res.getContentText()); } catch (err) {}
  var values = (data.embedding && data.embedding.values) || (data.embeddings && data.embeddings[0] && data.embeddings[0].values);
  if (!values) throw new Error('建立向量失敗：' + (data.error ? data.error.code + ' ' + data.error.message : res.getResponseCode()));
  return values.map(function (v) { return Math.round(v * 10000) / 10000; });
}

function noteText(n) {
  return 'title: ' + n.title + ' | text: ' + n.summary + ' ' + (n.tags || []).join(' ') + ' ' + String(n.text || '').slice(0, 2000);
}

// 幫一則筆記建立索引（存新筆記時會自動呼叫）
function indexNote(n) {
  vecSheet().appendRow([n.id, JSON.stringify(embed(noteText(n)))]);
}

// ---- 幫所有還沒有索引的舊筆記建立索引 ----
// 一次最多跑 4 分半鐘，沒跑完再執行一次就會接著做。
function buildIndex() {
  var have = loadVectors();
  var todo = readNotes().filter(function (n) { return !have[n.id]; });
  var start = Date.now(), done = 0;
  for (var i = 0; i < todo.length; i++) {
    if (Date.now() - start > 270000) { Logger.log('時間快到了，請再執行一次 buildIndex'); break; }
    indexNote(todo[i]);
    done++;
  }
  Logger.log('這次建立了 ' + done + ' 則索引，還剩 ' + (todo.length - done) + ' 則');
}

function loadVectors() {
  var sh = vecSheet(), last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    try { map[Number(r[0])] = JSON.parse(r[1]); } catch (err) {}
  });
  return map;
}

// 兩個向量有多像：1 表示方向完全相同，0 表示無關
function cosine(a, b) {
  var dot = 0, na = 0, nb = 0;
  for (var i = 0; i < a.length; i++) { dot += a[i] * b[i]; na += a[i] * a[i]; nb += b[i] * b[i]; }
  return na && nb ? dot / Math.sqrt(na * nb) : 0;
}

// 問題和每則筆記的相似度：{ 編號: 分數 }
function semanticScores(q) {
  var qv = embed('task: search result | query: ' + q);
  var vecs = loadVectors(), out = {};
  Object.keys(vecs).forEach(function (id) { out[id] = cosine(qv, vecs[id]); });
  return out;
}

// 和某一則筆記最相近的其他筆記
function relatedNotes(id, count) {
  var vecs = loadVectors();
  if (!vecs[id]) return [];
  var byId = {};
  readNotes().forEach(function (n) { byId[n.id] = n; });
  return Object.keys(vecs).filter(function (k) { return Number(k) !== id && byId[k]; })
    .map(function (k) { return { n: byId[k], s: cosine(vecs[id], vecs[k]) }; })
    .sort(function (a, b) { return b.s - a.s; }).slice(0, count || 5)
    .map(function (x) { var item = slim(x.n); item.score = x.s; return item; });
}

// ---- 用自己的筆記回答問題 ----
function answerQuestion(q) {
  var notes = readNotes();
  if (!notes.length) return { text: '筆記庫還是空的，先存幾則筆記再來問我。', sources: [] };
  var picked = [];
  try {
    var scores = semanticScores(q);
    picked = notes.filter(function (n) { return scores[n.id] !== undefined; })
      .sort(function (a, b) { return scores[b.id] - scores[a.id]; }).slice(0, 6);
  } catch (err) {
    Logger.log('語意搜尋失敗，改用關鍵字：' + err.message);
  }
  if (!picked.length) {   // 還沒建索引時，退回用關鍵字找
    var ids = searchNotes(q, {}).slice(0, 6).map(function (n) { return n.id; });
    picked = notes.filter(function (n) { return ids.indexOf(n.id) >= 0; });
  }
  if (!picked.length) return { text: '筆記裡找不到相關的內容。（如果還沒執行過 buildIndex，請先執行）', sources: [] };
  var context = picked.map(function (n) {
    return '[#' + n.id + '] ' + n.title + '（' + n.date + '，' + n.category + '）\n摘要：' + n.summary + '\n內容：' + n.text.slice(0, 1200);
  }).join('\n\n');
  var system = '你是使用者的「第二大腦」。只能根據下面提供的筆記回答，不可以加入筆記裡沒有的事實。' +
    '每個重點後面用 [#編號] 標出是根據哪一則筆記。如果筆記裡沒有答案，就直接說「筆記裡沒有這方面的內容」。' +
    '用繁體中文（台灣用語），條列重點，300 字以內。';
  var text = askGemini('我的筆記：\n\n' + context + '\n\n我的問題：' + q, system);
  var used = picked.filter(function (n) { return text.indexOf('#' + n.id) >= 0; });
  return { text: text, sources: (used.length ? used : []).map(function (n) { return { id: n.id, title: n.title }; }) };
}

// 在編輯器測試用
function testAsk() {
  var a = answerQuestion('我最近記了哪些重點？');
  Logger.log(a.text);
  Logger.log('來源：' + JSON.stringify(a.sources));
}
