// ===== 第二大腦｜階段 2：知識總站（搜尋網頁＋匯入舊資料）=====

// 用瀏覽器開啟 /exec 網址時會執行這裡
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('我的第二大腦')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 網頁上所有的操作都經過這裡。每次都要帶密碼，密碼錯 5 次會鎖 10 分鐘。
function api(pass, action, arg) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pw-fails') || 0);
  if (fails >= 5) return { error: '密碼錯太多次，請 10 分鐘後再試。' };
  if (!WEB_PASSWORD || String(pass) !== String(WEB_PASSWORD)) {
    cache.put('pw-fails', String(fails + 1), 600);
    return { error: '密碼不對' };
  }
  arg = arg || {};
  try {
    if (action === 'home') return home();
    if (action === 'search') return { notes: searchNotes(arg.q, arg) };
    if (action === 'note') return { note: readNotes().filter(function (n) { return n.id === Number(arg.id); })[0] || null };
    if (action === 'add') return { note: slim(captureText(arg.text)) };
    if (action === 'related') return { notes: typeof relatedNotes === 'function' ? relatedNotes(Number(arg.id), 5) : [] };
    if (action === 'ask') {
      if (typeof answerQuestion !== 'function') return { error: '問答功能要完成「階段 3」才會開啟。' };
      return answerQuestion(arg.q);
    }
    if (action === 'compose') {
      if (typeof composeDraft !== 'function') return { error: '產出功能要完成「階段 5」才會開啟。' };
      return composeDraft(arg);
    }
    return { error: '不認得的操作：' + action };
  } catch (err) {
    return { error: err.message };
  }
}

// 首頁要的資料：總數、各類別數量、常用標籤、最近 20 則
function home() {
  var notes = readNotes();
  var cats = {}, tags = {};
  notes.forEach(function (n) {
    cats[n.category] = (cats[n.category] || 0) + 1;
    n.tags.forEach(function (t) { tags[t] = (tags[t] || 0) + 1; });
  });
  var topTags = Object.keys(tags).sort(function (a, b) { return tags[b] - tags[a]; }).slice(0, 24)
    .map(function (t) { return { name: t, count: tags[t] }; });
  return {
    count: notes.length,
    categories: Object.keys(cats).map(function (c) { return { name: c, count: cats[c] }; }),
    tags: topTags,
    notes: notes.slice(-20).reverse().map(slim),
    canAsk: typeof answerQuestion === 'function',
    canCompose: typeof composeDraft === 'function',
    kinds: typeof composeKinds === 'function' ? composeKinds() : []
  };
}

// 傳給網頁的精簡版筆記（不含完整原文，比較快）
function slim(n) {
  return { id: n.id, date: n.date || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd'), type: n.type, title: n.title,
    summary: n.summary, tags: n.tags, category: n.category, link: n.link || '' };
}

// 搜尋：關鍵字出現在標題、標籤、摘要、原文各有不同分數。
// 階段 3 完成後，會再加上「語意相近」的分數，用不同的說法也找得到。
function searchNotes(q, opt) {
  opt = opt || {};
  q = String(q || '').trim();
  var words = q.toLowerCase().split(/\s+/).filter(String);
  var sem = {};
  if (q && typeof semanticScores === 'function') {
    try { sem = semanticScores(q); } catch (err) { Logger.log('語意搜尋失敗，只用關鍵字：' + err.message); }
  }
  var out = [];
  readNotes().forEach(function (n) {
    if (opt.category && n.category !== opt.category) return;
    if (opt.tag && n.tags.indexOf(opt.tag) < 0) return;
    var score = 0;
    words.forEach(function (w) {
      if (n.title.toLowerCase().indexOf(w) >= 0) score += 5;
      if (n.tags.join(' ').toLowerCase().indexOf(w) >= 0) score += 4;
      if (n.summary.toLowerCase().indexOf(w) >= 0) score += 3;
      if (n.text.toLowerCase().indexOf(w) >= 0) score += 1;
    });
    var s = sem[n.id] || 0;
    if (s > 0.55) score += (s - 0.55) * 20;
    if (!q) score = n.ms / 1e13;          // 沒有關鍵字：照時間新到舊
    if (q && score <= 0) return;
    var item = slim(n);
    item.score = score;
    out.push(item);
  });
  out.sort(function (a, b) { return b.score - a.score || b.id - a.id; });
  return out.slice(0, 50);
}

// ---- 匯入舊資料 ----
// 在試算表新增一個叫「匯入」的工作表，A 欄每一格貼一則舊筆記（或一個網址），再執行這個函式。
// 處理完的會在 B 欄寫上「已匯入 #編號」。一次最多跑 4 分半鐘，沒跑完再執行一次就會接著做。
function importSheet() {
  var sh = book().getSheetByName('匯入');
  if (!sh) throw new Error('找不到「匯入」工作表，請先新增並把舊筆記貼在 A 欄');
  var last = sh.getLastRow();
  if (last < 1) return;
  var rows = sh.getRange(1, 1, last, 2).getValues();
  var start = Date.now(), done = 0;
  for (var i = 0; i < rows.length; i++) {
    if (!String(rows[i][0]).trim() || String(rows[i][1]).trim()) continue;
    if (Date.now() - start > 270000) { Logger.log('時間快到了，請再執行一次 importSheet 接著匯入'); break; }
    try {
      var n = captureText(rows[i][0]);
      sh.getRange(i + 1, 2).setValue('已匯入 #' + n.id);
      done++;
    } catch (err) {
      sh.getRange(i + 1, 2).setValue('');
      Logger.log('第 ' + (i + 1) + ' 列失敗：' + err.message);
      break;
    }
  }
  Logger.log('這次匯入了 ' + done + ' 則');
}
