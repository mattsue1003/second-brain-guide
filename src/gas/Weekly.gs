// ===== 第二大腦｜階段 4：每週回顧信 =====

// ---- 寄出這一週的回顧（可以直接按「執行」測試）----
function weeklyReview() {
  var notes = readNotes();
  var weekAgo = Date.now() - 7 * 24 * 3600 * 1000;
  var week = notes.filter(function (n) { return n.ms >= weekAgo; });
  var old = pickOld(notes.filter(function (n) { return n.ms < weekAgo; }), week);

  var review;
  if (week.length) {
    var list = week.map(function (n) { return '[#' + n.id + '] ' + n.title + '（' + n.category + '）：' + n.summary; }).join('\n');
    var oldList = old.map(function (n) { return '[#' + n.id + '] ' + n.title + '（' + n.date + '）：' + n.summary; }).join('\n');
    review = askGemini(
      '這是我這一週存進筆記庫的內容：\n' + list + '\n\n' +
      (oldList ? '這是幾則以前的舊筆記：\n' + oldList + '\n\n' : '') +
      '請幫我寫一份每週回顧，用繁體中文（台灣用語），分成四段，每段開頭用【】標出段名：\n' +
      '【這週在想什麼】用 2 到 3 句話歸納這週筆記的主題。\n' +
      '【值得延伸的三件事】列 3 點，每點說明可以怎麼做，並標出根據的筆記 [#編號]。\n' +
      (oldList ? '【舊筆記的新用途】說明上面的舊筆記和這週的內容有什麼關聯、可以怎麼一起用。\n' : '') +
      '【可以發表的題目】建議 1 個可以寫成貼文或做成課程的題目，附一句開場白。\n' +
      '只根據筆記內容寫，不要加入筆記裡沒有的事實。',
      '你是協助使用者回顧筆記的助手，語氣務實、具體。');
  } else {
    review = '【這週沒有新筆記】\n這週還沒有存任何東西。想到什麼就傳給 LINE 機器人，一句話也可以。';
  }

  var subject = '第二大腦每週回顧｜' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'M/d') + '（本週 ' + week.length + ' 則）';
  var to = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: reviewHtml(review, week, old) });
  Logger.log('已寄到 ' + to);
  if (WEEKLY_LINE && typeof linePush === 'function') linePush(subject + '\n\n' + review);
  return review;
}

// 挑 3 則舊筆記：有語意索引就挑和這週最相關的，沒有就隨機挑
function pickOld(olds, week) {
  if (!olds.length) return [];
  if (typeof loadVectors === 'function' && week.length) {
    var vecs = loadVectors();
    var wv = week.map(function (n) { return vecs[n.id]; }).filter(Boolean);
    if (wv.length) {
      var scored = olds.filter(function (n) { return vecs[n.id]; }).map(function (n) {
        var best = 0;
        wv.forEach(function (v) { best = Math.max(best, cosine(v, vecs[n.id])); });
        return { n: n, s: best };
      }).sort(function (a, b) { return b.s - a.s; });
      if (scored.length) return scored.slice(0, 3).map(function (x) { return x.n; });
    }
  }
  return olds.slice().sort(function () { return Math.random() - 0.5; }).slice(0, 3);
}

function reviewHtml(review, week, old) {
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function rows(list) {
    return list.map(function (n) {
      return '<li style="margin-bottom:8px"><b>#' + n.id + ' ' + esc(n.title) + '</b> <span style="color:#667">' + n.date + '｜' + esc(n.category) + '</span><br>' + esc(n.summary) + '</li>';
    }).join('');
  }
  var body = esc(review).replace(/【([^】]+)】/g, '<h3 style="margin:18px 0 6px;color:#2F45C5">$1</h3>').replace(/\n/g, '<br>');
  return '<div style="font-family:sans-serif;max-width:640px;line-height:1.7;color:#1B2030">' + body +
    (week.length ? '<h3 style="margin:22px 0 6px">本週存了這些</h3><ul style="padding-left:18px">' + rows(week) + '</ul>' : '') +
    (old.length ? '<h3 style="margin:22px 0 6px">翻出來的舊筆記</h3><ul style="padding-left:18px">' + rows(old) + '</ul>' : '') +
    '<p style="color:#667;font-size:13px">這封信由你的第二大腦自動寄出。</p></div>';
}

// ---- 設定每週日晚上 8 點自動寄出（只要執行一次）----
function installWeekly() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'weeklyReview') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyReview').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(20).create();
  Logger.log('設定好了：每週日晚上 8 點到 9 點之間會自動寄出回顧信');
}

// 不想再收到：執行這個
function removeWeekly() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'weeklyReview') ScriptApp.deleteTrigger(t);
  });
  Logger.log('已取消每週回顧信');
}
