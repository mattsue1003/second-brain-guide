// ===== 第二大腦｜階段 6：AI 秘書（行事曆、待辦、記帳、存檔案）=====
// 貼上這個檔案後，LINE 傳來的文字會先經過這裡：
// 是行程、待辦、記帳就直接處理，其他的照舊存成筆記。金鑰一樣放在「設定.gs」，這裡不用填。

// true：每句話都先請 AI 判斷是行程、待辦、記帳還是筆記。
// false：只處理「行程」「待辦」「記帳」開頭的句子，其他一律存成筆記（比較省 AI 額度）。
var SECRETARY_AUTO = true;

// 秘書用的工作表和欄位，第一次用到時會自動建立。想加新功能，先在這裡加一行
var TABS = {
  '待辦': ['時間', '事項', '期限', '狀態', '完成時間'],
  '記帳': ['時間', '日期', '收支', '分類', '項目', '金額', '備註']
};

// ---- 判斷這句話要做什麼，然後真的去做。回傳空字串＝「這是筆記」，交回給「收集」存起來 ----
function secretaryText(text) {
  text = String(text).trim();
  if (/https?:\/\//.test(text)) return '';                    // 有網址的一律當筆記
  if (/^(?:筆記|記一下)\s*[:：\s]/.test(text)) return '';    // 開頭寫「筆記」就是要存成筆記
  // 1. 固定說法，不用問 AI
  if (/^待辦(?:清單|事項)?$/.test(text)) return listTodos();
  var todo = text.match(/^待辦(?:事項)?\s*[:：\s]\s*([\s\S]+)$/);
  if (todo) return addTodo({ content: todo[1] });
  var done = text.match(/^完成\s*[:：]?\s*(\d+)$/);
  if (done) return doneTodo({ keyword: done[1] });
  if (/^(?:記帳|查帳|帳目)$/.test(text)) return listExpenses({});
  if (/^(?:行程|行事曆)$/.test(text)) {
    var t = parseTime(todayText());
    return listEvents({ from: todayText(), to: dayText(new Date(t.getFullYear(), t.getMonth(), t.getDate() + 6)) });
  }
  // 2. 其他的請 AI 判斷
  var forced = /^(?:行程|行事曆|記帳)\s*[:：\s]/.test(text);   // 這幾個開頭代表一定是秘書的事
  if (!SECRETARY_AUTO && !forced) return '';
  var plan = askPlan(text);
  if (plan.error) {
    var quick = text.match(/^記帳\s*[:：\s]\s*(.+?)\s*(\d+(?:\.\d+)?)\s*元?$/);   // AI 不能用時，「記帳 午餐 120」還是記得進去
    if (quick) return addExpense({ item: quick[1], amount: quick[2] }) + '\n（AI 暫時不能用，分類先記成「未分類」）';
    return forced ? plan.error : '';   // 沒指定的句子，AI 不能用時就當筆記存
  }
  if (plan.action === 'add_event') return addEvent(plan);
  if (plan.action === 'list_events') return listEvents(plan);
  if (plan.action === 'add_todo') return addTodo(plan);
  if (plan.action === 'list_todos') return listTodos();
  if (plan.action === 'done_todo') return doneTodo(plan);
  if (plan.action === 'add_expense') return addExpense(plan);
  if (plan.action === 'list_expenses') return listExpenses(plan);
  return forced ? '我沒聽懂，請換個說法，例如「行程：明天下午三點開會」或「記帳 午餐 120」。' : '';
}

// ---- 行事曆 ----
function addEvent(p) {
  var start = parseTime(p.start);
  if (!p.title || !start) return '我沒聽懂時間或事情，請再說一次，例如「明天下午三點跟王主任開會」。';
  var cal = CalendarApp.getDefaultCalendar();
  var opt = { location: p.location || '', description: p.note || '' };
  if (String(p.start).length <= 10) {                 // 只有日期，沒有時間 → 整天的活動
    cal.createAllDayEvent(p.title, start, opt);
    return '已記到行事曆：\n' + fmtDate(start) + '（整天）\n' + p.title;
  }
  var end = parseTime(p.end);
  if (!end || end <= start) end = new Date(start.getTime() + 60 * 60 * 1000);   // 沒說結束時間就當作一小時
  cal.createEvent(p.title, start, end, opt);
  return '已記到行事曆：\n' + fmtDate(start) + ' ' + fmtTime(start) + '–' + fmtTime(end) + '\n' + p.title + (p.location ? '\n地點：' + p.location : '');
}

function listEvents(p) {
  var from = parseTime(p.from) || parseTime(todayText());
  var last = parseTime(p.to) || from;
  var to = new Date(last.getFullYear(), last.getMonth(), last.getDate() + 1);   // 包含結束那一天
  var range = fmtDate(from) + (last > from ? ' 到 ' + fmtDate(last) : '');
  var events = CalendarApp.getDefaultCalendar().getEvents(from, to);
  if (!events.length) return range + ' 行事曆上沒有安排。';
  return range + ' 的行事曆：\n' + events.slice(0, 20).map(function (ev) {
    return fmtDate(ev.getStartTime()) + ' ' + (ev.isAllDayEvent() ? '整天' : fmtTime(ev.getStartTime())) + '　' + ev.getTitle();
  }).join('\n');
}

// ---- 工作表：和筆記放在同一個「第二大腦」試算表 ----
function tab(name) {
  var sh = book().getSheetByName(name);
  if (!sh) {
    sh = book().insertSheet(name);
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function tabRows(name) {   // 讀出一個工作表的全部資料（不含標題列）
  var sh = tab(name), last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, TABS[name].length).getValues();
}

// ---- 待辦事項 ----
function openTodos() {   // 還沒完成的待辦，row 是它在試算表的第幾列
  var list = [];
  tabRows('待辦').forEach(function (r, i) {
    if (r[1] && r[3] !== '完成') list.push({ row: i + 2, content: String(r[1]), due: r[2] });
  });
  return list;
}

function todoLine(t, i) {
  var due = t.due ? '（期限 ' + (t.due instanceof Date ? fmtDate(t.due) : t.due) + '）' : '';
  return (i + 1) + '. ' + t.content + due;
}

function addTodo(p) {
  var content = String(p.content || '').trim();
  if (!content) return '我沒聽懂要做什麼，請再說一次，例如「待辦：打電話給王主任」。';
  var due = parseTime(p.due);
  tab('待辦').appendRow([new Date(), content, due || '', '未完成', '']);
  return '已加入待辦：\n' + content + (due ? '（期限 ' + fmtDate(due) + '）' : '') + '\n目前還有 ' + openTodos().length + ' 件沒完成。傳「待辦」可以看清單。';
}

function listTodos() {
  var list = openTodos();
  if (!list.length) return '目前沒有待辦事項。';
  return '待辦事項（' + list.length + ' 件）：\n' + list.map(todoLine).join('\n') + '\n\n做完了就傳「完成 1」。';
}

function doneTodo(p) {
  var list = openTodos();
  var key = String(p.keyword || '').trim();
  if (!list.length) return '目前沒有待辦事項。';
  var hits = /^\d+$/.test(key) ? (list[Number(key) - 1] ? [list[Number(key) - 1]] : []) : list.filter(function (t) { return key && t.content.indexOf(key) >= 0; });
  if (!hits.length) return '找不到這件待辦。目前的清單：\n' + list.map(todoLine).join('\n');
  if (hits.length > 1) return '有 ' + hits.length + ' 件很像，請用編號告訴我，例如「完成 1」：\n' + list.map(todoLine).join('\n');
  var sh = tab('待辦');
  sh.getRange(hits[0].row, 4).setValue('完成');
  sh.getRange(hits[0].row, 5).setValue(new Date());
  return '已完成：' + hits[0].content + '\n還有 ' + (list.length - 1) + ' 件沒完成。';
}

// ---- 記帳 ----
function money(n) { return String(Math.round(n * 100) / 100).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

function addExpense(p) {
  var amount = Number(String(p.amount == null ? '' : p.amount).replace(/[^\d.]/g, ''));
  if (!amount) return '我沒聽懂金額，請再說一次，例如「記帳 午餐 120」。';
  var type = p.type === '收入' ? '收入' : '支出';
  var day = parseTime(p.date) || parseTime(todayText());
  tab('記帳').appendRow([new Date(), day, type, p.category || '未分類', p.item || '', amount, p.note || '']);
  var sum = sumExpenses(new Date(day.getFullYear(), day.getMonth(), 1), new Date(day.getFullYear(), day.getMonth() + 1, 1));
  return '已記帳：\n' + fmtDate(day) + ' ' + type + ' ' + money(amount) + ' 元｜' + (p.category || '未分類') + '｜' + (p.item || '') +
    '\n' + (day.getMonth() + 1) + ' 月支出累計 ' + money(sum.out) + ' 元' + (sum.income ? '，收入 ' + money(sum.income) + ' 元' : '');
}

function sumExpenses(from, to) {   // 統計 from（含）到 to（不含）之間的帳
  var sum = { out: 0, income: 0, cats: {}, count: 0 };
  tabRows('記帳').forEach(function (r) {
    var d = new Date(r[1]), n = Number(r[5]) || 0;
    if (!(d >= from && d < to) || !n) return;
    sum.count++;
    if (r[2] === '收入') { sum.income += n; return; }
    sum.out += n;
    sum.cats[r[3]] = (sum.cats[r[3]] || 0) + n;
  });
  return sum;
}

function listExpenses(p) {
  var today = parseTime(todayText());
  var from = parseTime(p.from) || new Date(today.getFullYear(), today.getMonth(), 1);   // 沒說就看這個月
  var last = parseTime(p.to) || today;
  var sum = sumExpenses(from, new Date(last.getFullYear(), last.getMonth(), last.getDate() + 1));
  if (!sum.count) return fmtDate(from) + ' 到 ' + fmtDate(last) + ' 沒有記帳資料。';
  var cats = Object.keys(sum.cats).sort(function (a, b) { return sum.cats[b] - sum.cats[a]; });
  return fmtDate(from) + ' 到 ' + fmtDate(last) + '（' + sum.count + ' 筆）\n支出 ' + money(sum.out) + ' 元' + (sum.income ? '\n收入 ' + money(sum.income) + ' 元' : '') +
    (cats.length ? '\n\n支出分類：\n' + cats.map(function (c) { return c + ' ' + money(sum.cats[c]) + ' 元'; }).join('\n') : '');
}

// ---- 影片和檔案：存到雲端硬碟的「第二大腦附件」資料夾，並在「筆記」留一則，之後搜尋得到 ----
function saveAttachment(m) {
  var kind = m.type === 'video' ? ['影片', 'mp4'] : ['檔案', 'bin'];
  var name = m.fileName || (kind[0] + '_' + Utilities.formatDate(new Date(), tz(), 'yyyyMMdd_HHmmss') + '.' + kind[1]);
  var file = attachFolder().createFile(lineContent(m.id).setName(name));
  var note = { type: '檔案', title: name.slice(0, 40), summary: '存在雲端硬碟的' + kind[0], tags: [kind[0]], category: '其他', text: name, link: file.getUrl() };
  saveNote(note);
  return '已存入 #' + note.id + '｜檔案\n' + name + '\n' + file.getUrl();
}

// ---- 時間的小工具（以這個專案設定的時區為準）----
function tz() { return Session.getScriptTimeZone(); }
function dayText(d) { return Utilities.formatDate(d, tz(), 'yyyy-MM-dd'); }
function todayText() { return dayText(new Date()); }
function fmtDate(d) { return Utilities.formatDate(d, tz(), 'M/d') + '（' + '日一二三四五六'.charAt(d.getDay()) + '）'; }
function fmtTime(d) { return Utilities.formatDate(d, tz(), 'HH:mm'); }
function parseTime(s) {   // 接受 2026-11-25 或 2026-11-25T14:00
  var m = String(s || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0));
}

// ---- 請 Gemini 把一句話變成「要做什麼」的 JSON ----
function askPlan(text) {
  var now = new Date();
  var system = [
    '你是使用者的私人秘書。現在時間是 ' + Utilities.formatDate(now, tz(), 'yyyy-MM-dd HH:mm') + '，星期' + '日一二三四五六'.charAt(now.getDay()) + '，時區 ' + tz() + '。',
    '請判斷使用者這句話要做哪一件事，只輸出一個 JSON 物件，不要輸出其他文字：',
    '1. 新增行事曆：{"action":"add_event","title":"事情","start":"YYYY-MM-DDTHH:mm","end":"YYYY-MM-DDTHH:mm","location":"地點","note":"備註"}。沒說幾點就只給日期 "YYYY-MM-DD"；沒說結束時間就不要給 end。',
    '2. 查行事曆：{"action":"list_events","from":"YYYY-MM-DD","to":"YYYY-MM-DD"}',
    '3. 新增待辦事項：{"action":"add_todo","content":"要做的事","due":"期限 YYYY-MM-DD，沒說就空字串"}',
    '4. 看待辦清單：{"action":"list_todos"}',
    '5. 把待辦標成完成：{"action":"done_todo","keyword":"那件事的關鍵字，或清單上的編號"}',
    '6. 記帳：{"action":"add_expense","type":"支出 或 收入","category":"支出從 餐飲、交通、購物、醫療、娛樂、居家、教育、其他 選一個；收入從 薪資、獎金、其他 選一個","item":"項目","amount":數字,"date":"YYYY-MM-DD，沒說就今天","note":"備註"}',
    '7. 查帳或看花費統計：{"action":"list_expenses","from":"YYYY-MM-DD","to":"YYYY-MM-DD"}。沒說期間就兩個都給空字串。',
    '8. 以上都不是：{"action":"note"}',
    '判斷規則：日期要換算成實際日期，例如「明天」「下週三」。',
    '有明確時刻、要到場的事（開會、看診、上課）用 add_event；要去做、但沒有固定時刻的事（買東西、打電話、交報告、記得要…）用 add_todo。',
    '提到自己花了多少錢、買了什麼多少錢、收到多少錢，用 add_expense；一句話有好幾筆時只記第一筆。',
    '想法、心得、觀察、知識、教案點子、聯絡資訊、一般的問句，一律用 note，系統會把它存進筆記庫。拿不定主意時也用 note。'
  ].join('\n');
  var raw;
  try {
    raw = askGemini(text, system);
  } catch (err) {
    return { error: err.message };
  }
  try {
    var plan = JSON.parse(raw.slice(raw.indexOf('{'), raw.lastIndexOf('}') + 1));
    return plan && plan.action ? plan : { action: 'note' };
  } catch (err) {
    return { action: 'note' };   // AI 沒照格式回，就當成筆記
  }
}

// ---- 第一次請在編輯器執行這個：會要求授權行事曆，並建立「待辦」「記帳」工作表 ----
function testSecretary() {
  if (tz() !== 'Asia/Taipei') Logger.log('注意：這個專案的時區是 ' + tz() + '。在台灣使用請到左側「專案設定」改成 (GMT+08:00) 台北，不然記到行事曆的時間會跑掉。');
  tab('待辦'); tab('記帳');
  Logger.log('「待辦」「記帳」工作表準備好了');
  Logger.log(listEvents({}));
  Logger.log('AI 判斷測試：' + JSON.stringify(askPlan('明天下午三點跟王主任開會，地點在三樓會議室')));
}
