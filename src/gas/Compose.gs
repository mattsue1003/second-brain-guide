// ===== 第二大腦｜階段 5：把筆記變成成品 =====

// 想新增自己的成品類型：照格式在這裡加一行。名稱會出現在網頁的選單上。
var TEMPLATES = {
  'FB 貼文': '寫成一則臉書貼文：第一句要讓人想繼續看，中間分 3 個短段落，最後用一個問題邀請留言。不用表情符號，結尾附 3 到 5 個 hashtag。約 250 字。',
  '簡報大綱': '整理成一份 8 到 10 頁的簡報大綱。每一頁寫：頁面標題、3 個重點、建議放什麼圖或互動。第一頁是標題頁，最後一頁是帶得走的行動建議。',
  '教案初稿': '寫成一份 60 分鐘的團體活動教案，包含：活動名稱、對象、目標（3 點）、準備物品、流程（暖身 10 分鐘／主活動 35 分鐘／分享收尾 15 分鐘，各段寫帶領步驟）、分級調整（簡單版與進階版）、安全注意事項。',
  '衛教單張': '寫成一張給長輩和家屬看的衛教單張：一個大標題、3 到 4 個小標題，每段 2 到 3 句、句子短、用生活例子，最後列出「什麼情況要找醫師」。'
};

function composeKinds() {
  return Object.keys(TEMPLATES);
}

// arg = { ids: [筆記編號…], kind: '成品類型', extra: '補充要求' }
function composeDraft(arg) {
  var ids = (arg.ids || []).map(Number);
  var picked = readNotes().filter(function (n) { return ids.indexOf(n.id) >= 0; });
  if (!picked.length) return { error: '請先勾選至少一則筆記。' };
  var kind = TEMPLATES[arg.kind] ? arg.kind : composeKinds()[0];
  var context = picked.map(function (n) {
    return '[#' + n.id + '] ' + n.title + '\n摘要：' + n.summary + '\n內容：' + n.text.slice(0, 1500);
  }).join('\n\n');
  var text = askGemini(
    '素材（我自己的筆記）：\n\n' + context + '\n\n任務：' + TEMPLATES[kind] +
    (arg.extra ? '\n補充要求：' + arg.extra : '') +
    '\n\n只能使用素材裡有的事實與數字，素材不足的地方用「（待補：…）」標出來，不要自己編。用繁體中文（台灣用語）。',
    '你是使用者的寫作助手，根據他的筆記產出可以直接修改使用的初稿。');
  var sh = book().getSheetByName('產出');
  if (!sh) {
    sh = book().insertSheet('產出');
    sh.getRange(1, 1, 1, 4).setValues([['時間', '類型', '使用的筆記', '內容']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(4, 600);
  }
  sh.appendRow([new Date(), kind, picked.map(function (n) { return '#' + n.id; }).join(' '), text]);
  return { text: text, kind: kind, used: picked.map(function (n) { return { id: n.id, title: n.title }; }) };
}
