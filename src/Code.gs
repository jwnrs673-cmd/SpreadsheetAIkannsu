/**
 * Code.gs
 * ------------------------------------------------------------------
 * メニュー登録・イベント・テスト用ユーティリティ。
 * 実処理は Setup.gs / Anonymize.gs、設定は Config.gs 参照。
 * ------------------------------------------------------------------
 */

/** スプレッドシートを開いたときにメニューを追加。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('クレームAI分類')
    .addItem('① マスター系シートを作成/更新', 'setupMasters')
    .addItem('② メイン表に数式・プルダウンを適用', 'applyMainSheetFormulas')
    .addSeparator()
    .addItem('テストデータを投入（5ケース）', 'insertTestCases')
    .addItem('変更内容（計画）を表示', 'showSetupPlan')
    .addToUi();
}

function showSetupPlan() {
  SpreadsheetApp.getUi().alert(SETUP_PLAN.join('\n'));
}

/**
 * 依存プルダウン（任意）。メイン表で
 *  AK(確定大分類)を選ぶ → AL(確定中分類)候補を絞る
 *  AL(確定中分類)を選ぶ → AM(確定小分類)候補を絞る
 * 簡易トリガのため権限不要。失敗しても全件リストが有効で運用は止まらない。
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEETS.CLAIM) return;
    var col = e.range.getColumn(), row = e.range.getRow();
    if (row < 2) return;
    var ss = e.source;

    if (col === COL.FIX_DAI) {
      // 中分類マスター: A=大分類(1), B=中分類(2)
      setDependentValidation_(ss, sh, row, COL.FIX_CHU, SHEETS.RULE_CHU, 1, 2, e.value);
      sh.getRange(row, COL.FIX_SHO).clearDataValidations(); // 小分類はリセット
    }
    if (col === COL.FIX_CHU) {
      // 小分類マスター: A=大分類(1), B=中分類(2), C=小分類(3)
      var dai = sh.getRange(row, COL.FIX_DAI).getValue();
      setDependentTwoKeyValidation_(ss, sh, row, COL.FIX_SHO, SHEETS.RULE_SHO, 1, 2, 3, dai, e.value);
    }
  } catch (err) {
    Logger.log('onEdit: ' + err);
  }
}

function setDependentValidation_(ss, sh, row, targetCol, masterName, keyCol, valCol, keyVal) {
  var cell = sh.getRange(row, targetCol);
  if (!keyVal) { cell.clearDataValidations(); return; }
  var m = ss.getSheetByName(masterName);
  var last = getLastDataRow_(m, valCol);
  if (last < 2) return;
  var data = m.getRange(2, 1, last - 1, Math.max(keyCol, valCol)).getValues();
  var list = uniq_(data.filter(function (r) { return r[keyCol - 1] === keyVal; })
                       .map(function (r) { return r[valCol - 1]; }));
  applyExplicitList_(cell, list);
}

function setDependentTwoKeyValidation_(ss, sh, row, targetCol, masterName, k1Col, k2Col, valCol, k1, k2) {
  var cell = sh.getRange(row, targetCol);
  if (!k1 || !k2) { cell.clearDataValidations(); return; }
  var m = ss.getSheetByName(masterName);
  var last = getLastDataRow_(m, valCol);
  if (last < 2) return;
  var data = m.getRange(2, 1, last - 1, Math.max(k1Col, k2Col, valCol)).getValues();
  var list = uniq_(data.filter(function (r) { return r[k1Col - 1] === k1 && r[k2Col - 1] === k2; })
                       .map(function (r) { return r[valCol - 1]; }));
  applyExplicitList_(cell, list);
}

function uniq_(arr) {
  var seen = {}, out = [];
  arr.forEach(function (v) { if (v !== '' && !seen[v]) { seen[v] = 1; out.push(v); } });
  return out;
}

function applyExplicitList_(cell, list) {
  if (!list || !list.length) { cell.clearDataValidations(); return; }
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true).setAllowInvalid(true).build();
  cell.setDataValidation(rule);
}

/**
 * テストケースをメイン表の末尾に投入する（AB=原文, A=状況に"TEST"の目印）。
 * 投入後にメイン表の数式・プルダウンを適用する。※テスト行は後で削除可。
 */
function insertTestCases() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.CLAIM);
  if (!sh) { SpreadsheetApp.getUi().alert('メイン表「' + SHEETS.CLAIM + '」が見つかりません。'); return; }
  if (!ss.getSheetByName(SHEETS.RULE_CHU)) setupMasters();

  var cases = [
    'レジ担当者のネイルが長く、食品を扱う店として衛生面が気になった',                 // レジ接客/接客態度/身だしなみ
    'チラシの特売品を買おうとしたが棚になく、入荷予定も分からなかった',               // 欠品・特注/欠品/チラシ
    '棚の価格は198円だったが、レジでは248円になっていた',                           // アンマッチ/売価アンマッチ
    '売場で在庫を尋ねたのに、店員が調べてくれず対応が悪かった',                       // 外周り接客/売場・作業時/在庫確認・問い合わせ対応
    '返品できると言われたが、後日別の担当者から返品できないと言われた'                // 外周り接客/売場・作業時/返品・交換対応
  ];

  var last = getLastDataRow_(sh, COL.RAW);
  var start = Math.max(last + 1, 2);
  // A列(状況)に目印、AB列(原文)に本文
  var marks = cases.map(function () { return ['TEST']; });
  var texts = cases.map(function (t) { return [t]; });
  sh.getRange(start, 1, marks.length, 1).setValues(marks);
  sh.getRange(start, COL.RAW, texts.length, 1).setValues(texts);

  setupMainSheet_(ss, sh);
  setupValidations_(ss, sh);
  SpreadsheetApp.getActive().toast('テスト5件を投入し数式を適用しました（A列=TEST）。', 'クレームAI分類', 6);
}
