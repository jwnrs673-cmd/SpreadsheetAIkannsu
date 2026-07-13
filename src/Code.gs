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
    .addItem('① 初期セットアップ（全シート作成）', 'setupAll')
    .addItem('② 数式を全データ行へ適用', 'applyFormulasToAllRows')
    .addSeparator()
    .addItem('テストデータを投入（5ケース）', 'insertTestCases')
    .addItem('セットアップ計画を表示', 'showSetupPlan')
    .addToUi();
}

/** セットアップ計画をダイアログ表示。 */
function showSetupPlan() {
  SpreadsheetApp.getUi().alert(SETUP_PLAN.join('\n'));
}

/**
 * 依存プルダウン（任意）。
 * N（確定大分類）を選ぶと O（確定中分類）候補を、
 * N+O を選ぶと Y（確定小分類）候補を、その行だけ絞り込む。
 * 簡易トリガのため権限不要。失敗しても全件リストは有効なので運用は止まらない。
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEETS.CLAIM) return;
    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (row < 2) return;

    var ss = e.source;

    // N（確定大分類）変更 → O候補を大分類で絞る
    if (col === COL.FIX_DAI) {
      setDependentValidation_(ss, sh, row, COL.FIX_CHU, SHEETS.RULE_CHU, 1, 2, e.value);
    }
    // O（確定中分類）変更 → Y候補を大分類+中分類で絞る
    if (col === COL.FIX_CHU) {
      var daiVal = sh.getRange(row, COL.FIX_DAI).getValue();
      setDependentTwoKeyValidation_(ss, sh, row, COL.FIX_SHO, SHEETS.RULE_SHO, 1, 2, 3, daiVal, e.value);
    }
  } catch (err) {
    // 依存プルダウンはベストエフォート。エラーでも運用継続。
    Logger.log('onEdit: ' + err);
  }
}

/** 1キー（大分類）で絞ったリストを対象セルに設定。 */
function setDependentValidation_(ss, sh, row, targetCol, masterName, keyCol, valCol, keyVal) {
  var cell = sh.getRange(row, targetCol);
  if (!keyVal) { cell.clearDataValidations(); return; }
  var m = ss.getSheetByName(masterName);
  var last = getLastDataRow_(m, valCol);
  if (last < 2) return;
  var data = m.getRange(2, 1, last - 1, Math.max(keyCol, valCol)).getValues();
  var list = data.filter(function (r) { return r[keyCol - 1] === keyVal; })
                 .map(function (r) { return r[valCol - 1]; });
  applyExplicitList_(cell, list);
}

/** 2キー（大分類+中分類）で絞ったリストを対象セルに設定。 */
function setDependentTwoKeyValidation_(ss, sh, row, targetCol, masterName, key1Col, key2Col, valCol, k1, k2) {
  var cell = sh.getRange(row, targetCol);
  if (!k1 || !k2) { cell.clearDataValidations(); return; }
  var m = ss.getSheetByName(masterName);
  var last = getLastDataRow_(m, valCol);
  if (last < 2) return;
  var data = m.getRange(2, 1, last - 1, Math.max(key1Col, key2Col, valCol)).getValues();
  var list = data.filter(function (r) { return r[key1Col - 1] === k1 && r[key2Col - 1] === k2; })
                 .map(function (r) { return r[valCol - 1]; });
  applyExplicitList_(cell, list);
}

/** 明示リストのデータ検証を適用。 */
function applyExplicitList_(cell, list) {
  if (!list || !list.length) { cell.clearDataValidations(); return; }
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(true)
    .build();
  cell.setDataValidation(rule);
}

/**
 * テストケース（仕様14の5件）をクレーム記録に投入する。
 * A〜E のみ入れ、数式・AIは既存フローに任せる。既存データの下に追記。
 */
function insertTestCases() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.CLAIM);
  if (!sh) { SpreadsheetApp.getUi().alert('先に初期セットアップを実行してください。'); return; }

  var cases = [
    ['T-1', new Date(), 'テスト店', '電話', 'レジ担当者のネイルが長く、食品を扱う店として衛生面が気になった'],
    ['T-2', new Date(), 'テスト店', 'メール', '特売品を買おうとしたが棚になく、店員に聞いても入荷予定が分からなかった'],
    ['T-3', new Date(), 'テスト店', '電話', '棚の価格は198円だったが、レジでは248円になった'],
    ['T-4', new Date(), 'テスト店', '店頭', 'セルフレジで楽天Edyの支払いが完了していないのに、支払いが終わったと思って退店した'],
    ['T-5', new Date(), 'テスト店', '電話', '返品できると言われたが、後日別の担当者から返品できないと言われた']
  ];
  var last = getLastDataRow_(sh, COL.RAW);
  var start = Math.max(last + 1, 2);
  sh.getRange(start, 1, cases.length, 5).setValues(cases);

  // 追記した行に数式を適用
  applyFormulasToAllRows();
  SpreadsheetApp.getActive().toast('テスト5件を投入し数式を適用しました。', 'クレームAI分類', 5);
}
