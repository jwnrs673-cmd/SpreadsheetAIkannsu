/**
 * Code.gs
 * ------------------------------------------------------------------
 * メニュー登録・イベント・テスト用ユーティリティ。
 * 実処理は Setup.gs / Anonymize.gs、設定は Config.gs 参照。
 * ------------------------------------------------------------------
 */

/**
 * スプレッドシートを開いたときにメニューを追加。
 * ------------------------------------------------------------------
 * 【重要】Apps Script はプロジェクト内の全ファイルが1つのグローバル領域を
 * 共有するため、onOpen が2つあると後勝ちで片方のメニューが消える。
 * 他スクリプト（例：報告書リンク）と共存させるため、onOpen はこの1つだけにし、
 * 各メニューは「typeof で存在チェックしてから呼ぶ」独立ビルダー関数に分ける。
 * → 他スクリプト側は onOpen を持たず、buildXxxMenu_() を定義するだけでよい。
 * ------------------------------------------------------------------
 */
function onOpen() {
  // 本システムのメニュー
  buildClaimAiMenu_();

  // 他スクリプトのメニュー（存在すれば追加）。順序・有無に依存しない。
  try { if (typeof buildReportLinkMenu_ === 'function') buildReportLinkMenu_(); } catch (e) { Logger.log('buildReportLinkMenu_: ' + e); }
}

/** 本システム「クレームAI分類」メニューを構築。 */
function buildClaimAiMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('クレームAI分類')
    .addItem('① マスター系シートを作成/更新', 'setupMasters')
    .addItem('② メイン表に数式・プルダウンを適用', 'applyMainSheetFormulas')
    .addSeparator()
    .addItem('確定プルダウンを絞り込み直す（全行）', 'refreshDependentDropdowns')
    .addItem('店番から店舗情報を一括補完（全行）', 'fillAllStoreInfo')
    .addSeparator()
    .addItem('テストデータを投入（5ケース）', 'insertTestCases')
    .addItem('変更内容（計画）を表示', 'showSetupPlan')
    .addItem('マスターを既定値で再作成（上書き注意）', 'resetMastersToDefault')
    .addToUi();
}

function showSetupPlan() {
  SpreadsheetApp.getUi().alert(SETUP_PLAN.join('\n'));
}

/**
 * 全行の確定プルダウンを、現在の確定大分類/中分類にあわせて依存絞り込みし直す。
 * 一括貼り付けなどで onEdit が走らなかった場合の手動リフレッシュ用。
 */
function refreshDependentDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.CLAIM);
  if (!sh) { SpreadsheetApp.getUi().alert('メイン表「' + SHEETS.CLAIM + '」が見つかりません。'); return; }
  var res;
  try { res = resolveColumns_(sh); } catch (e) { res = staticRes_(); }
  var lastData = getLastDataRow_(sh, res.COLX.RAW);
  narrowExistingRows_(ss, sh, res, lastData);
  SpreadsheetApp.getActive().toast('確定プルダウンを絞り込み直しました（' + Math.max(lastData - 1, 0) + '行）。', 'クレームAI分類', 5);
}

/**
 * メイン表の編集トリガ（自動化の要）。
 *  1) AB(原文)を入力/貼付 → その行に全数式を自動反映（AI分類が自動で走る）
 *  2) AK(確定大分類)を選ぶ → AL(確定中分類)候補を絞る
 *  3) AL(確定中分類)を選ぶ → AM(確定小分類)候補を絞る
 * 簡易トリガのため追加の権限設定は不要。失敗しても運用は止まらない。
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEETS.CLAIM) return;
    var ss = e.source;
    var r1 = e.range.getRow(), r2 = e.range.getLastRow();
    var c1 = e.range.getColumn(), c2 = e.range.getLastColumn();

    // 見出しから列解決。失敗しても固定位置で続行（プルダウン絞り込みを止めない）
    var res;
    try { res = resolveColumns_(sh); } catch (e2) { res = staticRes_(); }
    var COLX = res.COLX;

    // 1) 原文(内容（原文）)が編集範囲に含まれる → 対象行に数式を自動反映
    if (c1 <= COLX.RAW && COLX.RAW <= c2) {
      autoFillRows_(sh, Math.max(r1, 2), r2, res);
    }

    // 1.5) 店番が編集範囲に含まれる → 店舗情報(店名/営業統括部/営業部/ブロック/エリア)を補完
    try {
      var keyColS = headerIndexMap_(sh).map[normHeader_(STORE_KEY_HEADER)];
      if (keyColS && c1 <= keyColS && keyColS <= c2) {
        fillStoreInfoRows_(ss, sh, Math.max(r1, 2), r2);
      }
    } catch (e3) { Logger.log('店舗補完: ' + e3); }

    // 2)(3) 確定プルダウンの依存絞り込み（単一セル編集時のみ）
    if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && r1 >= 2) {
      if (c1 === COLX.FIX_DAI) {
        setDependentValidation_(ss, sh, r1, COLX.FIX_CHU, SHEETS.RULE_CHU, 1, 2, e.value);
        sh.getRange(r1, COLX.FIX_SHO).clearDataValidations();
      }
      if (c1 === COLX.FIX_CHU) {
        var dai = sh.getRange(r1, COLX.FIX_DAI).getValue();
        setDependentTwoKeyValidation_(ss, sh, r1, COLX.FIX_SHO, SHEETS.RULE_SHO, 1, 2, 3, dai, e.value);
      }
    }
  } catch (err) {
    Logger.log('onEdit: ' + err);
  }
}

/**
 * 指定行範囲のうち、原文が入っている行に全数式を設定する。
 * onEdit から呼ばれ、原文入力だけで自動分類されるようにする。
 */
function autoFillRows_(sh, from, to, res) {
  var COLX = res.COLX;
  ACTIVE_CL = res.CLX; // buildFormula_ が解決済み列レターで数式を作る
  try {
    for (var r = from; r <= to; r++) {
      var ab = sh.getRange(r, COLX.RAW).getValue();
      if (String(ab).trim() === '') continue; // 空行はスキップ（AI("")防止）
      FORMULA_TARGET_COLS.forEach(function (key) {
        sh.getRange(r, COLX[key]).setFormula(buildFormula_(key, r));
      });
    }
  } finally {
    ACTIVE_CL = null;
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

  var res = resolveColumns_(sh); // 見出し名から列位置を解決
  var COLX = res.COLX;
  var last = getLastDataRow_(sh, COLX.RAW);
  var start = Math.max(last + 1, 2);
  // A列(状況)に目印、原文列に本文
  var marks = cases.map(function () { return ['TEST']; });
  var texts = cases.map(function (t) { return [t]; });
  sh.getRange(start, 1, marks.length, 1).setValues(marks);
  sh.getRange(start, COLX.RAW, texts.length, 1).setValues(texts);

  setupMainSheet_(sh, res);
  setupValidations_(ss, sh, res);
  SpreadsheetApp.getActive().toast('テスト5件を投入し数式を適用しました（A列=TEST）。', 'クレームAI分類', 6);
}
