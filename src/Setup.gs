/**
 * Setup.gs
 * ------------------------------------------------------------------
 * シート作成・ヘッダ・マスター投入・数式設定・プルダウン設定・精度検証構築。
 *
 * 【安全設計】
 *  - クレーム記録の A〜E（原文）と N/O/T/U/Y（人の確定値）は絶対に上書きしない。
 *  - 上書きするのは「システム管理列」の数式（F,G,H,I,J,K,L,M,P,Q,R,S,V,W,X）のみ。
 *  - マスターシートはシステム管理なので毎回作り直し（値を最新に保つ）。
 *  - 二重実行しても壊れない（ensureSheet_ で存在チェック、数式は再設定）。
 * ------------------------------------------------------------------
 */

/** どのシート・列を変更するかの事前サマリ（ログにも出す） */
const SETUP_PLAN = [
  '【作成/更新するシート】',
  '  ・クレーム記録        : ヘッダ(A〜Y) と 数式列(F,G,H,I,J,K,L,M,P,Q,R,S,V,W,X) を設定',
  '  ・仕訳ルールマスター_大分類 : 全面投入',
  '  ・仕訳ルールマスター_中分類 : 全面投入',
  '  ・仕訳ルールマスター_小分類 : 全面投入（★準備済み小分類で置換可）',
  '  ・発生場面マスター/ルール    : 全面投入',
  '  ・原因分類マスター/ルール    : 全面投入',
  '  ・大分類一覧          : 全面投入',
  '  ・設定              : 全面投入',
  '  ・精度検証           : 集計数式を設定',
  '【触れない列】クレーム記録 A,B,C,D,E,N,O,T,U,Y（原文と確定値）'
];

/* ============================ 実行エントリ ============================ */

/**
 * すべてを初期構築する。メニュー「クレームAI分類」→「初期セットアップ」から実行。
 */
function setupAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(SETUP_PLAN.join('\n'));
  try {
    // 1) マスター系（システム管理なので全面投入）
    writeMasterSheet_(ss, SHEETS.RULE_DAI,   MASTER_DAI);
    writeMasterSheet_(ss, SHEETS.RULE_CHU,   MASTER_CHU);
    writeMasterSheet_(ss, SHEETS.RULE_SHO,   MASTER_SHO);
    writeMasterSheet_(ss, SHEETS.SCENE_M,    MASTER_SCENE);
    writeMasterSheet_(ss, SHEETS.SCENE_RULE, MASTER_SCENE_RULE);
    writeMasterSheet_(ss, SHEETS.CAUSE_M,    MASTER_CAUSE);
    writeMasterSheet_(ss, SHEETS.CAUSE_RULE, MASTER_CAUSE_RULE);
    writeDaiList_(ss);
    writeConfigSheet_(ss);

    // 2) クレーム記録（安全に：ヘッダと数式のみ）
    setupClaimSheet_(ss);

    // 3) プルダウン
    setupValidations_(ss);

    // 4) 精度検証
    setupAccuracySheet_(ss);

    // 5) 設定シートに実行日時を記録
    stampConfig_(ss);

    SpreadsheetApp.getActive().toast('初期セットアップが完了しました。', 'クレームAI分類', 5);
    Logger.log('setupAll: 正常終了');
  } catch (e) {
    Logger.log('setupAll: エラー ' + e + '\n' + (e.stack || ''));
    SpreadsheetApp.getUi().alert('セットアップでエラー: ' + e.message);
    throw e;
  }
}

/**
 * 既存の全データ行に数式を（再）適用する。データを貼り付けた後に実行する。
 * メニュー「数式を全データ行へ適用」から実行。
 */
function applyFormulasToAllRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    setupClaimSheet_(ss);
    setupValidations_(ss);
    SpreadsheetApp.getActive().toast('数式を全データ行へ適用しました。', 'クレームAI分類', 5);
  } catch (e) {
    Logger.log('applyFormulasToAllRows: エラー ' + e);
    SpreadsheetApp.getUi().alert('数式適用でエラー: ' + e.message);
    throw e;
  }
}

/* ============================ シート基盤 ============================ */

/** シートを取得（無ければ作成）。 */
function ensureSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    Logger.log('シート作成: ' + name);
  }
  return sh;
}

/** マスターシートを全面投入（1行目=ヘッダ）。既存内容はクリアして最新化。 */
function writeMasterSheet_(ss, name, data) {
  var sh = ensureSheet_(ss, name);
  sh.clearContents();
  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#e8eaed');
  sh.autoResizeColumns(1, data[0].length);
  Logger.log('マスター投入: ' + name + '（' + (data.length - 1) + '件）');
}

/** 大分類一覧シート（プルダウン用に11種の名称のみ）。 */
function writeDaiList_(ss) {
  var names = MASTER_DAI.slice(1).map(function (r) { return [r[0]]; });
  var sh = ensureSheet_(ss, SHEETS.DAI_LIST);
  sh.clearContents();
  sh.getRange(1, 1, 1, 1).setValues([['大分類']]).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(2, 1, names.length, 1).setValues(names);
  sh.setFrozenRows(1);
  Logger.log('大分類一覧投入: ' + names.length + '件');
}

/** 設定シート。 */
function writeConfigSheet_(ss) {
  var sh = ensureSheet_(ss, SHEETS.CONFIG);
  sh.clearContents();
  sh.getRange(1, 1, CONFIG_ROWS.length, CONFIG_ROWS[0].length).setValues(CONFIG_ROWS);
  sh.getRange(1, 1, 1, CONFIG_ROWS[0].length).setFontWeight('bold').setBackground('#e8eaed');
  sh.autoResizeColumns(1, CONFIG_ROWS[0].length);
}

/** 設定シートに最終セットアップ日時を記録。 */
function stampConfig_(ss) {
  var sh = ss.getSheetByName(SHEETS.CONFIG);
  var finder = sh.createTextFinder('最終セットアップ日時').findNext();
  if (finder) sh.getRange(finder.getRow(), 2).setValue(new Date());
}

/* ============================ クレーム記録 ============================ */

/**
 * クレーム記録シートを安全に整備する。
 *  - ヘッダ(A〜Y)を設定
 *  - 数式列を設定（原文・確定列は触らない）
 */
function setupClaimSheet_(ss) {
  var sh = ensureSheet_(ss, SHEETS.CLAIM);

  // ヘッダ
  sh.getRange(1, 1, 1, CLAIM_HEADERS.length).setValues([CLAIM_HEADERS])
    .setFontWeight('bold').setBackground('#d9ead3');
  sh.setFrozenRows(1);

  // データ最終行（E列=原文で判定）
  var lastByE = getLastDataRow_(sh, COL.RAW);
  var helperLast = Math.max(lastByE, LIMITS.CLAIM_FORMULA_ROWS + 1); // 補助数式はテンプレとして多めに
  if (helperLast < 2) helperLast = 2;

  // 補助/結合文の数式（空行はIFで""になる＝安全）を 2..helperLast に設定
  setColumnFormulas_(sh, COL.AITEXT, 2, helperLast, buildFormula_.bind(null, 'F'));
  setColumnFormulas_(sh, COL.G_DAI,  2, helperLast, buildFormula_.bind(null, 'G'));
  setColumnFormulas_(sh, COL.I_CHU,  2, helperLast, buildFormula_.bind(null, 'I'));
  setColumnFormulas_(sh, COL.J_CHU,  2, helperLast, buildFormula_.bind(null, 'J'));
  setColumnFormulas_(sh, COL.L_REASON, 2, helperLast, buildFormula_.bind(null, 'L'));
  setColumnFormulas_(sh, COL.P_SCENE, 2, helperLast, buildFormula_.bind(null, 'P'));
  setColumnFormulas_(sh, COL.R_CAUSE, 2, helperLast, buildFormula_.bind(null, 'R'));
  setColumnFormulas_(sh, COL.V_SHO,   2, helperLast, buildFormula_.bind(null, 'V'));
  setColumnFormulas_(sh, COL.W_SHO,   2, helperLast, buildFormula_.bind(null, 'W'));

  // AI関数は「原文がある行だけ」に設定（空行でAI("")を呼ばないため／コスト対策）
  if (lastByE >= 2) {
    setColumnFormulas_(sh, COL.AI_DAI,    2, lastByE, buildFormula_.bind(null, 'H'));
    setColumnFormulas_(sh, COL.AI_CHU,    2, lastByE, buildFormula_.bind(null, 'K'));
    setColumnFormulas_(sh, COL.AI_REASON, 2, lastByE, buildFormula_.bind(null, 'M'));
    setColumnFormulas_(sh, COL.AI_SCENE,  2, lastByE, buildFormula_.bind(null, 'Q'));
    setColumnFormulas_(sh, COL.AI_CAUSE,  2, lastByE, buildFormula_.bind(null, 'S'));
    setColumnFormulas_(sh, COL.AI_SHO,    2, lastByE, buildFormula_.bind(null, 'X'));
    Logger.log('AI数式を ' + (lastByE - 1) + ' 行へ設定');
  } else {
    Logger.log('原文データが無いためAI数式は未設定（データ投入後に「数式を全データ行へ適用」を実行）');
  }
}

/** 指定列の from..to 行に、行ごとの数式を一括設定する。 */
function setColumnFormulas_(sh, col, from, to, builder) {
  var n = to - from + 1;
  if (n <= 0) return;
  var arr = [];
  for (var r = from; r <= to; r++) arr.push([builder(r)]);
  sh.getRange(from, col, n, 1).setFormulas(arr);
}

/** 指定列で値が入っている最終行を返す（ヘッダ行1は含まない）。 */
function getLastDataRow_(sh, col) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

/* ============================ 数式ビルダー ============================ */
/*
 * すべて「AI関数には単一セルのみ参照させる」方針。
 * 補助列(G/I/J/L/P/R/V/W)で結合文を作り、AI列(H/K/M/Q/S/X)は =AI(単一セル) だけ。
 */

// シート範囲文字列を作るヘルパ
function rng_(sheet, colFrom, colTo, rowFrom, rowTo) {
  return "'" + sheet + "'!$" + colFrom + "$" + rowFrom + ":$" + colTo + "$" + rowTo;
}

function buildFormula_(key, r) {
  var AIF = AI_FUNCTION_NAME; // 'AI' または 'Gemini'
  var NL = 'CHAR(10)';

  // 大分類ルール範囲（A..E, 2..LIMITS.DAI）
  var DAI_A = "'" + SHEETS.RULE_DAI + "'!$A$2:$A$" + LIMITS.DAI;
  var DAI_B = "'" + SHEETS.RULE_DAI + "'!$B$2:$B$" + LIMITS.DAI;
  var DAI_C = "'" + SHEETS.RULE_DAI + "'!$C$2:$C$" + LIMITS.DAI;
  var DAI_D = "'" + SHEETS.RULE_DAI + "'!$D$2:$D$" + LIMITS.DAI;
  var DAI_E = "'" + SHEETS.RULE_DAI + "'!$E$2:$E$" + LIMITS.DAI;

  // 中分類マスター範囲（A大分類,B中分類,C定義,D入る例,E除外例）
  var CHU = SHEETS.RULE_CHU;
  var CHU_A = "'" + CHU + "'!$A$2:$A$" + LIMITS.CHU;
  var CHU_B = "'" + CHU + "'!$B$2:$B$" + LIMITS.CHU;
  var CHU_C = "'" + CHU + "'!$C$2:$C$" + LIMITS.CHU;
  var CHU_D = "'" + CHU + "'!$D$2:$D$" + LIMITS.CHU;
  var CHU_E = "'" + CHU + "'!$E$2:$E$" + LIMITS.CHU;

  // 小分類マスター範囲（A大,B中,C小,D定義,E入る例,F除外例）
  var SHO = SHEETS.RULE_SHO;
  var SHO_A = "'" + SHO + "'!$A$2:$A$" + LIMITS.SHO;
  var SHO_B = "'" + SHO + "'!$B$2:$B$" + LIMITS.SHO;
  var SHO_C = "'" + SHO + "'!$C$2:$C$" + LIMITS.SHO;
  var SHO_D = "'" + SHO + "'!$D$2:$D$" + LIMITS.SHO;
  var SHO_E = "'" + SHO + "'!$E$2:$E$" + LIMITS.SHO;

  // 発生場面
  var SCN = SHEETS.SCENE_M;
  var SCN_A = "'" + SCN + "'!$A$2:$A$" + LIMITS.SCENE;
  var SCN_B = "'" + SCN + "'!$B$2:$B$" + LIMITS.SCENE;
  var SCN_RULE = "'" + SHEETS.SCENE_RULE + "'!$A$2:$A$" + LIMITS.SCENE;

  // 原因分類
  var CAU = SHEETS.CAUSE_M;
  var CAU_A = "'" + CAU + "'!$A$2:$A$" + LIMITS.CAUSE;
  var CAU_B = "'" + CAU + "'!$B$2:$B$" + LIMITS.CAUSE;
  var CAU_RULE = "'" + SHEETS.CAUSE_RULE + "'!$A$2:$A$" + LIMITS.CAUSE;

  switch (key) {

    // F: 匿名化（GASカスタム関数）
    case 'F':
      return '=IF($E' + r + '="","",ANONYMIZE($E' + r + '))';

    // G: 大分類用結合文
    case 'G':
      return '=IF($F' + r + '="","",' +
        '"あなたは食品スーパー・ディスカウントストアのクレーム分類担当です。次のクレーム内容を読み、下記の大分類候補の中から最も適切なものを1つだけ選び、候補の表記どおりに分類名のみを出力してください。説明・記号・前置きは不要です。"&' + NL + '&' +
        '"【判定方針】お客様が最初に強く不満を持った主因で判定する。安易に接客・応対やその他へ寄せない。"&' + NL + '&' +
        '"【大分類候補と定義】"&' + NL + '&' +
        'TEXTJOIN(' + NL + ',TRUE,ARRAYFORMULA(IF(' + DAI_A + '="","","■"&' + DAI_A + '&"｜定義:"&' + DAI_B + '&"｜入る例:"&' + DAI_C + '&"｜しない例:"&' + DAI_D + '&"｜着眼点:"&' + DAI_E + ')))&' + NL + '&' +
        '"【クレーム内容】"&' + NL + '&$F' + r + '&' + NL + '&' +
        '"【出力形式】大分類名を1つだけ。")';

    // H: AI大分類（単一セルのみ）
    case 'H':
      return '=' + AIF + '($G' + r + ')';

    // I: 中分類候補文（AI大分類でFILTER。範囲サイズは一致）
    case 'I':
      return '=IF($H' + r + '="","",' +
        'TEXTJOIN(' + NL + ',TRUE,IFERROR(FILTER(' +
        '"■"&' + CHU_B + '&"｜定義:"&' + CHU_C + '&"｜入る例:"&' + CHU_D + '&"｜しない例:"&' + CHU_E + ',' +
        CHU_A + '=$H' + r + '),"")))';

    // J: 中分類用結合文（候補・ルール・AI大分類・AI投入用テキスト）
    case 'J':
      return '=IF($I' + r + '="","",' +
        '"次のクレーム内容について、指定された大分類に属する中分類候補の中から最も適切なものを1つだけ選び、候補の表記どおりに中分類名のみを出力してください。説明は不要です。"&' + NL + '&' +
        '"【大分類】"&$H' + r + '&' + NL + '&' +
        '"【中分類候補】"&' + NL + '&$I' + r + '&' + NL + '&' +
        '"【判定の注意】従業員個人の爪・髪・制服・名札・清潔感の問題は身だしなみ不備。商品の異物混入や店舗全体の衛生は身だしなみ不備に含めない。"&' + NL + '&' +
        '"【クレーム内容】"&' + NL + '&$F' + r + '&' + NL + '&' +
        '"【出力形式】中分類名を1つだけ。")';

    // K: AI中分類
    case 'K':
      return '=' + AIF + '($J' + r + ')';

    // V: 小分類候補文（AI大分類＋AI中分類でFILTER。2条件は同サイズ）
    case 'V':
      return '=IF(OR($H' + r + '="",$K' + r + '=""),"",' +
        'TEXTJOIN(' + NL + ',TRUE,IFERROR(FILTER(' +
        '"■"&' + SHO_C + '&"｜定義:"&' + SHO_D + '&"｜入る例:"&' + SHO_E + ',' +
        SHO_A + '=$H' + r + ',' + SHO_B + '=$K' + r + '),"")))';

    // W: 小分類用結合文
    case 'W':
      return '=IF($V' + r + '="","",' +
        '"次のクレーム内容について、指定された大分類・中分類に属する小分類候補の中から最も適切なものを1つだけ選び、候補の表記どおりに小分類名のみを出力してください。該当が無ければ最も近いものを選ぶ。説明は不要です。"&' + NL + '&' +
        '"【大分類】"&$H' + r + '&"／【中分類】"&$K' + r + '&' + NL + '&' +
        '"【小分類候補】"&' + NL + '&$V' + r + '&' + NL + '&' +
        '"【クレーム内容】"&' + NL + '&$F' + r + '&' + NL + '&' +
        '"【出力形式】小分類名を1つだけ。")';

    // X: AI小分類
    case 'X':
      return '=' + AIF + '($W' + r + ')';

    // L: AI理由用結合文
    case 'L':
      return '=IF($F' + r + '="","",' +
        '"次のクレームが、なぜ下記の分類になるのかを40文字以内で簡潔に説明してください。分類名の列挙ではなく理由を述べる。"&' + NL + '&' +
        '"【大分類】"&$H' + r + '&"／【中分類】"&$K' + r + '&"／【小分類】"&$X' + r + '&' + NL + '&' +
        '"【クレーム内容】"&' + NL + '&$F' + r + '&' + NL + '&' +
        '"【出力形式】40文字以内の理由文。")';

    // M: AI理由
    case 'M':
      return '=' + AIF + '($L' + r + ')';

    // P: 発生場面候補文
    case 'P':
      return '=IF($F' + r + '="","",' +
        '"次のクレームで、お客様の不満が最初に強くなった発生場面を、下記候補から1つだけ選び候補の表記どおりに場面名のみ出力してください。説明は不要です。"&' + NL + '&' +
        '"【発生場面候補】"&' + NL + '&TEXTJOIN(' + NL + ',TRUE,ARRAYFORMULA(IF(' + SCN_A + '="","","・"&' + SCN_A + '&IF(' + SCN_B + '=""," "," : "&' + SCN_B + '))))&' + NL + '&' +
        '"【判定ルール】"&TEXTJOIN(" ",TRUE,' + SCN_RULE + ')&' + NL + '&' +
        '"【参考】大分類:"&$H' + r + '&"／中分類:"&$K' + r + '&' + NL + '&' +
        '"【クレーム内容】"&' + NL + '&$F' + r + '&' + NL + '&' +
        '"【出力形式】場面名を1つだけ。")';

    // Q: AI発生場面
    case 'Q':
      return '=' + AIF + '($P' + r + ')';

    // R: 原因分類候補文
    case 'R':
      return '=IF($F' + r + '="","",' +
        '"次のクレームについて、顧客が訴えた表面的な不満ではなく社内として直すべき背景要因を、下記候補から1つだけ選び候補の表記どおりに原因名のみ出力してください。説明は不要です。"&' + NL + '&' +
        '"【原因分類候補】"&' + NL + '&TEXTJOIN(' + NL + ',TRUE,ARRAYFORMULA(IF(' + CAU_A + '="","","・"&' + CAU_A + '&IF(' + CAU_B + '=""," "," : "&' + CAU_B + '))))&' + NL + '&' +
        '"【判定ルール】"&TEXTJOIN(" ",TRUE,' + CAU_RULE + ')&' + NL + '&' +
        '"【参考】大分類:"&$H' + r + '&"／中分類:"&$K' + r + '&"／発生場面:"&$Q' + r + '&' + NL + '&' +
        '"【クレーム内容】"&' + NL + '&$F' + r + '&' + NL + '&' +
        '"【出力形式】原因名を1つだけ。")';

    // S: AI原因分類
    case 'S':
      return '=' + AIF + '($R' + r + ')';

    default:
      throw new Error('未知の数式キー: ' + key);
  }
}

/* ============================ プルダウン ============================ */

/**
 * 確定列（N,O,T,U,Y）にプルダウンを設定する。
 * 依存プルダウンは onEdit（Code.gs）で任意対応。ここでは全件リストで確実に動く形にする。
 */
function setupValidations_(ss) {
  var sh = ss.getSheetByName(SHEETS.CLAIM);
  var rows = Math.max(getLastDataRow_(sh, COL.RAW), LIMITS.CLAIM_FORMULA_ROWS + 1);
  if (rows < 2) rows = 2;
  var n = rows - 1;

  // 各マスターの実データ範囲
  var daiRange   = rangeOfColumn_(ss, SHEETS.DAI_LIST, 1, 1);   // 大分類一覧 A
  var chuRange   = rangeOfColumn_(ss, SHEETS.RULE_CHU, 2, 1);   // 中分類 B
  var shoRange   = rangeOfColumn_(ss, SHEETS.RULE_SHO, 3, 1);   // 小分類 C
  var sceneRange = rangeOfColumn_(ss, SHEETS.SCENE_M, 1, 1);    // 発生場面 A
  var causeRange = rangeOfColumn_(ss, SHEETS.CAUSE_M, 1, 1);    // 原因分類 A

  applyListValidation_(sh, COL.FIX_DAI,   2, n, daiRange);
  applyListValidation_(sh, COL.FIX_CHU,   2, n, chuRange);
  applyListValidation_(sh, COL.FIX_SCENE, 2, n, sceneRange);
  applyListValidation_(sh, COL.FIX_CAUSE, 2, n, causeRange);
  applyListValidation_(sh, COL.FIX_SHO,   2, n, shoRange);
  Logger.log('プルダウン設定完了（' + n + '行）');
}

/** マスターシートの指定列の実データ範囲を返す（ヘッダ除く）。 */
function rangeOfColumn_(ss, sheetName, col, startOffsetRow) {
  var sh = ss.getSheetByName(sheetName);
  var last = getLastDataRow_(sh, col);
  if (last < 2) last = 2;
  return sh.getRange(2, col, last - 1, 1);
}

/** 指定列にリスト（範囲参照）のデータ検証を設定。 */
function applyListValidation_(sh, col, from, count, sourceRange) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(true) // AIの表記揺れを手修正できるよう警告のみ
    .build();
  sh.getRange(from, col, count, 1).setDataValidation(rule);
}

/* ============================ 精度検証 ============================ */

function setupAccuracySheet_(ss) {
  var sh = ensureSheet_(ss, SHEETS.ACCURACY);
  sh.clearContents();
  var C = "'" + SHEETS.CLAIM + "'";
  var R = 1000; // 集計対象の上限行

  // 列レター（クレーム記録）
  var H = C + '!$H$2:$H$' + R, N = C + '!$N$2:$N$' + R; // 大分類 AI/確定
  var K = C + '!$K$2:$K$' + R, O = C + '!$O$2:$O$' + R; // 中分類
  var X = C + '!$X$2:$X$' + R, Y = C + '!$Y$2:$Y$' + R; // 小分類
  var Q = C + '!$Q$2:$Q$' + R, T = C + '!$T$2:$T$' + R; // 発生場面
  var S = C + '!$S$2:$S$' + R, U = C + '!$U$2:$U$' + R; // 原因分類

  var rows = [];
  rows.push(['クレームAI分類 精度検証', '', '', '']);
  rows.push(['', '', '', '']);
  rows.push(['分類軸', '確定済み件数', '一致件数', '一致率']);

  // 一致率行（AI vs 確定、確定が入っている行のみ対象）
  function accRow(label, ai, fix) {
    var target = '=SUMPRODUCT((' + fix + '<>"")*1)';
    var match  = '=SUMPRODUCT((' + ai + '=' + fix + ')*(' + fix + '<>""))';
    var rate   = '=IFERROR(C{ROW}/B{ROW},"-")';
    return [label, target, match, rate];
  }
  rows.push(accRow('大分類', H, N));
  rows.push(accRow('中分類', K, O));
  rows.push(accRow('小分類', X, Y));
  rows.push(accRow('発生場面', Q, T));
  rows.push(accRow('原因分類', S, U));

  rows.push(['', '', '', '']);
  rows.push(['偏りチェック（AI大分類）', 'AI件数', '全体に対する割合', '']);
  rows.push(['接客・応対への偏り',
    '=COUNTIF(' + H + ',"接客・応対")',
    '=IFERROR(B{ROW}/COUNTA(' + H + '),"-")', '']);
  rows.push(['その他への偏り',
    '=COUNTIF(' + H + ',"その他")',
    '=IFERROR(B{ROW}/COUNTA(' + H + '),"-")', '']);

  // 書き込み（{ROW} を実行時の行番号へ置換）
  var startRow = 1;
  var values = rows.map(function (row, i) {
    var rowNo = startRow + i;
    return row.map(function (cell) {
      return (typeof cell === 'string') ? cell.replace(/\{ROW\}/g, rowNo) : cell;
    });
  });
  sh.getRange(startRow, 1, values.length, 4).setValues(values);

  // 見出し装飾
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.getRange(3, 1, 1, 4).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(4, 4, 5, 1).setNumberFormat('0.0%');
  sh.getRange(10, 3, 3, 1).setNumberFormat('0.0%');

  // 誤分類一覧（大分類）: 確定があり AI≠確定 の行を抽出
  var listTop = values.length + 3;
  sh.getRange(listTop, 1).setValue('■ 大分類 誤分類一覧（確定済みで AI≠確定）').setFontWeight('bold');
  sh.getRange(listTop + 1, 1, 1, 4).setValues([['受付No', 'クレーム内容原文', 'AI大分類', '確定大分類']])
    .setFontWeight('bold').setBackground('#f4cccc');
  var A = C + '!$A$2:$A$' + R, E = C + '!$E$2:$E$' + R;
  var filterFormula =
    '=IFERROR(FILTER({' + A + ',' + E + ',' + H + ',' + N + '},' +
    '(' + N + '<>"")*(' + H + '<>' + N + ')),"（不一致なし）")';
  sh.getRange(listTop + 2, 1).setFormula(filterFormula);

  // 中分類 誤分類一覧（縦に伸びるFILTER同士が衝突しないよう F〜I 列へ横並び配置）
  sh.getRange(listTop, 6).setValue('■ 中分類 誤分類一覧（確定済みで AI≠確定）').setFontWeight('bold');
  sh.getRange(listTop + 1, 6, 1, 4).setValues([['受付No', 'クレーム内容原文', 'AI中分類', '確定中分類']])
    .setFontWeight('bold').setBackground('#f4cccc');
  var filterFormula2 =
    '=IFERROR(FILTER({' + A + ',' + E + ',' + K + ',' + O + '},' +
    '(' + O + '<>"")*(' + K + '<>' + O + ')),"（不一致なし）")';
  sh.getRange(listTop + 2, 6).setFormula(filterFormula2);

  sh.autoResizeColumns(1, 4);
  sh.setColumnWidth(2, 320);
  sh.setColumnWidth(7, 320);
  Logger.log('精度検証シート構築完了');
}
