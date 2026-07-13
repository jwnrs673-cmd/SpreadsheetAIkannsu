/**
 * Setup.gs
 * ------------------------------------------------------------------
 * シート作成・マスター投入・数式設定・プルダウン設定・精度検証構築。
 *
 * 【安全設計（★本番シート「ご意見記録」対応）】
 *  - メイン表で触れるのは FORMULA_TARGET_COLS の列（数式）と AK〜AO のプルダウンのみ。
 *  - 原文AB・要約AC・確定AK〜AO(値)・PII(N/O/P)・社員/管理列(BB〜CI)は一切上書きしない。
 *  - 小分類の補助列 CJ/CK を末尾に新設（既存列を1つもずらさない）。
 *  - マスター系シートはシステム管理なので毎回最新化（メイン表とは別シート）。
 *  - 二重実行しても壊れない。原文が無い行にはAI数式を入れない（AI("")防止）。
 * ------------------------------------------------------------------
 */

/** 変更内容の事前サマリ */
const SETUP_PLAN = [
  '【マスター系シート（無ければ作成し最新化）】',
  '  仕訳ルールマスター_大分類 / _中分類 / _小分類',
  '  発生場面マスター / 発生場面ルールマスター',
  '  原因分類マスター / 原因分類ルールマスター',
  '  大分類一覧 / 設定 / 精度検証',
  '',
  '【メイン表「' + SHEETS.CLAIM + '」で触れる列だけ】',
  '  数式を設定: AC(要約), AD, AP, AE, AQ, AR, AF, CJ(新), CK(新), AG, AS, AJ, AT, AH, AU, AI, CL(新)',
  '  プルダウン: AK, AL, AM, AN, AO',
  '  新ヘッダ追加: CJ=小分類候補文, CK=小分類用結合文, CL=要約用結合文',
  '',
  '【絶対に触れない列】AB(原文) / AK〜AO(確定の値) / N,O,P(個人情報) / BB〜CI(社員・管理)',
  '  ※AC(内容・要約)はAI要約を入れるため上書き対象（既存の手入力要約がある場合は要注意）'
];

/* ============================ 実行エントリ ============================ */

/** マスター系シートを構築（メイン表には触れない。安全）。 */
function setupMasters() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(SETUP_PLAN.join('\n'));
  try {
    writeMasterSheet_(ss, SHEETS.RULE_DAI,   MASTER_DAI);
    writeMasterSheet_(ss, SHEETS.RULE_CHU,   MASTER_CHU);
    writeMasterSheet_(ss, SHEETS.RULE_SHO,   MASTER_SHO);
    writeMasterSheet_(ss, SHEETS.SCENE_M,    MASTER_SCENE);
    writeMasterSheet_(ss, SHEETS.SCENE_RULE, MASTER_SCENE_RULE);
    writeMasterSheet_(ss, SHEETS.CAUSE_M,    MASTER_CAUSE);
    writeMasterSheet_(ss, SHEETS.CAUSE_RULE, MASTER_CAUSE_RULE);
    writeDaiList_(ss);
    writeConfigSheet_(ss);
    setupAccuracySheet_(ss);
    stampConfig_(ss);
    SpreadsheetApp.getActive().toast('マスター系シートを構築しました。', 'クレームAI分類', 5);
  } catch (e) {
    Logger.log('setupMasters: エラー ' + e + '\n' + (e.stack || ''));
    SpreadsheetApp.getUi().alert('マスター構築でエラー: ' + e.message);
    throw e;
  }
}

/**
 * メイン表「ご意見記録」に数式・プルダウンを適用する（確認ダイアログ付き）。
 * マスターが未作成なら先に setupMasters を実行する。
 */
function applyMainSheetFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sh = ss.getSheetByName(SHEETS.CLAIM);
  if (!sh) {
    var names = ss.getSheets().map(function (s) { return '・' + s.getName(); }).join('\n');
    ui.alert('メイン表「' + SHEETS.CLAIM + '」が見つかりません。\n\n' +
      'Config.gs の SHEETS.CLAIM を、下の実在シート名のどれかに正確に合わせてください' +
      '（空白・全角半角も一致させる）。\n\n【このブック内のシート】\n' + names);
    return;
  }
  var res = ui.alert('メイン表への数式適用',
    '「' + SHEETS.CLAIM + '」の次の列だけを上書きします。\n' +
    '数式: AC(AI要約),AD,AP,AE,AQ,AR,AF,CJ,CK,AG,AS,AJ,AT,AH,AU,AI,CL\n' +
    'プルダウン: AK〜AO\n' +
    '（原文AB・確定値・個人情報・社員/管理列には触れません。\n' +
    '　AC=要約はAIで上書きします＝既存の手入力要約がある場合はご注意）\n\n実行しますか？',
    ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;

  try {
    if (!ss.getSheetByName(SHEETS.RULE_CHU)) setupMasters(); // マスター未作成なら作る
    setupMainSheet_(ss, sh);
    setupValidations_(ss, sh);
    SpreadsheetApp.getActive().toast('メイン表に数式・プルダウンを適用しました。', 'クレームAI分類', 5);
  } catch (e) {
    Logger.log('applyMainSheetFormulas: エラー ' + e + '\n' + (e.stack || ''));
    ui.alert('数式適用でエラー: ' + e.message);
    throw e;
  }
}

/* ============================ シート基盤 ============================ */

function ensureSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); Logger.log('シート作成: ' + name); }
  return sh;
}

function writeMasterSheet_(ss, name, data) {
  var sh = ensureSheet_(ss, name);
  sh.clearContents();
  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#e8eaed');
  sh.autoResizeColumns(1, data[0].length);
  Logger.log('マスター投入: ' + name + '（' + (data.length - 1) + '件）');
}

function writeDaiList_(ss) {
  var names = MASTER_DAI.slice(1).map(function (r) { return [r[0]]; });
  var sh = ensureSheet_(ss, SHEETS.DAI_LIST);
  sh.clearContents();
  sh.getRange(1, 1, 1, 1).setValues([['大分類']]).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(2, 1, names.length, 1).setValues(names);
  sh.setFrozenRows(1);
}

function writeConfigSheet_(ss) {
  var sh = ensureSheet_(ss, SHEETS.CONFIG);
  sh.clearContents();
  sh.getRange(1, 1, CONFIG_ROWS.length, CONFIG_ROWS[0].length).setValues(CONFIG_ROWS);
  sh.getRange(1, 1, 1, CONFIG_ROWS[0].length).setFontWeight('bold').setBackground('#e8eaed');
  sh.autoResizeColumns(1, CONFIG_ROWS[0].length);
}

function stampConfig_(ss) {
  var sh = ss.getSheetByName(SHEETS.CONFIG);
  var f = sh.createTextFinder('最終セットアップ日時').findNext();
  if (f) sh.getRange(f.getRow(), 2).setValue(new Date());
}

/* ============================ メイン表 ============================ */

/** メイン表に、許可された数式列だけを設定する（原文行のみ）。 */
function setupMainSheet_(ss, sh) {
  // 新規ヘッダ（CJ/CK）だけ設定（他のヘッダには触れない）
  NEW_HEADERS.forEach(function (h) {
    sh.getRange(1, h.col).setValue(h.name).setFontWeight('bold').setBackground('#d9ead3');
  });

  var lastData = getLastDataRow_(sh, COL.RAW); // AB(原文)で最終データ行を判定
  var templateLast = Math.max(lastData, LIMITS.TEMPLATE_ROWS + 1); // 空行にも補助数式を入れる範囲

  // 補助列(結合文等)はテンプレとして空行にも入れる（IFで""になり安全）。
  // AI列(=AI(単一セル))は原文がある行だけに入れる（AI("")防止）。
  FORMULA_TARGET_COLS.forEach(function (key) {
    if (AI_COLS.indexOf(key) >= 0) {
      if (lastData >= 2) setColumnFormulas_(sh, COL[key], 2, lastData, buildFormula_.bind(null, key));
    } else {
      setColumnFormulas_(sh, COL[key], 2, templateLast, buildFormula_.bind(null, key));
    }
  });

  var msg;
  if (lastData >= 2) {
    msg = '数式を設定しました。AI列: 2〜' + lastData + '行 ／ 補助列: 2〜' + templateLast + '行。';
  } else {
    msg = '補助列のテンプレ数式を 2〜' + templateLast + '行に入れました。' +
          'AB列(原文)に内容を入力し、もう一度②を実行するとAI列(AE〜AJ)も入ります。';
  }
  SpreadsheetApp.getActive().toast(msg, 'クレームAI分類', 8);
  Logger.log('setupMainSheet_: ' + msg);
}

function setColumnFormulas_(sh, col, from, to, builder) {
  var n = to - from + 1;
  if (n <= 0) return;
  var arr = [];
  for (var r = from; r <= to; r++) arr.push([builder(r)]);
  sh.getRange(from, col, n, 1).setFormulas(arr);
}

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
 * AI列は必ず =AI(単一セル) のみ。結合文は通常関数で作りAIの外側に置く。
 * 列参照は Config.gs の CL（列レター）を使用。
 */
function buildFormula_(key, r) {
  var AIF = AI_FUNCTION_NAME;
  var NL = 'CHAR(10)';

  // マスター範囲
  var D = SHEETS.RULE_DAI, Cn = SHEETS.RULE_CHU, Sn = SHEETS.RULE_SHO;
  var DAI_A="'"+D+"'!$A$2:$A$"+LIMITS.DAI, DAI_B="'"+D+"'!$B$2:$B$"+LIMITS.DAI,
      DAI_C="'"+D+"'!$C$2:$C$"+LIMITS.DAI, DAI_D="'"+D+"'!$D$2:$D$"+LIMITS.DAI,
      DAI_E="'"+D+"'!$E$2:$E$"+LIMITS.DAI;
  var CHU_A="'"+Cn+"'!$A$2:$A$"+LIMITS.CHU, CHU_B="'"+Cn+"'!$B$2:$B$"+LIMITS.CHU,
      CHU_C="'"+Cn+"'!$C$2:$C$"+LIMITS.CHU, CHU_D="'"+Cn+"'!$D$2:$D$"+LIMITS.CHU,
      CHU_E="'"+Cn+"'!$E$2:$E$"+LIMITS.CHU;
  var SHO_A="'"+Sn+"'!$A$2:$A$"+LIMITS.SHO, SHO_B="'"+Sn+"'!$B$2:$B$"+LIMITS.SHO,
      SHO_C="'"+Sn+"'!$C$2:$C$"+LIMITS.SHO, SHO_D="'"+Sn+"'!$D$2:$D$"+LIMITS.SHO,
      SHO_E="'"+Sn+"'!$E$2:$E$"+LIMITS.SHO;
  var SCN_A="'"+SHEETS.SCENE_M+"'!$A$2:$A$"+LIMITS.SCENE, SCN_B="'"+SHEETS.SCENE_M+"'!$B$2:$B$"+LIMITS.SCENE,
      SCN_R="'"+SHEETS.SCENE_RULE+"'!$A$2:$A$"+LIMITS.SCENE;
  var CAU_A="'"+SHEETS.CAUSE_M+"'!$A$2:$A$"+LIMITS.CAUSE, CAU_B="'"+SHEETS.CAUSE_M+"'!$B$2:$B$"+LIMITS.CAUSE,
      CAU_R="'"+SHEETS.CAUSE_RULE+"'!$A$2:$A$"+LIMITS.CAUSE;

  var AB=CL.RAW, AD=CL.AITEXT, AE=CL.AI_DAI, AF=CL.AI_CHU, AG=CL.AI_SHO, AH=CL.AI_SCENE,
      AP=CL.G_DAI, AQ=CL.I_CHU, AR=CL.J_CHU, AS=CL.L_REASON, AT=CL.P_SCENE, AU=CL.R_CAUSE,
      CJ=CL.V_SHO, CK=CL.W_SHO, CLc=CL.SUM_TXT;

  switch (key) {

    case 'AITEXT': // AD 匿名化
      return '=IF($'+AB+r+'="","",ANONYMIZE($'+AB+r+'))';

    case 'SUM_TXT': // CL 要約用結合文（AC=AI要約の入力プロンプト）
      return '=IF($'+AD+r+'="","",' +
        '"あなたはお客様相談窓口の担当者です。次のご意見を、お客様の立場・目線に立った自然な文章で200文字程度に要約してください。事実関係は変えず、経緯とご要望が伝わるようにする。箇条書きにせず、個人情報や固有の番号は含めない。"&'+NL+'&' +
        '"【ご意見内容】"&'+NL+'&$'+AD+r+'&'+NL+'&"【出力形式】200文字程度の要約文のみ。")';

    case 'AI_SUM': // AC AI要約（200字・顧客目線）
      return '='+AIF+'($'+CLc+r+')';

    case 'G_DAI': // AP 大分類用結合文
      return '=IF($'+AD+r+'="","",' +
        '"あなたは食品スーパー・ドラッグストアのご意見分類担当です。次のご意見を読み、下記の大分類候補から最も適切なものを1つだけ選び、候補の表記どおりに分類名のみを出力してください。説明・記号・前置きは不要です。"&'+NL+'&' +
        '"【判定方針】お客様が最初に強く不満を持った主因で判定する。安易にその他へ寄せない。"&'+NL+'&' +
        '"【大分類候補と定義】"&'+NL+'&' +
        'TEXTJOIN('+NL+',TRUE,ARRAYFORMULA(IF('+DAI_A+'="","","■"&'+DAI_A+'&"｜定義:"&'+DAI_B+'&"｜入る例:"&'+DAI_C+'&"｜しない例:"&'+DAI_D+'&"｜着眼点:"&'+DAI_E+')))&'+NL+'&' +
        '"【ご意見内容】"&'+NL+'&$'+AD+r+'&'+NL+'&"【出力形式】大分類名を1つだけ。")';

    case 'AI_DAI': // AE
      return '='+AIF+'($'+AP+r+')';

    case 'I_CHU': // AQ 中分類候補文（AI大分類でFILTER）
      return '=IF($'+AE+r+'="","",' +
        'TEXTJOIN('+NL+',TRUE,IFERROR(FILTER("■"&'+CHU_B+'&"｜定義:"&'+CHU_C+'&"｜入る例:"&'+CHU_D+'&"｜しない例:"&'+CHU_E+','+CHU_A+'=$'+AE+r+'),"")))';

    case 'J_CHU': // AR 中分類用結合文
      return '=IF($'+AQ+r+'="","",' +
        '"次のご意見について、指定された大分類に属する中分類候補から最も適切なものを1つだけ選び、候補の表記どおりに中分類名のみを出力してください。説明は不要です。"&'+NL+'&' +
        '"【大分類】"&$'+AE+r+'&'+NL+'&"【中分類候補】"&'+NL+'&$'+AQ+r+'&'+NL+'&' +
        '"【判定の注意】従業員個人の爪・髪・制服・名札・清潔感の問題は接客態度(身だしなみ)。店舗の清掃・衛生はクリンリネス。"&'+NL+'&' +
        '"【ご意見内容】"&'+NL+'&$'+AD+r+'&'+NL+'&"【出力形式】中分類名を1つだけ。")';

    case 'AI_CHU': // AF
      return '='+AIF+'($'+AR+r+')';

    case 'V_SHO': // CJ 小分類候補文（大分類×中分類でFILTER）
      return '=IF(OR($'+AE+r+'="",$'+AF+r+'=""),"",' +
        'TEXTJOIN('+NL+',TRUE,IFERROR(FILTER("■"&'+SHO_C+'&"｜定義:"&'+SHO_D+'&"｜入る例:"&'+SHO_E+','+SHO_A+'=$'+AE+r+','+SHO_B+'=$'+AF+r+'),"")))';

    case 'W_SHO': // CK 小分類用結合文
      return '=IF($'+CJ+r+'="","",' +
        '"次のご意見について、指定された大分類・中分類に属する小分類候補から最も適切なものを1つだけ選び、候補の表記どおりに小分類名のみを出力してください。該当が無ければ最も近いものを選ぶ。説明は不要です。"&'+NL+'&' +
        '"【大分類】"&$'+AE+r+'&"／【中分類】"&$'+AF+r+'&'+NL+'&"【小分類候補】"&'+NL+'&$'+CJ+r+'&'+NL+'&' +
        '"【ご意見内容】"&'+NL+'&$'+AD+r+'&'+NL+'&"【出力形式】小分類名を1つだけ。")';

    case 'AI_SHO': // AG
      return '='+AIF+'($'+CK+r+')';

    case 'L_REASON': // AS 補助セル（AI理由用）
      return '=IF($'+AD+r+'="","",' +
        '"次のご意見が、なぜ下記の分類になるのかを40文字以内で簡潔に説明してください。分類名の列挙ではなく理由を述べる。"&'+NL+'&' +
        '"【大分類】"&$'+AE+r+'&"／【中分類】"&$'+AF+r+'&"／【小分類】"&$'+AG+r+'&'+NL+'&' +
        '"【ご意見内容】"&'+NL+'&$'+AD+r+'&'+NL+'&"【出力形式】40文字以内の理由文。")';

    case 'AI_REASON': // AJ
      return '='+AIF+'($'+AS+r+')';

    case 'P_SCENE': // AT 発生場面候補文（結合文兼用）
      return '=IF($'+AD+r+'="","",' +
        '"次のご意見で、お客様の不満が最初に強くなった発生場面を、下記候補から1つだけ選び候補の表記どおりに場面名のみ出力してください。説明は不要です。"&'+NL+'&' +
        '"【発生場面候補】"&'+NL+'&TEXTJOIN('+NL+',TRUE,ARRAYFORMULA(IF('+SCN_A+'="","","・"&'+SCN_A+'&IF('+SCN_B+'=""," "," : "&'+SCN_B+'))))&'+NL+'&' +
        '"【判定ルール】"&TEXTJOIN(" ",TRUE,'+SCN_R+')&'+NL+'&' +
        '"【参考】大分類:"&$'+AE+r+'&"／中分類:"&$'+AF+r+'&'+NL+'&' +
        '"【ご意見内容】"&'+NL+'&$'+AD+r+'&'+NL+'&"【出力形式】場面名を1つだけ。")';

    case 'AI_SCENE': // AH
      return '='+AIF+'($'+AT+r+')';

    case 'R_CAUSE': // AU 原因分類候補文（結合文兼用）
      return '=IF($'+AD+r+'="","",' +
        '"次のご意見について、表面的な不満ではなく社内として直すべき背景要因を、下記候補から1つだけ選び候補の表記どおりに原因名のみ出力してください。説明は不要です。"&'+NL+'&' +
        '"【原因分類候補】"&'+NL+'&TEXTJOIN('+NL+',TRUE,ARRAYFORMULA(IF('+CAU_A+'="","","・"&'+CAU_A+'&IF('+CAU_B+'=""," "," : "&'+CAU_B+'))))&'+NL+'&' +
        '"【判定ルール】"&TEXTJOIN(" ",TRUE,'+CAU_R+')&'+NL+'&' +
        '"【参考】大分類:"&$'+AE+r+'&"／中分類:"&$'+AF+r+'&"／発生場面:"&$'+AH+r+'&'+NL+'&' +
        '"【ご意見内容】"&'+NL+'&$'+AD+r+'&'+NL+'&"【出力形式】原因名を1つだけ。")';

    case 'AI_CAUSE': // AI
      return '='+AIF+'($'+AU+r+')';

    default:
      throw new Error('未知の数式キー: ' + key);
  }
}

/* ============================ プルダウン ============================ */

function setupValidations_(ss, sh) {
  var lastData = getLastDataRow_(sh, COL.RAW);
  var last = Math.max(lastData, LIMITS.TEMPLATE_ROWS + 1); // 空行にもプルダウンを付ける
  var n = last - 1;

  applyListValidation_(sh, COL.FIX_DAI,   2, n, rangeOfColumn_(ss, SHEETS.DAI_LIST, 1));
  applyListValidation_(sh, COL.FIX_CHU,   2, n, rangeOfColumn_(ss, SHEETS.RULE_CHU, 2));
  applyListValidation_(sh, COL.FIX_SHO,   2, n, rangeOfColumn_(ss, SHEETS.RULE_SHO, 3));
  applyListValidation_(sh, COL.FIX_SCENE, 2, n, rangeOfColumn_(ss, SHEETS.SCENE_M, 1));
  applyListValidation_(sh, COL.FIX_CAUSE, 2, n, rangeOfColumn_(ss, SHEETS.CAUSE_M, 1));
  Logger.log('プルダウン設定完了（' + n + '行）');
}

function rangeOfColumn_(ss, sheetName, col) {
  var sh = ss.getSheetByName(sheetName);
  var last = getLastDataRow_(sh, col);
  if (last < 2) last = 2;
  return sh.getRange(2, col, last - 1, 1);
}

function applyListValidation_(sh, col, from, count, sourceRange) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(true) // 表記揺れを手修正できるよう警告のみ
    .build();
  sh.getRange(from, col, count, 1).setDataValidation(rule);
}

/* ============================ 精度検証 ============================ */

function setupAccuracySheet_(ss) {
  var sh = ensureSheet_(ss, SHEETS.ACCURACY);
  sh.clearContents();
  var C = "'" + SHEETS.CLAIM + "'";
  var R = LIMITS.ACC;
  function col(letter){ return C + '!$' + letter + '$2:$' + letter + '$' + R; }

  var AE=col('AE'),AK=col('AK'), AF=col('AF'),AL=col('AL'), AG=col('AG'),AM=col('AM'),
      AH=col('AH'),AN=col('AN'), AI=col('AI'),AO=col('AO'), W=col('W'), AC=col('AC');

  var rows = [];
  rows.push(['クレーム(ご意見)AI分類 精度検証', '', '', '']);
  rows.push(['', '', '', '']);
  rows.push(['分類軸', '確定済み件数', '一致件数', '一致率']);

  function accRow(label, ai, fix) {
    return [label,
      '=SUMPRODUCT(('+fix+'<>"")*1)',
      '=SUMPRODUCT(('+ai+'='+fix+')*('+fix+'<>""))',
      '=IFERROR(C{ROW}/B{ROW},"-")'];
  }
  rows.push(accRow('大分類', AE, AK));
  rows.push(accRow('中分類', AF, AL));
  rows.push(accRow('小分類', AG, AM));
  rows.push(accRow('発生場面', AH, AN));
  rows.push(accRow('原因分類', AI, AO));

  rows.push(['', '', '', '']);
  rows.push(['偏りチェック（AI大分類）', 'AI件数', '全体に対する割合', '']);
  rows.push(['レジ接客への偏り',   '=COUNTIF('+AE+',"レジ接客")',   '=IFERROR(B{ROW}/COUNTA('+AE+'),"-")', '']);
  rows.push(['外周り接客への偏り', '=COUNTIF('+AE+',"外周り接客")', '=IFERROR(B{ROW}/COUNTA('+AE+'),"-")', '']);
  rows.push(['その他への偏り',     '=COUNTIF('+AE+',"その他")',     '=IFERROR(B{ROW}/COUNTA('+AE+'),"-")', '']);

  var values = rows.map(function (row, i) {
    var rowNo = 1 + i;
    return row.map(function (cell) {
      return (typeof cell === 'string') ? cell.replace(/\{ROW\}/g, rowNo) : cell;
    });
  });
  sh.getRange(1, 1, values.length, 4).setValues(values);

  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.getRange(3, 1, 1, 4).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(4, 4, 5, 1).setNumberFormat('0.0%');
  sh.getRange(11, 3, 3, 1).setNumberFormat('0.0%');

  // 誤分類一覧（大分類）: 左 A〜D
  var top = values.length + 3;
  sh.getRange(top, 1).setValue('■ 大分類 誤分類一覧（確定済みで AI≠確定）').setFontWeight('bold');
  sh.getRange(top + 1, 1, 1, 4).setValues([['店名', '内容(要約)', 'AI大分類', '確定大分類']])
    .setFontWeight('bold').setBackground('#f4cccc');
  sh.getRange(top + 2, 1).setFormula(
    '=IFERROR(FILTER({'+W+','+AC+','+AE+','+AK+'},('+AK+'<>"")*('+AE+'<>'+AK+')),"（不一致なし）")');

  // 誤分類一覧（中分類）: 右 F〜I（縦スピル衝突回避）
  sh.getRange(top, 6).setValue('■ 中分類 誤分類一覧（確定済みで AI≠確定）').setFontWeight('bold');
  sh.getRange(top + 1, 6, 1, 4).setValues([['店名', '内容(要約)', 'AI中分類', '確定中分類']])
    .setFontWeight('bold').setBackground('#f4cccc');
  sh.getRange(top + 2, 6).setFormula(
    '=IFERROR(FILTER({'+W+','+AC+','+AF+','+AL+'},('+AL+'<>"")*('+AF+'<>'+AL+')),"（不一致なし）")');

  sh.setColumnWidth(2, 320); sh.setColumnWidth(7, 320);
  Logger.log('精度検証シート構築完了');
}
