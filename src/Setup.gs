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

/**
 * マスター系シートを構築（メイン表には触れない。安全）。
 * ★分類マスターは「無ければ作成、あれば温存」＝手編集した分類は上書きしません。
 *   既定値で作り直したい場合はメニュー「マスターを既定値で再作成」を使用。
 */
function setupMasters() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(SETUP_PLAN.join('\n'));
  try {
    // 分類定義マスターは非破壊（既存の手編集を温存）
    ensureMasterSheet_(ss, SHEETS.RULE_DAI,   MASTER_DAI);
    ensureMasterSheet_(ss, SHEETS.RULE_CHU,   MASTER_CHU);
    ensureMasterSheet_(ss, SHEETS.RULE_SHO,   MASTER_SHO);
    ensureMasterSheet_(ss, SHEETS.SCENE_M,    MASTER_SCENE);
    ensureMasterSheet_(ss, SHEETS.SCENE_RULE, MASTER_SCENE_RULE);
    ensureMasterSheet_(ss, SHEETS.CAUSE_M,    MASTER_CAUSE);
    ensureMasterSheet_(ss, SHEETS.CAUSE_RULE, MASTER_CAUSE_RULE);
    // 補助シートは常に最新化（利用者の手編集は想定しない）
    writeDaiList_(ss);
    writeConfigSheet_(ss);
    setupAccuracySheet_(ss);
    stampConfig_(ss);
    SpreadsheetApp.getActive().toast('マスター系シートを構築しました（既存の分類は温存）。', 'クレームAI分類', 5);
  } catch (e) {
    Logger.log('setupMasters: エラー ' + e + '\n' + (e.stack || ''));
    SpreadsheetApp.getUi().alert('マスター構築でエラー: ' + e.message);
    throw e;
  }
}

/**
 * 分類マスターを Config.gs の既定値で「強制的に」作り直す（確認ダイアログ付き）。
 * ★手編集した分類はすべて破棄されます。既定値へ戻したいときだけ使用。
 */
function resetMastersToDefault() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('マスターを既定値で再作成',
    '分類マスター（大/中/小分類・発生場面・原因分類）を Config.gs の既定値で上書きします。\n' +
    '★シート上で手編集した分類はすべて失われます。よろしいですか？',
    ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;
  try {
    writeMasterSheet_(ss, SHEETS.RULE_DAI,   MASTER_DAI);
    writeMasterSheet_(ss, SHEETS.RULE_CHU,   MASTER_CHU);
    writeMasterSheet_(ss, SHEETS.RULE_SHO,   MASTER_SHO);
    writeMasterSheet_(ss, SHEETS.SCENE_M,    MASTER_SCENE);
    writeMasterSheet_(ss, SHEETS.SCENE_RULE, MASTER_SCENE_RULE);
    writeMasterSheet_(ss, SHEETS.CAUSE_M,    MASTER_CAUSE);
    writeMasterSheet_(ss, SHEETS.CAUSE_RULE, MASTER_CAUSE_RULE);
    writeDaiList_(ss);
    SpreadsheetApp.getActive().toast('マスターを既定値で再作成しました。', 'クレームAI分類', 5);
  } catch (e) {
    Logger.log('resetMastersToDefault: エラー ' + e);
    ui.alert('再作成でエラー: ' + e.message);
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
    var res = resolveColumns_(sh); // ★見出し名から列位置を解決（並び替えに対応）
    setupMainSheet_(sh, res);
    setupValidations_(ss, sh, res);
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

/** マスターを強制投入（既存内容をクリアして最新化）。 */
function writeMasterSheet_(ss, name, data) {
  var sh = ensureSheet_(ss, name);
  sh.clearContents();
  sh.getRange(1, 1, data.length, data[0].length).setValues(data);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#e8eaed');
  sh.autoResizeColumns(1, data[0].length);
  Logger.log('マスター投入: ' + name + '（' + (data.length - 1) + '件）');
}

/** マスターを「無ければ作成／あれば温存」で用意（手編集を壊さない）。 */
function ensureMasterSheet_(ss, name, data) {
  var sh = ss.getSheetByName(name);
  if (sh && getLastDataRow_(sh, 1) >= 2) {   // 既にデータあり → 温存
    Logger.log('マスター温存: ' + name);
    return;
  }
  writeMasterSheet_(ss, name, data);         // 無い/空 → 既定値を投入
}

/** 大分類一覧（参考用）を大分類ルールマスターの現在の内容から再構築。 */
function writeDaiList_(ss) {
  var src = ss.getSheetByName(SHEETS.RULE_DAI);
  var names = null;
  if (src) {
    var last = getLastDataRow_(src, 1);
    if (last >= 2) names = src.getRange(2, 1, last - 1, 1).getValues();
  }
  if (!names || !names.length) names = MASTER_DAI.slice(1).map(function (r) { return [r[0]]; });

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

/* ============================ 列の見出し解決 ============================ */

/** 見出しテキストの正規化（前後空白・内部空白の除去、全角括弧→半角）。 */
function normHeader_(s) {
  return String(s).trim().replace(/\s+/g, '').replace(/（/g, '(').replace(/）/g, ')');
}

/** 列番号(1始まり)→列レター（例: 28→'AB'）。 */
function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m - 1) / 26); }
  return s;
}

/** 1行目の見出しから {正規化見出し: 列番号} を作る。 */
function headerIndexMap_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {}, nonEmpty = 0;
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (String(h).trim() !== '') { nonEmpty++; var k = normHeader_(h); if (map[k] === undefined) map[k] = i + 1; }
  }
  return { map: map, nonEmpty: nonEmpty, lastCol: lastCol };
}

/**
 * ★メイン表の列位置を「見出し名」から解決する。
 *  - 見つかった見出しはその位置を使う（列を並び替えても追従）。
 *  - システム補助列(小分類候補文/小分類用結合文/要約用結合文)が無ければ末尾に新設。
 *  - 見出しが実質空のシートは、従来の固定位置(Config.gs)にフォールバック。
 *  - 必須見出しが見つからない場合はエラー（見出し名の確認を促す）。
 * @return {{COLX:Object, CLX:Object, dynamic:boolean}}
 */
function resolveColumns_(sh) {
  var hi = headerIndexMap_(sh);
  var rawKey = normHeader_(HEADER_TEXT.RAW);

  // 見出しがほぼ無い（新規/空シート）→ 固定位置にフォールバック
  if (hi.nonEmpty === 0 || (hi.map[rawKey] === undefined && hi.nonEmpty < 5)) {
    var COLX0 = {}, CLX0 = {};
    Object.keys(HEADER_TEXT).forEach(function (key) { COLX0[key] = COL[key]; CLX0[key] = colLetter_(COL[key]); });
    Logger.log('resolveColumns_: 見出し未検出のため固定位置を使用');
    return { COLX: COLX0, CLX: CLX0, dynamic: false };
  }

  var COLX = {}, CLX = {}, missing = [];
  var nextCol = hi.lastCol + 1;
  Object.keys(HEADER_TEXT).forEach(function (key) {
    var idx = hi.map[normHeader_(HEADER_TEXT[key])];
    if (idx === undefined) {
      if (SYSTEM_HELPER_KEYS.indexOf(key) >= 0) {
        sh.getRange(1, nextCol).setValue(HEADER_TEXT[key]).setFontWeight('bold').setBackground('#d9ead3');
        idx = nextCol; nextCol++;
        Logger.log('補助列を新設: ' + HEADER_TEXT[key] + '（列' + colLetter_(idx) + '）');
      } else {
        missing.push(HEADER_TEXT[key]);
        return;
      }
    }
    COLX[key] = idx; CLX[key] = colLetter_(idx);
  });

  if (missing.length) {
    throw new Error('メイン表「' + SHEETS.CLAIM + '」に必要な見出しが見つかりません。\n' +
      '1行目の見出し名を確認してください（並び順は自由ですが名称は一致が必要）:\n・' + missing.join('\n・'));
  }
  return { COLX: COLX, CLX: CLX, dynamic: true };
}

/** 固定位置(Config.gsのCOL/CL)から解決マップを作る（フォールバック用）。 */
function staticRes_() {
  var COLX = {}, CLX = {};
  Object.keys(HEADER_TEXT).forEach(function (k) { COLX[k] = COL[k]; CLX[k] = colLetter_(COL[k]); });
  return { COLX: COLX, CLX: CLX, dynamic: false };
}

/* ============================ メイン表 ============================ */

/** 数式ビルダーが参照する列レターマップ（実行時に resolveColumns_ の結果をセット）。 */
var ACTIVE_CL = null;

/** メイン表に、許可された数式列だけを設定する（AI列は原文行のみ）。 */
function setupMainSheet_(sh, res) {
  var COLX = res.COLX;
  ACTIVE_CL = res.CLX; // buildFormula_ がこの列レターで数式を作る

  var lastData = getLastDataRow_(sh, COLX.RAW);
  var templateLast = Math.max(lastData, LIMITS.TEMPLATE_ROWS + 1);

  FORMULA_TARGET_COLS.forEach(function (key) {
    if (AI_COLS.indexOf(key) >= 0) {
      if (lastData >= 2) setColumnFormulas_(sh, COLX[key], 2, lastData, buildFormula_.bind(null, key));
    } else {
      setColumnFormulas_(sh, COLX[key], 2, templateLast, buildFormula_.bind(null, key));
    }
  });

  ACTIVE_CL = null; // 後始末

  var msg;
  if (lastData >= 2) {
    msg = '数式を設定しました。AI列: 2〜' + lastData + '行 ／ 補助列: 2〜' + templateLast + '行'
        + (res.dynamic ? '（見出し名で列を自動解決）' : '（固定位置）') + '。';
  } else {
    msg = '補助列のテンプレ数式を 2〜' + templateLast + '行に入れました。'
        + 'AB(原文)を入力すると、その行が自動分類されます（onEdit）。';
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

  var L = ACTIVE_CL || CL; // 実行時に解決した列レター（無ければ固定既定）
  var AB=L.RAW, AD=L.AITEXT, AE=L.AI_DAI, AF=L.AI_CHU, AG=L.AI_SHO, AH=L.AI_SCENE,
      AP=L.G_DAI, AQ=L.I_CHU, AR=L.J_CHU, AS=L.L_REASON, AT=L.P_SCENE, AU=L.R_CAUSE,
      CJ=L.V_SHO, CK=L.W_SHO, CLc=L.SUM_TXT;

  switch (key) {

    case 'AITEXT': // AD 匿名化（カスタム関数に依存しない純粋数式＝スクリプト不具合の影響を受けない）
      return '=IF($'+AB+r+'="","",' +
        'REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(' +
        '$'+AB+r+',' +
        '"[A-Za-z0-9._%+\\-]+@[A-Za-z0-9.\\-]+\\.[A-Za-z]{2,}","[メール]"),' +
        '"https?://[^\\s　、。]+","[URL]"),' +
        '"〒?\\s?\\d{3}[-－]\\d{4}","[郵便番号]"),' +
        '"0\\d{1,4}[-－(（]?\\d{1,4}[-－)）]?\\d{3,4}","[電話番号]"),' +
        '"(会員番号|会員No|注文番号|受注番号|レシート番号|伝票番号)[ 　]*[:：]?[ 　]*[A-Za-z0-9\\-－]{3,}","$1：[番号]"),' +
        '"(氏名|お名前|お客様名|顧客名|名前)[ 　]*[:：][ 　]*[^ 　、。]{1,20}","$1：[氏名]"),' +
        '"(ご住所|住所|お届け先|届け先)[ 　]*[:：][ 　]*[^ 　、。]{1,40}","$1：[住所]"),' +
        '"\\d{7,}","[番号]"))';

    case 'SUM_TXT': // CL 要約用結合文（AC=AI要約の入力プロンプト）
      return '=IF($'+AD+r+'="","",' +
        '"あなたはお客様相談窓口の担当者です。次のご意見の要点だけを、お客様の立場で自然な文章に要約してください。次を厳守してください。(1)原文に書かれていない情報・推測・「適切に対応いたします」等の定型的な決まり文句を足さない。(2)文字数を埋めるための水増しをしない。内容が短ければ一文で構わない。(3)長くなる場合でも200文字以内（下限なし・短い方が良い）。(4)事実は変えない、箇条書きにしない、個人情報や固有の番号は含めない。"&'+NL+'&' +
        '"【ご意見内容】"&'+NL+'&$'+AD+r+'&'+NL+'&"【出力形式】要点のみの要約文。水増し禁止。最大200文字。")';

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

function setupValidations_(ss, sh, res) {
  var COLX = res.COLX;
  var lastData = getLastDataRow_(sh, COLX.RAW);
  var last = Math.max(lastData, LIMITS.TEMPLATE_ROWS + 1); // 空行にもプルダウンを付ける
  var n = last - 1;

  // 基本（フルリスト）: 親未選択でも何か選べるように全件を入れておく
  applyListValidation_(sh, COLX.FIX_DAI,   2, n, rangeOfColumn_(ss, SHEETS.RULE_DAI, 1));
  applyListValidation_(sh, COLX.FIX_CHU,   2, n, rangeOfColumn_(ss, SHEETS.RULE_CHU, 2));
  applyListValidation_(sh, COLX.FIX_SHO,   2, n, rangeOfColumn_(ss, SHEETS.RULE_SHO, 3));
  applyListValidation_(sh, COLX.FIX_SCENE, 2, n, rangeOfColumn_(ss, SHEETS.SCENE_M, 1));
  applyListValidation_(sh, COLX.FIX_CAUSE, 2, n, rangeOfColumn_(ss, SHEETS.CAUSE_M, 1));

  // 既存行を「現在の確定大分類/中分類」で依存絞り込み（onEditに頼らず即反映）
  narrowExistingRows_(ss, sh, res, lastData);
  Logger.log('プルダウン設定完了（' + n + '行）');
}

/**
 * 既存の各行について、確定大分類(AK)で中分類(AL)候補を、
 * 確定中分類(AL)で小分類(AM)候補を、その行だけに絞り込む。
 */
function narrowExistingRows_(ss, sh, res, lastData) {
  if (lastData < 2) return;
  var COLX = res.COLX;
  var akv = sh.getRange(2, COLX.FIX_DAI, lastData - 1, 1).getValues();
  var alv = sh.getRange(2, COLX.FIX_CHU, lastData - 1, 1).getValues();
  for (var i = 0; i < akv.length; i++) {
    var row = i + 2;
    var dai = String(akv[i][0]).trim(), chu = String(alv[i][0]).trim();
    if (dai) setDependentValidation_(ss, sh, row, COLX.FIX_CHU, SHEETS.RULE_CHU, 1, 2, dai);
    if (dai && chu) setDependentTwoKeyValidation_(ss, sh, row, COLX.FIX_SHO, SHEETS.RULE_SHO, 1, 2, 3, dai, chu);
  }
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

/* ============================ 店番→店舗情報の補完 ============================ */

/** 店番を突合用に正規化（数字のみなら4桁ゼロ埋め。先頭ゼロの取りこぼしを防ぐ）。 */
function normStoreNo_(v) {
  var s = String(v == null ? '' : v).trim();
  if (s === '') return '';
  if (/^\d+$/.test(s)) return s.length < 4 ? ('0000' + s).slice(-4) : s;
  return s;
}

/**
 * 店舗マスタから {正規化店番: {見出し: 値, ...}} の対応表を作る。
 * 店舗マスタが無い/キー列が無い場合は null。
 */
function buildStoreLookup_(ss) {
  var st = ss.getSheetByName(SHEETS.STORE);
  if (!st) return null;
  var hi = headerIndexMap_(st);
  var keyCol = hi.map[normHeader_(STORE_KEY_HEADER)];
  if (!keyCol) return null;

  var srcCol = {};
  STORE_FILL_HEADERS.forEach(function (h) {
    var c = hi.map[normHeader_(h)];
    if (c) srcCol[h] = c;
  });

  var last = getLastDataRow_(st, keyCol);
  if (last < 2) return {};
  var data = st.getRange(2, 1, last - 1, st.getLastColumn()).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var no = normStoreNo_(data[i][keyCol - 1]);
    if (no === '') continue;
    var rec = {};
    STORE_FILL_HEADERS.forEach(function (h) { if (srcCol[h]) rec[h] = data[i][srcCol[h] - 1]; });
    map[no] = rec;
  }
  return map;
}

/**
 * 受付表の from..to 行について、店番から店舗情報(値)を書き込む。
 *  - 店番が空/店舗マスタに該当なしの行は既存値を維持（触れない）。
 *  - dest列は見出しで解決するので列の並び替えに強い。列単位のバッチ書き込みで高速。
 */
function fillStoreInfoRows_(ss, sh, from, to) {
  if (to < from) return;
  var lookup = buildStoreLookup_(ss);
  if (!lookup) return; // 店舗マスタ無し → 何もしない（分類機能は止めない）

  var hi = headerIndexMap_(sh);
  var keyCol = hi.map[normHeader_(STORE_KEY_HEADER)];
  if (!keyCol) return;

  var n = to - from + 1;
  var keys = sh.getRange(from, keyCol, n, 1).getValues();

  STORE_FILL_HEADERS.forEach(function (h) {
    var col = hi.map[normHeader_(h)];
    if (!col) return; // 受付表に該当見出しが無ければスキップ
    var cur = sh.getRange(from, col, n, 1).getValues(); // 既存値を保持
    var changed = false;
    for (var i = 0; i < n; i++) {
      var no = normStoreNo_(keys[i][0]);
      if (no === '') continue;
      var rec = lookup[no];
      if (!rec) continue;
      var v = rec[h];
      if (v !== undefined && v !== '' && cur[i][0] !== v) { cur[i][0] = v; changed = true; }
    }
    if (changed) sh.getRange(from, col, n, 1).setValues(cur);
  });
}

/** 全データ行の店舗情報を店番から一括補完（インポート後などに使用）。 */
function fillAllStoreInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sh = ss.getSheetByName(SHEETS.CLAIM);
  if (!sh) { ui.alert('メイン表「' + SHEETS.CLAIM + '」が見つかりません。'); return; }
  if (!ss.getSheetByName(SHEETS.STORE)) { ui.alert('「' + SHEETS.STORE + '」シートが見つかりません。'); return; }
  var hi = headerIndexMap_(sh);
  var keyCol = hi.map[normHeader_(STORE_KEY_HEADER)];
  if (!keyCol) { ui.alert('受付表に「' + STORE_KEY_HEADER + '」列（見出し）が見つかりません。'); return; }
  var last = getLastDataRow_(sh, keyCol);
  if (last < 2) { SpreadsheetApp.getActive().toast('店番が入力された行がありません。', 'クレームAI分類', 5); return; }
  fillStoreInfoRows_(ss, sh, 2, last);
  SpreadsheetApp.getActive().toast('店舗情報を補完しました（' + (last - 1) + '行）。', 'クレームAI分類', 5);
}

/* ============================ 精度検証 ============================ */

function setupAccuracySheet_(ss) {
  var sh = ensureSheet_(ss, SHEETS.ACCURACY);
  sh.clearContents();
  var C = "'" + SHEETS.CLAIM + "'";
  var R = LIMITS.ACC;

  // メイン表の列レターを見出しから解決（無ければ固定既定にフォールバック）
  var LET = {}; // key -> 列レター
  Object.keys(HEADER_TEXT).forEach(function (k) { LET[k] = colLetter_(COL[k]); });
  var storeLetter = REF_FALLBACK_LETTER.STORE;
  var claim = ss.getSheetByName(SHEETS.CLAIM);
  if (claim) {
    try {
      var hi = headerIndexMap_(claim);
      Object.keys(HEADER_TEXT).forEach(function (k) {
        var idx = hi.map[normHeader_(HEADER_TEXT[k])];
        if (idx) LET[k] = colLetter_(idx);
      });
      var sIdx = hi.map[normHeader_(REF_HEADER_TEXT.STORE)];
      if (sIdx) storeLetter = colLetter_(sIdx);
    } catch (e) { Logger.log('精度検証: 見出し解決失敗のため固定位置を使用 ' + e); }
  }
  function col(letter){ return C + '!$' + letter + '$2:$' + letter + '$' + R; }

  var AE=col(LET.AI_DAI),AK=col(LET.FIX_DAI), AF=col(LET.AI_CHU),AL=col(LET.FIX_CHU),
      AG=col(LET.AI_SHO),AM=col(LET.FIX_SHO), AH=col(LET.AI_SCENE),AN=col(LET.FIX_SCENE),
      AI=col(LET.AI_CAUSE),AO=col(LET.FIX_CAUSE), W=col(storeLetter), AC=col(LET.AI_SUM);

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
