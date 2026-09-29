/**
 * Formatter.gs
 * ------------------------------------------------------------------
 * 段落整形（旧：要約）まわり。見出し「内容（要約）」の列（以下、整形列）は
 * 「内容（原文）を改行だけ調整した全文」を入れる列として使う。
 *   - 整形列       : =AI(整形用結合文)   … AI関数は単独でセルに置く
 *   - 整形用結合文 : 段落整形設定の指示文 ＋ 原文（通常の数式）
 *   - 整形チェック : 原文と整形列を「CR/LFだけ除いて」EXACT照合（通常の数式）
 * GASの役割：設定シートの用意／選択行の照合（差分位置の表示）／表示調整。
 * AIの「生成して挿入」「更新して挿入」は利用者が画面で押す（GASからは実行しない）。
 *
 * 他スクリプトと同じプロジェクトに入るため、この機能の名前はすべて pf で始める。
 * 原文・整形結果は Logger に出力しない。
 * ------------------------------------------------------------------
 */

/* ============================ 照合（純粋関数） ============================ */

/** 比較用：CRとLFだけを除く。trim・空白除去・全角半角変換・Unicode正規化はしない。 */
function pfStripNewlines_(text) {
  return String(text).replace(/[\r\n]/g, '');
}

/** セルの表示値がスプレッドシートのエラー表示か。 */
function pfIsErrorText_(text) {
  return /^#(N\/A|REF!|VALUE!|ERROR!|NAME\?|DIV\/0!|NUM!|NULL!|SPILL!|CALC!)/.test(String(text));
}

/**
 * 原文と整形結果を照合する。
 * @return {{status:string, pos?:number, rawCtx?:string, fmtCtx?:string}}
 *   status: '対象外' | '未生成' | 'OK' | 'NG'
 */
function pfCompareText_(raw, fmt) {
  var r = (raw === null || raw === undefined) ? '' : String(raw);
  var f = (fmt === null || fmt === undefined) ? '' : String(fmt);
  if (r === '') return { status: '対象外' };
  if (f === '') return { status: '未生成' };
  var a = pfStripNewlines_(r), b = pfStripNewlines_(f);
  if (a === b) return { status: 'OK' };
  var d = pfFirstDiff_(a, b);
  return { status: 'NG', pos: d.pos, rawCtx: d.rawCtx, fmtCtx: d.fmtCtx };
}

/** 最初に異なる位置（1始まり・改行除去後の文字数。絵文字も1文字）と前後の抜粋。 */
function pfFirstDiff_(a, b) {
  var A = Array.from(a), B = Array.from(b);
  var i = 0;
  while (i < A.length && i < B.length && A[i] === B[i]) i++;
  var from = Math.max(0, i - 5);
  var show = function (arr) {
    var s = arr.slice(from, i + 6).join('');
    return s === '' ? '（ここで終わり）' : s;
  };
  return { pos: i + 1, rawCtx: show(A), fmtCtx: show(B) };
}

/** 選択範囲（複数可）から、見出し行を除いた行番号の昇順リストを作る。 */
function pfRowsFromRanges_(ranges) {
  var seen = {}, rows = [];
  ranges.forEach(function (rg) {
    for (var r = rg.getRow(); r <= rg.getLastRow(); r++) {
      if (r >= 2 && !seen[r]) { seen[r] = 1; rows.push(r); }
    }
  });
  return rows.sort(function (x, y) { return x - y; });
}

/* ============================ 段落整形設定シート ============================ */

var PF_SETTINGS_HEADER = ['項目', '値', '説明'];

/** 設定項目（A列の項目名で読み取る。B列の値は利用者が自由に変更してよい）。 */
function pfSettingDefs_() {
  return [
    ['AIへの指示文', PF_DEFAULT_PROMPT,
      '整形用結合文がこのセルを参照します。変更後は「内容（要約）」列を選んで「更新して挿入」で再生成してください。「改行以外を変更しない」条件は緩めないでください。'],
    ['整形列の列幅（px）', 480, '「内容（要約）」列の列幅（メニュー「表示を再調整」で適用）。仮の初期値です。']
  ];
}

/**
 * 段落整形設定シートを「無ければ作成／あれば不足項目だけ追記」で用意する。
 * 既存の値は上書きしない。見出しが想定と違う同名シートは触らずに false を返す。
 * @return {{ok:boolean, msg:string}}
 */
function pfEnsureSettingsSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.FMT_CONFIG);
  var defs = pfSettingDefs_();
  if (!sh) {
    sh = ss.insertSheet(SHEETS.FMT_CONFIG);
    var rows = [PF_SETTINGS_HEADER].concat(defs);
    sh.getRange(1, 1, rows.length, 3).setValues(rows);
    pfStyleSettingsSheet_(sh);
    return { ok: true, msg: '「' + SHEETS.FMT_CONFIG + '」を作成しました。' };
  }
  var head = sh.getRange(1, 1, 1, 3).getValues()[0].map(function (v) { return String(v).trim(); });
  if (sh.getLastRow() === 0) {
    var rows2 = [PF_SETTINGS_HEADER].concat(defs);
    sh.getRange(1, 1, rows2.length, 3).setValues(rows2);
    pfStyleSettingsSheet_(sh);
    return { ok: true, msg: '「' + SHEETS.FMT_CONFIG + '」に既定値を入れました。' };
  }
  if (head.join('|') !== PF_SETTINGS_HEADER.join('|')) {
    return { ok: false, msg: '「' + SHEETS.FMT_CONFIG + '」という名前のシートがありますが、1行目が「項目｜値｜説明」ではないため変更しませんでした。\n' +
      'シート名を変えるか、1行目を整えてから再実行してください。' };
  }
  var labels = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(function (r) { return String(r[0]).trim(); });
  var add = defs.filter(function (d) { return labels.indexOf(d[0]) < 0; });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  return { ok: true, msg: add.length ? ('設定項目を' + add.length + '件追記しました（既存の値はそのまま）。') : '設定シートは最新です（既存の値はそのまま）。' };
}

function pfStyleSettingsSheet_(sh) {
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 520); sh.setColumnWidth(3, 420);
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).setWrap(true).setVerticalAlignment('top');
}

/** 項目名→{row, value} を読む（シートが無ければ既定値）。 */
function pfReadSettings_(ss) {
  var out = {};
  pfSettingDefs_().forEach(function (d) { out[d[0]] = { row: 0, value: d[1] }; });
  var sh = ss.getSheetByName(SHEETS.FMT_CONFIG);
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  vals.forEach(function (r, i) {
    var k = String(r[0]).trim();
    if (out[k] && out[k].row === 0) {
      out[k] = { row: i + 2, value: (r[1] === '' || r[1] === null) ? out[k].value : r[1] };
    }
  });
  return out;
}

function pfNum_(settings, key) {
  var n = Number(settings[key].value);
  return (isFinite(n) && n > 0) ? n : Number(pfSettingDefs_().filter(function (d) { return d[0] === key; })[0][1]);
}

/** 整形用結合文が参照する指示文セルの絶対参照（例: '段落整形設定'!$B$2）。 */
function pfPromptRef_(ss) {
  var row = 2;
  try {
    var s = pfReadSettings_(ss)['AIへの指示文'];
    if (s && s.row) row = s.row;
  } catch (e) { Logger.log('pfPromptRef_: ' + e); }
  return "'" + SHEETS.FMT_CONFIG + "'!$B$" + row;
}

/** メニュー：段落整形設定シートを用意して開く。 */
function pfOpenSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = pfEnsureSettingsSheet_(ss);
  if (!r.ok) { SpreadsheetApp.getUi().alert(r.msg); return; }
  ss.setActiveSheet(ss.getSheetByName(SHEETS.FMT_CONFIG));
  SpreadsheetApp.getActive().toast(r.msg, '段落整形', 5);
}

/** メニュー：段落整形設定の「AIへの指示文」を、コード内の初期値（PF_DEFAULT_PROMPT）に戻す。 */
function pfResetPrompt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var r = pfEnsureSettingsSheet_(ss);
  if (!r.ok) { ui.alert(r.msg); return; }
  var row = pfReadSettings_(ss)['AIへの指示文'].row || 2;
  var ans = ui.alert('指示文を初期値に戻す',
    '「' + SHEETS.FMT_CONFIG + '」B' + row + 'のAIへの指示文を、コードに書かれた初期値で置き換えます。\n' +
    '（いまの指示文は上書きされます。原文・整形結果には触れません）\n\n実行しますか？', ui.ButtonSet.OK_CANCEL);
  if (ans !== ui.Button.OK) return;
  ss.getSheetByName(SHEETS.FMT_CONFIG).getRange(row, 2).setValue(PF_DEFAULT_PROMPT);
  ui.alert('指示文を初期値に戻しました。',
    '整形列（内容（要約））の範囲を選んで「更新して挿入」を押すと、新しい指示文で作り直されます。', ui.ButtonSet.OK);
}

/* ============================ 選択行の照合 ============================ */

/** ご意見記録で選択中の行を返す（別シートなら null）。 */
function pfSelectedClaimRows_(ss) {
  var sh = ss.getActiveSheet();
  if (sh.getName() !== SHEETS.CLAIM) return null;
  var list = ss.getActiveRangeList();
  var ranges = list ? list.getRanges() : [ss.getActiveRange()];
  return { sheet: sh, rows: pfRowsFromRanges_(ranges) };
}

/**
 * 指定行の原文・整形結果をその場で読み取り照合する（過去の表示は使わない）。
 * @return {Array<{row:number, status:string, reason:string}>}
 */
function pfInspectRows_(sh, rows, res) {
  var col = res.CLX.AI_FMT + '列';
  var out = [];
  rows.forEach(function (r) {
    var rawCell = sh.getRange(r, res.COLX.RAW), fmtCell = sh.getRange(r, res.COLX.AI_FMT);
    var rawV = rawCell.getValue(), rawD = rawCell.getDisplayValue();
    var fmtV = fmtCell.getValue(), fmtD = fmtCell.getDisplayValue();
    if (pfIsErrorText_(rawD) || pfIsErrorText_(fmtD) || pfIsErrorText_(fmtV)) {
      out.push({ row: r, status: 'エラー', reason: '数式エラー（' + (pfIsErrorText_(fmtD) ? fmtD : rawD) + '）。' + col + 'を選んで「更新して挿入」を試してください。' });
      return;
    }
    var c = pfCompareText_(rawV, fmtV);
    var reason = {
      '対象外': '原文が空です。',
      '未生成': '整形結果がまだありません。' + col + 'を選んで「生成して挿入」を押してください。',
      'OK': '改行以外の文字は原文と一致（段落の自然さは目視で確認してください）。'
    }[c.status] || ('改行以外に差があります：' + c.pos + '文字目付近　原文「' + c.rawCtx + '」→ 結果「' + c.fmtCtx + '」。' + col + 'を選んで「更新して挿入」で再生成してください。');
    out.push({ row: r, status: c.status, reason: reason });
  });
  return out;
}

/** メニュー：選択行の整形結果を原文と照合し、結果をダイアログで示す（シートには書き込まない）。 */
function pfCheckSelectedRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sel = pfSelectedClaimRows_(ss);
  if (!sel) { ui.alert('「' + SHEETS.CLAIM + '」シートで、確認したい行を選択してから実行してください（列はどこでも構いません）。'); return; }
  if (!sel.rows.length) { ui.alert('見出し行（1行目）以外を選択してください。'); return; }
  var res;
  try { res = resolveColumns_(sel.sheet); } catch (e) { ui.alert(e.message); return; }
  var list = pfInspectRows_(sel.sheet, sel.rows, res);
  var count = {};
  list.forEach(function (x) { count[x.status] = (count[x.status] || 0) + 1; });
  var lines = list.filter(function (x) { return x.status !== 'OK' && x.status !== '対象外'; })
    .map(function (x) { return '行' + x.row + '【' + x.status + '】' + x.reason; });
  var summary = ['OK', 'NG', '未生成', 'エラー', '対象外'].filter(function (k) { return count[k]; })
    .map(function (k) { return k + ' ' + count[k] + '件'; }).join('　');
  ui.alert('原文一致チェック（選択 ' + list.length + '行）',
    summary + '\n\n' + (lines.length ? lines.slice(0, 30).join('\n') + (lines.length > 30 ? '\n…ほか' + (lines.length - 30) + '件' : '') :
      '要対応の行はありません。') +
    '\n\n※OKは「改行以外の文字が一致」という意味だけです。段落の自然さは目視で確認してください。',
    ui.ButtonSet.OK);
}

/* ============================ 表示の再調整 ============================ */

/** メニュー：整形列の列幅・折り返し・上詰めを設定値に合わせる。本文は変更しない。 */
function pfAdjustDisplay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sh = ss.getSheetByName(SHEETS.CLAIM);
  if (!sh) { ui.alert('「' + SHEETS.CLAIM + '」シートが見つかりません。'); return; }
  var res;
  try { res = resolveColumns_(sh); } catch (e) { ui.alert(e.message); return; }
  var col = res.COLX.AI_FMT;
  sh.setColumnWidth(col, pfNum_(pfReadSettings_(ss), '整形列の列幅（px）'));
  var last = getLastDataRow_(sh, res.COLX.RAW);
  if (last >= 2) {
    sh.getRange(2, col, last - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top');
  }
  SpreadsheetApp.getActive().toast(res.CLX.AI_FMT + '列：列幅・折り返し・上詰めを適用しました。', '段落整形', 5);
}

/* ============================ 整形チェック列の色分け ============================ */

/** 整形チェック列に条件付き書式（NG=赤, エラー=橙, 未生成=灰）。再実行しても重複しない。 */
function pfApplyCheckFormatting_(sh, col, lastRow) {
  var marks = ['NG', 'エラー', '未生成'];
  var kept = sh.getConditionalFormatRules().filter(function (rule) {
    var ranges = rule.getRanges();
    var mine = ranges.length === 1 && ranges[0].getColumn() === col && ranges[0].getNumColumns() === 1;
    var cond = rule.getBooleanCondition();
    var vals = cond ? cond.getCriteriaValues() : [];
    return !(mine && vals.length && marks.indexOf(String(vals[0])) >= 0);
  });
  var rg = sh.getRange(2, col, Math.max(lastRow - 1, 1), 1);
  var make = function (text, bg, fg) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(text)
      .setBackground(bg).setFontColor(fg).setBold(true).setRanges([rg]).build();
  };
  kept.push(make('NG', '#f4cccc', '#990000'), make('エラー', '#fce5cd', '#b45f06'), make('未生成', '#eeeeee', '#666666'));
  sh.setConditionalFormatRules(kept);
}
