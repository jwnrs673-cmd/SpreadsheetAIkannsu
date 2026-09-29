/**
 * Formatter.gs
 * ------------------------------------------------------------------
 * 段落整形（旧：要約）まわり。AC列は「原文ABを改行だけ調整した全文」。
 *   - AC  : =AI(CL)            … AI関数は単独でセルに置く
 *   - CL  : 整形用結合文        … 段落整形設定の指示文 ＋ 原文AB（通常の数式）
 *   - CM  : 整形チェック        … 原文とACを「CR/LFだけ除いて」EXACT照合（通常の数式）
 * GASの役割：設定シートの用意／選択行の照合（差分位置の表示）／会議用シートへの値転記／表示調整。
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

/** 連続する行番号を「2〜5, 8」のような表記にまとめる。 */
function pfRowsLabel_(rows) {
  var out = [], i = 0;
  while (i < rows.length) {
    var j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
    out.push(i === j ? String(rows[i]) : rows[i] + '〜' + rows[j]);
    i = j + 1;
  }
  return out.join(', ');
}

/* ============================ 段落整形設定シート ============================ */

var PF_SETTINGS_HEADER = ['項目', '値', '説明'];

/** 設定項目（A列の項目名で読み取る。B列の値は利用者が自由に変更してよい）。 */
function pfSettingDefs_() {
  return [
    ['AIへの指示文', PF_DEFAULT_PROMPT,
      '整形用結合文(CL)がこのセルを参照します。変更後はAC列を選んで「更新して挿入」で再生成してください。「改行以外を変更しない」条件は緩めないでください。'],
    ['整形列の列幅（px）', 480, 'ご意見記録のAC列（内容・整形）の列幅。仮の初期値です。'],
    ['会議用：受付日時の列幅（px）', 150, '会議用シートA列の列幅。仮の初期値です。'],
    ['会議用：店舗名の列幅（px）', 160, '会議用シートB列の列幅。仮の初期値です。'],
    ['会議用：本文の列幅（px）', 760, '会議用シートC列の列幅。スクリーンの幅に合わせて調整してください。'],
    ['会議用：フォントサイズ', 14, '会議用シート本文のフォントサイズ。仮の初期値です。'],
    ['長文注意の目安（文字数）', 600, 'これを超える本文は、1画面に収まらない可能性として注意表示します。'],
    ['転記元：受付日時の見出し', '受電日,受付日時,受付日,日付', 'ご意見記録の1行目で、この順に探します（カンマ区切りで複数可）。会議用シートA列に転記する日付です。'],
    ['転記元：店舗名の見出し', '店名,店舗名', 'ご意見記録の1行目で、この順に探します（カンマ区切りで複数可）。']
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
  var empty = sh.getLastRow() === 0;
  if (empty) {
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

/** 整形用結合文(CL)が参照する指示文セルの絶対参照（例: '段落整形設定'!$B$2）。 */
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

/* ============================ 選択行の読み取り・照合 ============================ */

/** ご意見記録で選択中の行を返す（別シートなら null）。 */
function pfSelectedClaimRows_(ss) {
  var sh = ss.getActiveSheet();
  if (sh.getName() !== SHEETS.CLAIM) return null;
  var list = ss.getActiveRangeList();
  var ranges = list ? list.getRanges() : [ss.getActiveRange()];
  return { sheet: sh, rows: pfRowsFromRanges_(ranges) };
}

/** 見出し候補（カンマ区切り）から最初に見つかった列番号。無ければ 0。 */
function pfFindHeaderCol_(sh, candidates) {
  var map = headerIndexMap_(sh).map;
  var list = String(candidates).split(/[,、，]/).map(function (s) { return s.trim(); }).filter(String);
  for (var i = 0; i < list.length; i++) {
    var idx = map[normHeader_(list[i])];
    if (idx) return idx;
  }
  return 0;
}

/**
 * 指定行の原文・整形結果をその場で読み取り照合する（過去の表示は使わない）。
 * @return {Array<{row:number, status:string, reason:string, fmt:string}>}
 */
function pfInspectRows_(sh, rows, COLX) {
  var out = [];
  rows.forEach(function (r) {
    var rawCell = sh.getRange(r, COLX.RAW), fmtCell = sh.getRange(r, COLX.AI_FMT);
    var rawV = rawCell.getValue(), rawD = rawCell.getDisplayValue();
    var fmtV = fmtCell.getValue(), fmtD = fmtCell.getDisplayValue();
    if (pfIsErrorText_(rawD) || pfIsErrorText_(fmtD) || pfIsErrorText_(fmtV)) {
      out.push({ row: r, status: 'エラー', reason: '数式エラー（' + (pfIsErrorText_(fmtD) ? fmtD : rawD) + '）。AC列を選んで「更新して挿入」を試してください。', fmt: '' });
      return;
    }
    var c = pfCompareText_(rawV, fmtV);
    var reason = {
      '対象外': '原文が空です。',
      '未生成': '整形結果がまだありません。AC列を選んで「生成して挿入」を押してください。',
      'OK': '改行以外の文字は原文と一致（段落の自然さは目視で確認してください）。'
    }[c.status] || ('改行以外に差があります：' + c.pos + '文字目付近　原文「' + c.rawCtx + '」→ 結果「' + c.fmtCtx + '」。再生成してください。');
    out.push({ row: r, status: c.status, reason: reason, fmt: c.status === 'OK' ? String(fmtV) : '' });
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
  var list = pfInspectRows_(sel.sheet, sel.rows, res.COLX);
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

/* ============================ 会議用シート ============================ */

var PF_MEETING_HEADER = ['受付日時', '店舗名', '本文（段落整形・確認済み）', '元の行（' + SHEETS.CLAIM + '）'];

/** メニュー：選択行を再照合し、OKの行だけを会議用シートへ値として転記する。 */
function pfBuildMeetingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sel = pfSelectedClaimRows_(ss);
  if (!sel) { ui.alert('「' + SHEETS.CLAIM + '」シートで、会議資料に採用する行を選択してから実行してください（列はどこでも構いません）。'); return; }
  if (!sel.rows.length) { ui.alert('見出し行（1行目）以外を選択してください。'); return; }
  var res;
  try { res = resolveColumns_(sel.sheet); } catch (e) { ui.alert(e.message); return; }
  var settings = pfReadSettings_(ss);
  // 設定の見出し名で見つからなければ、初期値の候補でも探す
  var defs = {};
  pfSettingDefs_().forEach(function (d) { defs[d[0]] = d[1]; });
  var dateCol = pfFindHeaderCol_(sel.sheet, settings['転記元：受付日時の見出し'].value + ',' + defs['転記元：受付日時の見出し']);
  var storeCol = pfFindHeaderCol_(sel.sheet, settings['転記元：店舗名の見出し'].value + ',' + defs['転記元：店舗名の見出し']);

  // 転記直前にその時点の値で再照合（過去のOK表示は使わない）
  var inspected = pfInspectRows_(sel.sheet, sel.rows, res.COLX);
  var ok = inspected.filter(function (x) { return x.status === 'OK'; });
  var ng = inspected.filter(function (x) { return x.status !== 'OK'; });
  var ngText = ng.map(function (x) { return '行' + x.row + '【' + x.status + '】' + x.reason; });

  if (!ok.length) {
    ui.alert('会議用シートを作成できません',
      '転記できる行（原文一致チェックがOKの行）がありません。\n\n' + ngText.slice(0, 30).join('\n'), ui.ButtonSet.OK);
    return;
  }

  var meeting = ss.getSheetByName(SHEETS.MEETING);
  var existing = 0;
  if (meeting) {
    var chk = pfCheckMeetingSheet_(meeting);
    if (!chk.ok) { ui.alert(chk.msg); return; }
    existing = chk.dataRows;
  }

  var msg = '転記する行：' + ok.length + '件（行 ' + pfRowsLabel_(ok.map(function (x) { return x.row; })) + '）\n';
  if (ng.length) msg += '転記しない行：' + ng.length + '件\n' + ngText.slice(0, 20).join('\n') + (ng.length > 20 ? '\n…ほか' + (ng.length - 20) + '件' : '') + '\n';
  if (!dateCol) msg += '\n※受付日時の見出しが見つからないため、受付日時は空欄になります（段落整形設定で見出し名を変更できます）。';
  if (!storeCol) msg += '\n※店舗名の見出しが見つからないため、店舗名は空欄になります（段落整形設定で見出し名を変更できます）。';
  msg += '\n\n' + (existing ? '★「' + SHEETS.MEETING + '」の既存データ ' + existing + '件を、上の内容で置き換えます。' : '「' + SHEETS.MEETING + '」に書き出します。') +
    '\n（' + SHEETS.CLAIM + ' などほかのシートは変更しません）\n\n実行しますか？';
  if (ui.alert('会議用シートの作成', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  var body = ok.map(function (x) {
    return [
      dateCol ? sel.sheet.getRange(x.row, dateCol).getDisplayValue() : '',
      storeCol ? sel.sheet.getRange(x.row, storeCol).getDisplayValue() : '',
      x.fmt
    ];
  });

  if (!meeting) meeting = ss.insertSheet(SHEETS.MEETING);
  var result = pfWriteMeetingSheet_(meeting, body, ok.map(function (x) { return x.row; }));
  pfApplyMeetingLayout_(meeting, settings, body.length);

  var longRows = [];
  var limit = pfNum_(settings, '長文注意の目安（文字数）');
  body.forEach(function (b, i) { if (Array.from(b[2]).length > limit) longRows.push(i + 3); });

  var done = '「' + SHEETS.MEETING + '」に ' + body.length + '件を値として書き出しました（' + result.stamp + ' 時点のスナップショット）。';
  if (result.mismatch.length) done += '\n\n★書き込み後の読み戻しで文字が一致しない行があります（会議用の行 ' + result.mismatch.join(', ') + '）。その行を目視で確認してください。';
  if (ng.length) done += '\n\n転記しなかった行：' + ng.length + '件（行 ' + pfRowsLabel_(ng.map(function (x) { return x.row; })) + '）';
  if (longRows.length) done += '\n\n※長文のため1画面に収まらない可能性がある行（会議用の行 ' + longRows.join(', ') + '）。スクリーン表示で最後まで見えるか確認してください。';
  ss.setActiveSheet(meeting);
  ui.alert('会議用シートを作成しました', done, ui.ButtonSet.OK);
}

/** 会議用シートが本システムの形か確認（違えば触らない）。 */
function pfCheckMeetingSheet_(sh) {
  if (sh.getLastRow() === 0) return { ok: true, dataRows: 0 };
  var a1 = String(sh.getRange(1, 1).getValue()).trim();
  var head = sh.getRange(2, 1, 1, PF_MEETING_HEADER.length).getValues()[0].map(function (v) { return String(v).trim(); });
  if (a1 !== '作成日時' || head.join('|') !== PF_MEETING_HEADER.join('|')) {
    return { ok: false, msg: '「' + SHEETS.MEETING + '」という名前のシートがありますが、このシステムで作った形ではないため変更しませんでした。\n' +
      'シート名を変えるか削除してから再実行してください。' };
  }
  return { ok: true, dataRows: Math.max(sh.getLastRow() - 2, 0) };
}

/**
 * 会議用シートに書き出す。本文は数式として解釈されないようリッチテキスト（文字列）で書き、
 * 読み戻して一致しないセルは先頭アポストロフィ付きで再書き込みし、なお違えば報告する。
 */
function pfWriteMeetingSheet_(sh, body, srcRows) {
  var tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Asia/Tokyo';
  var stamp = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm');
  sh.clear();
  var meta = [['作成日時', stamp, 'この表は作成時点の値のコピー（スナップショット）です。「' + SHEETS.CLAIM + '」を後で変更しても自動では追随しません。', '']];
  sh.getRange(1, 1, 1, 4).setValues(meta);
  sh.getRange(2, 1, 1, 4).setValues([PF_MEETING_HEADER]);

  var rt = function (s) { return SpreadsheetApp.newRichTextValue().setText(String(s)).build(); };
  sh.getRange(3, 3, body.length, 1).setRichTextValues(body.map(function (row) { return [rt(row[2])]; })); // 本文（OKの行は必ず空でない）
  body.forEach(function (row, i) {
    for (var j = 0; j < 2; j++) if (String(row[j]) !== '') sh.getRange(3 + i, 1 + j).setRichTextValue(rt(row[j]));
  });
  sh.getRange(3, 4, body.length, 1).setValues(srcRows.map(function (r) { return [r]; }));

  var mismatch = [];
  var back = sh.getRange(3, 1, body.length, 3).getValues();
  for (var i = 0; i < body.length; i++) {
    for (var j = 0; j < 3; j++) {
      if (String(back[i][j]) === String(body[i][j])) continue;
      var cell = sh.getRange(3 + i, 1 + j);
      cell.setValue("'" + body[i][j]);
      if (String(cell.getValue()) !== String(body[i][j]) && mismatch.indexOf(3 + i) < 0) mismatch.push(3 + i);
    }
  }
  return { stamp: stamp, mismatch: mismatch };
}

function pfApplyMeetingLayout_(sh, settings, n) {
  sh.setFrozenRows(2);
  sh.getRange(1, 1, 1, 4).setFontColor('#5f6368');
  sh.getRange(1, 1).setFontWeight('bold');
  sh.getRange(2, 1, 1, 4).setFontWeight('bold').setBackground('#e8eaed');
  sh.setColumnWidth(1, pfNum_(settings, '会議用：受付日時の列幅（px）'));
  sh.setColumnWidth(2, pfNum_(settings, '会議用：店舗名の列幅（px）'));
  sh.setColumnWidth(3, pfNum_(settings, '会議用：本文の列幅（px）'));
  sh.setColumnWidth(4, 90);
  if (n > 0) {
    var data = sh.getRange(3, 1, n, 4);
    data.setFontSize(pfNum_(settings, '会議用：フォントサイズ'))
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');
    sh.getRange(3, 4, n, 1).setFontSize(10).setFontColor('#5f6368');
    sh.autoResizeRows(3, n);
  }
}

/* ============================ 表示の再調整 ============================ */

/** メニュー：AC列（整形）と会議用シートの表示を設定値に合わせる。本文は変更しない。 */
function pfAdjustDisplay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var settings = pfReadSettings_(ss);
  var notes = [];

  var sh = ss.getSheetByName(SHEETS.CLAIM);
  if (sh) {
    var res;
    try { res = resolveColumns_(sh); } catch (e) { ui.alert(e.message); return; }
    var col = res.COLX.AI_FMT;
    sh.setColumnWidth(col, pfNum_(settings, '整形列の列幅（px）'));
    var last = getLastDataRow_(sh, res.COLX.RAW);
    if (last >= 2) {
      sh.getRange(2, col, last - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top');
    }
    notes.push('「' + SHEETS.CLAIM + '」' + res.CLX.AI_FMT + '列：列幅・折り返し・上詰めを適用');
  }

  var meeting = ss.getSheetByName(SHEETS.MEETING);
  if (meeting) {
    var chk = pfCheckMeetingSheet_(meeting);
    if (chk.ok) {
      pfApplyMeetingLayout_(meeting, settings, chk.dataRows);
      notes.push('「' + SHEETS.MEETING + '」：列幅・フォントサイズ・折り返し・上詰め・行の高さを適用');
      var limit = pfNum_(settings, '長文注意の目安（文字数）');
      if (chk.dataRows) {
        var longRows = [];
        meeting.getRange(3, 3, chk.dataRows, 1).getValues().forEach(function (v, i) {
          if (Array.from(String(v[0])).length > limit) longRows.push(i + 3);
        });
        if (longRows.length) notes.push('※長文の行（' + longRows.join(', ') + '）は1画面に収まらない可能性があります。スクリーン表示で最後まで見えるか確認してください。');
      }
    } else {
      notes.push(chk.msg);
    }
  }
  ui.alert('表示を調整しました', notes.join('\n') || '対象シートがありません。', ui.ButtonSet.OK);
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
