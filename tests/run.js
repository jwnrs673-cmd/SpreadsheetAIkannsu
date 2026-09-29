#!/usr/bin/env node
// 段落整形まわりのテスト。dist の統合版を模擬 Apps Script 環境で読み込んで検証する。
//   node tests/run.js
// 模擬環境は実在する Apps Script のメソッド名だけを用意し、それ以外を呼ぶと失敗する。
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const DIST = path.join(ROOT, 'dist', 'クレームAI分類_統合版.gs');
const CASES = JSON.parse(fs.readFileSync(path.join(ROOT, 'testdata', '段落整形_架空ご意見.json'), 'utf8'));

let pass = 0, fail = 0;
function t(name, fn) {
  try { fn(); pass++; console.log('  ok  ' + name); }
  catch (e) { fail++; console.log('  NG  ' + name + '\n        ' + (e && e.message)); }
}
function eq(a, b, msg) { if (a !== b) throw new Error((msg || '') + ' 期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); }
function ok(c, msg) { if (!c) throw new Error(msg || '条件が偽'); }

/* ============================ 模擬 Apps Script ============================ */

const LOGS = [];
function makeEnv() {
  const ui = {
    alerts: [], next: 'OK',
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' },
    Button: { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO' },
    alert(a, b) { this.alerts.push(b === undefined ? String(a) : String(a) + '\n' + String(b)); return this.next; },
    createMenu() { const m = { addItem: () => m, addSeparator: () => m, addSubMenu: () => m, addToUi: () => m }; return m; }
  };

  class Sheet {
    constructor(name, ss) { this.name = name; this.ss = ss; this.v = new Map(); this.f = new Map(); this.cf = []; this.formulaWrites = []; }
    k(r, c) { return r + ',' + c; }
    getName() { return this.name; }
    getParent() { return this.ss; }
    _cells() { const s = new Set(); for (const [k, x] of this.v) if (x !== '') s.add(k); for (const [k, x] of this.f) if (x !== '') s.add(k); return [...s].map((k) => k.split(',').map(Number)); }
    getLastRow() { return this._cells().reduce((m, [r]) => Math.max(m, r), 0); }
    getLastColumn() { return this._cells().reduce((m, [, c]) => Math.max(m, c), 0); }
    getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
    get(r, c) { return this.v.has(this.k(r, c)) ? this.v.get(this.k(r, c)) : ''; }
    formula(r, c) { return this.f.get(this.k(r, c)) || ''; }
    put(r, c, x) { // 利用者の入力と同じ解釈（' で始まれば文字列、= で始まれば数式）
      const k = this.k(r, c);
      if (typeof x === 'string' && x.startsWith("'")) { this.v.set(k, x.slice(1)); this.f.set(k, ''); }
      else if (typeof x === 'string' && x.startsWith('=')) { this.f.set(k, x); this.v.set(k, ''); }
      else { this.v.set(k, x); this.f.set(k, ''); }
    }
    gen(r, c, text) { this.v.set(this.k(r, c), text); } // AIの「生成して挿入」を模擬（数式は残る）
    setFrozenRows() {} setColumnWidth() {} autoResizeRows() {} autoResizeColumns() {}
    getConditionalFormatRules() { return this.cf.slice(); }
    setConditionalFormatRules(rules) { this.cf = rules.slice(); }
    clear() { this.v.clear(); this.f.clear(); }
    clearContents() { this.v.clear(); this.f.clear(); }
    insertRowBefore(r) { // テスト用：行挿入（下へずらす）
      const nv = new Map(), nf = new Map();
      for (const [k, x] of this.v) { const [rr, cc] = k.split(',').map(Number); nv.set((rr >= r ? rr + 1 : rr) + ',' + cc, x); }
      for (const [k, x] of this.f) { const [rr, cc] = k.split(',').map(Number); nf.set((rr >= r ? rr + 1 : rr) + ',' + cc, x); }
      this.v = nv; this.f = nf;
    }
  }

  class Range {
    constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
    getRow() { return this.r; } getColumn() { return this.c; }
    getLastRow() { return this.r + this.nr - 1; } getLastColumn() { return this.c + this.nc - 1; }
    getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
    _map(fn) { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(fn(this.r + i, this.c + j)); out.push(row); } return out; }
    getValues() { return this._map((r, c) => this.sh.get(r, c)); }
    getValue() { return this.sh.get(this.r, this.c); }
    getDisplayValues() { return this._map((r, c) => String(this.sh.get(r, c))); }
    getDisplayValue() { return String(this.sh.get(this.r, this.c)); }
    getFormulas() { return this._map((r, c) => this.sh.formula(r, c)); }
    getFormula() { return this.sh.formula(this.r, this.c); }
    setValues(a) { a.forEach((row, i) => row.forEach((x, j) => this.sh.put(this.r + i, this.c + j, x))); return this; }
    setValue(x) { this.sh.put(this.r, this.c, x); return this; }
    setFormulas(a) { a.forEach((row, i) => row.forEach((x, j) => { this.sh.put(this.r + i, this.c + j, x); this.sh.formulaWrites.push([this.r + i, this.c + j]); })); return this; }
    setFormula(x) { this.sh.put(this.r, this.c, x); this.sh.formulaWrites.push([this.r, this.c]); return this; }
    setRichTextValues(a) { // リッチテキストは文字列そのまま（数式として解釈しない）
      a.forEach((row, i) => row.forEach((rt, j) => { const k = this.sh.k(this.r + i, this.c + j); this.sh.v.set(k, rt.getText()); this.sh.f.set(k, ''); }));
      return this;
    }
    setRichTextValue(rt) { return this.setRichTextValues([[rt]]); }
    setFontWeight() { return this; } setBackground() { return this; } setWrap() { return this; }
    setWrapStrategy() { return this; } setVerticalAlignment() { return this; } setFontSize() { return this; }
    setFontColor() { return this; } setNumberFormat() { return this; }
    clearDataValidations() { return this; } setDataValidation() { return this; }
  }

  const sheets = new Map();
  let active = null, selection = [];
  const ss = {
    toasts: [],
    getSheetByName: (n) => sheets.get(n) || null,
    insertSheet: (n) => { const s = new Sheet(n, ss); sheets.set(n, s); return s; },
    getSheets: () => [...sheets.values()],
    getActiveSheet: () => active,
    setActiveSheet: (s) => { active = s; return s; },
    getActiveRangeList: () => ({ getRanges: () => selection }),
    getActiveRange: () => selection[0],
    toast(m) { this.toasts.push(m); },
    _select(sh, ranges) { active = sh; selection = ranges.map(([r1, r2, c]) => sh.getRange(r1, c || 1, r2 - r1 + 1, 1)); }
  };

  const ruleBuilder = () => {
    const st = {};
    const b = {
      whenTextStartsWith: (t) => { st.v = [t]; return b; },
      setBackground: () => b, setFontColor: () => b, setBold: () => b,
      setRanges: (r) => { st.r = r; return b; },
      build: () => ({ getRanges: () => st.r, getBooleanCondition: () => ({ getCriteriaValues: () => st.v }) })
    };
    return b;
  };

  const ctx = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss, getActive: () => ss, getUi: () => ui,
      newRichTextValue: () => { let s = ''; const b = { setText: (x) => { s = x; return b; }, build: () => ({ getText: () => s }) }; return b; },
      newConditionalFormatRule: ruleBuilder,
      newDataValidation: () => { const b = { requireValueInList: () => b, requireValueInRange: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; },
      WrapStrategy: { WRAP: 'WRAP', CLIP: 'CLIP', OVERFLOW: 'OVERFLOW' }
    },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: { formatDate: (d) => d.toISOString().slice(0, 16).replace('T', ' ') },
    Logger: { log: (m) => LOGS.push(String(m)) },
    console
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(DIST, 'utf8'), ctx, { filename: 'dist.gs' });
  const g = (expr) => vm.runInContext(expr, ctx);
  return { ctx, g, ss, ui, Sheet };
}

/* ============================ シート数式の簡易評価器 ============================ */
// 整形チェック列の数式を、Sheets と同じ意味で評価する（IF/OR/ISERROR/EXACT/SUBSTITUTE/CHAR と = / & のみ）。
const ERR = { err: '#ERROR!' };
function evalSheetFormula(formula, cells) {
  const src = formula.replace(/^=/, '');
  const toks = []; let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"') { let s = ''; i++; for (;;) { if (src[i] === '"' && src[i + 1] === '"') { s += '"'; i += 2; } else if (src[i] === '"') { i++; break; } else s += src[i++]; } toks.push({ t: 's', v: s }); }
    else if ('(),=&'.includes(ch)) { toks.push({ t: ch }); i++; }
    else if (/[0-9]/.test(ch)) { let n = ''; while (/[0-9]/.test(src[i])) n += src[i++]; toks.push({ t: 'n', v: Number(n) }); }
    else if (/[A-Z$]/.test(ch)) { let w = ''; while (i < src.length && /[A-Z0-9$]/.test(src[i])) w += src[i++]; toks.push({ t: 'w', v: w }); }
    else throw new Error('未対応の文字: ' + ch);
  }
  let p = 0;
  const peek = () => toks[p], take = (t) => { const x = toks[p++]; if (t && x.t !== t) throw new Error('構文: ' + t); return x; };
  function expr() { let a = concat(); if (peek() && peek().t === '=') { take('='); const b = concat(); return { op: '=', a, b }; } return a; }
  function concat() { let a = prim(); while (peek() && peek().t === '&') { take('&'); a = { op: '&', a, b: prim() }; } return a; }
  function prim() {
    const x = take();
    if (x.t === 's' || x.t === 'n') return { lit: x.v };
    if (x.t === 'w') {
      if (peek() && peek().t === '(') { take('('); const args = []; if (peek().t !== ')') { args.push(expr()); while (peek().t === ',') { take(','); args.push(expr()); } } take(')'); return { fn: x.v, args }; }
      return { ref: x.v.replace(/\$/g, '').replace(/[0-9]+$/, '') };
    }
    throw new Error('構文エラー');
  }
  const ast = expr();
  if (p !== toks.length) throw new Error('余分なトークン');
  const isErr = (v) => v === ERR;
  function ev(n) {
    if ('lit' in n) return n.lit;
    if (n.ref) return cells[n.ref];
    if (n.op) { const a = ev(n.a), b = ev(n.b); if (isErr(a)) return a; if (isErr(b)) return b; return n.op === '&' ? String(a) + String(b) : String(a).toLowerCase() === String(b).toLowerCase(); }
    const A = n.args;
    switch (n.fn) {
      case 'IF': { const c = ev(A[0]); if (isErr(c)) return c; return c ? ev(A[1]) : ev(A[2]); }
      case 'OR': { const vs = A.map(ev); const e = vs.find(isErr); return e || vs.some(Boolean); }
      case 'ISERROR': return isErr(ev(A[0]));
      case 'EXACT': { const a = ev(A[0]), b = ev(A[1]); if (isErr(a)) return a; if (isErr(b)) return b; return String(a) === String(b); }
      case 'SUBSTITUTE': { const s = ev(A[0]); if (isErr(s)) return s; return String(s).split(String(ev(A[1]))).join(String(ev(A[2]))); }
      case 'CHAR': return String.fromCharCode(ev(A[0]));
      default: throw new Error('未対応の関数: ' + n.fn);
    }
  }
  return ev(ast);
}

function balanced(formula) {
  let depth = 0, inStr = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"') { if (inStr && formula[i + 1] === '"') { i++; continue; } inStr = !inStr; continue; }
    if (inStr) continue;
    if (ch === '(') depth++; if (ch === ')') depth--;
    if (depth < 0) return false;
  }
  return depth === 0 && !inStr;
}

/* ============================ 本番シートの模擬 ============================ */

// ご意見記録の見出し（旧見出し「内容（要約）」「要約用結合文」のまま＝差し替え直後の状態）
function makeClaimSheet(env) {
  const sh = env.ss.insertSheet('ご意見記録');
  const HT = env.g('HEADER_TEXT');
  const heads = ['状況', '受付日時', '店名'];
  Object.keys(HT).forEach((k) => {
    if (k === 'FMT_CHK') return;                        // 整形チェック列はまだ無い
    if (k === 'AI_FMT') heads.push('内容（要約）');     // 旧見出し
    else if (k === 'FMT_TXT') heads.push('要約用結合文'); // 旧見出し
    else heads.push(HT[k]);
  });
  sh.getRange(1, 1, 1, heads.length).setValues([heads]);
  return sh;
}
function colOf(sh, header) {
  const row = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const i = row.indexOf(header); if (i < 0) throw new Error('見出しなし: ' + header); return i + 1;
}
// 各ケースを2行目から投入。AC には =AI(CL) を入れたうえで、生成結果（模擬）を値として持たせる。
function loadCases(env, sh, res) {
  CASES.forEach((c, i) => {
    const r = i + 2;
    sh.getRange(r, colOf(sh, '受付日時')).setValue(c.受付日時);
    sh.getRange(r, colOf(sh, '店名')).setValue(c.店名);
    if (c.原文 !== '') sh.getRange(r, res.COLX.RAW).setRichTextValues([[{ getText: () => c.原文 }]]);
    if (c.原文 !== '' || c.模擬整形結果 !== '') {
      sh.getRange(r, res.COLX.AI_FMT).setFormula('=AI($' + res.CLX.FMT_TXT + r + ')');
      sh.gen(r, res.COLX.AI_FMT, c.模擬整形結果);
    }
  });
}

/* ============================ テスト ============================ */

console.log('\n[1] 照合ロジック（GAS側）');
{
  const { g } = makeEnv();
  CASES.filter((c) => c.期待 !== 'エラー').forEach((c) => {
    t(c.id + ' ' + c.説明 + ' → ' + c.期待, () => {
      const r = g('pfCompareText_')(c.原文, c.模擬整形結果);
      eq(r.status, c.期待);
    });
  });
  t('比較は改行(CR/LF)だけを除く（trim・空白除去をしない）', () => {
    eq(g('pfStripNewlines_')(' a\r\n\tb　\n '), ' a\tb　 ');
  });
  t('両方空は OK にならない', () => { eq(g('pfCompareText_')('', '').status, '対象外'); });
  t('NG の差分位置は絵文字も1文字として数える', () => {
    const r = g('pfCompareText_')('😊あいう', '😊あえう');
    eq(r.status, 'NG'); eq(r.pos, 3);
  });
  t('エラー表示の判定', () => {
    const f = g('pfIsErrorText_');
    ok(f('#ERROR!') && f('#N/A') && f('#REF!') && f('#NAME?'), 'エラーを検出できない');
    ok(!f('＃ではない') && !f('#タグ付きの文章'), '通常の文章をエラー扱いした');
  });
}

console.log('\n[2] 整形チェック列の数式（シート側）');
{
  const { g } = makeEnv();
  const build = g('buildFormula_');
  const formula = build('FMT_CHK', 2);
  t('数式に AI 関数が含まれない', () => ok(!/\bAI\(/.test(formula)));
  t('括弧と引用符の対応が正しい', () => ok(balanced(formula)));
  const expectFormula = (c) => ({ OK: 'OK（文字一致）', NG: 'NG（改行以外に差あり）', 未生成: '未生成', エラー: 'エラー',
    対象外: c.模擬整形結果 === '' ? '' : '対象外（原文が空）' })[c.期待];
  CASES.forEach((c) => {
    t(c.id + ' ' + c.説明 + ' → 「' + expectFormula(c) + '」', () => {
      const cells = { AB: c.原文, AC: c.模擬整形結果 === '#ERROR!' ? ERR : c.模擬整形結果 };
      eq(evalSheetFormula(formula, cells), expectFormula(c));
    });
  });
  t('原文がエラーでも「エラー」', () => eq(evalSheetFormula(formula, { AB: ERR, AC: 'x' }), 'エラー'));
}

console.log('\n[3] 数式の組み立て');
{
  const { g } = makeEnv();
  const build = g('buildFormula_');
  t('AC は =AI($CL2) の単独（他の関数の内側に入れない）', () => eq(build('AI_FMT', 2), '=AI($CL2)'));
  t('CL は原文ABを参照し、指示文は段落整形設定シートのセルを参照する', () => {
    const f = build('FMT_TXT', 7);
    ok(f.startsWith('=IF($AB7="","",'), f);
    ok(f.includes("'段落整形設定'!$B$2"), '指示文セル参照がない');
    ok(!f.includes('$AD7'), '匿名化列ADを参照している');
    ok(balanced(f), '括弧不一致');
  });
  t('すべての数式列が組み立て可能（undefined・括弧不一致なし）', () => {
    g('FORMULA_TARGET_COLS').forEach((k) => {
      const f = build(k, 5);
      ok(!/undefined/.test(f), k + ' に undefined'); ok(balanced(f), k + ' 括弧不一致');
    });
  });
  t('AI列の数式はすべて =AI(単一セル)', () => {
    g('AI_COLS').forEach((k) => ok(/^=AI\(\$[A-Z]+5\)$/.test(build(k, 5)), k + ': ' + build(k, 5)));
  });
}

console.log('\n[4] 列の解決（旧見出しからの移行）');
{
  const env = makeEnv(); const { g } = env;
  const sh = makeClaimSheet(env);
  const res = g('resolveColumns_')(sh);
  t('「内容（要約）」のままでも AC(整形) として見つかる', () => eq(sh.getRange(1, res.COLX.AI_FMT).getValue(), '内容（要約）'));
  t('利用者の列「内容（要約）」は改名しない', () => ok(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].includes('内容（要約）')));
  t('システム補助列「要約用結合文」は「整形用結合文」へ改名', () => {
    eq(sh.getRange(1, res.COLX.FMT_TXT).getValue(), '整形用結合文');
    ok(!sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].includes('要約用結合文'));
  });
  t('「整形チェック」列が末尾に新設される', () => {
    eq(sh.getRange(1, res.COLX.FMT_CHK).getValue(), '整形チェック');
    eq(res.COLX.FMT_CHK, sh.getLastColumn());
  });
  t('再実行しても列が増えない', () => {
    const before = sh.getLastColumn(); g('resolveColumns_')(sh); eq(sh.getLastColumn(), before);
  });
}

console.log('\n[5] 段落整形設定シート');
{
  const env = makeEnv(); const { g, ss } = env;
  const ensure = g('pfEnsureSettingsSheet_');
  t('初回に作成される（指示文はB2）', () => {
    ok(ensure(ss).ok);
    const sh = ss.getSheetByName('段落整形設定');
    eq(sh.getRange(2, 1).getValue(), 'AIへの指示文');
    ok(String(sh.getRange(2, 2).getValue()).includes('変更してよいのは改行だけです。'));
  });
  t('再実行しても利用者が変えた値を保持する', () => {
    const sh = ss.getSheetByName('段落整形設定');
    sh.getRange(2, 2).setValue('利用者が編集した指示文');
    sh.getRange(3, 2).setValue(600);
    ensure(ss);
    eq(sh.getRange(2, 2).getValue(), '利用者が編集した指示文');
    eq(sh.getRange(3, 2).getValue(), 600);
  });
  t('消えた項目だけを追記する', () => {
    const sh = ss.getSheetByName('段落整形設定');
    const last = sh.getLastRow();
    sh.getRange(last, 1, 1, 3).setValues([['', '', '']]);
    const r = ensure(ss);
    ok(/1件追記/.test(r.msg), r.msg);
    eq(sh.getRange(2, 2).getValue(), '利用者が編集した指示文');
  });
  t('行を挿入して指示文の行がずれても、参照先が追従する', () => {
    const sh = ss.getSheetByName('段落整形設定');
    sh.insertRowBefore(2);
    eq(g('pfPromptRef_')(ss), "'段落整形設定'!$B$3");
  });
  t('同名で形の違うシートは変更せず停止', () => {
    const env2 = makeEnv();
    const other = env2.ss.insertSheet('段落整形設定');
    other.getRange(1, 1, 1, 2).setValues([['メモ', '大事なデータ']]);
    const r = env2.g('pfEnsureSettingsSheet_')(env2.ss);
    ok(!r.ok); eq(other.getRange(1, 2).getValue(), '大事なデータ'); eq(other.getLastRow(), 1);
  });
}

console.log('\n[6] ②数式の適用（手入力値・生成済みを守る）');
{
  const env = makeEnv(); const { g, ss } = env;
  const sh = makeClaimSheet(env);
  g('pfEnsureSettingsSheet_')(ss);
  const res = g('resolveColumns_')(sh);
  loadCases(env, sh, res);
  const manualRow = CASES.length + 2; // 手入力で整形文を入れた行
  sh.getRange(manualRow, res.COLX.RAW).setValue('手入力の原文です。');
  sh.getRange(manualRow, res.COLX.AI_FMT).setValue('手入力の原文です。');
  const blankRow = manualRow + 1;     // 原文あり・AC空（未配置）
  sh.getRange(blankRow, res.COLX.RAW).setValue('未配置の原文です。');
  sh.formulaWrites = [];
  g('setupMainSheet_')(sh, res);
  const acWrites = sh.formulaWrites.filter(([, c]) => c === res.COLX.AI_FMT).map(([r]) => r);
  t('手入力の整形文は上書きしない', () => { eq(sh.formula(manualRow, res.COLX.AI_FMT), ''); eq(sh.get(manualRow, res.COLX.AI_FMT), '手入力の原文です。'); });
  t('同じ =AI 数式が入っている行は書き換えない（生成済み結果を保つ）', () => {
    ok(!acWrites.includes(2), 'row2 を書き換えた'); eq(sh.get(2, res.COLX.AI_FMT), CASES[0].模擬整形結果);
  });
  t('未配置の行には =AI 数式を入れる', () => eq(sh.formula(blankRow, res.COLX.AI_FMT), '=AI($' + res.CLX.FMT_TXT + blankRow + ')'));
  t('原文が空の行（見出し除く）には AI 数式を入れない', () => {
    const emptyRow = blankRow + 1;
    eq(sh.formula(emptyRow, res.COLX.AI_FMT), '');
  });
  t('原文ABは一切変更しない', () => CASES.forEach((c, i) => { if (c.原文) eq(sh.get(i + 2, res.COLX.RAW), c.原文); }));
  t('整形チェック列に色分けルール（NG/エラー/未生成）', () => {
    const vals = sh.cf.map((r) => r.getBooleanCondition().getCriteriaValues()[0]).sort().join(',');
    eq(vals, 'NG,エラー,未生成');
  });
  t('再実行しても色分けルールが重複しない', () => { g('setupMainSheet_')(sh, res); eq(sh.cf.length, 3); });
  t('原文入力時の自動反映（onEdit）でも手入力の整形文を守る', () => {
    g('autoFillRows_')(sh, manualRow, blankRow, res);
    eq(sh.formula(manualRow, res.COLX.AI_FMT), ''); eq(sh.get(manualRow, res.COLX.AI_FMT), '手入力の原文です。');
    eq(sh.formula(blankRow, res.COLX.FMT_TXT).startsWith('=IF($' + res.CLX.RAW + blankRow), true);
  });
}

console.log('\n[7] 会議用シートの作成');
{
  const env = makeEnv(); const { g, ss, ui } = env;
  const sh = makeClaimSheet(env);
  g('pfEnsureSettingsSheet_')(ss);
  const res = g('resolveColumns_')(sh);
  loadCases(env, sh, res);
  const last = CASES.length + 1;
  const okCases = CASES.filter((c) => c.期待 === 'OK');

  ss._select(sh, [[2, last, 5]]);
  ui.next = 'CANCEL';
  g('pfBuildMeetingSheet')();
  t('確認でキャンセルすると何も作らない', () => eq(ss.getSheetByName('会議用'), null));
  t('確認画面に転記件数・転記しない行と理由が出る', () => {
    const m = ui.alerts[ui.alerts.length - 1];
    ok(m.includes('転記する行：' + okCases.length + '件'), m);
    ok(/【NG】/.test(m) && /【未生成】/.test(m) && /【エラー】/.test(m), m);
  });

  ui.next = 'OK';
  g('pfBuildMeetingSheet')();
  const mt = ss.getSheetByName('会議用');
  t('OKの行だけを順番どおり転記（受付日時・店名・本文）', () => {
    eq(mt.getLastRow(), 2 + okCases.length);
    okCases.forEach((c, i) => {
      eq(mt.get(3 + i, 1), c.受付日時); eq(mt.get(3 + i, 2), c.店名); eq(mt.get(3 + i, 3), c.模擬整形結果);
    });
  });
  t('作成日時とスナップショットである旨を表示', () => {
    eq(mt.get(1, 1), '作成日時'); ok(String(mt.get(1, 2)).length >= 10); ok(/スナップショット/.test(mt.get(1, 3)));
  });
  t('本文が「=」で始まっても数式にならず文字が変わらない', () => {
    const i = okCases.findIndex((c) => c.id === 'C12');
    eq(mt.formula(3 + i, 3), ''); eq(mt.get(3 + i, 3), okCases[i].模擬整形結果);
  });
  t('元の行番号を残す', () => eq(mt.get(3, 4), CASES.findIndex((c) => c.id === okCases[0].id) + 2));
  t('ご意見記録は変更しない', () => CASES.forEach((c, i) => { if (c.原文) eq(sh.get(i + 2, res.COLX.RAW), c.原文); }));

  // 照合後に原文を変更 → 転記時の再照合で弾かれる
  const c02row = CASES.findIndex((c) => c.id === 'C02') + 2;
  ss._select(sh, [[c02row, c02row, 1]]);
  g('pfCheckSelectedRows')();
  t('変更前はチェックOK', () => ok(/OK 1件/.test(ui.alerts[ui.alerts.length - 1])));
  sh.getRange(c02row, res.COLX.RAW).setValue('レジの方の笑顔が素敵でした。ありがとう。');
  ui.alerts.length = 0;
  g('pfBuildMeetingSheet')();
  t('照合後に原文を変えると、転記時の再照合でNGになり転記しない', () => {
    const m = ui.alerts.join('\n');
    ok(m.includes('転記できる行') && m.includes('行' + c02row + '【NG】'), m);
    eq(mt.getLastRow(), 2 + okCases.length, '会議用が上書きされた');
  });

  t('既存の会議用データを置き換える前に件数を示す', () => {
    ss._select(sh, [[2, 3, 1]]);
    ui.next = 'CANCEL'; ui.alerts.length = 0;
    g('pfBuildMeetingSheet')();
    ok(ui.alerts[0].includes('既存データ ' + okCases.length + '件'), ui.alerts[0]);
  });

  t('同名で形の違う「会議用」シートは変更せず停止', () => {
    const env2 = makeEnv(); const sh2 = makeClaimSheet(env2);
    env2.g('pfEnsureSettingsSheet_')(env2.ss);
    const res2 = env2.g('resolveColumns_')(sh2);
    loadCases(env2, sh2, res2);
    const other = env2.ss.insertSheet('会議用');
    other.getRange(1, 1).setValue('手作りの資料');
    env2.ss._select(sh2, [[2, 3, 1]]); env2.ui.next = 'OK';
    env2.g('pfBuildMeetingSheet')();
    eq(other.get(1, 1), '手作りの資料'); eq(other.getLastRow(), 1);
  });

  t('ご意見記録以外のシートで実行したら案内して止まる', () => {
    ss._select(ss.getSheetByName('段落整形設定'), [[2, 3, 1]]);
    ui.alerts.length = 0;
    g('pfBuildMeetingSheet')();
    ok(/選択してから実行/.test(ui.alerts[0]));
  });
}

console.log('\n[8] その他');
{
  const { g } = makeEnv();
  t('選択行の解析：複数範囲・重複・見出し行を処理', () => {
    const R = (a, b) => ({ getRow: () => a, getLastRow: () => b });
    eq(JSON.stringify(g('pfRowsFromRanges_')([R(5, 7), R(1, 3), R(6, 9)])), '[2,3,5,6,7,8,9]');
    eq(g('pfRowsLabel_')([2, 3, 5, 6, 7, 9]), '2〜3, 5〜7, 9');
  });
  t('ログに原文・整形結果を出力していない', () => {
    const leaked = LOGS.filter((l) => CASES.some((c) => c.原文.length > 8 && l.includes(c.原文.slice(0, 8))));
    eq(leaked.length, 0, leaked.join(' / '));
  });
  t('dist の統合版が src と一致', () => {
    require('child_process').execFileSync('node', [path.join(ROOT, 'tools', 'build.js'), '--check'], { stdio: 'pipe' });
  });
  t('報告書リンク共存版と同じプロジェクトに入れても名前が衝突しない', () => {
    const both = fs.readFileSync(path.join(ROOT, 'dist', '報告書リンク_共存版.gs'), 'utf8') + '\n' + fs.readFileSync(DIST, 'utf8');
    vm.runInNewContext(both, { Logger: { log() {} } });
  });
}

console.log('\n結果: ' + pass + '件成功 / ' + fail + '件失敗');
process.exit(fail ? 1 : 0);
