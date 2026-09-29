#!/usr/bin/env node
// src/*.gs を結合して dist/クレームAI分類_統合版.gs を生成する。
//   node tools/build.js          … 生成
//   node tools/build.js --check  … dist が src と一致しているか検査（不一致なら終了コード1）
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const PARTS = ['Config', 'Anonymize', 'Setup', 'Formatter', 'Code'];
const OUT = path.join(ROOT, 'dist', 'クレームAI分類_統合版.gs');

function build() {
  const head = [
    '/**',
    ' * ============================================================',
    ' * クレーム(ご意見)AI自動分類システム 【統合版・1ファイル】',
    ' * ------------------------------------------------------------',
    ' * この1ファイルだけをApps Scriptに貼り付ければ動きます。',
    ' * （src/ の ' + PARTS.map((p) => p + '.gs').join(' + ') + ' を tools/build.js で結合）',
    ' * ============================================================',
    ' */',
    ''
  ].join('\n');
  const body = PARTS.map((p, i) => {
    const src = fs.readFileSync(path.join(ROOT, 'src', p + '.gs'), 'utf8');
    return '/* ==================== ' + (i + 1) + '/' + PARTS.length + ' ' + p + ' ==================== */\n' + src;
  }).join('\n');
  return head + '\n' + body;
}

const out = build();
if (process.argv.includes('--check')) {
  const cur = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
  if (cur !== out) { console.error('dist が src と一致しません。node tools/build.js を実行してください。'); process.exit(1); }
  console.log('dist は src と一致しています。');
} else {
  fs.writeFileSync(OUT, out);
  console.log('生成: ' + path.relative(ROOT, OUT) + '（' + out.split('\n').length + '行）');
}
