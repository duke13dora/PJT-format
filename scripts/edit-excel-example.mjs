// edit-excel-example.mjs — adm-zip Excel編集の最小動作サンプル
// Usage: node edit-excel-example.mjs
//
// このスクリプトは adm-zip によるExcel XML直接編集のパターンを示す。
// 実際のプロジェクトでは、対象ファイル・シート・セルに合わせてカスタマイズすること。

import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const AdmZip = require('adm-zip');

// ============================================================
// 設定
// ============================================================
const INPUT_FILE = 'input.xlsx';
const OUTPUT_FILE = 'output.xlsx';
const TARGET_SHEET = 'xl/worksheets/sheet1.xml';

// ============================================================
// Step 1: ZIP展開・XML読み取り
// ============================================================
const zip = new AdmZip(INPUT_FILE);
let ssXml = zip.readAsText('xl/sharedStrings.xml');
let sheetXml = zip.readAsText(TARGET_SHEET);

// ============================================================
// Step 2: sharedStrings.xml に文字列を追加
// ============================================================
function addSharedString(xml, text) {
  // 現在のuniqueCount取得
  const countMatch = xml.match(/uniqueCount="(\d+)"/);
  const currentCount = parseInt(countMatch[1]);
  const newIndex = currentCount;

  // <si><t>テキスト</t></si> を追加
  const newEntry = `<si><t>${escapeXml(text)}</t></si>`;
  xml = xml.replace('</sst>', newEntry + '</sst>');

  // uniqueCount と count をインクリメント
  xml = xml.replace(/uniqueCount="\d+"/, `uniqueCount="${currentCount + 1}"`);
  xml = xml.replace(/count="\d+"/, (m) => {
    const c = parseInt(m.match(/\d+/)[0]);
    return `count="${c + 1}"`;
  });

  return { xml, index: newIndex };
}

function escapeXml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// 使用例: "新しいテキスト" をsharedStringsに追加
const result = addSharedString(ssXml, '新しいテキスト');
ssXml = result.xml;
const newStringIndex = result.index;

// ============================================================
// Step 3: ワークシートのセルを編集
// ============================================================
// 既存セルの値を変更する例（A5セルの値をnewStringIndexに変更）
// t="s" は共有文字列への参照を意味する
//
// sheetXml = sheetXml.replace(
//   /<c r="A5"[^>]*><v>\d+<\/v><\/c>/,
//   `<c r="A5" t="s" s="4"><v>${newStringIndex}</v></c>`
// );

// ============================================================
// Step 4: 保存
// ============================================================
zip.updateFile('xl/sharedStrings.xml', Buffer.from(ssXml));
zip.updateFile(TARGET_SHEET, Buffer.from(sheetXml));
zip.writeZip(OUTPUT_FILE);

console.log(`[OK] ${OUTPUT_FILE} に保存しました`);

// ============================================================
// Step 5: JSZipで検証（adm-zipは自分の出力を再読不可）
// ============================================================
// import JSZip from 'jszip';
// import fs from 'fs';
// const data = fs.readFileSync(OUTPUT_FILE);
// const verifyZip = await JSZip.loadAsync(data);
// const verifySheet = await verifyZip.file(TARGET_SHEET).async('string');
// console.log('検証OK:', verifySheet.includes('新しいテキスト'));
