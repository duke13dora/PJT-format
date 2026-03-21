// build-agenda-sheet-format.mjs
// PJT-YP/TSのアジェンダシートをベースに、フォーマットテンプレートを生成
//
// Usage: node scripts/build-agenda-sheet-format.mjs
//
// 修正v2:
// - sheet4 XMLエラー修正（shared formula範囲、dataValidation）
// - ハイパーリンク全除去
// - マスタシート追加（yet/ing/done プルダウン用）
// - ステータスのデータ入力規則（プルダウン）追加
// - 条件付き書式（色ルール）追加: done=灰色, yet=黄色, ing=ピンク
// - シート名修正: すり合わせ用メモ（元のまま）
// - 列幅をYPオリジナルから保持
// - #列ヘッダーに「#」記号

import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

const YP_SOURCE =
  'C:/PJT-YP/drive/50.打ち合わせ関連/【YP様】アジェンダシート.xlsx';
const scriptDir = path.dirname(
  new URL(import.meta.url).pathname.replace(/^\/([A-Z]:)/, '$1')
);
const OUTPUT = path.join(
  scriptDir,
  '..',
  'templates',
  'agenda-sheet-format.xlsx'
);

function escapeXml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

(async () => {
  console.log('[1/8] ソースファイル読み込み...');
  const ypData = fs.readFileSync(YP_SOURCE);
  const ypZip = await JSZip.loadAsync(ypData);

  // ============================================================
  // Step 1: sharedStrings に必要な文字列を追加
  // ============================================================
  console.log('[2/8] sharedStrings更新...');
  let ssXml = await ypZip.file('xl/sharedStrings.xml').async('string');
  const ucMatch = ssXml.match(/uniqueCount="(\d+)"/);
  let uniqueCount = parseInt(ucMatch[1]);

  function findOrAddString(text) {
    const siAll = ssXml.match(/<si>[\s\S]*?<\/si>/g) || [];
    for (let i = 0; i < siAll.length; i++) {
      const t = siAll[i]
        .replace(/<[^>]+>/g, '')
        .replace(/\s+/g, ' ')
        .trim();
      if (t === text) return i;
    }
    const idx = uniqueCount;
    ssXml = ssXml.replace('</sst>', `<si><t>${escapeXml(text)}</t></si></sst>`);
    uniqueCount++;
    ssXml = ssXml.replace(/uniqueCount="\d+"/, `uniqueCount="${uniqueCount}"`);
    ssXml = ssXml.replace(
      /count="\d+"/,
      (m) => `count="${parseInt(m.match(/\d+/)[0]) + 1}"`
    );
    return idx;
  }

  const xxxIdx = findOrAddString('×××');
  const tantouIdx = findOrAddString('担当');
  const kigenIdx = findOrAddString('期限');
  const yetIdx = findOrAddString('yet');
  const ingIdx = findOrAddString('ing');
  const doneIdx = findOrAddString('done');
  const statusIdx = findOrAddString('ステータス');
  const hashIdx = findOrAddString('#');
  const kihyoubiIdx = findOrAddString('起票日');
  const kihyoushaIdx = findOrAddString('起票者');
  const gaiyouIdx = findOrAddString('概要');
  const shousaiIdx = findOrAddString('詳細');
  const bikouIdx = findOrAddString('備考');

  console.log(
    `  ×××=${xxxIdx} 担当=${tantouIdx} 期限=${kigenIdx} yet=${yetIdx} ing=${ingIdx} done=${doneIdx}`
  );

  ypZip.file('xl/sharedStrings.xml', ssXml);

  // ============================================================
  // Step 2: styles.xml に条件付き書式用のdxfを追加
  // ============================================================
  console.log('[3/8] 条件付き書式スタイル追加...');
  let stylesXml = await ypZip.file('xl/styles.xml').async('string');

  // done=灰色(#D8D8D8), yet=黄色(#FEF1CC), ing=ピンク(#FAD9D6), 行done=薄灰(#EFEFEF)
  const newDxfs = [
    // dxf0: done (gray) - for status cell
    '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFD8D8D8"/><bgColor rgb="FFD8D8D8"/></patternFill></fill></dxf>',
    // dxf1: yet (yellow) - for status cell
    '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFEF1CC"/><bgColor rgb="FFFEF1CC"/></patternFill></fill></dxf>',
    // dxf2: ing (pink) - for status cell
    '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFAD9D6"/><bgColor rgb="FFFAD9D6"/></patternFill></fill></dxf>',
    // dxf3: row done (light gray) - for entire row when done
    '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFEFEFEF"/><bgColor rgb="FFEFEFEF"/></patternFill></fill></dxf>',
  ];

  // Get existing dxf count
  const dxfCountMatch = stylesXml.match(/<dxfs count="(\d+)"/);
  let dxfStartIdx;
  if (dxfCountMatch) {
    dxfStartIdx = parseInt(dxfCountMatch[1]);
    stylesXml = stylesXml.replace(
      /<dxfs count="\d+"/,
      `<dxfs count="${dxfStartIdx + newDxfs.length}"`
    );
    stylesXml = stylesXml.replace('</dxfs>', newDxfs.join('') + '</dxfs>');
  } else {
    dxfStartIdx = 0;
    stylesXml = stylesXml.replace(
      '</styleSheet>',
      `<dxfs count="${newDxfs.length}">${newDxfs.join('')}</dxfs></styleSheet>`
    );
  }

  const dxfDone = dxfStartIdx;
  const dxfYet = dxfStartIdx + 1;
  const dxfIng = dxfStartIdx + 2;
  const dxfRowDone = dxfStartIdx + 3;

  ypZip.file('xl/styles.xml', stylesXml);

  // ============================================================
  // Step 3: 各シートのデータクリア + 機能追加
  // ============================================================
  console.log('[4/8] 各シートのデータクリア...');
  const today = Math.floor(
    (Date.now() - new Date('1899-12-30').getTime()) / 86400000
  );

  // --- Sheet1: アジェンダ ---
  // Status column = H, data validation + conditional formatting
  let sheet1 = await ypZip.file('xl/worksheets/sheet1.xml').async('string');
  sheet1 = rebuildSheet(sheet1, {
    headerRow: 1,
    sampleRow: 2,
    keepEmptyRowsUpTo: 1000,
    sampleCells: [
      { col: 'A', formula: 'row()-row($A$1)', value: '1', style: '5' },
      { col: 'B', date: today, style: '6' },
      { col: 'C', ssIdx: xxxIdx, style: '7' },
      { col: 'D', ssIdx: xxxIdx, style: '7' },
      { col: 'E', ssIdx: xxxIdx, style: '8' },
      { col: 'F', ssIdx: xxxIdx, style: '8' },
      { col: 'G', ssIdx: xxxIdx, style: '9' },
      { col: 'H', ssIdx: yetIdx, style: '7' },
    ],
    statusCol: 'H',
    statusRange: 'H2:H1000',
    conditionalFormatRange: 1000,
    dxfDone,
    dxfYet,
    dxfIng,
    dxfRowDone,
    dataCols: 'A:H',
  });
  ypZip.file('xl/worksheets/sheet1.xml', sheet1);

  // --- Sheet2: 決定事項 ---
  let sheet2 = await ypZip.file('xl/worksheets/sheet2.xml').async('string');
  sheet2 = rebuildSheet(sheet2, {
    headerRow: 1,
    sampleRow: 2,
    keepEmptyRowsUpTo: 0,
    sampleCells: [
      { col: 'A', formula: 'row()-row($A$1)', value: '1', style: '21' },
      { col: 'B', date: today, style: '22' },
      { col: 'C', ssIdx: xxxIdx, style: '21' },
      { col: 'D', ssIdx: xxxIdx, style: '21' },
      { col: 'E', ssIdx: xxxIdx, style: '21' },
      { col: 'F', ssIdx: xxxIdx, style: '21' },
    ],
  });
  ypZip.file('xl/worksheets/sheet2.xml', sheet2);

  // --- Sheet3: ネクスト（PJ管理ツール使用版）---
  let sheet3 = await ypZip.file('xl/worksheets/sheet3.xml').async('string');
  sheet3 = rebuildSheet(sheet3, {
    headerRow: 1,
    sampleRow: 2,
    keepEmptyRowsUpTo: 0,
    sampleCells: [{ col: 'A', ssIdx: xxxIdx, style: '24' }],
  });
  ypZip.file('xl/worksheets/sheet3.xml', sheet3);

  // --- Sheet4: 事実一覧 ---
  let sheet4 = await ypZip.file('xl/worksheets/sheet4.xml').async('string');
  // Remove YP-specific dataValidation, fix shared formula range
  sheet4 = rebuildSheet(sheet4, {
    headerRow: 1,
    sampleRow: 2,
    keepEmptyRowsUpTo: 0,
    sampleCells: [
      { col: 'A', formula: 'row()-row($A$1)', value: '1', style: '21' },
      { col: 'B', date: today, style: '22' },
      { col: 'C', ssIdx: xxxIdx, style: '21' },
      { col: 'D', ssIdx: xxxIdx, style: '21' },
      { col: 'E', ssIdx: xxxIdx, style: '21' },
    ],
    removeDataValidation: true,
  });
  ypZip.file('xl/worksheets/sheet4.xml', sheet4);

  // --- Sheet5: すり合わせ用メモ (空シート) ---
  // No changes needed

  // --- Sheet6: リンク類 → リンク集 ---
  let sheet6 = await ypZip.file('xl/worksheets/sheet6.xml').async('string');
  sheet6 = rebuildSheet(sheet6, {
    headerRow: 1,
    sampleRow: 2,
    keepEmptyRowsUpTo: 1000,
    sampleCells: [
      { col: 'A', formula: 'row()-row($A$1)', value: '1', style: '5' },
      { col: 'B', date: today, style: '30' },
      { col: 'C', ssIdx: xxxIdx, style: '31' },
    ],
  });
  ypZip.file('xl/worksheets/sheet6.xml', sheet6);

  // ============================================================
  // Step 4: ネクスト（PJ管理ツール不使用）シート + マスタシート追加
  // ============================================================
  console.log('[5/8] 新規シート作成（ネクスト不使用版 + マスタ）...');

  // Sheet7: ネクスト（PJ管理ツール不使用版）
  // TS sheet3 column widths + structure
  const sheet7Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1.0" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>
<sheetFormatPr customHeight="1" defaultColWidth="14.43" defaultRowHeight="15.0"/>
<cols><col customWidth="1" min="1" max="1" width="4.63"/><col customWidth="1" min="2" max="3" width="11.09"/><col customWidth="1" min="4" max="4" width="35.27"/><col customWidth="1" min="5" max="6" width="53.09"/><col customWidth="1" min="7" max="7" width="18.73"/><col customWidth="1" min="8" max="9" width="11.09"/></cols>
<sheetData>
<row r="1" ht="18.0" customHeight="1">
<c r="A1" s="2" t="s"><v>${hashIdx}</v></c>
<c r="B1" s="2" t="s"><v>${kihyoubiIdx}</v></c>
<c r="C1" s="2" t="s"><v>${kihyoushaIdx}</v></c>
<c r="D1" s="2" t="s"><v>${gaiyouIdx}</v></c>
<c r="E1" s="2" t="s"><v>${shousaiIdx}</v></c>
<c r="F1" s="2" t="s"><v>${bikouIdx}</v></c>
<c r="G1" s="2" t="s"><v>${tantouIdx}</v></c>
<c r="H1" s="2" t="s"><v>${kigenIdx}</v></c>
<c r="I1" s="2" t="s"><v>${statusIdx}</v></c>
</row>
<row r="2" ht="15.75" customHeight="1">
<c r="A2" s="21"><f t="shared" ref="A2:A2" si="0">row()-row($A$1)</f><v>1</v></c>
<c r="B2" s="22"><v>${today}.0</v></c>
<c r="C2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="D2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="E2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="F2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="G2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="H2" s="22"><v>${today}.0</v></c>
<c r="I2" s="21" t="s"><v>${yetIdx}</v></c>
</row>
</sheetData>
<conditionalFormatting sqref="A2:I1000"><cfRule type="expression" dxfId="${dxfRowDone}" priority="1"><formula>$I2="done"</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="I2:I1000"><cfRule type="cellIs" dxfId="${dxfDone}" priority="2" operator="equal"><formula>"done"</formula></cfRule><cfRule type="cellIs" dxfId="${dxfYet}" priority="3" operator="equal"><formula>"yet"</formula></cfRule><cfRule type="cellIs" dxfId="${dxfIng}" priority="4" operator="equal"><formula>"ing"</formula></cfRule></conditionalFormatting>
<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="I2:I1000"><formula1>マスタ!$B$2:$B$4</formula1></dataValidation></dataValidations>
</worksheet>`;
  ypZip.file('xl/worksheets/sheet7.xml', sheet7Xml);

  // Sheet8: マスタ
  const sheet8Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetFormatPr customHeight="1" defaultColWidth="14.43" defaultRowHeight="15.0"/>
<cols><col customWidth="1" min="1" max="1" width="3.86"/><col customWidth="1" min="2" max="2" width="12.57"/></cols>
<sheetData>
<row r="1" ht="18.0" customHeight="1">
<c r="A1" s="2" t="s"><v>${hashIdx}</v></c>
<c r="B1" s="2" t="s"><v>${statusIdx}</v></c>
</row>
<row r="2"><c r="A2" s="21"><v>1</v></c><c r="B2" s="21" t="s"><v>${yetIdx}</v></c></row>
<row r="3"><c r="A3" s="21"><v>2</v></c><c r="B3" s="21" t="s"><v>${ingIdx}</v></c></row>
<row r="4"><c r="A4" s="21"><v>3</v></c><c r="B4" s="21" t="s"><v>${doneIdx}</v></c></row>
</sheetData>
</worksheet>`;
  ypZip.file('xl/worksheets/sheet8.xml', sheet8Xml);

  // ============================================================
  // Step 5: ハイパーリンク除去（全シートのrels）
  // ============================================================
  console.log('[6/8] ハイパーリンク・リレーション除去...');
  for (let i = 1; i <= 6; i++) {
    const relsPath = `xl/worksheets/_rels/sheet${i}.xml.rels`;
    const relsFile = ypZip.file(relsPath);
    if (relsFile) {
      let relsXml = await relsFile.async('string');
      relsXml = relsXml.replace(/<Relationship[^>]*hyperlink[^>]*\/>/g, '');
      ypZip.file(relsPath, relsXml);
    }
  }

  // ============================================================
  // Step 6: workbook.xml更新
  // ============================================================
  console.log('[7/8] workbook.xml・rels・ContentTypes更新...');
  let wbXml = await ypZip.file('xl/workbook.xml').async('string');

  // シート名変更
  wbXml = wbXml.replace(
    'name="ネクスト"',
    'name="ネクスト（PJ管理ツール使用）"'
  );
  wbXml = wbXml.replace('name="リンク類"', 'name="リンク集"');
  // すり合わせ用メモはそのまま維持

  // 新シート2つ追加: ネクスト不使用(rId11) + マスタ(rId12)
  const newSheet7Tag =
    '<sheet state="visible" name="ネクスト（PJ管理ツール不使用）" sheetId="7" r:id="rId11"/>';
  const newSheet8Tag =
    '<sheet state="visible" name="マスタ" sheetId="8" r:id="rId12"/>';

  // 事実一覧の前にネクスト不使用を挿入
  wbXml = wbXml.replace(
    /<sheet [^>]*name="事実一覧"[^>]*\/>/,
    (match) => newSheet7Tag + match
  );
  // </sheets>の前にマスタを追加
  wbXml = wbXml.replace('</sheets>', newSheet8Tag + '</sheets>');

  ypZip.file('xl/workbook.xml', wbXml);

  // workbook.xml.rels
  let wbRels = await ypZip.file('xl/_rels/workbook.xml.rels').async('string');
  wbRels = wbRels.replace(
    '</Relationships>',
    '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet7.xml"/>' +
      '<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet8.xml"/>' +
      '</Relationships>'
  );
  ypZip.file('xl/_rels/workbook.xml.rels', wbRels);

  // Content_Types.xml
  let ctXml = await ypZip.file('[Content_Types].xml').async('string');
  ctXml = ctXml.replace(
    '</Types>',
    '<Override PartName="/xl/worksheets/sheet7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet8.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>'
  );
  ypZip.file('[Content_Types].xml', ctXml);

  // ============================================================
  // Step 7: 保存・検証
  // ============================================================
  console.log('[8/8] 保存・検証...');
  const buf = await ypZip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(OUTPUT, buf);
  console.log(`[完了] ${OUTPUT}`);

  // 検証
  const verifyData = fs.readFileSync(OUTPUT);
  const verifyZip = await JSZip.loadAsync(verifyData);
  const vWb = await verifyZip.file('xl/workbook.xml').async('string');
  const sheets = vWb.match(/<sheet [^>]+>/g) || [];
  console.log('\n検証: シート一覧');
  sheets.forEach((s, i) => {
    const name = s.match(/name="([^"]+)"/)?.[1];
    console.log(`  ${i + 1}. ${name}`);
  });

  for (let i = 1; i <= 8; i++) {
    const f = verifyZip.file(`xl/worksheets/sheet${i}.xml`);
    if (f) {
      const xml = await f.async('string');
      const rows = xml.match(/<row /g) || [];
      const hasCF = xml.includes('conditionalFormatting');
      const hasDV = xml.includes('dataValidation');
      console.log(
        `  sheet${i}: ${rows.length} rows${hasCF ? ' [CF]' : ''}${hasDV ? ' [DV]' : ''}`
      );
    }
  }
})();

// ============================================================
// ヘルパー: シート再構築
// ============================================================
function rebuildSheet(sheetXml, opts) {
  const {
    sampleCells,
    keepEmptyRowsUpTo,
    statusCol,
    statusRange,
    conditionalFormatRange,
    dxfDone,
    dxfYet,
    dxfIng,
    dxfRowDone,
    removeDataValidation,
    dataCols,
  } = opts;

  // ヘッダー行を保持
  const rows = sheetXml.match(/<row [^>]*>[\s\S]*?<\/row>/g) || [];
  if (rows.length === 0) return sheetXml;
  const headerRow = rows[0];

  // サンプル行を構築
  const sampleRowCells = sampleCells
    .map((c) => {
      if (c.formula) {
        return `<c r="${c.col}2" s="${c.style}"><f t="shared" ref="${c.col}2:${c.col}2" si="0">${c.formula}</f><v>${c.value}</v></c>`;
      } else if (c.date !== undefined) {
        return `<c r="${c.col}2" s="${c.style}"><v>${c.date}.0</v></c>`;
      } else {
        return `<c r="${c.col}2" s="${c.style}" t="s"><v>${c.ssIdx}</v></c>`;
      }
    })
    .join('');
  const sampleRow = `<row r="2" ht="15.75" customHeight="1">${sampleRowCells}</row>`;

  // 空行を構築
  let emptyRows = '';
  if (keepEmptyRowsUpTo > 0) {
    for (let r = 2; r < rows.length; r++) {
      const rowNum = rows[r].match(/r="(\d+)"/)?.[1];
      if (!rowNum || parseInt(rowNum) > keepEmptyRowsUpTo) break;
      if (parseInt(rowNum) <= 2) continue; // skip row 2 (replaced by sample)

      // 空行: セルのデータ(<v>, <f>, t属性)を削除、スタイルのみ残す
      let cleanedRow = rows[r];
      cleanedRow = cleanedRow.replace(/<v>[^<]*<\/v>/g, '');
      cleanedRow = cleanedRow.replace(/ t="[^"]*"/g, '');
      cleanedRow = cleanedRow.replace(/<f[^>]*>[\s\S]*?<\/f>/g, '');
      cleanedRow = cleanedRow.replace(/<f[^>]*\/>/g, '');
      emptyRows += cleanedRow + '\n';
    }
  }

  const sheetDataContent = headerRow + '\n' + sampleRow + '\n' + emptyRows;
  sheetXml = sheetXml.replace(
    /<sheetData>[\s\S]*<\/sheetData>/,
    `<sheetData>${sheetDataContent}</sheetData>`
  );

  // ============================================================
  // 旧要素を全て除去（hyperlinks, autoFilter, 旧conditionalFormatting, 旧dataValidations）
  // ============================================================
  sheetXml = sheetXml.replace(/<hyperlinks>[\s\S]*?<\/hyperlinks>/g, '');
  sheetXml = sheetXml.replace(/<autoFilter[\s\S]*?<\/autoFilter>/g, '');
  sheetXml = sheetXml.replace(/<autoFilter[^>]*\/>/g, '');
  sheetXml = sheetXml.replace(
    /<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g,
    ''
  );
  sheetXml = sheetXml.replace(
    /<dataValidations[\s\S]*?<\/dataValidations>/g,
    ''
  );

  // ステータスのデータ入力規則 + 条件付き書式を追加
  // OOXML順序: conditionalFormatting → dataValidations → drawing
  if (statusCol && statusRange) {
    const maxRow = conditionalFormatRange || 1000;
    const cfXml =
      `<conditionalFormatting sqref="A2:${statusCol}${maxRow}"><cfRule type="expression" dxfId="${dxfRowDone}" priority="1"><formula>$${statusCol}2="done"</formula></cfRule></conditionalFormatting>` +
      `<conditionalFormatting sqref="${statusCol}2:${statusCol}${maxRow}"><cfRule type="cellIs" dxfId="${dxfDone}" priority="2" operator="equal"><formula>"done"</formula></cfRule><cfRule type="cellIs" dxfId="${dxfYet}" priority="3" operator="equal"><formula>"yet"</formula></cfRule><cfRule type="cellIs" dxfId="${dxfIng}" priority="4" operator="equal"><formula>"ing"</formula></cfRule></conditionalFormatting>`;

    const dvXml = `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="${statusRange}"><formula1>マスタ!$B$2:$B$4</formula1></dataValidation></dataValidations>`;

    // drawing要素の前に挿入（drawingがあれば）
    if (sheetXml.includes('<drawing')) {
      sheetXml = sheetXml.replace(/<drawing/, cfXml + dvXml + '<drawing');
    } else {
      sheetXml = sheetXml.replace(
        '</worksheet>',
        cfXml + dvXml + '</worksheet>'
      );
    }
  }

  return sheetXml;
}
