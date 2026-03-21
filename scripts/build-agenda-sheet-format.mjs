// build-agenda-sheet-format.mjs
// PJT-YP/TSのアジェンダシートをベースに、フォーマットテンプレートを生成
//
// Usage: node scripts/build-agenda-sheet-format.mjs
//
// ソース:
//   YP: C:\PJT-YP\drive\50.打ち合わせ関連\【YP様】アジェンダシート.xlsx
//   TS: C:\PJT-TS\【ダッシュボード】アジェンダシート_2508~2603.xlsx
// 出力: C:\PJT-format\templates\agenda-sheet-format.xlsx

import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

const YP_SOURCE =
  'C:/PJT-YP/drive/50.打ち合わせ関連/【YP様】アジェンダシート.xlsx';
const TS_SOURCE = 'C:/PJT-TS/【ダッシュボード】アジェンダシート_2508~2603.xlsx';
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
  console.log('[1/6] ソースファイル読み込み...');
  const ypData = fs.readFileSync(YP_SOURCE);
  const ypZip = await JSZip.loadAsync(ypData);

  // ============================================================
  // Step 1: sharedStrings に ×××を追加
  // ============================================================
  console.log('[2/6] sharedStrings更新...');
  let ssXml = await ypZip.file('xl/sharedStrings.xml').async('string');

  // 現在のuniqueCount取得
  const ucMatch = ssXml.match(/uniqueCount="(\d+)"/);
  let uniqueCount = parseInt(ucMatch[1]);

  // 追加する文字列と対応インデックス
  const newStrings = ['×××', '担当', '期限'];
  const newIndices = {};
  for (const s of newStrings) {
    // 既存チェック
    const siMatches = ssXml.match(/<si>[\s\S]*?<\/si>/g) || [];
    let found = -1;
    for (let i = 0; i < siMatches.length; i++) {
      const text = siMatches[i].replace(/<[^>]+>/g, '').trim();
      if (text === s) {
        found = i;
        break;
      }
    }
    if (found >= 0) {
      newIndices[s] = found;
    } else {
      newIndices[s] = uniqueCount;
      ssXml = ssXml.replace('</sst>', `<si><t>${escapeXml(s)}</t></si></sst>`);
      uniqueCount++;
      ssXml = ssXml.replace(
        /uniqueCount="\d+"/,
        `uniqueCount="${uniqueCount}"`
      );
      ssXml = ssXml.replace(/count="\d+"/, (m) => {
        const c = parseInt(m.match(/\d+/)[0]);
        return `count="${c + 1}"`;
      });
    }
  }
  console.log(
    `  ×××: index ${newIndices['×××']}, 担当: index ${newIndices['担当']}, 期限: index ${newIndices['期限']}`
  );

  // 既存の重要インデックスを確認
  const siAll = ssXml.match(/<si>[\s\S]*?<\/si>/g) || [];
  // ステータス用のインデックスを見つける
  let statusIndices = {};
  for (let i = 0; i < siAll.length; i++) {
    const text = siAll[i].replace(/<[^>]+>/g, '').trim();
    if (text === 'ステータス') statusIndices['ステータス'] = i;
    if (text === 'yet') statusIndices['yet'] = i;
  }
  // yetがなければ追加
  if (statusIndices['yet'] === undefined) {
    newIndices['yet'] = uniqueCount;
    ssXml = ssXml.replace('</sst>', `<si><t>yet</t></si></sst>`);
    uniqueCount++;
    ssXml = ssXml.replace(/uniqueCount="\d+"/, `uniqueCount="${uniqueCount}"`);
    ssXml = ssXml.replace(/count="\d+"/, (m) => {
      const c = parseInt(m.match(/\d+/)[0]);
      return `count="${c + 1}"`;
    });
  } else {
    newIndices['yet'] = statusIndices['yet'];
  }

  const xxxIdx = newIndices['×××'];

  ypZip.file('xl/sharedStrings.xml', ssXml);

  // ============================================================
  // Step 2: 各シートのデータクリア
  // ============================================================
  console.log('[3/6] 各シートのデータクリア...');

  // --- Sheet1: アジェンダ ---
  // Keep row 1 (header) and row 2 (sample → replace with ×××)
  // Keep rows 3-1000 as empty styled rows
  let sheet1 = await ypZip.file('xl/worksheets/sheet1.xml').async('string');
  sheet1 = clearSheetData(sheet1, 1, {
    sampleRow: 2,
    keepEmptyRowsUpTo: 1000,
    xxxIdx,
    columnsToReplace: ['B', 'C', 'D', 'E', 'F', 'G', 'H'], // A is formula
    statusColumn: 'H',
    statusValue: xxxIdx,
  });
  ypZip.file('xl/worksheets/sheet1.xml', sheet1);

  // --- Sheet2: 決定事項 ---
  let sheet2 = await ypZip.file('xl/worksheets/sheet2.xml').async('string');
  sheet2 = clearSheetData(sheet2, 2, {
    sampleRow: 2,
    keepEmptyRowsUpTo: 0, // no extra empty rows
    xxxIdx,
    columnsToReplace: ['B', 'C', 'D', 'E', 'F'], // A is formula
    dateColumns: ['B'],
  });
  ypZip.file('xl/worksheets/sheet2.xml', sheet2);

  // --- Sheet3: ネクスト (PJ管理ツール使用版) ---
  let sheet3 = await ypZip.file('xl/worksheets/sheet3.xml').async('string');
  sheet3 = clearSheetData(sheet3, 3, {
    sampleRow: 2,
    keepEmptyRowsUpTo: 0,
    xxxIdx,
    columnsToReplace: ['A'], // only A has data
  });
  ypZip.file('xl/worksheets/sheet3.xml', sheet3);

  // --- Sheet4: 事実一覧 ---
  let sheet4 = await ypZip.file('xl/worksheets/sheet4.xml').async('string');
  sheet4 = clearSheetData(sheet4, 4, {
    sampleRow: 2,
    keepEmptyRowsUpTo: 0,
    xxxIdx,
    columnsToReplace: ['B', 'C', 'D', 'E'], // A is formula
    dateColumns: ['B'],
  });
  ypZip.file('xl/worksheets/sheet4.xml', sheet4);

  // --- Sheet5: すり合わせ用メモ → メモ (空シート、そのまま) ---
  // No changes needed - already empty

  // --- Sheet6: リンク類 ---
  let sheet6 = await ypZip.file('xl/worksheets/sheet6.xml').async('string');
  sheet6 = clearSheetData(sheet6, 6, {
    sampleRow: 2,
    keepEmptyRowsUpTo: 1000,
    xxxIdx,
    columnsToReplace: ['B', 'C'], // A is formula
    dateColumns: ['B'],
  });
  ypZip.file('xl/worksheets/sheet6.xml', sheet6);

  // ============================================================
  // Step 3: ネクスト（PJ管理ツール不使用）シートを追加
  // ============================================================
  console.log('[4/6] ネクスト（PJ管理ツール不使用）シート追加...');

  // 既存のステータスインデックスを取得
  const updatedSs = ssXml;
  const allSi = updatedSs.match(/<si>[\s\S]*?<\/si>/g) || [];
  let headerIndices = {};
  for (let i = 0; i < allSi.length; i++) {
    const text = allSi[i].replace(/<[^>]+>/g, '').trim();
    if (
      ['#', '起票日', '起票者', '概要', '詳細', '備考', 'ステータス'].includes(
        text
      )
    ) {
      headerIndices[text] = i;
    }
  }
  // #のインデックス（sharedStringsの97番 in YP）
  // 起票日=0, 起票者=1, 概要=3, 詳細=4, 備考=5, ステータス=6

  // TS-style Next sheet: # / 起票日 / 起票者 / 概要 / 詳細 / 備考 / 担当 / 期限 / ステータス
  // Use YP's style IDs: header=s2(bold), data=s21(normal text), date=s22
  const sheet7Xml = createTsNextSheet({
    headerIndices,
    xxxIdx,
    tantouIdx: newIndices['担当'],
    kigenIdx: newIndices['期限'],
    yetIdx: newIndices['yet'],
  });
  ypZip.file('xl/worksheets/sheet7.xml', sheet7Xml);

  // ============================================================
  // Step 4: workbook.xml更新（シート名変更・追加）
  // ============================================================
  console.log('[5/6] workbook.xml・rels・ContentTypes更新...');

  let wbXml = await ypZip.file('xl/workbook.xml').async('string');

  // シート名変更
  wbXml = wbXml.replace(
    'name="ネクスト"',
    'name="ネクスト（PJ管理ツール使用）"'
  );
  wbXml = wbXml.replace('name="すり合わせ用メモ"', 'name="メモ"');
  wbXml = wbXml.replace('name="リンク類"', 'name="リンク集"');

  // 新しいシートを追加（sheet9の前、つまり事実一覧の後）
  // 現在: アジェンダ(rId4), 決定事項(rId5), ネクスト(rId6), 事実一覧(rId7), メモ(rId8), リンク集(rId9)
  // 追加: ネクスト不使用(rId11) をrId6の後に挿入
  const newSheetTag =
    '<sheet state="visible" name="ネクスト（PJ管理ツール不使用）" sheetId="7" r:id="rId11"/>';
  wbXml = wbXml.replace(
    /<sheet [^>]*name="事実一覧"[^>]*\/>/,
    (match) => newSheetTag + match
  );

  // definedNames: フィルター範囲をヘッダー+サンプル行に縮小
  wbXml = wbXml.replace(
    /('アジェンダ'!\$H\$1:\$H\$)\d+/,
    '$11000' // keep 1000 for filter range
  );

  ypZip.file('xl/workbook.xml', wbXml);

  // workbook.xml.rels に新しいシートのRelationship追加
  let wbRels = await ypZip.file('xl/_rels/workbook.xml.rels').async('string');
  wbRels = wbRels.replace(
    '</Relationships>',
    '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet7.xml"/></Relationships>'
  );
  ypZip.file('xl/_rels/workbook.xml.rels', wbRels);

  // Content_Types.xml に新しいシートを追加
  let ctXml = await ypZip.file('[Content_Types].xml').async('string');
  ctXml = ctXml.replace(
    '</Types>',
    '<Override PartName="/xl/worksheets/sheet7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
  );
  ypZip.file('[Content_Types].xml', ctXml);

  // ============================================================
  // Step 5: 保存
  // ============================================================
  console.log('[6/6] 保存・検証...');
  const buf = await ypZip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(OUTPUT, buf);
  console.log(`[完了] ${OUTPUT}`);

  // ============================================================
  // 検証
  // ============================================================
  const verifyData = fs.readFileSync(OUTPUT);
  const verifyZip = await JSZip.loadAsync(verifyData);
  const vWb = await verifyZip.file('xl/workbook.xml').async('string');
  const sheets = vWb.match(/<sheet [^>]+>/g) || [];
  console.log('\n検証: シート一覧');
  sheets.forEach((s, i) => {
    const name = s.match(/name="([^"]+)"/)?.[1];
    console.log(`  ${i + 1}. ${name}`);
  });

  // Verify each sheet has rows
  for (let i = 1; i <= 7; i++) {
    const sheetFile = `xl/worksheets/sheet${i}.xml`;
    const file = verifyZip.file(sheetFile);
    if (file) {
      const xml = await file.async('string');
      const rows = xml.match(/<row /g) || [];
      console.log(`  sheet${i}.xml: ${rows.length} rows`);
    }
  }
})();

// ============================================================
// ヘルパー: シートのデータクリア
// ============================================================
function clearSheetData(sheetXml, sheetNum, opts) {
  const {
    sampleRow,
    keepEmptyRowsUpTo,
    xxxIdx,
    columnsToReplace,
    dateColumns,
    statusColumn,
    statusValue,
  } = opts;

  // 全行を抽出
  const rows = sheetXml.match(/<row [^>]*>[\s\S]*?<\/row>/g) || [];
  if (rows.length === 0) return sheetXml;

  // Row 1 (header) はそのまま保持
  const headerRow = rows[0];

  // Row 2 (sample) のセルデータを×××に置換
  let sampleRowXml = rows.length > 1 ? rows[1] : '';
  if (sampleRowXml && columnsToReplace) {
    for (const col of columnsToReplace) {
      const cellRef = `${col}${sampleRow}`;
      // 日付列の場合はそのまま保持（日付シリアル値のまま）
      if (dateColumns && dateColumns.includes(col)) {
        // 日付をサンプル値（今日）に
        const today = Math.floor(
          (Date.now() - new Date('1899-12-30').getTime()) / 86400000
        );
        const cellRegex = new RegExp(
          `<c r="${cellRef}"[^>]*>([\\s\\S]*?)<\\/c>`
        );
        sampleRowXml = sampleRowXml.replace(cellRegex, (match, inner) => {
          return match.replace(/<v>[^<]*<\/v>/, `<v>${today}.0</v>`);
        });
        continue;
      }
      // 文字列セルを×××に
      const cellRegex = new RegExp(
        `<c r="${cellRef}"[^>]*>([\\s\\S]*?)<\\/c>|<c r="${cellRef}"[^>]*\\/>`
      );
      sampleRowXml = sampleRowXml.replace(cellRegex, (match) => {
        // スタイルIDを保持
        const sAttr = match.match(/s="([^"]+)"/)?.[0] || '';
        return `<c r="${cellRef}" ${sAttr} t="s"><v>${xxxIdx}</v></c>`;
      });
    }
  }

  // 残りの行を構築
  let newRows = [headerRow, sampleRowXml];

  // 空行を保持する場合（Sheet1, Sheet6: 1000行）
  if (keepEmptyRowsUpTo > 0) {
    for (let r = 2; r < rows.length; r++) {
      // 行番号を取得
      const rowNum = rows[r].match(/r="(\d+)"/)?.[1];
      if (!rowNum || parseInt(rowNum) > keepEmptyRowsUpTo) break;

      // セルのデータをクリア（値を削除、スタイルのみ残す）
      let cleanedRow = rows[r];
      // <v>...</v>を削除
      cleanedRow = cleanedRow.replace(/<v>[^<]*<\/v>/g, '');
      // t="s" や t="shared" を削除（空セルにする）
      cleanedRow = cleanedRow.replace(/ t="[^"]*"/g, '');
      // <f>...</f> (数式) を削除
      cleanedRow = cleanedRow.replace(/<f[^>]*>[\s\S]*?<\/f>/g, '');
      cleanedRow = cleanedRow.replace(/<f[^>]*\/>/g, '');
      newRows.push(cleanedRow);
    }
  }

  // sheetDataを再構築
  const sheetDataContent = newRows.filter(Boolean).join('\n');
  sheetXml = sheetXml.replace(
    /<sheetData>[\s\S]*<\/sheetData>/,
    `<sheetData>${sheetDataContent}</sheetData>`
  );

  return sheetXml;
}

// ============================================================
// ヘルパー: TS方式ネクストシートの作成
// ============================================================
function createTsNextSheet(opts) {
  const { headerIndices, xxxIdx, tantouIdx, kigenIdx, yetIdx } = opts;

  // ヘッダー: # / 起票日 / 起票者 / 概要 / 詳細 / 備考 / 担当 / 期限 / ステータス
  // YPのスタイルID: ヘッダー=s2(bold), データテキスト=s21, データ日付=s22, 連番数式=s21
  // Sheet2のスタイルをベースにする

  // #のsharedStringインデックス
  const hashIdx = headerIndices['#'] || 97; // YPでは97

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
<row r="1" spans="1:9">
<c r="A1" s="2" t="s"><v>${hashIdx}</v></c>
<c r="B1" s="2" t="s"><v>${headerIndices['起票日'] || 0}</v></c>
<c r="C1" s="2" t="s"><v>${headerIndices['起票者'] || 1}</v></c>
<c r="D1" s="2" t="s"><v>${headerIndices['概要'] || 3}</v></c>
<c r="E1" s="2" t="s"><v>${headerIndices['詳細'] || 4}</v></c>
<c r="F1" s="2" t="s"><v>${headerIndices['備考'] || 5}</v></c>
<c r="G1" s="2" t="s"><v>${tantouIdx}</v></c>
<c r="H1" s="2" t="s"><v>${kigenIdx}</v></c>
<c r="I1" s="2" t="s"><v>${headerIndices['ステータス'] || 6}</v></c>
</row>
<row r="2" spans="1:9">
<c r="A2" s="21"><v>1</v></c>
<c r="B2" s="22"><v>${Math.floor((Date.now() - new Date('1899-12-30').getTime()) / 86400000)}.0</v></c>
<c r="C2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="D2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="E2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="F2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="G2" s="21" t="s"><v>${xxxIdx}</v></c>
<c r="H2" s="22"><v>${Math.floor((Date.now() - new Date('1899-12-30').getTime()) / 86400000)}.0</v></c>
<c r="I2" s="21" t="s"><v>${yetIdx}</v></c>
</row>
</sheetData>
</worksheet>`;
}
