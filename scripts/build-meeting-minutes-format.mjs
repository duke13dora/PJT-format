// build-meeting-minutes-format.mjs
// PJT-YPの議事メモ(docx)をベースに、中身をプレースホルダーに置換したフォーマットを生成
//
// Usage: node scripts/build-meeting-minutes-format.mjs
//
// ソース: C:\PJT-YP\drive\50.打ち合わせ関連\【YP様】議事メモ.docx
// 出力:   C:\PJT-format\templates\meeting-minutes-format.docx

import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

const SOURCE = 'C:/PJT-YP/drive/50.打ち合わせ関連/【YP様】議事メモ.docx';
const OUTPUT = path.join(
  path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Z]:)/, '$1')),
  '..',
  'templates',
  'meeting-minutes-format.docx'
);

// ============================================================
// テンプレートに残す段落構成
// ============================================================
// Para 0: [Heading1] YYMMDD_×××
// Para 1: [区切り線] (border)
// Para 2: [ilvl=0]   決定事項
// Para 3: [ilvl=1]   ×××
// Para 5: [ilvl=0]   ネクスト
// Para 6: [ilvl=1]   ×××
// Para 8: [ilvl=0]   議事メモ
// Para 9: [ilvl=1]   ×××（議題見出し）
// Para 10:[ilvl=2]   ×××（議事内容）
//
// Para 4, 7 (リンク付き行) は削除
// Para 11+ (実データ) は全削除

(async () => {
  console.log('[1/4] ソースファイル読み込み...');
  const data = fs.readFileSync(SOURCE);
  const zip = await JSZip.loadAsync(data);
  let docXml = await zip.file('word/document.xml').async('string');

  // ============================================================
  // Step 1: 段落を抽出
  // ============================================================
  console.log('[2/4] 段落解析・テンプレート構築...');

  // <w:body>...</w:body> の中身を操作
  const bodyMatch = docXml.match(/<w:body>([\s\S]*)<\/w:body>/);
  if (!bodyMatch) throw new Error('w:body not found');
  const bodyContent = bodyMatch[1];

  // 段落を分割（<w:p ...>...</w:p> or <w:p/>）
  const paraRegex = /<w:p[ >\/][\s\S]*?(?:<\/w:p>|<w:p\/>)/g;
  const allParas = bodyContent.match(paraRegex) || [];
  console.log(`  段落数: ${allParas.length}`);

  // 残す段落のインデックス: 0,1,2,3,5,6,8,9,10
  const keepIndices = [0, 1, 2, 3, 5, 6, 8, 9, 10];
  const keptParas = keepIndices.map((i) => allParas[i]).filter(Boolean);

  // ============================================================
  // Step 2: テキストをプレースホルダーに置換
  // ============================================================
  function replaceAllText(paraXml, newText) {
    // rPh（フリガナ）を除去
    let cleaned = paraXml.replace(/<w:rPh[\s\S]*?<\/w:rPh>/g, '');

    // ハイパーリンクを通常テキストに変換
    cleaned = cleaned.replace(/<w:hyperlink[\s\S]*?<\/w:hyperlink>/g, '');

    // 全ての<w:r>...</w:r>を削除して、1つだけ新しいランを作成
    // まず既存のランからrPr（ランスタイル）を取得
    const rPrMatch = cleaned.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    const rPr = rPrMatch ? `<w:rPr>${rPrMatch[1]}</w:rPr>` : '';

    // 全ての<w:r>を削除
    cleaned = cleaned.replace(/<w:r>[\s\S]*?<\/w:r>/g, '');
    cleaned = cleaned.replace(/<w:r [\s\S]*?<\/w:r>/g, '');

    // </w:pPr>の後に新しいランを挿入
    if (cleaned.includes('</w:pPr>')) {
      cleaned = cleaned.replace(
        '</w:pPr>',
        `</w:pPr><w:r>${rPr}<w:t>${escapeXml(newText)}</w:t></w:r>`
      );
    } else {
      // pPrがない場合は<w:p>の直後に挿入
      cleaned = cleaned.replace(
        /<w:p([^>]*)>/,
        `<w:p$1><w:r>${rPr}<w:t>${escapeXml(newText)}</w:t></w:r>`
      );
    }

    return cleaned;
  }

  function escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // 各段落の置換マッピング
  // keepIndices: [0, 1, 2, 3, 5, 6, 8, 9, 10]
  const replacements = {
    0: 'YYMMDD_×××', // Heading1タイトル
    // 1: そのまま（区切り線）
    // 2: そのまま（「決定事項」見出し）
    3: '×××', // 決定事項の内容
    // 5→index 4: そのまま（「ネクスト」見出し）
    6: '×××', // ネクストの内容
    // 8→index 6: そのまま（「議事メモ」見出し）
    9: '×××', // 議題見出し
    10: '×××', // 議事内容
  };

  // keptParas配列のインデックスとkeepIndices配列の対応
  for (let ki = 0; ki < keepIndices.length; ki++) {
    const origIdx = keepIndices[ki];
    if (replacements[origIdx] !== undefined) {
      keptParas[ki] = replaceAllText(keptParas[ki], replacements[origIdx]);
    }
  }

  // ============================================================
  // Step 3: document.xmlを再構築
  // ============================================================
  // <w:body>の後のsectPr（セクションプロパティ）を保持
  const sectPrMatch = bodyContent.match(/<w:sectPr[\s\S]*<\/w:sectPr>/);
  const sectPr = sectPrMatch ? sectPrMatch[0] : '';

  const newBody = keptParas.join('\n') + '\n' + sectPr;
  docXml = docXml.replace(
    /<w:body>[\s\S]*<\/w:body>/,
    `<w:body>${newBody}</w:body>`
  );

  // ============================================================
  // Step 4: ハイパーリンクのRelationshipを削除
  // ============================================================
  console.log('[3/4] ハイパーリンク・リレーション除去...');
  let relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
  // hyperlink typeのRelationshipを全て削除
  relsXml = relsXml.replace(/<Relationship[^>]*hyperlink[^>]*\/>/g, '');

  // ============================================================
  // Step 5: 保存
  // ============================================================
  console.log('[4/4] 保存...');
  zip.file('word/document.xml', docXml);
  zip.file('word/_rels/document.xml.rels', relsXml);

  const buf = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(OUTPUT, buf);

  console.log(`[完了] ${OUTPUT}`);

  // ============================================================
  // 検証: 生成ファイルの段落を確認
  // ============================================================
  const verifyData = fs.readFileSync(OUTPUT);
  const verifyZip = await JSZip.loadAsync(verifyData);
  const verifyDoc = await verifyZip.file('word/document.xml').async('string');
  const verifyParas =
    verifyDoc.match(/<w:p[ >\/][\s\S]*?(?:<\/w:p>|<w:p\/>)/g) || [];

  console.log(`\n検証: ${verifyParas.length}段落`);
  for (let i = 0; i < verifyParas.length; i++) {
    const p = verifyParas[i];
    const styleMatch = p.match(/<w:pStyle w:val="([^"]+)"/);
    const style = styleMatch ? styleMatch[1] : '-';
    const ilvlMatch = p.match(/<w:ilvl w:val="([^"]+)"/);
    const ilvl = ilvlMatch ? ilvlMatch[1] : '-';
    const hasBorder = p.includes('w:pBdr');
    // Extract text
    const cleaned = p.replace(/<w:rPh[\s\S]*?<\/w:rPh>/g, '');
    const tMatches = cleaned.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
    const text = tMatches
      ? tMatches.map((t) => t.replace(/<[^>]+>/g, '')).join('')
      : '';
    console.log(
      `  Para ${i}: style=${style} ilvl=${ilvl}${hasBorder ? ' [BORDER]' : ''} | ${text}`
    );
  }
})();
