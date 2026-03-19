/**
 * 提案書 PPTX生成スクリプト（汎用ヘルパー）
 *
 * 使い方:
 *   1. このファイルをプロジェクトルートにコピー
 *   2. 定数セクション（FONT_FACE, PRIMARY_COLOR等）をプロジェクトに合わせて変更
 *   3. 「スライドデータ定義」セクションに各スライドのデータを記述
 *   4. node build-proposal.mjs で実行
 *
 * 前提:
 *   - npm install adm-zip が必要
 *   - テンプレートPPTX（templates/flux-presentation-template.pptx）が必要
 *
 * 最終報告書版との違い:
 *   - buildCoverSlide() / buildAgendaSlide() が追加
 *   - 提案書は7セクション構成（最終報告書は5章構成）
 *   - 検討プロセスセクション（MDのみ、スライドには含めない）
 *
 * 詳細: guides/14-proposal-workflow.md
 */

import AdmZip from 'adm-zip';

// ============================================================
// プロジェクト固有の定数（ここを変更する）
// ============================================================
const FONT_FACE = 'メイリオ';        // 日本語フォント
const PRIMARY_COLOR = '0055FF';       // プライマリカラー（テーブルヘッダー等）
const ACCENT_COLOR = 'ED7D31';        // アクセントカラー（コールアウト等）
const EVEN_ROW_BG = 'F2F7FF';         // テーブル偶数行の背景色
const CALLOUT_BG = 'D3E2FF';          // コールアウトボックスの背景色（青系）
const CALLOUT_BG_ORANGE = 'FFF3E0';   // コールアウトボックスの背景色（橙系: 期待効果等に使用）

const templatePath = './templates/flux-presentation-template.pptx';  // テンプレートPPTX
const outputPath = './提案書_v1.pptx';                            // 出力先

// ============================================================
// テンプレート読み込み
// ============================================================
const zip = new AdmZip(templatePath);

// テンプレートスライドの読み込み（テンプレートにスライドがある場合）
// ※テンプレートのスライド番号はプロジェクトごとに確認すること
// const coverTemplate = zip.readAsText('ppt/slides/slide1.xml');
// const coverRels = zip.readAsText('ppt/slides/_rels/slide1.xml.rels');
// const agendaTemplate = zip.readAsText('ppt/slides/slide2.xml');
// const agendaRels = zip.readAsText('ppt/slides/_rels/slide2.xml.rels');
// const sectionTemplate = zip.readAsText('ppt/slides/slide3.xml');
// const sectionRels = zip.readAsText('ppt/slides/_rels/slide3.xml.rels');
// const contentTemplate = zip.readAsText('ppt/slides/slide4.xml');
// const contentRels = zip.readAsText('ppt/slides/_rels/slide4.xml.rels');

// ============================================================
// ユーティリティ
// ============================================================
function escXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const emu = (inches) => Math.round(inches * 914400);

function stripBodyShapes(slideXml) {
  // プレースホルダ以外のシェイプを除去
  let result = slideXml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    if (content.includes('<p:ph')) return match;
    return '';
  });
  result = result.replace(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g, '');
  result = result.replace(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g, '');
  result = result.replace(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g, '');
  return result;
}

// notesSlide参照の除去
function stripNotesRef(relsXml) {
  return relsXml.replace(/<Relationship[^>]*notesSlide[^>]*\/>/g, '');
}

// ============================================================
// 要素生成ヘルパー
// ============================================================

/**
 * テキストボックスを生成
 * @param {number} id - シェイプID
 * @param {number} x - X位置（インチ）
 * @param {number} y - Y位置（インチ）
 * @param {number} w - 幅（インチ）
 * @param {number} h - 高さ（インチ）
 * @param {string} text - テキスト（\nで改行）
 * @param {object} opts - オプション（fontSize, bold, color, align, bullet）
 */
function makeTextBox(id, x, y, w, h, text, opts = {}) {
  const fontSize = opts.fontSize || 1000;
  const bold = opts.bold ? ' b="1"' : '';
  const color = opts.color || '000000';
  const align = opts.align || 'l';

  const lines = text.split('\n');
  const paragraphs = lines.map(line => {
    const bulletAttr = opts.bullet ? '<a:buChar char="•"/>' : '<a:buNone/>';
    const indent = opts.bullet ? ' marL="228600" indent="-228600"' : '';
    return `<a:p><a:pPr algn="${align}"${indent}>${bulletAttr}</a:pPr><a:r><a:rPr lang="ja-JP" sz="${fontSize}"${bold} dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/><a:ea typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/></a:rPr><a:t>${escXml(line)}</a:t></a:r></a:p>`;
  }).join('');

  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

const STD_ROW_H = 0.28; // 標準行高さ（単一行テーブル共通）
const LINE_H = 0.18;    // 追加行あたりの高さ

/**
 * テーブルを生成
 * @param {number} id - シェイプID
 * @param {number} x - X位置（インチ）
 * @param {number} y - Y位置（インチ）
 * @param {number[]} colWidths - 各列の幅（インチ）
 * @param {string[]} headers - ヘッダー行
 * @param {string[][]} rows - データ行
 * @param {object} opts - オプション（fontSize, headerColor）
 */
function makeTable(id, x, y, colWidths, headers, rows, opts = {}) {
  const fontSize = opts.fontSize || 1000;
  const headerColor = opts.headerColor || PRIMARY_COLOR;
  const totalW = colWidths.reduce((a, b) => a + b, 0);

  const gridCols = colWidths.map(w => `<a:gridCol w="${emu(w)}"/>`).join('');

  function makeCell(text, isHeader, rowIdx) {
    const bgColor = isHeader ? headerColor : (rowIdx % 2 === 1 ? EVEN_ROW_BG : 'FFFFFF');
    const textColor = isHeader ? 'FFFFFF' : '000000';
    const bold = isHeader ? ' b="1"' : '';
    const align = isHeader ? 'ctr' : 'l';
    const lines = String(text).split('\n');
    const paras = lines.map(l =>
      `<a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="ja-JP" sz="${fontSize}"${bold} dirty="0"><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill><a:latin typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/><a:ea typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/></a:rPr><a:t>${escXml(l)}</a:t></a:r></a:p>`
    ).join('');
    // ★重要: テーブルセルは <a:txBody> で <a:p> を囲む（必須）
    return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paras}</a:txBody><a:tcPr marL="68580" marR="68580" marT="34290" marB="34290" anchor="ctr"><a:solidFill><a:srgbClr val="${bgColor}"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></a:tcPr></a:tc>`;
  }

  function calcRowH(row) {
    const maxLines = Math.max(...row.map(cell => String(cell).split('\n').length));
    if (maxLines <= 1) return STD_ROW_H;
    return STD_ROW_H + (maxLines - 1) * LINE_H;
  }

  const headerRow = '<a:tr h="' + emu(STD_ROW_H) + '">' + headers.map(h => makeCell(h, true, 0)).join('') + '</a:tr>';
  const rowHeights = rows.map(row => calcRowH(row));
  const dataRows = rows.map((row, idx) =>
    '<a:tr h="' + emu(rowHeights[idx]) + '">' + row.map(cell => makeCell(cell, false, idx)).join('') + '</a:tr>'
  ).join('');
  const totalH = STD_ROW_H + rowHeights.reduce((a, b) => a + b, 0);

  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(totalW)}" cy="${emu(totalH)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:noFill/></a:tblPr><a:tblGrid>${gridCols}</a:tblGrid>${headerRow}${dataRows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

/**
 * コールアウトボックスを生成（角丸四角形）
 * 色の使い分け:
 *   - 青系（デフォルト）: bgColor=D3E2FF, textColor=PRIMARY_COLOR → 結論・まとめ
 *   - 橙系: bgColor=FFF3E0, textColor=ED7D31 → 期待効果・注意事項
 */
function makeCalloutBox(id, x, y, w, h, text, opts = {}) {
  const bgColor = opts.bgColor || CALLOUT_BG;
  const textColor = opts.textColor || PRIMARY_COLOR;
  const fontSize = opts.fontSize || 1000;
  // adj=5000: 角丸の曲率。10000で最大丸、0で直角。5000は控えめな角丸
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Callout ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 5000"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="${bgColor}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="ja-JP" sz="${fontSize}" b="1" dirty="0"><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill><a:latin typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/><a:ea typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/></a:rPr><a:t>${escXml(text)}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

/**
 * サブセクション見出し + 箇条書きテキスト
 * @returns {{ elements: string[], nextY: number, nextId: number }}
 */
function makeBodySection(idStart, y, heading, bullets) {
  const elements = [];
  elements.push(makeTextBox(idStart, 0.4, y, 9.2, 0.22, heading, { fontSize: 950, bold: true, color: PRIMARY_COLOR }));
  const bulletText = bullets.map(b => '• ' + b).join('\n');
  const bulletH = bullets.length * 0.17 + 0.08;
  elements.push(makeTextBox(idStart + 1, 0.5, y + 0.22, 9.0, bulletH, bulletText, { fontSize: 850 }));
  return { elements, nextY: y + 0.22 + bulletH + 0.08, nextId: idStart + 2 };
}

// ============================================================
// スライドビルド関数
// ============================================================

/**
 * コンテンツスライドを生成
 * テンプレートのコンテンツスライドを基に、プレースホルダを置換しボディ要素を追加
 */
function buildContentSlide(contentTemplate, sectionLabel, title, lead, pageNum, bodyElements) {
  let xml = stripBodyShapes(contentTemplate);

  // プレースホルダテキストを置換
  xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    if (content.includes('type="subTitle"') && content.includes('idx="2"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(sectionLabel) + '</a:t>');
    }
    if (content.includes('type="title"') && !content.match(/idx="\d+"/)) {
      let first = true;
      return match.replace(/<a:t>[^<]*<\/a:t>/g, () => {
        if (first) { first = false; return '<a:t>' + escXml(title) + '</a:t>'; }
        return '<a:t></a:t>';
      });
    }
    if (content.includes('type="body"') && content.includes('idx="1"')) {
      return ''; // リードプレースホルダは除去（テキストボックスで追加）
    }
    if (content.includes('type="sldNum"') || content.includes('idx="12"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(String(pageNum)) + '</a:t>');
    }
    return match;
  });

  // 灰色を黒に強制変換
  xml = xml.replace(/srgbClr val="434343"/g, 'srgbClr val="000000"');
  xml = xml.replace(/srgbClr val="333333"/g, 'srgbClr val="000000"');

  // リードテキストとボディ要素を挿入
  const leadBox = lead ? makeTextBox(899, 0.37, 0.85, 9.3, 0.5, lead, { fontSize: 1200, color: '000000' }) : '';
  const bodyXml = leadBox + bodyElements.join('');
  xml = xml.replace('</p:spTree>', bodyXml + '</p:spTree>');

  return xml;
}

/**
 * セクション扉スライドを生成
 */
function buildSectionSlide(sectionTemplate, sectionNum, sectionTitle, pageNum) {
  let xml = sectionTemplate;
  xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    if (content.includes('type="subTitle"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(sectionNum) + '</a:t>');
    }
    if (content.includes('type="title"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(sectionTitle) + '</a:t>');
    }
    if (!content.includes('<p:ph') && content.match(/<a:t>\d+<\/a:t>/)) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(String(pageNum)) + '</a:t>');
    }
    return match;
  });
  return xml;
}

/**
 * 表紙スライドを生成（提案書固有）
 * @param {string} coverTemplate - 表紙スライドのXML
 * @param {string} title - 提案タイトル
 * @param {string} clientName - クライアント名（「○○ 御中」の部分）
 * @param {string} date - 日付（例: "2026.XX.XX"）
 */
function buildCoverSlide(coverTemplate, title, clientName, date) {
  let xml = coverTemplate;
  xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    // メインタイトル（type="title" で idx="2" でないもの）
    if (content.includes('type="title"') && !content.includes('idx="2"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(title) + '</a:t>');
    }
    // 日付フィールド（type="title" idx="2"）
    if (content.includes('type="title"') && content.includes('idx="2"')) {
      let first = true;
      return match.replace(/<a:t>[^<]*<\/a:t>/g, () => {
        if (first) { first = false; return '<a:t>' + escXml(date) + '</a:t>'; }
        return '<a:t></a:t>';
      });
    }
    // クライアント名シェイプ（プレースホルダなし、「御中」を含む）
    if (!content.includes('<p:ph') && content.includes('御中')) {
      let first = true;
      return match.replace(/<a:t>[^<]*<\/a:t>/g, () => {
        if (first) { first = false; return '<a:t>' + escXml(clientName) + '</a:t>'; }
        return '<a:t></a:t>';
      });
    }
    return match;
  });
  return xml;
}

/**
 * アジェンダスライドを生成（提案書固有）
 *
 * ★重要: <a:t> の逐次置換は禁止。<p:txBody> ごと再構築すること。
 * 理由: テンプレートのアジェンダ段落は1つの <a:p> 内に複数の <a:t> を持つことがあり、
 *       逐次置換すると段落の区切りを超えてテキストが結合される。
 *
 * @param {string} agendaTemplate - アジェンダスライドのXML
 * @param {string[]} items - アジェンダ項目の配列（例: ['エグゼクティブサマリ', '現状分析', ...]）
 * @param {number} pageNum - ページ番号
 */
function buildAgendaSlide(agendaTemplate, items, pageNum) {
  let xml = agendaTemplate;
  xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    // タイトル
    if (content.includes('type="title"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml('アジェンダ') + '</a:t>');
    }
    // ボディ（番号付きリスト）— <a:p> 単位で段落を再構築
    if (content.includes('type="body"') && content.includes('idx="1"')) {
      const newParas = items.map(item =>
        `<a:p><a:pPr marL="425450" lvl="0" indent="-285750" algn="l" rtl="0">` +
          `<a:lnSpc><a:spcPct val="115000"/></a:lnSpc>` +
          `<a:spcBef><a:spcPts val="0"/></a:spcBef>` +
          `<a:spcAft><a:spcPts val="0"/></a:spcAft>` +
          `<a:buSzPts val="1400"/>` +
          `<a:buFont typeface="Noto Sans JP"/>` +
          `<a:buAutoNum type="arabicPeriod"/>` +  // ← 自動番号（1. 2. 3. ...）
        `</a:pPr>` +
        `<a:r><a:rPr lang="ja-JP" sz="1600" dirty="0">` +
          `<a:solidFill><a:srgbClr val="000000"/></a:solidFill>` +
          `<a:latin typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/>` +
          `<a:ea typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/>` +
          `<a:cs typeface="Noto Sans JP"/>` +
          `<a:sym typeface="Noto Sans JP"/>` +
        `</a:rPr><a:t>${escXml(item)}</a:t></a:r>` +
        `<a:endParaRPr sz="1600" dirty="0">` +
          `<a:latin typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/>` +
          `<a:ea typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/>` +
          `<a:cs typeface="Noto Sans JP"/>` +
          `<a:sym typeface="Noto Sans JP"/>` +
        `</a:endParaRPr></a:p>`
      ).join('');
      // <p:txBody> ごと置換（段落の逐次置換は禁止）
      return match.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/,
        `<p:txBody><a:bodyPr/><a:lstStyle/>${newParas}</p:txBody>`);
    }
    // ページ番号
    if (content.includes('type="sldNum"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(String(pageNum)) + '</a:t>');
    }
    return match;
  });
  return xml;
}

// ============================================================
// TODO: ここに各PJTのスライドデータを定義
// ============================================================
//
// 提案書の7セクション構成:
//   — 表紙+アジェンダ (cover, content)
//   01 エグゼクティブサマリ (section, content)
//   02 現状分析 (section, content/data ×2-4)
//   03 提案戦略 (section, content/data ×3-6)
//   04 効果見込み (section, data ×3-5)
//   05 実行計画 (section, data, content)
//   06 ネクストステップ (content, action)
//   07 参考資料 (content)
//
// 例:
// const coverTemplate = zip.readAsText('ppt/slides/slide1.xml');
// const coverRels = zip.readAsText('ppt/slides/_rels/slide1.xml.rels');
// const agendaTemplate = zip.readAsText('ppt/slides/slide2.xml');
// const agendaRels = zip.readAsText('ppt/slides/_rels/slide2.xml.rels');
// const sectionTemplate = zip.readAsText('ppt/slides/slide3.xml');
// const sectionRels = zip.readAsText('ppt/slides/_rels/slide3.xml.rels');
// const contentTemplate = zip.readAsText('ppt/slides/slide4.xml');
// const contentRels = zip.readAsText('ppt/slides/_rels/slide4.xml.rels');
//
// let n = 900; // shapeIDカウンター
// const slides = [];
//
// // Slide 1: 表紙
// slides.push({ xml: buildCoverSlide(coverTemplate, '提案タイトル', 'クライアント名 御中', '2026.XX.XX'), rels: coverRels });
//
// // Slide 2: アジェンダ
// slides.push({ xml: buildAgendaSlide(agendaTemplate, ['エグゼクティブサマリ', '現状分析', '提案戦略', '効果見込み', '実行計画', 'ネクストステップ'], 2), rels: agendaRels });
//
// // Slide 3: セクション扉 01
// slides.push({ xml: buildSectionSlide(sectionTemplate, '01', 'エグゼクティブサマリ', 3), rels: sectionRels });
//
// // Slide 4: エグゼクティブサマリ本体
// slides.push({
//   xml: buildContentSlide(contentTemplate, '01 エグゼクティブサマリ', '結論を言い切るタイトル',
//     'リード文（数字を含む）', 4, [
//       makeTable(n++, 0.4, 1.5, [3.0, 6.2], ['項目', '数値'], [['目標', 'XXX']], {}),
//     ]),
//   rels: contentRels
// });

const slides = [];

if (slides.length === 0) {
  console.log('スライドデータが定義されていません。');
  console.log('このファイルの「TODO: ここに各PJTのスライドデータを定義」セクションを編集してください。');
  console.log('詳細: guides/14-proposal-workflow.md');
  process.exit(0);
}

// ============================================================
// PPTX組み立て
// ============================================================
console.log('Building PPTX with ' + slides.length + ' slides...');

// 既存スライドを削除（テンプレートにスライドがある場合）
const existingEntries = zip.getEntries().map(e => e.entryName);
const existingSlides = existingEntries.filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e));
existingSlides.forEach(entry => {
  const num = entry.match(/slide(\d+)/)[1];
  zip.deleteFile(entry);
  try { zip.deleteFile(`ppt/slides/_rels/slide${num}.xml.rels`); } catch(e) {}
  try { zip.deleteFile(`ppt/notesSlides/notesSlide${num}.xml`); } catch(e) {}
  try { zip.deleteFile(`ppt/notesSlides/_rels/notesSlide${num}.xml.rels`); } catch(e) {}
});

// 新スライドを追加
slides.forEach((s, i) => {
  const num = i + 1;
  zip.addFile('ppt/slides/slide' + num + '.xml', Buffer.from(s.xml, 'utf-8'));
  zip.addFile('ppt/slides/_rels/slide' + num + '.xml.rels', Buffer.from(stripNotesRef(s.rels), 'utf-8'));
});

// presentation.xml 更新
let presXml = zip.readAsText('ppt/presentation.xml');
presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, () => {
  let refs = '<p:sldIdLst>';
  slides.forEach((s, i) => {
    refs += `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`;
  });
  refs += '</p:sldIdLst>';
  return refs;
});
zip.updateFile('ppt/presentation.xml', Buffer.from(presXml, 'utf-8'));

// presentation.xml.rels 更新
let presRels = zip.readAsText('ppt/_rels/presentation.xml.rels');
presRels = presRels.replace(/<Relationship[^>]*Target="slides\/slide\d+\.xml"[^>]*\/>/g, '');
const newRels = slides.map((s, i) =>
  `<Relationship Id="rId${100+i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`
).join('');
presRels = presRels.replace('</Relationships>', newRels + '</Relationships>');
zip.updateFile('ppt/_rels/presentation.xml.rels', Buffer.from(presRels, 'utf-8'));

// [Content_Types].xml 更新
let contentTypes = zip.readAsText('[Content_Types].xml');
contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, '');
contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/notesSlides\/notesSlide\d+\.xml"[^>]*\/>/g, '');
const newOverrides = slides.map((s, i) =>
  `<Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
).join('');
contentTypes = contentTypes.replace('</Types>', newOverrides + '</Types>');
zip.updateFile('[Content_Types].xml', Buffer.from(contentTypes, 'utf-8'));

// 出力
zip.writeZip(outputPath);
console.log('Done! Output: ' + outputPath);
console.log('Total slides: ' + slides.length);
