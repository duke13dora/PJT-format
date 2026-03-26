/**
 * ヒアリング資料 PPTX生成スクリプト（汎用ヘルパー）
 *
 * 使い方:
 *   1. このファイルをプロジェクトルートにコピー
 *   2. 定数セクション（FONT_FACE, PRIMARY_COLOR等）をプロジェクトに合わせて変更
 *   3. 「スライドデータ定義」セクションに各スライドのデータを記述
 *   4. node build-hearing.mjs で実行
 *
 * 前提:
 *   - npm install adm-zip が必要
 *   - テンプレートPPTX（templates/flux-presentation-template.pptx）が必要
 *   - フロー図PNG（Mermaid等で事前生成）が必要
 *
 * 提案書版との違い:
 *   - buildFlowSlide() が追加（画像埋め込み + 質問ボックス）
 *   - 画像レジストリ（registerImage/makePicture）で複数フロー図を管理
 *   - QUESTION_BG / HIGHLIGHT_BG カラー定義
 *   - 仮説マーク対応（opts.hypothesis）
 *
 * 重要: テンプレートには ppt/media/image1-4.png（ロゴ等）が存在する。
 * フロー図のファイル名は flow_1.png 等のプレフィックス付きにすること。
 * image1.png 等にするとテンプレート画像を上書きし、
 * 全スライドにゴースト画像が表示される。
 *
 * 詳細: guides/15-hearing-workflow.md
 */

import AdmZip from 'adm-zip';
import { readFileSync } from 'fs';
import path from 'path';

// ============================================================
// プロジェクト固有の定数（ここを変更する）
// ============================================================
const FONT_FACE = 'メイリオ';         // 日本語フォント
const PRIMARY_COLOR = '1B3A5C';       // プライマリカラー（テーブルヘッダー等）
const ACCENT_COLOR = 'E85D3A';        // アクセントカラー
const EVEN_ROW_BG = 'EDF2F7';         // テーブル偶数行の背景色
const HIGHLIGHT_BG = 'FFF3E0';        // 仮説マーク等のハイライト色
const QUESTION_BG = 'F0F4FF';         // 質問ボックスの背景色

const templatePath = './templates/flux-presentation-template.pptx';  // テンプレートPPTX
const outputPath = './ヒアリング資料_v0.1.pptx';                     // 出力先
const flowImageDir = './tmp-mermaid';                                 // フロー図PNGの格納先

// ============================================================
// テンプレート読み込み
// ============================================================
const zip = new AdmZip(templatePath);

// テンプレートのスライドXMLを読み込み（レイアウト別）
const coverTemplate = zip.readAsText('ppt/slides/slide1.xml');       // 表紙
const coverRels = zip.readAsText('ppt/slides/_rels/slide1.xml.rels');
const agendaTemplate = zip.readAsText('ppt/slides/slide2.xml');      // アジェンダ
const agendaRels = zip.readAsText('ppt/slides/_rels/slide2.xml.rels');
const contentTemplate = zip.readAsText('ppt/slides/slide5.xml');     // コンテンツ（汎用）
const contentRels = zip.readAsText('ppt/slides/_rels/slide5.xml.rels');

// ============================================================
// ユーティリティ
// ============================================================
function escXml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
const emu = (inches) => Math.round(inches * 914400);

/** テンプレートスライドXMLからプレースホルダー以外の図形を削除 */
function stripBodyShapes(slideXml) {
  let result = slideXml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    if (content.includes('<p:ph')) return match;
    return '';
  });
  result = result.replace(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g, '');
  result = result.replace(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g, '');
  result = result.replace(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g, '');
  return result;
}

function stripNotesRef(relsXml) { return relsXml.replace(/<Relationship[^>]*notesSlide[^>]*\/>/g, ''); }

let shapeId = 900;

// ============================================================
// 要素生成ヘルパー
// ============================================================

/** テキストボックス（透明背景） */
function makeTextBox(x, y, w, h, text, opts = {}) {
  const id = shapeId++;
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

/** 塗りつぶしボックス（角丸） */
function makeFilledBox(x, y, w, h, text, bgColor, textColor, opts = {}) {
  const id = shapeId++;
  const fontSize = opts.fontSize || 900;
  const bold = opts.bold !== false ? ' b="1"' : '';
  const align = opts.align || 'l';
  const lines = text.split('\n');
  const paragraphs = lines.map(line =>
    `<a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="ja-JP" sz="${fontSize}"${bold} dirty="0"><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill><a:latin typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/><a:ea typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/></a:rPr><a:t>${escXml(line)}</a:t></a:r></a:p>`
  ).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Box ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 5000"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="${bgColor}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

const STD_ROW_H = 0.28;
const LINE_H = 0.16;

/** テーブル */
function makeTable(x, y, colWidths, headers, rows, opts = {}) {
  const id = shapeId++;
  const fontSize = opts.fontSize || 900;
  const headerColor = opts.headerColor || PRIMARY_COLOR;
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  const gridCols = colWidths.map(w => `<a:gridCol w="${emu(w)}"/>`).join('');
  function makeCell(text, isHeader, rowIdx) {
    const bgColor = isHeader ? headerColor : (rowIdx % 2 === 1 ? EVEN_ROW_BG : 'FFFFFF');
    const textColor = isHeader ? 'FFFFFF' : '000000';
    const bld = isHeader ? ' b="1"' : '';
    const align = isHeader ? 'ctr' : 'l';
    const lines = String(text).split('\n');
    const paras = lines.map(l =>
      `<a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="ja-JP" sz="${fontSize}"${bld} dirty="0"><a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill><a:latin typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/><a:ea typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/></a:rPr><a:t>${escXml(l)}</a:t></a:r></a:p>`
    ).join('');
    return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paras}</a:txBody><a:tcPr marL="68580" marR="68580" marT="34290" marB="34290" anchor="ctr"><a:solidFill><a:srgbClr val="${bgColor}"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></a:tcPr></a:tc>`;
  }
  function calcRowH(row) {
    const maxLines = Math.max(...row.map(cell => String(cell).split('\n').length));
    return maxLines <= 1 ? STD_ROW_H : STD_ROW_H + (maxLines - 1) * LINE_H;
  }
  const headerRow = '<a:tr h="' + emu(STD_ROW_H) + '">' + headers.map(h => makeCell(h, true, 0)).join('') + '</a:tr>';
  const rowHeights = rows.map(row => calcRowH(row));
  const dataRows = rows.map((row, idx) =>
    '<a:tr h="' + emu(rowHeights[idx]) + '">' + row.map(cell => makeCell(cell, false, idx)).join('') + '</a:tr>'
  ).join('');
  const totalH = STD_ROW_H + rowHeights.reduce((a, b) => a + b, 0);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(totalW)}" cy="${emu(totalH)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:noFill/></a:tblPr><a:tblGrid>${gridCols}</a:tblGrid>${headerRow}${dataRows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

// ============================================================
// 画像埋め込み
// ============================================================
const imageRegistry = [];
let imageCounter = 0;

// 画像のrId管理: slideIndex → [{rId, mediaName}]
const slideImages = {};

/**
 * フロー図をスライドに登録。rIdを返す。
 * 重要: mediaNameは flow_N.png 形式にする（image1.png等はテンプレートと衝突）
 */
function registerImage(slideIdx, pngName) {
  imageCounter++;
  const mediaName = `flow_${imageCounter}.png`;
  const rId = `rImg${imageCounter}`;
  if (!slideImages[slideIdx]) slideImages[slideIdx] = [];
  slideImages[slideIdx].push({ rId, mediaName, pngName });
  return rId;
}

/** 画像要素（p:pic） */
function makePicture(x, y, w, h, rId) {
  const id = shapeId++;
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

// ============================================================
// スライドビルダー
// ============================================================

/** 汎用コンテンツスライド */
function buildContentSlide(sectionLabel, title, lead, pageNum, bodyElements) {
  let xml = stripBodyShapes(contentTemplate);
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
    if (content.includes('type="body"') && content.includes('idx="1"')) { return ''; }
    if (content.includes('type="sldNum"') || content.includes('idx="12"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(String(pageNum)) + '</a:t>');
    }
    return match;
  });
  xml = xml.replace(/srgbClr val="434343"/g, 'srgbClr val="000000"');
  xml = xml.replace(/srgbClr val="333333"/g, 'srgbClr val="000000"');
  const leadBox = lead ? makeTextBox(0.37, 0.85, 9.3, 0.45, lead, { fontSize: 1050, color: '333333' }) : '';
  const bodyXml = leadBox + bodyElements.join('');
  xml = xml.replace('</p:spTree>', bodyXml + '</p:spTree>');
  return xml;
}

/** 表紙スライド */
function buildCoverSlide(title, clientName, date) {
  let xml = coverTemplate;
  xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    if (content.includes('type="title"') && !content.includes('idx="2"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(title) + '</a:t>');
    }
    if (content.includes('type="title"') && content.includes('idx="2"')) {
      let first = true;
      return match.replace(/<a:t>[^<]*<\/a:t>/g, () => {
        if (first) { first = false; return '<a:t>' + escXml(date) + '</a:t>'; }
        return '<a:t></a:t>';
      });
    }
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

/** アジェンダスライド: テンプレートの非プレースホルダーSPの段落を置換 */
function buildAgendaSlide(items, pageNum) {
  let xml = agendaTemplate;
  xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    if (content.includes('type="title"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml('本日のアジェンダ') + '</a:t>');
    }
    if (content.includes('type="sldNum"') || content.includes('idx="12"')) {
      return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(String(pageNum)) + '</a:t>');
    }
    // 非プレースホルダーで「エグゼクティブ」or「01」を含む → アジェンダ本体
    if (!content.includes('<p:ph') && (content.includes('エグゼクティブ') || content.includes('01'))) {
      const newParas = items.map((item, i) => {
        const num = String(i + 1).padStart(2, '0');
        return `<a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="400"/></a:spcBef></a:pPr><a:r><a:rPr lang="ja-JP" sz="1600" dirty="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/><a:ea typeface="${FONT_FACE}" panose="020B0604030504040204" pitchFamily="50" charset="-128"/></a:rPr><a:t>${escXml(num + '   ' + item)}</a:t></a:r></a:p>`;
      }).join('');
      return match.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/,
        `<p:txBody><a:bodyPr/><a:lstStyle/>${newParas}</p:txBody>`);
    }
    return match;
  });
  return xml;
}

/**
 * フロースライド（画像 + 質問ボックス）
 * ヒアリング資料特有のレイアウト: 上半分にフロー図、下半分に質問リスト
 */
function buildFlowSlide(sectionLabel, title, lead, pageNum, imageRId, questions, opts = {}) {
  const elements = [];

  // フロー画像
  elements.push(makePicture(0.5, 1.35, 5.5, 2.3, imageRId));

  // 仮説マーク
  if (opts.hypothesis) {
    elements.push(makeFilledBox(6.2, 1.35, 3.3, 0.3, '※ 仮説ベース — 実態を確認したい', HIGHLIGHT_BG, ACCENT_COLOR, { fontSize: 800 }));
  }

  // 質問ボックス
  const qStartY = 3.8;
  elements.push(makeTextBox(0.4, qStartY - 0.3, 9.2, 0.28, '▼ 確認したいこと', { fontSize: 950, bold: true, color: PRIMARY_COLOR }));
  const qText = questions.join('\n');
  const qH = Math.max(questions.length * 0.19 + 0.1, 0.8);
  elements.push(makeFilledBox(0.4, qStartY, 9.2, qH, qText, QUESTION_BG, '333333', { fontSize: 850, bold: false }));

  return buildContentSlide(sectionLabel, title, lead, pageNum, elements);
}

// ============================================================
// スライドデータ定義（ここをプロジェクトに合わせて変更する）
// ============================================================
const slides = [];

// --- Slide 1: 表紙 ---
slides.push({
  xml: buildCoverSlide(
    'ヒアリング資料タイトル\nプロジェクト名',
    'クライアント名 御中',
    '20XX年XX月'
  ),
  rels: coverRels
});

// --- Slide 2: アジェンダ ---
slides.push({
  xml: buildAgendaSlide([
    'プロジェクト概要・課題の全体像',
    '現状フロー確認',
    'あるべき姿の提示',
    '過渡期の進め方',
    '確認・ディスカッション事項',
    '今後のスケジュール',
  ], 2),
  rels: agendaRels
});

// --- Slide 3: プロジェクト概要 ---
slides.push({
  xml: buildContentSlide('01 概要', 'プロジェクト概要',
    'プロジェクトの背景と目的を1-2行で。', 3, [
      makeTable(0.4, 1.5, [2.0, 7.2], ['項目', '内容'], [
        ['背景', '（ここに背景を記述）'],
        ['目的', '（ここに目的を記述）'],
        ['スコープ', '（ここにスコープを記述）'],
        ['アプローチ', '（ここにアプローチを記述）'],
      ]),
    ]),
  rels: contentRels
});

// --- Slide 4-N: 現状フロー（フロースライドの例）---
// ※ 実際のプロジェクトでは業務パターンの数だけ繰り返す
const flowSlideExample = {
  label: '02 現状フロー',
  title: '現状フロー: パターン名（拠点名）',
  lead: 'このフローの概要を1-2行で。',
  png: 'flow-current-example',  // tmp-mermaid/ 内のファイル名（拡張子なし）
  questions: [
    'Q1: 質問文をここに記述',
    'Q2: 質問文をここに記述',
    'Q3: 質問文をここに記述',
  ],
  hypothesis: false,  // trueで仮説マーク表示
};

// フロースライドの登録例（実際は配列でループ）
{
  const slideIdx = slides.length;
  const rId = registerImage(slideIdx, flowSlideExample.png);
  slides.push({
    xml: buildFlowSlide(
      flowSlideExample.label, flowSlideExample.title, flowSlideExample.lead,
      4, rId, flowSlideExample.questions, { hypothesis: flowSlideExample.hypothesis }
    ),
    rels: contentRels
  });
}

// --- 以降、あるべき姿フロー、過渡期、確認事項、スケジュール、Appendix ---
// ※ buildContentSlide() / buildFlowSlide() を使って同様に追加

// ============================================================
// PPTX組み立て（この部分は原則変更不要）
// ============================================================
console.log('Building PPTX with ' + slides.length + ' slides...');

// 画像をzipに追加
for (const [slideIdxStr, imgs] of Object.entries(slideImages)) {
  for (const img of imgs) {
    const pngData = readFileSync(path.join(flowImageDir, img.pngName + '.png'));
    zip.addFile('ppt/media/' + img.mediaName, pngData);
  }
}

// 既存スライド削除
const existingEntries = zip.getEntries().map(e => e.entryName);
existingEntries.filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e)).forEach(entry => {
  const num = entry.match(/slide(\d+)/)[1];
  zip.deleteFile(entry);
  try { zip.deleteFile(`ppt/slides/_rels/slide${num}.xml.rels`); } catch(e) {}
  try { zip.deleteFile(`ppt/notesSlides/notesSlide${num}.xml`); } catch(e) {}
  try { zip.deleteFile(`ppt/notesSlides/_rels/notesSlide${num}.xml.rels`); } catch(e) {}
});

// スライド追加
slides.forEach((s, i) => {
  const num = i + 1;
  zip.addFile('ppt/slides/slide' + num + '.xml', Buffer.from(s.xml, 'utf-8'));

  let rels = stripNotesRef(s.rels);
  if (slideImages[i]) {
    const imgRels = slideImages[i].map(img =>
      `<Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${img.mediaName}"/>`
    ).join('');
    rels = rels.replace('</Relationships>', imgRels + '</Relationships>');
  }
  zip.addFile('ppt/slides/_rels/slide' + num + '.xml.rels', Buffer.from(rels, 'utf-8'));
});

// presentation.xml
let presXml = zip.readAsText('ppt/presentation.xml');
presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, () => {
  let refs = '<p:sldIdLst>';
  slides.forEach((s, i) => { refs += `<p:sldId id="${256 + i}" r:id="rId${100 + i}"/>`; });
  refs += '</p:sldIdLst>';
  return refs;
});
zip.updateFile('ppt/presentation.xml', Buffer.from(presXml, 'utf-8'));

// presentation.xml.rels
let presRels = zip.readAsText('ppt/_rels/presentation.xml.rels');
presRels = presRels.replace(/<Relationship[^>]*Target="slides\/slide\d+\.xml"[^>]*\/>/g, '');
const newRels = slides.map((s, i) =>
  `<Relationship Id="rId${100+i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`
).join('');
presRels = presRels.replace('</Relationships>', newRels + '</Relationships>');
zip.updateFile('ppt/_rels/presentation.xml.rels', Buffer.from(presRels, 'utf-8'));

// [Content_Types].xml
let contentTypes = zip.readAsText('[Content_Types].xml');
contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, '');
contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/ppt\/notesSlides\/notesSlide\d+\.xml"[^>]*\/>/g, '');
const newOverrides = slides.map((s, i) =>
  `<Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
).join('');
if (!contentTypes.includes('image/png')) {
  contentTypes = contentTypes.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
}
contentTypes = contentTypes.replace('</Types>', newOverrides + '</Types>');
zip.updateFile('[Content_Types].xml', Buffer.from(contentTypes, 'utf-8'));

zip.writeZip(outputPath);
console.log('Done! Output: ' + outputPath);
console.log('Total slides: ' + slides.length);
