# 既存PPTX XML直接編集ガイド

> 既存のPPTXファイルをadm-zip/JSZipで直接編集する手法。pptxgenjs（新規生成）では対応できない、テンプレートベースのスライド生成・既存スライドの修正に使用する。PJT-TSで確立。

---

## 新規生成 vs 既存ファイル編集の使い分け

| シナリオ | 推奨ツール | 理由 |
|---------|-----------|------|
| ガントチャートを一から作る | pptxgenjs（03参照） | 新規生成が得意。テーマ・レイアウトを自由に設計可 |
| 既存テンプレートにスライド追加 | adm-zip XML直接編集 | レイアウト・ロゴ・フッターを保持したまま操作可 |
| 既存スライドのテキスト差し替え | adm-zip XML直接編集 | プレースホルダの中身だけ変更 |
| 赤線（コネクタ）の位置移動 | adm-zip/JSZip XML直接編集 | 座標値のみ変更 |

---

## 基本パイプライン

### adm-zipで読み書き

```javascript
import AdmZip from 'adm-zip';

// 1. テンプレートPPTXを読み込み
const zip = new AdmZip('template.pptx');

// 2. XMLを読み取り
const slideXml = zip.readAsText('ppt/slides/slide1.xml');
const presXml = zip.readAsText('ppt/presentation.xml');
const presRels = zip.readAsText('ppt/_rels/presentation.xml.rels');
const contentTypes = zip.readAsText('[Content_Types].xml');

// 3. XMLを編集（後述）
// ...

// 4. 更新して保存
zip.updateFile('ppt/slides/slide1.xml', Buffer.from(updatedXml, 'utf-8'));
zip.writeZip('output.pptx');
```

### 検証にはJSZipを使用

adm-zipで書いたファイルをadm-zipで再読み込みすると `No descriptor present` エラーが発生する。検証にはJSZipを使用:

```javascript
import JSZip from 'jszip';
import fs from 'fs';

const data = fs.readFileSync('output.pptx');
const verifyZip = await JSZip.loadAsync(data);
const verifySlide = await verifyZip.file('ppt/slides/slide1.xml').async('string');
// 内容を確認
```

---

## スライドの追加・削除

### 追加に必要な更新箇所（4ファイル）

#### 1. スライドXML + rels の追加

```javascript
// スライド本体
zip.addFile('ppt/slides/slide25.xml', Buffer.from(slideXml, 'utf-8'));
// スライドのリレーション
zip.addFile('ppt/slides/_rels/slide25.xml.rels', Buffer.from(slideRels, 'utf-8'));
```

#### 2. presentation.xml の `<p:sldIdLst>` 更新

```javascript
let presXml = zip.readAsText('ppt/presentation.xml');
presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, () => {
  let refs = '<p:sldIdLst>';
  for (let i = 0; i < slideCount; i++) {
    const id = 256 + i;           // sldId（256以上の連番）
    const rId = 'rId' + (100 + i); // ★ rId100+i で既存と衝突回避
    refs += `<p:sldId id="${id}" r:id="${rId}"/>`;
  }
  refs += '</p:sldIdLst>';
  return refs;
});
```

#### 3. presentation.xml.rels の Relationship 更新

```javascript
let presRels = zip.readAsText('ppt/_rels/presentation.xml.rels');
// 既存のスライド参照を除去
presRels = presRels.replace(/<Relationship[^>]*Target="slides\/slide\d+\.xml"[^>]*\/>/g, '');
// 新しいスライド参照を追加
const newRels = slides.map((s, i) => {
  const rId = 'rId' + (100 + i);
  return `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`;
}).join('');
presRels = presRels.replace('</Relationships>', newRels + '</Relationships>');
```

#### 4. [Content_Types].xml の Override 更新

```javascript
let ct = zip.readAsText('[Content_Types].xml');
// 既存のスライドOverrideを除去
ct = ct.replace(/<Override[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, '');
// 新しいOverrideを追加
const overrides = slides.map((s, i) =>
  `<Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
).join('');
ct = ct.replace('</Types>', overrides + '</Types>');
```

### rId衝突回避（★重要）

テンプレートPPTXには既にrIdが割り当てられている:
- rId1〜rId25: 既存スライド、slideMaster、theme等
- rId26: notesMaster
- rId27+: fonts

**新しいスライドにはrId100+iを使用する**ことで、既存のrIdと衝突しない。

### notesSlide の処理

```javascript
// スライドrelsからnotesSlide参照を除去
function stripNotesRef(relsXml) {
  return relsXml.replace(/<Relationship[^>]*notesSlide[^>]*\/>/g, '');
}

// 孤立したnotesSlideファイルも削除
for (let i = 1; i <= existingSlideCount; i++) {
  try { zip.deleteFile('ppt/notesSlides/notesSlide' + i + '.xml'); } catch(e) {}
  try { zip.deleteFile('ppt/notesSlides/_rels/notesSlide' + i + '.xml.rels'); } catch(e) {}
}

// [Content_Types].xmlからもnotesSlideのOverrideを除去
ct = ct.replace(/<Override[^>]*PartName="\/ppt\/notesSlides\/notesSlide\d+\.xml"[^>]*\/>/g, '');
```

---

## プレースホルダの操作

### callback-based replace（★必須パターン）

```javascript
// 正しい方法: callback-based replace
xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
  if (content.includes('type="title"')) {
    return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(newTitle) + '</a:t>');
  }
  if (content.includes('type="subTitle"')) {
    return match.replace(/<a:t>[^<]*<\/a:t>/g, '<a:t>' + escXml(sectionLabel) + '</a:t>');
  }
  return match;
});
```

### ボディシェイプの削除（プレースホルダ以外を除去）

```javascript
function stripBodyShapes(slideXml) {
  let result = slideXml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
    if (content.includes('<p:ph')) return match; // プレースホルダは残す
    return ''; // それ以外は削除
  });
  result = result.replace(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g, '');
  result = result.replace(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g, '');
  result = result.replace(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g, '');
  return result;
}
```

---

## テーブル・テキストボックス・コールアウトボックスの生成

### テキストボックス

```javascript
function makeTextBox(id, x, y, w, h, text, opts = {}) {
  const fontSize = opts.fontSize || 1000;
  const bold = opts.bold ? ' b="1"' : '';
  const color = opts.color || '000000';
  const fontFace = '{FONT_FACE}'; // メイリオ等

  return `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/>
      <p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm>
      <a:off x="${emu(x)}" y="${emu(y)}"/>
      <a:ext cx="${emu(w)}" cy="${emu(h)}"/>
    </a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
    <p:txBody><a:bodyPr wrap="square" anchor="t"><a:noAutofit/></a:bodyPr>
      <a:lstStyle/>
      <a:p><a:r><a:rPr lang="ja-JP" sz="${fontSize}"${bold} dirty="0">
        <a:solidFill><a:srgbClr val="${color}"/></a:solidFill>
        <a:latin typeface="${fontFace}"/><a:ea typeface="${fontFace}"/>
      </a:rPr><a:t>${escXml(text)}</a:t></a:r></a:p>
    </p:txBody>
  </p:sp>`;
}
```

### テーブル

```javascript
function makeTable(id, x, y, colWidths, headers, rows, opts = {}) {
  const headerColor = opts.headerColor || '{PRIMARY_COLOR}';
  // テーブルは<p:graphicFrame>でラップ（★<p:xfrm>を使用、<a:xfrm>ではない）
  return `<p:graphicFrame>
    <p:nvGraphicFramePr>...</p:nvGraphicFramePr>
    <p:xfrm>
      <a:off x="${emu(x)}" y="${emu(y)}"/>
      <a:ext cx="${emu(totalW)}" cy="${emu(totalH)}"/>
    </p:xfrm>
    <a:graphic><a:graphicData uri="...table">
      <a:tbl>
        <a:tblPr firstRow="1" bandRow="1"><a:noFill/></a:tblPr>
        <a:tblGrid>${gridCols}</a:tblGrid>
        ${headerRow}${dataRows}
      </a:tbl>
    </a:graphicData></a:graphic>
  </p:graphicFrame>`;
}
```

**テーブルセルの必須構造**:
```xml
<a:tc>
  <a:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:pPr algn="l"/><a:r><a:rPr .../><a:t>テキスト</a:t></a:r></a:p>
  </a:txBody>
  <a:tcPr marL="68580" marR="68580" marT="34290" marB="34290" anchor="ctr">
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln>
  </a:tcPr>
</a:tc>
```

### コールアウトボックス

```javascript
function makeCalloutBox(id, x, y, w, h, text, opts = {}) {
  const bgColor = opts.bgColor || 'D3E2FF';  // 薄い背景
  const textColor = opts.textColor || '{PRIMARY_COLOR}';
  // 角丸四角形 roundRect adj=5000
  return `<p:sp>...<a:prstGeom prst="roundRect">
    <a:avLst><a:gd name="adj" fmla="val 5000"/></a:avLst>
  </a:prstGeom>...</p:sp>`;
}
```

---

## OOXML正規表現の致命的バグパターン（★厳守）

### 使用禁止パターン

```javascript
// ❌ 絶対禁止: 最初の<p:sp>からターゲットまで全シェイプを巻き込む
xml = xml.replace(/<p:sp>[\s\S]*?name="TargetName"[\s\S]*?<\/p:sp>/, '');
```

### 正しいパターン

```javascript
// ✅ callback-based replace
xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
  if (content.includes('name="TargetName"')) return ''; // 削除
  return match;
});
```

この問題は `<p:sp>`, `<p:cxnSp>`, `<p:graphicFrame>`, `<w:p>` 等、全てのOOXML要素に適用される。

---

## EMU座標計算Tips

### 基本変換

```javascript
const emu = (inches) => Math.round(inches * 914400);
// 1 inch = 914400 EMU
// 1 cm ≈ 360000 EMU
```

### 標準スライドサイズ

| 項目 | EMU | インチ |
|------|-----|-------|
| 幅（16:9） | 9144000 | 10.0 |
| 高さ（16:9） | 5143500 | 5.625 |
| 幅（4:3） | 9144000 | 10.0 |
| 高さ（4:3） | 6858000 | 7.5 |

### レーン均等化の計算

```javascript
// スライド高さからヘッダー高を引いて等分
const laneHeight = (slideHeight - headerHeight) / laneCount;
// シェイプを中央配置
const shapeY = laneY + laneHeight / 2 - shapeHeight / 2;
```

---

## コネクタ座標移動パターン（赤線等）

### 基本構造

```xml
<p:cxnSp>
  <p:spPr>
    <a:xfrm>
      <a:off x="7195072" y="1541512"/>   <!-- ← x座標を変更 -->
      <a:ext cx="0" cy="3179700"/>        <!-- 垂直線: cx=0 -->
    </a:xfrm>
    <a:ln w="19050">
      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
    </a:ln>
  </p:spPr>
</p:cxnSp>
```

### 移動計算

```javascript
// 週幅が一定の場合: 次週 = 現在のx + 週幅(EMU)
const weekWidth = 320732; // EMU（プロジェクトごとに実測）
const newX = currentX + weekWidth;
```

---

## 精度限界の明示（★重要）

XML直接編集で構造的な変更（テキスト置換、スタイル変更、行追加）は**概ね正確**に行える。ただし以下の点で**完璧にはできない**ため、**最終的に人間がファイルを開いて目視確認・微修正する前提**で作業すること:

- レイアウト微調整（座標のピクセル単位のズレ、余白、文字切れ）
- フォントの再適用（XMLで指定しても反映されないケースあり）
- コネクタの接続点ずれ（座標計算だけでは完璧に合わない）
- マージセルの見た目崩れ
- 条件付き書式・データバリデーション等の複合機能への影響

---

## adm-zip vs JSZip の使い分け（★PJT-TD教訓）

### PPTX書き込みにはJSZipを使う

adm-zipで書き込んだPPTXは**adm-zipで再読み込みできない**（`No descriptor present` エラー）。

| ファイル形式 | 読み取り | 書き込み | 理由 |
|-------------|---------|---------|------|
| Excel (.xlsx) | adm-zip ✅ | adm-zip ✅ | Excelはadm-zipで問題なし |
| PPTX (.pptx) | adm-zip ✅ | **JSZip ✅** | adm-zip書込後に再読込不可 |
| Word (.docx) | JSZip ✅ | JSZip ✅ | 一貫してJSZip |

**PPTX書き込みパイプライン（JSZip使用）**:

```javascript
import JSZip from 'jszip';
import fs from 'fs';

// 1. 読み取り
const data = fs.readFileSync('input.pptx');
const zip = await JSZip.loadAsync(data);
const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');

// 2. XML編集
const updatedXml = slideXml.replace(/x="(\d+)"/, (match, x) => {
  return 'x="' + newX + '"';
});

// 3. 書き込み
zip.file('ppt/slides/slide1.xml', updatedXml);
const buf = await zip.generateAsync({ type: 'nodebuffer' });
fs.writeFileSync('output.pptx', buf);
```

---

## ガントバーの2シェイプ構成（★PJT-TD教訓）

ガントチャートのバーは「**色付き矩形（バー本体）＋テキストラベル**」の2シェイプで構成されている。

### 失敗パターン

テキストラベルだけ移動して、バー矩形は元の位置のまま → ラベルだけがずれて壊れる。

### 正しいパターン

y座標のレンジ（バンド）でフィルタし、その帯にある**全シェイプ**を一括移動する:

```javascript
// y座標がこの範囲にあるシェイプを全て移動
const BAND_TOP = targetY - tolerance;
const BAND_BOTTOM = targetY + tolerance;

xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
  const yMatch = content.match(/<a:off[^>]+y="(\d+)"/);
  if (yMatch) {
    const y = parseInt(yMatch[1]);
    if (y >= BAND_TOP && y <= BAND_BOTTOM) {
      // このバンド内のシェイプ → x座標を更新
      return match.replace(/<a:off([^>]+)x="(\d+)"/, (m, attrs, oldX) => {
        return '<a:off' + attrs + 'x="' + newX + '"';
      });
    }
  }
  return match;
});
```

**教訓**: ガントチャート上のシェイプを移動する場合は、名前やタイプではなく**座標レンジで選択**し、同じ行にある全要素を一括操作する。

---

## 番号付きリスト（アジェンダ等）の段落再構築

### `<a:t>` 逐次置換がNGな理由

テンプレートのアジェンダスライドでは、1つの段落 `<a:p>` 内に複数の `<a:t>` が存在することがある（例: ボールド部分と通常部分で分割）。`<a:t>` を1つずつ置換すると:

1. 項目数がテンプレートの `<a:t>` 数と一致しない場合に破綻する
2. 同一 `<a:p>` 内の複数 `<a:t>` を跨いで意図しない置換が発生する
3. 番号リストの書式（`<a:buAutoNum>` 等）が消える場合がある

### 正しいパターン: `<p:txBody>` ごと再構築

bodyプレースホルダの `<p:txBody>` を丸ごと差し替える:

```javascript
// 1. bodyプレースホルダを特定
xml = xml.replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (match, content) => {
  if (!content.includes('idx="1"') || !content.includes('<p:ph')) return match;

  // 2. 段落を全て新規構築
  const paragraphs = items.map(item =>
    `<a:p>
       <a:pPr marL="342900" indent="-342900">
         <a:buAutoNum type="arabicPeriod"/>
       </a:pPr>
       <a:r><a:rPr lang="ja-JP" sz="1400" dirty="0">
         <a:latin typeface="{FONT_FACE}"/><a:ea typeface="{FONT_FACE}"/>
       </a:rPr><a:t>${escXml(item)}</a:t></a:r>
     </a:p>`
  ).join('');

  // 3. <p:txBody>を丸ごと差し替え
  return match.replace(
    /<p:txBody>[\s\S]*?<\/p:txBody>/,
    `<p:txBody>
       <a:bodyPr wrap="square" anchor="t"><a:noAutofit/></a:bodyPr>
       <a:lstStyle/>
       ${paragraphs}
     </p:txBody>`
  );
});
```

### `<a:buAutoNum>` の主なtype値

| type | 表示例 |
|------|-------|
| `arabicPeriod` | 1. 2. 3. |
| `arabicParenR` | 1) 2) 3) |
| `alphaLcPeriod` | a. b. c. |
| `romanUcPeriod` | I. II. III. |

---

## コールアウトボックスの色バリエーション

### 用途別カラーパレット

| 用途 | 背景色 | テキスト色 | 使用場面 |
|------|--------|-----------|---------|
| 結論・まとめ | `D3E2FF`（薄青） | `{PRIMARY_COLOR}` | エグゼクティブサマリ、各セクションのまとめ |
| 期待効果・注意 | `FFF3E0`（薄橙） | `ED7D31` | 効果見込み、リスク注意喚起 |

### 実装

```javascript
// 青系（デフォルト）
const CALLOUT_BG_BLUE = 'D3E2FF';
const CALLOUT_TEXT_BLUE = '{PRIMARY_COLOR}'; // 例: 0055FF

// 橙系
const CALLOUT_BG_ORANGE = 'FFF3E0';
const CALLOUT_TEXT_ORANGE = 'ED7D31';

// 使用例
makeCalloutBox(id, x, y, w, h, text, {
  bgColor: CALLOUT_BG_ORANGE,
  textColor: CALLOUT_TEXT_ORANGE
});
```

### 角丸パラメータ

`adj=5000` は角丸の半径を制御する値（OOXMLの1/50000単位）。値が大きいほど角丸が大きくなる。

```xml
<a:prstGeom prst="roundRect">
  <a:avLst><a:gd name="adj" fmla="val 5000"/></a:avLst>
</a:prstGeom>
```

- `adj=0`: 角丸なし（矩形と同じ）
- `adj=5000`: 軽い角丸（推奨値）
- `adj=16667`: デフォルト値（角丸が大きすぎる場合あり）

---

## 関連ファイル

- `guides/03-pptx-generation.md` — pptxgenjsでの新規生成ガイド
- `guides/08-final-report.md` — 最終報告書の作り方（XML直接編集の実践例）
- `guides/14-proposal-workflow.md` — 提案書の作り方（XML直接編集の実践例）
- `guides/02-excel-editing.md` — Excel XML直接編集（同じadm-zipパイプライン）
- `scripts/build-final-report-example.mjs` — 最終報告書PPTX生成スクリプト（実装例）
- `scripts/build-proposal-example.mjs` — 提案書PPTX生成スクリプト（実装例）
