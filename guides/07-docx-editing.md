# Word編集ガイド（JSZip XML直接編集）

> docxファイル（議事メモ等）をプログラムで編集する際の手順。レイアウト・スタイルを保持する。

---

## 基本パイプライン

```javascript
import JSZip from 'jszip';
import fs from 'fs';

// 1. 読み取り
const data = fs.readFileSync('input.docx');
const zip = await JSZip.loadAsync(data);
const docXml = await zip.file('word/document.xml').async('string');

// 2. XMLを解析して段落追加・編集
// ...

// 3. 保存
zip.file('word/document.xml', updatedXml);
const buf = await zip.generateAsync({ type: 'nodebuffer' });
fs.writeFileSync('output.docx', buf);
```

---

## 段落の操作

### 既存段落のスタイルを保持して内容だけ変更

```xml
<!-- 既存段落の構造 -->
<w:p>
  <w:pPr>
    <!-- 段落スタイル（フォント、インデント、行間等）-->
    <w:pStyle w:val="ListParagraph"/>
    <w:numPr>
      <w:ilvl w:val="0"/>
      <w:numId w:val="1"/>
    </w:numPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <!-- ランスタイル（フォントサイズ、太字等）-->
    </w:rPr>
    <w:t>テキスト内容</w:t>
  </w:r>
</w:p>
```

**鉄則**: `w:pPr` と `w:rPr` をそのままコピーし、`w:t` の中身だけ変更する。

### 新しい段落の挿入

1. 既存の類似段落をテンプレートとしてコピー
2. テキスト内容だけ差し替え
3. 挿入位置を特定して `</w:body>` の前に追加

---

## 注意事項

### numbering.xml の破損

- 番号付きリスト（`w:numPr`）を含む段落を操作する場合、`word/numbering.xml` の整合性に注意
- 新しいリストを追加する場合は既存のnumIdを再利用するのが安全

### フォーマット厳守原則

> 基本原則: CLAUDE.md「フォーマット厳守原則」参照。

- 出席者セクション等、従来書いていないセクションを勝手に追加しない
- タイトル形式: `YYMMDD_{会議名}`（余計な文言は付けない）

### 改ページの挿入

会議エントリの区切り等に改ページを入れる場合:

```xml
<w:p><w:r><w:br w:type="page"/></w:r></w:p>
```

見出し（Heading1等）の直前に挿入する。

### ハイパーリンクの追加

1. `word/_rels/document.xml.rels` にRelationshipを追加:

```xml
<Relationship Id="rIdXX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://..." TargetMode="External"/>
```

2. 本文XMLでハイパーリンクを参照:

```xml
<w:hyperlink r:id="rIdXX">
  <w:r>
    <w:rPr>
      <w:color w:val="0000FF"/>
      <w:u w:val="single"/>
    </w:rPr>
    <w:t>リンクテキスト</w:t>
  </w:r>
</w:hyperlink>
```

### 色の明示指定

`<w:color w:val="000000"/>` を付けないと、スタイル定義のデフォルト色に依存する。意図しない色（灰色等）になる場合があるため、黒色テキストには明示的に `w:val="000000"` を指定すること。

### フリガナ（rPh）の処理

docx内のセルやテキストにフリガナ情報（`<w:rPh>` 要素）が含まれている場合がある。

- **テキスト抽出時**: `<w:rPh>...</w:rPh>` を除去してからテキストを取得すること（フリガナがテキストに混入する）
- **新規段落作成時**: rPh要素は不要（付けなくてよい）

### 先頭挿入の注意

議事メモ等の時系列ドキュメントでは、新しいエントリは先頭に挿入（最新が上、最古が下）。`<w:body>`直後の最初の段落の前に新しいXMLを挿入する。

### 箇条書きレベルの構成例

議事メモ等で使用する箇条書きレベルの例:

```
ilvl=0: ● 大見出し（numId参照）
ilvl=1: ○ 中見出し
ilvl=2: ■ 詳細
ilvl=3: ● サブ詳細
```

`w:numPr` の `w:ilvl` と `w:numId` でレベルとリストIDを指定。新しいリストを追加する場合は既存のnumIdを再利用するのが安全。

### docx内のXMLファイル構成

```
word/
├── document.xml        # 本文（メインの編集対象）
├── styles.xml          # スタイル定義
├── numbering.xml       # 番号付きリスト定義
├── settings.xml        # 文書設定
├── fontTable.xml       # フォントテーブル
└── _rels/
    └── document.xml.rels  # リレーション
```

---

## 議事メモのdocx構成（参考: PJT-YP）

```
YYMMDD_{会議名}定例
├── 決定事項
│   └─ 箇条書き
├── ネクスト
│   └─ 箇条書き
└── 議事メモ
    └─ 箇条書き（議論の要点サマリ）
```

- 出席者セクションは書かない
- 各回の区切りは見出しスタイルで

---

## 関連ファイル

- `guides/06-meeting-workflow.md` — 定例後ワークフロー（docx更新の運用フロー）
- `guides/02-excel-editing.md` — Excel編集（adm-zipパイプライン）
