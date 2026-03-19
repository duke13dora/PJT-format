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
- **フォント、罫線、インデント、スタイルは既存エントリと完全一致させる**
- 出席者セクション等、従来書いていないセクションを勝手に追加しない
- タイトル形式: `YYMMDD_{会議名}`（余計な文言は付けない）

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
