# Excel編集ガイド（adm-zip XML直接編集）

> PJT-TD/YPで検証済みのExcel書き込みパイプライン。条件付き書式・フィルター・数式を破壊しない唯一の方法。

---

## ツール選定（厳守）

| 操作 | 推奨ツール | 禁止ツール | 理由 |
|------|-----------|-----------|------|
| **書き込み** | adm-zip XML直接編集 | ExcelJS, xlsx(SheetJS) | 条件付き書式・オートフィルター・数式を破壊する |
| **読み取り** | exceljs（読み取り専用）or JSZip | — | 読み取り専用なら安全 |
| **書き込み後の検証** | JSZip | adm-zip | adm-zipは自分の出力を再読み込みすると `No descriptor present` エラー |

---

## 基本パイプライン

### 1. adm-zipでXMLを読み取り

```javascript
import AdmZip from 'adm-zip';
const zip = new AdmZip('input.xlsx');

// 共有文字列テーブル
const ssXml = zip.readAsText('xl/sharedStrings.xml');
// ワークシート（sheet1 = sheetN.xml）
const sheetXml = zip.readAsText('xl/worksheets/sheet1.xml');
```

### 2. sharedStrings.xmlに文字列を追加

```javascript
// 既存の文字列数を取得（count属性とuniqueCount属性）
// 新しい<si><t>テキスト</t></si>要素を</sst>の前に挿入
// count, uniqueCountをインクリメント

const newIndex = currentUniqueCount; // 0始まりのインデックス
const newEntry = `<si><t>新しいテキスト</t></si>`;
updatedSsXml = ssXml.replace('</sst>', newEntry + '</sst>');
// count/uniqueCount属性を更新
```

### 3. worksheetのセルを編集

```javascript
// セル参照: r="A5" 等
// 文字列型セル: t="s" で sharedStringsのインデックスを<v>に指定
// 数値型セル: t属性なし、<v>に数値を直接記載
// 日付型セル: <v>に日付シリアル値（例: 46096.0）

// 行の追加: 既存行をコピーして値だけ変更するのが安全
// スタイルID（s属性）は既存行からそのままコピー
```

### 4. writeZipで保存

```javascript
zip.updateFile('xl/sharedStrings.xml', Buffer.from(updatedSsXml));
zip.updateFile('xl/worksheets/sheet1.xml', Buffer.from(updatedSheetXml));
zip.writeZip('output.xlsx');
```

### 5. JSZipで検証

```javascript
import JSZip from 'jszip';
import fs from 'fs';

const data = fs.readFileSync('output.xlsx');
const verifyZip = await JSZip.loadAsync(data);
const verifySheet = await verifyZip.file('xl/worksheets/sheet1.xml').async('string');
// 内容を確認
```

---

## 注意事項・教訓

### 共有数式（shared formula）
- 行を追加した場合、shared formulaの `ref` 属性の範囲を新行まで拡張すること
- 例: `ref="A49:A61"` → `ref="A49:A71"`（10行追加した場合）
- 忘れると新行の数式が動作しない

### ht値・日付値のサフィックス
- ht（行高さ）値には `.0` サフィックスを付ける（例: `ht="15.75"` → 既存と同じ形式で）
- 日付シリアル値にも `.0` を付ける（例: `46096.0`）
- 既存行と形式を完全に合わせることが重要

### 行高さの自動計算
```
max(15.75, 推定行数 × 18pt)
```
- 日本語文字 = 幅2（半角換算）
- 英数字 = 幅1
- 対象列幅（例: 62.63）→ 1行あたり約69半角文字

### 列・行の非表示は禁止
- `hidden="1"` を使わない → オートフィルターで絞り込む
- 非表示にすると表示復帰操作がユーザーにとって煩雑

### adm-zip → JSZip のリレー
- adm-zipで書き込んだファイルをadm-zipで再読みすると `No descriptor present` エラー
- 修正が必要な場合: JSZipで読み取り → 修正 → adm-zipで再書き込み

### フォーマット厳守
- **フォント、罫線、インデント、スタイルは既存エントリと完全一致させる**
- XML直接編集なら既存要素をコピーして値だけ変更するアプローチが安全
- スタイルID（s属性）を間違えると見た目が崩れる

---

## スタイルID早見表（参考: PJT-YPアジェンダシート）

実際のスタイルIDはプロジェクトごとに異なるため、既存行から読み取ること。

```
s=4  → 標準テキスト（A/C/H列等）
s=6  → 中央揃え（C/D列等）
s=7  → 左揃え（E/F列等）
s=8  → 折り返しテキスト（F列等）
s=9  → ハイパーリンク（G列等）
s=12 → 日付形式（B列等）
```

---

## 関連ファイル

- `scripts/edit-excel-example.mjs` — 最小動作サンプル
- `guides/06-meeting-workflow.md` — アジェンダシート更新の実運用フロー
