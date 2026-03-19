# 13. Google Spreadsheet → JSON Export（GAS経由）

> PJT-TD（東京ドーム案件）で確立。会社Google Workspaceのセキュリティ制限によりMCP直接接続が不可な場合の代替手段。

---

## 1. フロー概要

```
Google Spreadsheet（外部組織のDrive上）
  ↓ GASスクリプトで読み取り
JSON ファイル生成
  ↓ Google Driveの指定フォルダに出力
Drive Desktop 同期
  ↓ junction経由でローカルから参照
c:\PJT-{CODE}\drive\99.data-export\spreadsheet1.json
```

**メリット**:
- セキュリティ制限でMCP（Google Drive MCP等）が使えない環境でも利用可能
- スプシの構造をそのまま保持（ヘッダー重複・空文字でもデータ消滅しない）
- ローカルファイルとして高速に読み取り可能

---

## 2. GASスクリプトテンプレート

```javascript
// gas-spreadsheet-exporter.js
// GASエディタにコピペして使用

var OUTPUT_FOLDER_ID = '{YOUR_DRIVE_FOLDER_ID}'; // 出力先フォルダのID

var SPREADSHEETS = [
  { id: '{SPREADSHEET_1_ID}', name: 'spreadsheet1' },
  { id: '{SPREADSHEET_2_ID}', name: 'spreadsheet2' }
];

function exportAllData() {
  var folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);

  for (var s = 0; s < SPREADSHEETS.length; s++) {
    var config = SPREADSHEETS[s];
    var ss = SpreadsheetApp.openById(config.id);
    var result = { exportedAt: new Date().toISOString(), sheets: [] };

    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetData = {
        name: sheet.getName(),
        rows: []
      };

      var range = sheet.getDataRange();
      var values = range.getValues();

      for (var r = 0; r < values.length; r++) {
        sheetData.rows.push(values[r]);
      }

      result.sheets.push(sheetData);
    }

    var json = JSON.stringify(result, null, 2);
    var fileName = config.name + '.json';

    // 既存ファイルがあれば削除してから作成（上書き）
    var existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    folder.createFile(fileName, json, MimeType.PLAIN_TEXT);
    Logger.log('Exported: ' + fileName + ' (' + result.sheets.length + ' sheets)');
  }

  Logger.log('=== Export complete ===');
}
```

---

## 3. データ形式

### 生2D配列（rowsプロパティ）

```json
{
  "exportedAt": "2026-03-18T10:00:00.000Z",
  "sheets": [
    {
      "name": "シート1",
      "rows": [
        ["ヘッダー1", "ヘッダー2", "ヘッダー3"],
        ["データ1", "データ2", "データ3"],
        ["データ4", "", "データ6"]
      ]
    }
  ]
}
```

**重要**: ヘッダー行とデータ行の区別はない（全て `rows` 配列の要素）。ヘッダーが重複していてもデータが空文字でも、そのまま保持される。

### ローカルでの読み取り

```javascript
const fs = require('fs');
const data = JSON.parse(fs.readFileSync('drive/99.data-export/spreadsheet1.json', 'utf-8'));

// シート名で検索
const targetSheet = data.sheets.find(s => s.name === 'ターゲットシート名');

// ヘッダー行（rows[0]）をキーにしてオブジェクト配列に変換
const headers = targetSheet.rows[0];
const records = targetSheet.rows.slice(1).map(row => {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
});
```

---

## 4. 運用上の注意点

### 同期タイミング

- GASスクリプト実行 → Drive出力は即時
- Drive Desktop同期 → ローカル反映は **数秒〜数分のラグ** がある
- 急ぎの場合は Drive Desktop の「今すぐ同期」を手動実行

### ファイル名規則

- `{スプシ識別名}.json` で統一（例: `spreadsheet1.json`, `営業リスト.json`）
- 日付付きにしない（毎回上書き方式のため）

### GASの実行制限

- GASの実行時間制限: **6分**（無料アカウント）/ **30分**（Workspace）
- 大量のシート・セルがある場合はスプシを分割するか、対象シートを絞る

### 出力先フォルダ

- プロジェクトのDriveフォルダ内に `99.data-export/` を作成して使用
- 出力先フォルダIDは `OUTPUT_FOLDER_ID` で指定

### var統一

GASスクリプトなので、**const/let禁止、var統一**（12-gas-forms.md 参照）。
