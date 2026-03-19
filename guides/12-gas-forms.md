# 12. Google Forms + GAS操作ガイド

> PJT-TD（東京ドーム案件）で確立。Google Formsをプログラムで操作する際の教訓集。

---

## 1. 前提

- Google Formsは **直接編集不可**。GAS（Google Apps Script）スクリプト経由のみ
- フォームがFC側（外部組織）のGoogle Drive上にある場合、権限付与が必要
- 操作フロー: **項目テキスト出力スクリプト → 構成確認 → 項目追加/削除スクリプト**

---

## 2. GAS API制約

### var統一（const/let禁止）

GASのRhino環境ではES6構文が使えない場合がある。安全のため **全スクリプトvar統一**。

```javascript
// ❌ エラーになる可能性
const form = FormApp.openById(FORM_ID);
let items = form.getItems();

// ✅ 安全
var form = FormApp.openById(FORM_ID);
var items = form.getItems();
```

### addFileUploadItem() は存在しない

GAS APIにファイルアップロード項目の作成メソッドは存在しない。

| 操作 | 対応方法 |
|------|---------|
| FILE_UPLOADの新規作成 | **不可** → スクリプト実行後に手動追加（ログに手順出力） |
| FILE_UPLOADの移動 | **可能** → `form.moveItem(item, toIndex)` |
| FILE_UPLOADの削除 | **可能** → `form.deleteItem(index)` |

### 利用可能なメソッド（確認済み）

| メソッド | 用途 |
|---------|------|
| `form.addTextItem()` | テキスト入力項目 |
| `form.addMultipleChoiceItem()` | 選択肢項目 |
| `form.addSectionHeaderItem()` | セクションヘッダー |
| `form.addPageBreakItem()` | ページ区切り |
| `form.moveItem(item, toIndex)` | 項目移動（途中挿入） |
| `form.deleteItem(index)` | 項目削除 |
| `item.setHelpText(text)` | 説明文の設定（PAGE_BREAK, SECTION_HEADER等で使用可能） |

---

## 3. 冪等操作パターン

スクリプトは **何度実行しても同じ結果** になるよう設計する。

### 追加の冪等チェック

```javascript
// 追加対象が既にあればスキップ
var items = form.getItems();
for (var i = 0; i < items.length; i++) {
  if (items[i].getTitle().indexOf('ターゲットタイトル') !== -1) {
    Logger.log('Already exists. Skip.');
    return;
  }
}
// 存在しなければ追加処理へ
```

### 削除の冪等チェック

```javascript
// 削除対象が無ければスキップ
var targetIndex = -1;
var items = form.getItems();
for (var i = 0; i < items.length; i++) {
  if (items[i].getTitle() === '削除対象のタイトル') {
    targetIndex = i;
    break;
  }
}
if (targetIndex === -1) {
  Logger.log('Target not found. Skip deletion.');
  return;
}
```

---

## 4. 項目追加・削除の手順

### 削除は高インデックスから（★重要）

複数項目を削除する場合、低インデックスから削除するとインデックスがずれる。

```javascript
// ❌ インデックスがずれる
form.deleteItem(5);  // 削除後、旧index6が5になる
form.deleteItem(6);  // 意図しない項目を削除

// ✅ 高インデックスから削除
form.deleteItem(6);  // 先にこちらを削除
form.deleteItem(5);  // インデックスがずれない
```

### 操作後は必ず getItems() で再取得（★重要）

```javascript
// 削除実行
form.deleteItem(targetIndex);

// ❌ 古い参照を使い続ける
// var item = items[nextIndex]; // ← ずれている可能性

// ✅ 必ず再取得
items = form.getItems();
```

---

## 5. moveItemの活用

### FILE_UPLOADの移動

addFileUploadItem() は不可だが、既存のFILE_UPLOAD項目を別の位置に移動することは可能。

```javascript
// 身分証明書（FILE_UPLOAD）を事業者情報セクションに移動
var items = form.getItems();
var fileUploadItem = null;
var targetIndex = -1;

for (var i = 0; i < items.length; i++) {
  if (items[i].getTitle().indexOf('身分証明書') !== -1) {
    fileUploadItem = items[i];
  }
  if (items[i].getTitle().indexOf('事業者の区分') !== -1) {
    targetIndex = items[i].getIndex() + 1; // 区分の直後
  }
}

if (fileUploadItem && targetIndex !== -1) {
  form.moveItem(fileUploadItem, targetIndex);
  Logger.log('Moved FILE_UPLOAD to index ' + targetIndex);
}

// ★ 操作後は必ず再取得
items = form.getItems();
```

### 途中挿入

新規項目を追加した後、`moveItem` で目的の位置に移動する。

```javascript
// 項目追加（末尾に追加される）
var newItem = form.addMultipleChoiceItem()
  .setTitle('居住地域')
  .setChoiceValues(['国内', '国外'])
  .setRequired(true);

// 目的の位置に移動
var emailIndex = -1;
var items = form.getItems();
for (var i = 0; i < items.length; i++) {
  if (items[i].getTitle().indexOf('メールアドレス') !== -1) {
    emailIndex = items[i].getIndex();
    break;
  }
}
if (emailIndex !== -1) {
  form.moveItem(newItem, emailIndex + 1);
}
```

---

## 6. PageBreakとナビゲーション

### ナビゲーション参照中のPageBreakは直接削除不可

PageBreakが他の項目のナビゲーション先として参照されている場合、削除するとエラーになる。

```javascript
// ❌ ナビゲーション参照がある場合エラー
form.deleteItem(pageBreakIndex);

// ✅ 先にナビゲーションをCONTINUEにリセット
var bizTypeItem = items[bizTypeIndex].asMultipleChoiceItem();
var choices = bizTypeItem.getChoices();
var newChoices = [];
for (var c = 0; c < choices.length; c++) {
  newChoices.push(bizTypeItem.createChoice(
    choices[c].getValue(),
    FormApp.PageNavigationType.CONTINUE
  ));
}
bizTypeItem.setChoices(newChoices);

// ナビゲーション解除後に削除
form.deleteItem(pageBreakIndex);

// 再取得してからナビゲーションを再設定
items = form.getItems();
// 法人 → 法人情報ページ, 個人事業主 → SUBMIT 等
```

### ナビゲーション設定パターン

```javascript
var bizTypeItem = items[bizTypeIndex].asMultipleChoiceItem();
var houjinPB = items[houjinPBIndex].asPageBreakItem();

bizTypeItem.setChoices([
  bizTypeItem.createChoice('法人', houjinPB),
  bizTypeItem.createChoice('個人事業主', FormApp.PageNavigationType.SUBMIT)
]);
```

**注意**: SECTION_HEADERはナビゲーション参照されないので直接削除可能。

---

## 7. 典型的な実装パターン

### フォーム全項目スキャン＆ログ

```javascript
function scanForm() {
  var form = FormApp.openById(FORM_ID);
  var items = form.getItems();

  Logger.log('=== Form Items (' + items.length + ') ===');
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    Logger.log('[' + i + '] ' + item.getType() + ': ' + item.getTitle());
  }
}
```

### 手動追加手順のログ出力

GAS APIで作成不可な項目（FILE_UPLOAD等）は、スクリプト末尾に手動追加手順を出力:

```javascript
Logger.log('');
Logger.log('★★★ 手動追加が必要です ★★★');
Logger.log('→ フォームエディタで以下の手順で追加してください:');
Logger.log('  1. 「インボイス番号」の後にファイルアップロード項目を追加');
Logger.log('  2. タイトル: 身分証明書（運転免許証、パスポート、健康保険証等）');
Logger.log('  3. 説明: ご本人様確認のため、身分証明書の画像をアップロードしてください。');
Logger.log('  4. 必須をONに設定');
```

---

## 過去の失敗と教訓

| 失敗 | 原因 | 対策 |
|------|------|------|
| 身分証明書を誤削除 | 「マイナンバー」検索が「マイナンバーカード」にもヒット | **完全一致タイトルで削除対象を特定** |
| PageBreak削除時にエラー | ナビゲーション参照中 | **先にCONTINUEリセット→削除→ナビ再設定** |
| addFileUploadItem()不在 | GAS APIにメソッドなし | **moveItemで移動、または手動追加** |
| const/letでエラー | GAS Rhino環境 | **全スクリプトvar統一** |
| moveItem後にインデックスずれ | 古い参照を使用 | **操作後は必ずgetItems()再取得** |
