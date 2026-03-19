# Google Driveマウント手順

> プロジェクトフォルダからGoogle Drive上の共有フォルダに直接アクセスするための設定。

---

## 前提

- Google Drive Desktop がインストール済み
- 共有ドライブまたは共有フォルダへのアクセス権がある

---

## 方法1: junction（推奨）

管理者権限のコマンドプロンプトで実行:

```cmd
mklink /J "c:\PJT-{CODE}\drive" "H:\.shortcut-targets-by-id\{FOLDER_ID}\{FOLDER_NAME}"
```

### Google Driveのパスの確認方法

1. Google Drive Desktopでフォルダを右クリック → 「パスをコピー」
2. または、エクスプローラでGoogle Driveのドライブレターを確認（通常 G: または H:）
3. 共有フォルダの場合、`.shortcut-targets-by-id/{ID}/{フォルダ名}` のパスになる

### 確認

```cmd
dir "c:\PJT-{CODE}\drive"
```

---

## 方法2: マイドライブにショートカットを追加（権限不足時の代替）

junction作成に管理者権限がない場合:

1. Google Driveのブラウザ版で対象フォルダを開く
2. フォルダを右クリック → 「マイドライブにショートカットを追加」
3. マイドライブ配下にショートカットが作成される
4. Google Drive Desktopで同期 → ローカルからアクセス可能に

```cmd
mklink /J "c:\PJT-{CODE}\drive" "H:\マイドライブ\{ショートカット名}"
```

---

## 注意事項

### git checkout によるクラウドファイル巻き戻り

**重大リスク**: `git checkout -- drive/` や `git restore drive/` を実行すると、Google Drive上のファイルも巻き戻る。

- drive/ はjunction（シンボリックリンク）であり、ローカルファイルのように見えるがクラウドの実体を指している
- gitの復元コマンドはjunction先のファイルを上書きする
- これによりGoogle Drive上の最新版が失われる

**対策**: drive/ は `.gitignore` に追加すること。

### .gitignore の設定

```
drive/
*.xlsx
*.docx
*.pptx
```

### Google Drive Desktop の自動同期

- drive/ 配下のファイルを編集すると自動的にクラウドに同期される
- 大きなファイルの書き込み中に同期が走ると競合する可能性がある
- 対策: 一時ファイルに書き込み → 完成後にdrive/にコピー

---

## 標準フォルダ構成

drive/ 配下の標準ディレクトリ体系:

```
drive/
├── 00.提案資料/
├── 10.納品成果物/
├── 20.契約書関連/
├── 30.全体設計/
├── 40.WORK/
├── 50.打ち合わせ関連/
├── 60.受領資料/
└── 99.参考資料/
```

---

## 関連ファイル

- `setup.md` — 新規PJセットアップ手順（junction作成含む）
