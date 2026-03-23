# PJT-format — プロジェクト運営テンプレート

Claude Codeを活用したコンサルティングプロジェクトの運営テンプレート。
新規プロジェクト開始時にこのフォルダをまるごとコピーして使う。

PJT-YP（ヤマトプロテック DXプロジェクト）、PJT-TS（トヨタエムズシステム データ利活用プロジェクト）、PJT-SE（スポーツ&エンターテインメント案件）で蓄積した知見を汎用化したもの。

---

## このフォルダの使い方

1. フォルダごとコピーして新規プロジェクトを作る
2. プレースホルダーを実値に置換する
3. 必要なテンプレートをルートにコピーして運用開始

詳細は後述の「セットアップ手順」を参照。

---

## フォルダ構成

```
PJT-format/
│
├── README.md                ← 本ファイル（全体説明 + セットアップ手順）
├── CLAUDE.md                  Claude Code設定テンプレート（コアルール集）
├── MEMORY-TEMPLATE.md         Claude auto-memory の初期構造
├── .mcp.json.example          MCP接続設定のサンプル
├── package.json               npm依存定義
│
├── guides/                    知見ドキュメント（ナレッジ集）
│   ├── 01-kickoff.md            キックオフ資料の作り方
│   ├── 02-excel-editing.md      Excel編集（adm-zip XML直接編集）
│   ├── 03-pptx-generation.md    PowerPoint生成（pptxgenjs）
│   ├── 04-google-drive-mount.md Google Driveマウント手順
│   ├── 05-mcp-setup.md          MCP設定ガイド
│   ├── 06-meeting-workflow.md   定例運用フロー
│   ├── 07-docx-editing.md       Word編集（JSZip）
│   ├── 08-final-report.md      最終報告書の作り方
│   ├── 09-pptx-xml-editing.md  既存PPTX XML直接編集
│   ├── 10-backlog-operations.md Backlog運用ガイド
│   ├── 11-asana-operations.md   Asana運用ガイド
│   ├── 12-gas-forms.md          Google Forms + GAS操作ガイド
│   ├── 13-gas-spreadsheet-export.md Google Spreadsheet → JSON Export
│   └── 14-proposal-workflow.md  提案書の作り方
│
├── templates/                 コピーして使う雛形ファイル
│   ├── kickoff-format.md        キックオフMD（6セクション構造）
│   ├── PROJECT-REPORT.md        プロジェクトレポート
│   ├── SCOPE-SUMMARY.md         全体計画書
│   ├── DECISIONS.md             決定事項ログ
│   ├── FACTS.md                 確定事実管理
│   ├── MEETING-FOLLOWUP-HANDOFF.md 定例後ワークフロー
│   ├── research-plan.md        リサーチ計画＋情報パッケージ（Phase 0用）
│   ├── proposal-format.md      提案書テンプレート（7セクション構成）
│   ├── final-report-format.md  最終報告書テンプレート（5章構成）
│   ├── flux-presentation-template.pptx FLUX汎用PPTXテンプレート（5種レイアウト）
│   └── communication-style-example.md コミュニケーションスタイル例
│
├── scripts/                   再利用スクリプト
│   ├── convert-to-text.mjs      受領資料テキスト変換（xlsx/pptx/pdf → txt）
│   ├── create-schedule-pptx.mjs ガントチャートPPTX生成
│   ├── create-kickoff-notion.mjs キックオフNotionページ作成
│   ├── edit-excel-example.mjs   adm-zip Excel編集サンプル
│   ├── build-final-report-example.mjs 最終報告書PPTX生成（汎用ヘルパー）
│   └── build-proposal-example.mjs 提案書PPTX生成（汎用ヘルパー）
│
└── drive/                     Google Drive junction先（空フォルダ構成）
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

## 知見ガイド一覧

| # | ガイド | 概要 |
|---|--------|------|
| 01 | [キックオフ資料の作り方](guides/01-kickoff.md) | 受領資料 → キックオフMD → すり合わせ → PPTX の全フロー |
| 02 | [Excel編集](guides/02-excel-editing.md) | adm-zip XML直接編集パイプライン。ExcelJS/SheetJS書き込み禁止の理由と代替手法 |
| 03 | [PowerPoint生成](guides/03-pptx-generation.md) | pptxgenjsでガントチャート自動生成。赤線制御・テーマカラー |
| 04 | [Google Driveマウント](guides/04-google-drive-mount.md) | junction作成、マイドライブショートカット、git巻き戻りリスク |
| 05 | [MCP設定](guides/05-mcp-setup.md) | Windows環境の教訓、ツール別設定テンプレート |
| 06 | [定例運用フロー](guides/06-meeting-workflow.md) | 定例前後のワークフロー、アジェンダシート更新、議事メモ更新 |
| 07 | [Word編集](guides/07-docx-editing.md) | JSZip XML直接編集、スタイル保持のコツ |
| 08 | [最終報告書の作り方](guides/08-final-report.md) | MD作成→すり合わせ→PPTX生成の全フロー。5章構成の設計原則 |
| 09 | [PPTX XML直接編集](guides/09-pptx-xml-editing.md) | adm-zip/JSZipで既存PPTXを編集。rId衝突回避・OOXML正規表現バグ |
| 10 | [Backlog運用](guides/10-backlog-operations.md) | メンション・コメント書式・資料共有フロー・親子チケット運用 |
| 11 | [Asana運用](guides/11-asana-operations.md) | Backlog運用ガイドのAsana版。PJT-TDで確立 |
| 12 | [Google Forms + GAS操作](guides/12-gas-forms.md) | Google Formsをプログラムで操作する際の教訓集 |
| 13 | [Google Spreadsheet Export](guides/13-gas-spreadsheet-export.md) | GAS経由でSpreadsheet → JSON Export。MCP直接接続不可時の代替手段 |
| 14 | [提案書の作り方](guides/14-proposal-workflow.md) | MD作成→すり合わせ→PPTX生成の全フロー。7セクション構成の設計原則 |

---

## テンプレート一覧

| ファイル | 用途 | 運用タイミング |
|---------|------|--------------|
| [kickoff-format.md](templates/kickoff-format.md) | キックオフMD（6セクション） | PJ開始時。受領資料を流し込んですり合わせ→PPTX化 |
| [PROJECT-REPORT.md](templates/PROJECT-REPORT.md) | プロジェクトレポート | 随時更新。新規参画者・経営報告用 |
| [SCOPE-SUMMARY.md](templates/SCOPE-SUMMARY.md) | 全体計画書 | テーマのSingle Source of Truth |
| [DECISIONS.md](templates/DECISIONS.md) | 決定事項ログ | 定例ごとに追記 |
| [FACTS.md](templates/FACTS.md) | 確定事実管理 | ヒアリング・調査のたびに追記 |
| [MEETING-FOLLOWUP-HANDOFF.md](templates/MEETING-FOLLOWUP-HANDOFF.md) | 定例後ワークフロー | Claude Codeへの引き継ぎ手順書 |
| [research-plan.md](templates/research-plan.md) | リサーチ計画＋情報パッケージ（Phase 0用） | 提案前。情報収集→リサーチ→構造化 |
| [proposal-format.md](templates/proposal-format.md) | 提案書テンプレート（7セクション構成） | 提案時。MD→すり合わせ→PPTX化 |
| [final-report-format.md](templates/final-report-format.md) | 最終報告書テンプレート（5章構成） | PJ終了時。MD→すり合わせ→PPTX化 |
| [flux-presentation-template.pptx](templates/flux-presentation-template.pptx) | FLUX汎用PPTXテンプレート（表紙/アジェンダ/扉/テーブル/テキスト） | 最終報告書・キックオフ・提案書等、全PPTX資料のベース |
| [communication-style-example.md](templates/communication-style-example.md) | コミュニケーションスタイル例 | MEMORY.mdの別ファイルとして配置・カスタマイズ |

---

## CLAUDE.md のコアルール（主要なもの）

- **3ストライクルール**: 同じ操作で3回連続失敗 → 即停止、プランモードで根本原因分析
- **動的専門家ペルソナ**: レポート作成時、タスクに最適な2名の専門家でディスカッション
- **曖昧語禁止**: 「エラーが多発」→「15分間に3回発生」。具体的な数値・名称を必ず使う
- **Office操作ガイド**: Excel書き込みはadm-zip XML直接編集（ExcelJS/SheetJS禁止）
- **影響範囲チェックリスト**: 変更を加えたら参照/依存箇所を全て確認
- **ファイル修正・再生成禁止**: 明示的指示なしにファイルを変更しない
- **憶測回答禁止**: 確信がない情報は「確認が必要」と明示。根拠を示して回答

---

## セットアップ手順

### 1. フォルダコピー

```cmd
xcopy /E /I "c:\PJT-format" "c:\PJT-{CODE}"
```

`{CODE}` はプロジェクトの識別コード（例: YP, TS, NSD 等）。

### 2. Google Drive junction作成

管理者権限ありの場合:
```cmd
mklink /J "c:\PJT-{CODE}\drive" "H:\.shortcut-targets-by-id\{FOLDER_ID}\{FOLDER_NAME}"
```

権限不足の場合:
1. Google Driveブラウザ版で「マイドライブにショートカットを追加」
2. `mklink /J "c:\PJT-{CODE}\drive" "H:\マイドライブ\{ショートカット名}"`

詳細: [guides/04-google-drive-mount.md](guides/04-google-drive-mount.md)

### 3. MCP設定

```cmd
copy ".mcp.json.example" ".mcp.json"
```

`.mcp.json` のプレースホルダーをトークン実値に置換。
詳細: [guides/05-mcp-setup.md](guides/05-mcp-setup.md)

### 4. CLAUDE.mdのプレースホルダー置換

| プレースホルダー | 置換先 | 例 |
|-----------------|--------|-----|
| `{PROJECT_NAME}` | プロジェクト名 | ヤマトプロテック DXプロジェクト |
| `{CLIENT_NAME}` | 顧客名 | ヤマトプロテック株式会社 |
| `{PJ_ROOT}` | プロジェクトルートパス | c:\PJT-YP |
| `{DRIVE_PATH}` | Google Driveのパス | H:\...2512~｜ヤマトプロテック |

VSCodeの検索置換（Ctrl+Shift+H）で一括置換が便利。

### 5. MEMORY-TEMPLATE.md の配置

```cmd
mkdir "C:\Users\{USER}\.claude\projects\c--PJT-{CODE}\memory"
copy "MEMORY-TEMPLATE.md" "C:\Users\{USER}\.claude\projects\c--PJT-{CODE}\memory\MEMORY.md"
```

### 6. npm install

```cmd
npm install
```

### 7. テンプレートファイルの配置

templates/ から必要なファイルをルートにコピー:

```cmd
copy "templates\DECISIONS.md" "DECISIONS.md"
copy "templates\PROJECT-REPORT.md" "PROJECT-REPORT.md"
copy "templates\SCOPE-SUMMARY.md" "SCOPE-SUMMARY.md"
copy "templates\FACTS.md" "FACTS.md"
```

### 8. .gitignore の設定

```
drive/
.mcp.json
node_modules/
*.xlsx
*.docx
*.pptx
```

---

## セットアップ チェックリスト

- [ ] フォルダコピー完了
- [ ] Google Drive junction 作成・動作確認
- [ ] .mcp.json 設定・接続テスト
- [ ] CLAUDE.md プレースホルダー置換
- [ ] MEMORY-TEMPLATE.md → Claude memory配置
- [ ] npm install
- [ ] テンプレートファイル配置
- [ ] CLAUDE.md に用語集・MCP接続状況を記入
- [ ] MEMORY.md にプロジェクト概要・キーパーソンを記入
