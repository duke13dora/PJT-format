# MCP設定ガイド

> Windows環境でのMCPサーバー設定手順とハマりどころ。PJT-YPで蓄積した教訓。

---

## 設定ファイル

`.mcp.json` をプロジェクトルートに配置する。テンプレートは `.mcp.json.example` を参照。

```bash
cp .mcp.json.example .mcp.json
# トークンを実値に置換
```

---

## Windows環境の教訓

### 1. `cmd /c npx` はタイムアウトする

```json
// ❌ 遅い・タイムアウトリスク
{ "command": "cmd", "args": ["/c", "npx", "-y", "@some/mcp-server"] }

// ✅ バイナリ直接指定が確実
{ "command": "C:\\path\\to\\mcp-server.exe", "args": ["--transport", "stdio"] }

// ✅ node直接指定も安定
{ "command": "node", "args": ["C:/path/to/mcp-server/build/index.js"] }
```

### 2. `claude mcp add` コマンドのバグ

`claude mcp add` は引数の `/c` を `C:/` に変換するバグがある。

**対策**: `.mcp.json` を手動編集する（`claude mcp add` は使わない）。

### 3. 設定変更前のテスト

設定変更前に必ず `node -e` でMCPプロトコル通信テストを行い、認証成功を確認してから再起動する:

```bash
node -e "
const { spawn } = require('child_process');
const p = spawn('node', ['path/to/server.js'], { env: { ...process.env, TOKEN: 'xxx' } });
p.stdout.on('data', d => console.log(d.toString()));
p.stderr.on('data', d => console.error(d.toString()));
setTimeout(() => p.kill(), 5000);
"
```

---

## ツール別設定テンプレート

### Backlog

```json
{
  "backlog": {
    "command": "node",
    "args": ["C:/Users/{USER}/AppData/Roaming/npm/node_modules/backlog-mcp-server/build/index.js"],
    "env": {
      "BACKLOG_API_KEY": "{YOUR_BACKLOG_API_KEY}",
      "BACKLOG_DOMAIN": "{SPACE}.backlog.com"
    }
  }
}
```

- **注意**: 環境変数は `BACKLOG_DOMAIN`（ホスト名のみ）。ドキュメントの `BACKLOG_SPACE_URL` は誤り
- インストール: `npm install -g backlog-mcp-server`

### Notion

```json
{
  "notion": {
    "command": "cmd",
    "args": ["/c", "npx", "-y", "@notionhq/notion-mcp-server"],
    "env": {
      "NOTION_TOKEN": "{YOUR_NOTION_TOKEN}"
    }
  }
}
```

- Notionトークン取得: Settings → Integrations → Internal Integration Token
- 組織のNotionはオーナー権限がないと制限あり → ローカルDL方式も検討

### Asana

```json
{
  "asana": {
    "command": "cmd",
    "args": ["/c", "npx", "-y", "@roychri/mcp-server-asana"],
    "env": {
      "ASANA_ACCESS_TOKEN": "{YOUR_ASANA_ACCESS_TOKEN}"
    }
  }
}
```

- パッケージ: `@roychri/mcp-server-asana`（npx経由で自動インストール）
- トークン取得: Asana → Settings → Apps → Personal access tokens
- **MCP制限あり**: セクション配置やメンション付きコメントはREST API直接呼出しが必要 → 詳細は `guides/11-asana-operations.md` 参照

### Slack

```json
{
  "slack": {
    "type": "stdio",
    "command": "C:\\path\\to\\slack-mcp-server.exe",
    "args": ["--transport", "stdio"],
    "env": {
      "SLACK_MCP_XOXC_TOKEN": "{YOUR_XOXC_TOKEN}",
      "SLACK_MCP_XOXD_TOKEN": "{YOUR_XOXD_TOKEN}"
    }
  }
}
```

- xoxc/xoxd トークンはセッションベースで短命 → 切れたら再取得が必要
- ブラウザからコピーすると `%2F` 等URLエンコード済み → **デコードしない**（そのまま使う）

---

## Notion API + 日本語（Windows環境の注意）

Windows環境のcurlはShift_JIS（CP932）でリクエストボディを送信するため、Notion APIに日本語テキストを送ると**文字化けする**。

### NG: curl直叩き

```bash
# ❌ Windows curlはShift_JISで送信 → Notion上で文字化け
curl -X POST https://api.notion.com/v1/blocks/xxx/children \
  -H "Content-Type: application/json" \
  -d '{"children":[{"paragraph":{"rich_text":[{"text":{"content":"日本語テキスト"}}]}}]}'
```

### OK: Node.js (https モジュール)

```javascript
// ✅ Node.jsはUTF-8で送信 → 文字化けしない
import https from 'https';

const body = JSON.stringify({
  children: [{
    paragraph: {
      rich_text: [{ text: { content: '日本語テキスト' } }]
    }
  }]
});

const req = https.request({
  hostname: 'api.notion.com',
  path: '/v1/blocks/xxx/children',
  method: 'PATCH',
  headers: {
    'Authorization': `Bearer ${process.env.NOTION_TOKEN}`,
    'Content-Type': 'application/json',
    'Notion-Version': '2022-06-28',
    'Content-Length': Buffer.byteLength(body)
  }
}, (res) => { /* ... */ });

req.write(body);
req.end();
```

**教訓**: Windows環境でNotion APIに日本語を含むリクエストを送る場合は、必ず**Node.js（httpsモジュール）**を使用する。MCP経由のNotion操作は内部でNode.jsを使うため問題ないが、デバッグや手動リクエストでcurlを使う場合に注意。

---

## トークン管理

- **ハードコード禁止**: スクリプト内にトークンを直接書かない
- **環境変数で管理**: `.mcp.json` の `env` セクション、または `process.env` で参照
- **APIキーの安全な保管**: `.mcp.json` は `.gitignore` に追加する

```
# .gitignore
.mcp.json
```

---

## 関連ファイル

- `.mcp.json.example` — 設定テンプレート
- `setup.md` — セットアップ手順（MCP設定含む）
