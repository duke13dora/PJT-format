# 11. Asana運用ガイド

> PJT-TD（東京ドーム案件）で確立。Backlog運用ガイド（10-backlog-operations.md）のAsana版。

---

## 1. API操作の基本ルール

- **Node.js fetch を使用**（curlはWindows UTF-8文字化けリスク）
- **MCP（`@roychri/mcp-server-asana`）はコメント投稿に制限あり** → 後述のREST API直接呼出しを併用
- **投稿・編集後の必須確認**: APIレスポンス確認 + Asana UI目視確認の2段階

```javascript
// 基本的なAPI呼出しパターン
const ASANA_TOKEN = process.env.ASANA_ACCESS_TOKEN; // .mcp.json の env から取得
const WORKSPACE_GID = '{YOUR_WORKSPACE_GID}';

async function asanaApi(path, method = 'GET', body = null) {
  const opts = {
    method,
    headers: {
      Authorization: 'Bearer ' + ASANA_TOKEN,
      'Content-Type': 'application/json',
    },
  };
  if (body) opts.body = JSON.stringify({ data: body });
  const res = await fetch('https://app.asana.com/api/1.0' + path, opts);
  return res.json();
}
```

---

## 2. メンション・通知

### プロフィールURL方式（★正解）

Asanaでメンションを実現するには、コメントの `text` パラメータにプロフィールURLを含める。Asanaが自動でメンションに変換する。

```
https://app.asana.com/1/{workspace_gid}/profile/{membership_gid}
```

- `workspace_gid`: ワークスペースのGID
- `membership_gid`: 各メンバーのワークスペースメンバーシップGID（※ユーザーGIDとは別物）

### メンバーシップGIDの調べ方

```javascript
// ワークスペースメンバー一覧からGIDを取得
const members = await asanaApi(
  '/workspaces/{workspace_gid}/workspace_memberships?opt_fields=user.name'
);
```

**重要**: メンバーシップGIDはプロジェクトMEMORY.mdに記録しておく。毎回APIで取得するのは非効率。

### cc書式

メイン宛先を先頭、次行に「cc」＋スペース区切りで列挙。ccには対象以外の全メンバーを入れる。

```javascript
const text =
  profileUrl('{MAIN_RECIPIENT_GID}') +
  '\ncc ' +
  profileUrl('{MEMBER2_GID}') +
  ' ' +
  profileUrl('{MEMBER3_GID}') +
  ' {手動名前}' +
  '\n\n' +
  '本文をここに記載';
```

**注意**: 休暇中のメンバーなど、メンションすべきでない人は漢字テキストで記載（URLにしない）。

---

## 3. html_text方式の廃止警告

### ⚠ 使用禁止

以下の方式は **2026-03時点で動作しない**:

```javascript
// ❌ 禁止: html_textでメンション
headers: { 'Asana-Enable': 'new_rich_text' }
body: { data: { html_text: '<a data-asana-gid="..." data-asana-type="user">@名前</a>' } }
// → HTMLタグがそのまま表示される
```

```javascript
// ❌ 禁止: MCPのtext にHTMLタグを含める
create_task_story({ text: '<b>太字</b>テキスト' });
// → 文字化けする
```

### ✅ 正しい方法

`text` パラメータにプレーンテキスト + プロフィールURLのみ使用:

```javascript
const res = await fetch(
  'https://app.asana.com/api/1.0/tasks/' + taskId + '/stories',
  {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      data: { text: profileUrl('{GID}') + '\n\n本文テキスト' },
    }),
  }
);
```

---

## 4. MCP制限と回避策

`@roychri/mcp-server-asana` のMCP経由では以下の制限がある:

| 操作                   | MCP対応                         | 回避策                                 |
| ---------------------- | ------------------------------- | -------------------------------------- |
| タスク作成             | ✅ 可能                         | —                                      |
| コメント投稿           | ✅ 可能（ただしメンション不可） | REST API直接呼出し                     |
| セクション配置         | ❌ 不可                         | `POST /sections/{section_gid}/addTask` |
| notes改行              | ⚠ `&#10;` がリテラル表示        | REST API経由で `\n` 使用               |
| カスタムフィールド設定 | ✅ 可能                         | —                                      |

### セクション配置のREST API呼出し

MCPでタスクを作成するとデフォルトセクションに配置される。正しいセクションに移動するには:

```javascript
await fetch('https://app.asana.com/api/1.0/sections/{section_gid}/addTask', {
  method: 'POST',
  headers: {
    Authorization: 'Bearer ' + token,
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({ data: { task: '{task_gid}' } }),
});
```

### セクション一覧の取得

```javascript
const sections = await asanaApi('/projects/{project_gid}/sections');
// sections.data → [{ gid, name }, ...]
```

---

## 5. コメント書式

### 基本構成

```
@宛先
cc @メンバー1 @メンバー2

（本文）
```

### 長文コメントのパターン

```
@宛先
cc @メンバー1 @メンバー2

まだ詳細を詰めている段階だと思うのでそれ故のズレもある認識ですが、
いくつか確認させてください。

■1. タイトル1
現状の事実の説明。
確認事項: Yes/Noまたは二択で回答しやすい形式にする。

■2. タイトル2
現状の事実の説明。
→ 具体的なアクション提案。

以上2点になります。お手すきの際にご確認いただけますと幸いです。
```

**ポイント**:

- 見出しは `■番号. タイトル` 形式
- 冒頭に免責的クッション（「まだ詳細を詰めている段階だと思うので…」等）
- 質問はYes/No or 二択で回答しやすく
- 末尾は「以上N点になります。お手すきの際にご確認いただけますと幸いです。」

---

## 6. 投稿後検証

コメント投稿後は **必ず2段階の検証** を行う:

### Step 1: APIレスポンス確認

```javascript
const result = await res.json();
if (result.data) {
  console.log('SUCCESS - GID:', result.data.gid);
  console.log('Text preview:', result.data.text.substring(0, 200));
} else {
  console.log('ERROR:', JSON.stringify(result));
}
```

### Step 2: GETで投稿内容を再取得

```javascript
const verify = await asanaApi('/stories/' + result.data.gid);
console.log(verify.data.text); // 全文確認
```

### Step 3: Asana UI上で目視確認

- メンションが正しく変換されているか（クリック可能なリンクになっているか）
- 文字化けがないか
- 改行が意図通りか

**⚠ UI確認を完了してからユーザーに報告すること**（お客様に通知が飛ぶため、文字化けは取り返しがつかない）

---

## 7. AI感排除原則

外部向けコメント（Asana、メール等）では、自動化ツールの存在を一切言及しない。

### 禁止表現

| ❌ 禁止                    | ✅ 代替                |
| -------------------------- | ---------------------- |
| 「スクリプトで対応します」 | 「対応進めていきます」 |
| 「ツールで自動化しました」 | 「修正しました」       |
| 「APIで取得した結果」      | 「確認したところ」     |
| 「自動生成しました」       | 「作成しました」       |

### 人名表記

チームメンバーの名前は日本語表記で統一。英語名・ローマ字名は禁止。

| ❌ 禁止               | ✅ 正しい                |
| --------------------- | ------------------------ |
| Kai, Kai Shitashima   | 下島（自分の場合は省略） |
| fukuoka, FC福岡氏     | 福岡さん                 |
| Masuda, Fuhito Masuda | 桝田さん                 |

---

## 8. 言い回しルール（口調模倣ガイドライン）

外部コメントはプロジェクト担当者の口調を模倣する。以下はPJT-TDで確立したパターン:

| ルール                 | 例                                                         |
| ---------------------- | ---------------------------------------------------------- |
| 「了解」→「承知」      | 「承知しました」「承知です」                               |
| クッション言葉を入れる | 「またしつこくて申し訳ないんですが」「させていただきたく」 |
| 断定を避ける           | 「〜ようです」「〜余地がある」「〜残りそう」               |
| 注釈は※で              | 「※一方で〜」「※もし〜」                                   |
| 文末の余韻は「。。」   | 「〜しかなさそうですが。。」                               |
| 改行を多めに           | 1文ずつ改行。密な段落は避ける                              |
| 語尾カジュアル混在OK   | 「〜かと思います」「〜ですかね」                           |

**注意**: 口調パターンはプロジェクトごとに異なる。新規プロジェクト開始時に担当者の過去コメントを分析し、MEMORY.mdに記録すること。

---

## 9. チケット起票ルール

### 必須設定項目

- **セクション**: MCPデフォルトではなく正しいセクションに配置（REST API使用）
- **カスタムフィールド**: ステータス等のカスタムフィールドを設定

```javascript
// カスタムフィールド付きタスク作成
await asanaApi('/tasks', 'POST', {
  name: 'タスク名',
  projects: ['{project_gid}'],
  custom_fields: {
    '{status_field_gid}': '{status_value_gid}', // 例: 未着手
  },
});
```

### 起票前チェック

- 既存チケットの重複確認（`search_tasks` で検索）
- 親タスクとの関係確認（サブタスクにすべきか、独立タスクにすべきか）

---

## 10. 定例ネクストからのチケット作成パターン

定例会議後のネクストアイテムをAsanaチケットに変換する際のルール:

### サブタスク作成時

各サブタスクに「MM/DD定例ネクスト：（チケット名と同内容）」のコメントを追加。
コメント冒頭に担当者のメンションリンクを入れる。

```javascript
// サブタスク作成 + 初期コメント投稿のパターン
const task = await asanaApi('/tasks/{parent_gid}/subtasks', 'POST', {
  name: 'サブタスク名',
});

await fetch(
  'https://app.asana.com/api/1.0/tasks/' + task.data.gid + '/stories',
  {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      data: {
        text:
          profileUrl('{ASSIGNEE_GID}') + '\n\nMM/DD定例ネクスト：サブタスク名',
      },
    }),
  }
);
```

### 既存チケットへのコメント追加

新規起票ではなく既存チケットにコメントで指示を追加するケース:

```javascript
await fetch('https://app.asana.com/api/1.0/tasks/{existing_task_gid}/stories', {
  method: 'POST',
  headers: {
    Authorization: 'Bearer ' + token,
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({
    data: {
      text: profileUrl('{ASSIGNEE_GID}') + '\n\nMM/DD定例ネクスト：指示内容',
    },
  }),
});
```

---

## 11. 事前つつきコメント運用

→ 運用ルール・フォーマット: [guides/06-meeting-workflow.md](06-meeting-workflow.md)「事前つつきコメント運用」参照

Asana固有の補足: メンションはプロフィールURL方式（本ガイド セクション2参照）を使用すること。

---

## トークン管理

- Asanaトークンは `.mcp.json` の `env` で管理
- ハードコード禁止
- トークンの読取は `.mcp.json` をパースして取得:

```javascript
const fs = require('fs');
const mcpConfig = JSON.parse(fs.readFileSync('.mcp.json', 'utf-8'));
const token = mcpConfig.mcpServers.asana.env.ASANA_ACCESS_TOKEN;
```
