// create-kickoff-notion.mjs — キックオフMDの内容をNotionページに展開
// Usage: node create-kickoff-notion.mjs
//
// 前提:
// - NOTION_TOKEN 環境変数にトークンを設定
// - PAGE_ID を対象ページのIDに書き換え
//
// テンプレート構造:
// キックオフMD（templates/kickoff-format.md）の6セクションに対応する
// Notionブロックを生成する。内容はプロジェクトに合わせて書き換えること。

const NOTION_TOKEN = process.env.NOTION_TOKEN;
const PAGE_ID = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'; // 対象ページID

if (!NOTION_TOKEN) {
  console.error('ERROR: NOTION_TOKEN 環境変数を設定してください');
  process.exit(1);
}

// ============================================================
// Block helpers（汎用 — そのまま再利用可能）
// ============================================================
const t = (content, ann = {}) => ({
  type: 'text', text: { content },
  ...(Object.keys(ann).length ? { annotations: ann } : {}),
});
const b = (content) => t(content, { bold: true });
const it = (content) => t(content, { italic: true });

const h1 = (text) => ({ type: 'heading_1', heading_1: { rich_text: [t(text)] } });
const h2 = (text) => ({ type: 'heading_2', heading_2: { rich_text: [t(text)] } });
const h3 = (text) => ({ type: 'heading_3', heading_3: { rich_text: [t(text)] } });
const p = (rt) => ({
  type: 'paragraph',
  paragraph: { rich_text: Array.isArray(rt) ? rt : (rt ? [t(rt)] : []) },
});
const bullet = (rt) => ({
  type: 'bulleted_list_item',
  bulleted_list_item: { rich_text: Array.isArray(rt) ? rt : [t(rt)] },
});
const num = (rt) => ({
  type: 'numbered_list_item',
  numbered_list_item: { rich_text: Array.isArray(rt) ? rt : [t(rt)] },
});
const divider = () => ({ type: 'divider', divider: {} });
const quote = (rt) => ({
  type: 'quote',
  quote: { rich_text: Array.isArray(rt) ? rt : [t(rt)] },
});
const callout = (emoji, text, color = 'gray_background') => ({
  type: 'callout',
  callout: {
    rich_text: Array.isArray(text) ? text : [t(text)],
    icon: { type: 'emoji', emoji },
    color,
  },
});
const tbl = (width, hasHeader, rows) => ({
  type: 'table',
  table: {
    table_width: width,
    has_column_header: hasHeader,
    has_row_header: false,
    children: rows.map((cells) => ({
      type: 'table_row',
      table_row: {
        cells: cells.map((cell) => {
          if (!Array.isArray(cell)) return [t(String(cell))];
          const flat = cell.flat();
          return flat.map((item) => (typeof item === 'string' ? t(item) : item));
        }),
      },
    })),
  },
});

// Status badges
const done = [t('済', { bold: true, color: 'green' })];
const notDone = [t('未', { bold: true, color: 'orange' })];

// Schedule cell helpers
const BAR = [t('■', { color: 'blue' })];
const STAR = [t('★', { color: 'orange' })];
const EMPTY = [t('')];
const wk = (flag) => (flag === 1 ? BAR : flag === 2 ? STAR : EMPTY);

// ============================================================
// Section builders — プロジェクトに合わせて書き換える
// ============================================================

function buildSection01() {
  return [
    h1('01. プロジェクト概要'),
    callout('🎯', [
      b('{ミッション文}'),
    ], 'blue_background'),

    h3('背景'),
    p('<!-- Pain → Gain のストーリー -->'),

    h3('目的・スコープ'),
    bullet('スコープ内: '),
    bullet('スコープ外: '),

    h3('プロジェクト体制'),
    tbl(4, true, [
      [[b('組織')], [b('氏名')], [b('役割')], [b('備考')]],
      ['自社', '', '', ''],
      ['顧客', '', '', ''],
    ]),

    divider(),
  ];
}

function buildSection02() {
  return [
    h1('02. 弊社理解（現状認識とゴール）'),
    callout('📊', [
      t('弊社の現状認識とゴールを整理し、認識を合わせる。'),
    ], 'green_background'),

    h2('現状認識'),
    p([b('As-Is:')]),
    bullet('<!-- 現状1 -->'),

    h2('ゴール'),
    p([b('To-Be:')]),
    bullet('<!-- あるべき姿1 -->'),

    divider(),
  ];
}

function buildSection03() {
  return [
    h1('03. 論点'),
    callout('❓', [
      t('プロジェクトを進める上で判断が必要な事項を整理する。'),
    ], 'yellow_background'),

    p([b('大論点1 — ')]),
    bullet([b('1-1. '), t('<!-- 小論点 -->')]),

    divider(),
  ];
}

function buildSection04() {
  return [
    h1('04. アプローチ・実施事項・アウトプット'),

    h3('フェーズと実施事項'),
    tbl(6, true, [
      [[b('#')], [b('フェーズ')], [b('実施事項')], [b('済/未')], [b('備考')], [b('対応論点')]],
      ['1', '', '', notDone, '', ''],
    ]),

    h3('IPO整理'),
    tbl(5, true, [
      [[b('#')], [b('実施事項')], [b('インプット')], [b('プロセス')], [b('アウトプット')]],
      ['1', '', '', '', ''],
    ]),

    divider(),
  ];
}

function buildSection05() {
  return [
    h1('05. マイルストーン'),
    tbl(3, true, [
      [[b('時期')], [b('マイルストーン')], [b('備考')]],
      ['', '', ''],
    ]),

    divider(),
  ];
}

function buildSection06() {
  return [
    h1('06. リスクと前提条件'),

    h3('前提条件'),
    bullet('<!-- 前提条件1 -->'),

    h3('リスク'),
    tbl(5, true, [
      [[b('#')], [b('リスク')], [b('影響度')], [b('発生可能性')], [b('対応方針')]],
      ['1', '', '', '', ''],
    ]),

    divider(),
  ];
}

// ============================================================
// Notion API helpers
// ============================================================
async function notionFetch(path, method = 'GET', body = null) {
  const res = await fetch(`https://api.notion.com/v1${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${NOTION_TOKEN}`,
      'Notion-Version': '2022-06-28',
      'Content-Type': 'application/json',
    },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Notion API ${method} ${path}: ${res.status} ${err}`);
  }
  return res.json();
}

async function deleteAllChildren(pageId) {
  const res = await notionFetch(`/blocks/${pageId}/children?page_size=100`);
  for (const block of res.results) {
    await notionFetch(`/blocks/${block.id}`, 'DELETE');
  }
}

async function appendBlocks(pageId, blocks) {
  // Notion API は1回あたり100ブロックまで
  for (let i = 0; i < blocks.length; i += 100) {
    const chunk = blocks.slice(i, i + 100);
    await notionFetch(`/blocks/${pageId}/children`, 'PATCH', { children: chunk });
  }
}

// ============================================================
// Main
// ============================================================
async function main() {
  console.log('Deleting existing blocks...');
  await deleteAllChildren(PAGE_ID);

  const blocks = [
    ...buildSection01(),
    ...buildSection02(),
    ...buildSection03(),
    ...buildSection04(),
    ...buildSection05(),
    ...buildSection06(),
  ];

  console.log(`Appending ${blocks.length} blocks...`);
  await appendBlocks(PAGE_ID, blocks);
  console.log('[OK] Notion page updated');
}

main().catch(console.error);
