#!/usr/bin/env node
// ============================================================
// create-schedule-pptx.mjs — 全体スケジュール ペライチ パワポ生成
// Usage: node create-schedule-pptx.mjs
// Output: schedule.pptx
//
// カスタマイズ箇所:
//   1. weekHeaders — 週の開始日（月曜日ベース）
//   2. rawData — タスク定義（#/区分/アプローチ/実施事項/週フラグ）
//   3. THEME_COLORS — テーマ別カラー
//   4. todayWeekFloat — 赤線位置（weekIndex + dayOffset/7）
//   5. GW_COL_INDEX — GW週のインデックス
//   6. プロジェクト名・フッター
// ============================================================
import PptxGenJS from 'pptxgenjs';

// ============================================================
// Schedule Data — プロジェクトに合わせて書き換える
// ============================================================
const weekHeaders = [
  // 月曜日の日付を列挙（例: 20週分）
  '4/7','4/14','4/21','4/28','5/5','5/12','5/19','5/26',
  '6/2','6/9','6/16','6/23','6/30','7/7','7/14','7/21',
  '7/28','8/4','8/11','8/18',
];

const GW_COL_INDEX = 4; // GW週のインデックス（0始まり）。GWがない場合は -1

// [#, 区分, アプローチ, 実施事項, ...週フラグ]
// フラグ: 0=空, 1=■(バー), 2=★(マイルストーン)
const rawData = [
  // 例:
  [1,  'テーマA', '現状整理',   'ヒアリング実施',          1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
  [2,  'テーマA', '要件定義',   'データ項目の定義',        0,0,1,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
  ['', 'テーマA', '—',         '★ 要件定義完了',          0,0,0,0,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
  [3,  'テーマA', '設計・実装', 'システム構築',             0,0,0,0,0,1,1,1,1,1,0,0,0,0,0,0,0,0,0,0],
  ['', 'テーマA', '—',         '★ 実装完了',              0,0,0,0,0,0,0,0,0,2,0,0,0,0,0,0,0,0,0,0],
  [4,  'テーマA', '展開',       '全拠点展開・FB収集',      0,0,0,0,0,0,0,0,0,0,1,1,1,1,0,0,0,0,0,0],
  ['', 'テーマA', '—',         '★ 最終報告',              0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2],
];

// ============================================================
// 赤線の位置 — 定例前に書き換える
// ============================================================
// weekIndex + dayOffset/7
// 月曜=0, 火曜=1/7, 水曜=2/7, 木曜=3/7, 金曜=5/7
const todayWeekFloat = 0; // ← ここを書き換える

// ============================================================
// Colors
// ============================================================
const THEME_COLORS = {
  'テーマA': { bar: '4472C4', bg: 'D6E4F0', text: '1F3864' },  // Blue
  'テーマB': { bar: '70AD47', bg: 'E2EFDA', text: '375623' },  // Green
  'テーマC': { bar: 'ED7D31', bg: 'FCE4D6', text: '843C0C' },  // Orange
};
const MILESTONE_COLOR = 'FF0000';
const HEADER_BG = '333333';
const HEADER_TEXT = 'FFFFFF';
const GW_HEADER_BG = 'BFBFBF';
const GW_CELL_BG = 'E0E0E0';
const GW_TEXT_COLOR = '999999';
const BORDER = { pt: 0.5, color: 'CCCCCC' };

// ============================================================
// Build PPTX
// ============================================================
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'FLUX-G';

const slide = pptx.addSlide();

// タイトル
slide.addText('{PROJECT_NAME} 全体スケジュール（概要版）', {
  x: 0.3, y: 0.15, w: 12, h: 0.4,
  fontSize: 14, fontFace: 'Meiryo', bold: true, color: '333333',
});

// ============================================================
// レイアウト定数
// ============================================================
const LEFT = 0.3;
const TOP = 0.7;
const COL_NUM_W = 0.35;
const COL_CAT_W = 0.6;
const COL_APPR_W = 0.85;
const COL_TASK_W = 2.8;
const LABEL_COLS_W = COL_NUM_W + COL_CAT_W + COL_APPR_W + COL_TASK_W;
const WEEK_W = (13.33 - LEFT * 2 - LABEL_COLS_W) / weekHeaders.length;
const ROW_H = 0.28;
const HDR_H = 0.35;

// ============================================================
// ヘッダー行
// ============================================================
const headerCols = [
  { text: '#', w: COL_NUM_W },
  { text: '区分', w: COL_CAT_W },
  { text: 'アプローチ', w: COL_APPR_W },
  { text: '実施事項', w: COL_TASK_W },
];

let x = LEFT;
for (const col of headerCols) {
  slide.addText(col.text, {
    x, y: TOP, w: col.w, h: HDR_H,
    fontSize: 7, fontFace: 'Meiryo', bold: true,
    color: HEADER_TEXT, fill: { color: HEADER_BG },
    align: 'center', valign: 'middle',
    border: [BORDER, BORDER, BORDER, BORDER],
  });
  x += col.w;
}

// 月ヘッダー
const months = {};
weekHeaders.forEach((wh, i) => {
  const m = wh.split('/')[0] + '月';
  if (!months[m]) months[m] = { start: i, count: 0 };
  months[m].count++;
});

// 月ラベル
for (const [month, info] of Object.entries(months)) {
  slide.addText(month, {
    x: LEFT + LABEL_COLS_W + info.start * WEEK_W,
    y: TOP - 0.2,
    w: info.count * WEEK_W,
    h: 0.2,
    fontSize: 7, fontFace: 'Meiryo', bold: true,
    color: '333333', align: 'center', valign: 'middle',
  });
}

// 週ヘッダー
weekHeaders.forEach((wh, i) => {
  const isGW = i === GW_COL_INDEX;
  slide.addText(wh, {
    x: LEFT + LABEL_COLS_W + i * WEEK_W,
    y: TOP,
    w: WEEK_W,
    h: HDR_H,
    fontSize: 6, fontFace: 'Meiryo', bold: true,
    color: isGW ? GW_TEXT_COLOR : HEADER_TEXT,
    fill: { color: isGW ? GW_HEADER_BG : HEADER_BG },
    align: 'center', valign: 'middle',
    border: [BORDER, BORDER, BORDER, BORDER],
  });
});

// ============================================================
// データ行
// ============================================================
rawData.forEach((row, rowIdx) => {
  const [num, cat, approach, task, ...flags] = row;
  const isMilestone = num === '';
  const theme = THEME_COLORS[cat] || { bar: '999999', bg: 'F0F0F0', text: '333333' };
  const y = TOP + HDR_H + rowIdx * ROW_H;

  // ラベル列
  const labelData = [
    { text: String(num), w: COL_NUM_W, align: 'center' },
    { text: cat, w: COL_CAT_W, align: 'center' },
    { text: approach, w: COL_APPR_W, align: 'center' },
    { text: task, w: COL_TASK_W, align: 'left' },
  ];

  let lx = LEFT;
  for (const ld of labelData) {
    slide.addText(ld.text, {
      x: lx, y, w: ld.w, h: ROW_H,
      fontSize: 6, fontFace: 'Meiryo',
      color: isMilestone ? MILESTONE_COLOR : theme.text,
      bold: isMilestone,
      fill: { color: isMilestone ? 'FFF2CC' : theme.bg },
      align: ld.align, valign: 'middle',
      border: [BORDER, BORDER, BORDER, BORDER],
    });
    lx += ld.w;
  }

  // 週セル
  flags.forEach((flag, fi) => {
    const isGW = fi === GW_COL_INDEX;
    let cellText = '';
    let cellColor = '333333';
    let cellFill = isGW ? GW_CELL_BG : 'FFFFFF';

    if (flag === 1) {
      cellText = '■';
      cellColor = theme.bar;
    } else if (flag === 2) {
      cellText = '★';
      cellColor = MILESTONE_COLOR;
      cellFill = 'FFF2CC';
    }

    slide.addText(cellText, {
      x: LEFT + LABEL_COLS_W + fi * WEEK_W,
      y, w: WEEK_W, h: ROW_H,
      fontSize: 8, fontFace: 'Meiryo', bold: true,
      color: isGW && flag === 0 ? GW_TEXT_COLOR : cellColor,
      fill: { color: cellFill },
      align: 'center', valign: 'middle',
      border: [BORDER, BORDER, BORDER, BORDER],
    });
  });
});

// ============================================================
// 赤線（Today）
// ============================================================
if (todayWeekFloat > 0) {
  const totalH = HDR_H + rawData.length * ROW_H;
  const todayX = LEFT + LABEL_COLS_W + todayWeekFloat * WEEK_W;

  slide.addShape(pptx.shapes.LINE, {
    x: todayX, y: TOP - 0.05, w: 0, h: totalH + 0.1,
    line: { color: 'FF0000', width: 1.5 },
    beginArrowType: 'oval',
    endArrowType: 'oval',
  });
}

// ============================================================
// 凡例
// ============================================================
const legendY = TOP + HDR_H + rawData.length * ROW_H + 0.15;
slide.addText([
  { text: '■', options: { color: '4472C4', bold: true } },
  { text: ' 実施期間　', options: { fontSize: 7 } },
  { text: '★', options: { color: 'FF0000', bold: true } },
  { text: ' マイルストーン　', options: { fontSize: 7 } },
  { text: '|', options: { color: 'FF0000', bold: true } },
  { text: ' 今日', options: { fontSize: 7 } },
], {
  x: LEFT, y: legendY, w: 6, h: 0.25,
  fontSize: 8, fontFace: 'Meiryo',
});

// フッター
slide.addText('Confidential — {PROJECT_NAME}', {
  x: 0, y: 7.1, w: 13.33, h: 0.3,
  fontSize: 6, fontFace: 'Meiryo', color: '999999', align: 'center',
});

// ============================================================
// Save
// ============================================================
await pptx.writeFile({ fileName: 'schedule.pptx' });
console.log('[OK] schedule.pptx generated');
