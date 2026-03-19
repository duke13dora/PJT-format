# PowerPoint生成ガイド（pptxgenjs）

> ガントチャート（スケジュール）PPTXの自動生成パターン。PJT-YPで確立した手法。

---

## 基本構成

```javascript
import PptxGenJS from 'pptxgenjs';

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'WIDE', width: 13.33, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = 'FLUX-G';

const slide = pptx.addSlide();
```

---

## ガントチャートのデータ形式

### rawData配列

```javascript
// [#, 区分, アプローチ, 実施事項, ...週フラグ]
// フラグ: 0=空, 1=■(バー), 2=★(マイルストーン)
const rawData = [
  [1, 'テーマA', '現状整理', 'ヒアリング実施',  1,1,0,0,0,0,0,0],
  [2, 'テーマA', '要件定義', 'データ項目定義',  0,0,1,1,1,0,0,0],
  ['','テーマA', '—',       '★ 中間報告',     0,0,0,0,2,0,0,0],
  // ...
];
```

### weekHeaders

```javascript
const weekHeaders = ['4/7','4/14','4/21','4/28','5/5','5/12','5/19','5/26'];
```

### テーマカラー

```javascript
const THEME_COLORS = {
  'テーマA': { bar: '4472C4', bg: 'D6E4F0', text: '1F3864' },  // Blue
  'テーマB': { bar: '70AD47', bg: 'E2EFDA', text: '375623' },  // Green
  'テーマC': { bar: 'ED7D31', bg: 'FCE4D6', text: '843C0C' },  // Orange
};
```

---

## 赤線（Today線）の制御

### todayWeekFloat

```javascript
// weekIndex + dayOffset/7
// 月曜=0, 火曜=1/7, 水曜=2/7, ..., 金曜=5/7, 日曜=6/7
const todayWeekFloat = 4 + 5/7; // 第4週の金曜日
```

### 描画

```javascript
// X座標の計算
const todayX = LEFT_MARGIN + LABEL_COLS_WIDTH + todayWeekFloat * WEEK_COL_WIDTH;

slide.addShape(pptx.shapes.LINE, {
  x: todayX,
  y: topY - 0.05,
  w: 0,
  h: totalHeight + 0.1,
  line: { color: 'FF0000', width: 1.5 },
  beginArrowType: 'oval',  // 丸ポチ付き端
  endArrowType: 'oval',
});
```

### 注意事項
- 赤線の端は `beginArrowType:'oval', endArrowType:'oval'` を使用（丸ポチ付き）
- 「Today」テキスト等を勝手に追加しない
- スケジュールデータ（rawData）自体は変更しない（赤線の位置だけ変える）

---

## GW・祝日の処理

```javascript
const GW_COL_INDEX = 12; // GW週のインデックス（0始まり）
const GW_HEADER_BG = 'BFBFBF';
const GW_CELL_BG = 'E0E0E0';
const GW_TEXT_COLOR = '999999';

// GW列はグレーアウトして視覚的に区別
```

---

## フォント

- 日本語: `Meiryo` / `メイリオ` を指定
- pptxgenjsのデフォルトフォントは英語フォントのため、明示的に日本語フォントを指定する

---

## 出力・共有

```javascript
await pptx.writeFile({ fileName: 'schedule.pptx' });
```

生成後、Google Drive の共有フォルダにコピー:
```bash
cp schedule.pptx "drive/50.打ち合わせ関連/【XX様】全体スケジュール（概要版）.pptx"
```

---

## 関連ファイル

- `scripts/create-schedule-pptx.mjs` — 汎用版スケジュール生成スクリプト
- `guides/01-kickoff.md` — キックオフフロー（スケジュールはセクション5）
