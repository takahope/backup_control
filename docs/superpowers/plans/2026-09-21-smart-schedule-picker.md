# 智慧排程時間連動選擇器與雙軌標準化語法實作計畫 (Implementation Plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在備份管制前端表單建構智慧排程時間複合選擇器，並實作雙軌（完整/差異）標準字串合成與歷史資料反向解析正規化，消滅非法輸入並保證跨專案 100% 正則解析相容。

**Architecture:** 前端將原文字輸入框替換為結構化複合元件（星期/日期/月份下拉 ＋ 原生 `<input type="time">`），由前端控制器依週期單位動態切換元件可見度。透過純函式合成器（`composeDualTrackSchedule`）與解析器（`parseScheduleString`）將 UI 狀態同步至隱藏輸入欄位，後端儲存格式維持透明相容。

**Tech Stack:** HTML5 / Tailwind CSS / Vanilla JavaScript (ES6) / Node.js Test Runner

**Spec:** [`docs/superpowers/specs/2026-09-21-smart-schedule-picker-design.md`](file:///home/itsutaka/biobank_manager/backup_control/docs/superpowers/specs/2026-09-21-smart-schedule-picker-design.md)

## Global Constraints

- 單軌輸出格式：天為 `每日 HH:mm`、週為 `每週{一~日} HH:mm`、月為 `每月 {1~31} 日 HH:mm` 或 `每月月底最後一日 HH:mm`、年為 `每年 {1~12} 月 {1~31} 日 HH:mm` 或 `每年 {1~12} 月月底最後一日 HH:mm`、版本為 `每次版本發布`。
- 本地雙軌輸出格式：同時有完整與有效差異備份時為 `完整:{完整排程} / 差異:{差異排程}`；單軌時為單一排程字串。
- 系統日誌檔固定為天週期模式（預設 `每日 00:00`）。
- 異地備份對稱採用相同的單軌時間格式；未啟用時為 `無`。
- 編輯載入遇非標準歷史舊值時，一律自動提取時間並依照週期單位正規化為標準預設值（選項 A）。
- 嚴格遵守 [`~/.agents/PLAYBOOK.md`](file:///home/itsutaka/.agents/PLAYBOOK.md) 規範，絕不主動執行 `clasp push`。

---

### Task 1: 排程標準語法合成與反向解析核心引擎 (Schedule Engine & Unit Tests)

**Files:**
- Create: `tests/schedule-picker.test.cjs`
- Modify: `Index.html`（在 `<script>` 區塊開頭加入純核心函式）

**Interfaces:**
- Produces:
  - `formatTrackSchedule(unit, time, options)`: 依週期單位產出單軌標準字串
  - `composeDualTrackSchedule(options)`: 本地雙軌/單軌字串合成函式
  - `parseScheduleString(scheduleStr, defaultUnit)`: 反向解析排程字串為結構化物件（含容錯正規化）

- [x] **Step 1: 建立失敗的單元測試檔案**

建立 `tests/schedule-picker.test.cjs`，覆蓋所有週期單位、雙軌合成與非標準歷史資料解析：

```javascript
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const vm = require('vm');

// 載入 Index.html 內部的核心純函式
const html = fs.readFileSync('Index.html', 'utf8');
const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/i);
if (!scriptMatch) throw new Error('Index.html script tag not found');

const sandbox = {};
vm.createContext(sandbox);
vm.runInContext(scriptMatch[1], sandbox);

const { formatTrackSchedule, composeDualTrackSchedule, parseScheduleString } = sandbox;

test('formatTrackSchedule 產生單軌標準時間格式', () => {
  assert.equal(formatTrackSchedule('天', '02:30', {}), '每日 02:30');
  assert.equal(formatTrackSchedule('週', '03:00', { weekday: '日' }), '每週日 03:00');
  assert.equal(formatTrackSchedule('個月', '04:00', { day: '1' }), '每月 1 日 04:00');
  assert.equal(formatTrackSchedule('個月', '23:00', { day: '月底最後一日' }), '每月月底最後一日 23:00');
  assert.equal(formatTrackSchedule('年', '05:00', { month: '12', day: '31' }), '每年 12 月 31 日 05:00');
  assert.equal(formatTrackSchedule('年', '02:00', { month: '6', day: '月底最後一日' }), '每年 6 月月底最後一日 02:00');
  assert.equal(formatTrackSchedule('版本', '00:00', {}), '每次版本發布');
});

test('composeDualTrackSchedule 支援雙軌與單軌合成', () => {
  // 雙軌情境
  const dualResult = composeDualTrackSchedule({
    hasDiff: true,
    fullSchedule: '每週日 02:00',
    diffSchedule: '每日 02:00'
  });
  assert.equal(dualResult, '完整:每週日 02:00 / 差異:每日 02:00');

  // 僅完整軌
  const fullOnly = composeDualTrackSchedule({
    hasDiff: false,
    fullSchedule: '每週日 02:00',
    diffSchedule: '每日 02:00'
  });
  assert.equal(fullOnly, '每週日 02:00');

  // 僅差異軌
  const diffOnly = composeDualTrackSchedule({
    hasDiff: true,
    fullSchedule: '',
    diffSchedule: '每日 02:00'
  });
  assert.equal(diffOnly, '每日 02:00');
});

test('parseScheduleString 支援反向解析與歷史舊資料正規化 (選項 A)', () => {
  // 1. 雙軌字串解析
  const dualParsed = parseScheduleString('完整:每週日 02:00 / 差異:每日 00:00', '週');
  assert.equal(dualParsed.isDual, true);
  assert.equal(dualParsed.full.weekday, '日');
  assert.equal(dualParsed.full.time, '02:00');
  assert.equal(dualParsed.diff.time, '00:00');

  // 2. 單軌月解析
  const monthParsed = parseScheduleString('每月月底最後一日 23:00', '個月');
  assert.equal(monthParsed.full.day, '月底最後一日');
  assert.equal(monthParsed.full.time, '23:00');

  // 3. 非標準舊資料正規化 (提取到時間)
  const legacyParsed = parseScheduleString('凌晨 02:30 執行備份', '週');
  assert.equal(legacyParsed.full.time, '02:30');
  assert.equal(legacyParsed.full.weekday, '日'); // 預設週日

  // 4. 完全無時間的舊資料正規化
  const emptyLegacy = parseScheduleString('無排程', '天');
  assert.equal(emptyLegacy.full.time, '00:00');
});
```

- [x] **Step 2: 執行測試驗證其失敗 (Red)**

Run: `node --test tests/schedule-picker.test.cjs`
Expected: FAIL (函式未定義 `formatTrackSchedule is not a function`)

- [x] **Step 3: 在 `Index.html` 實作核心純函式 (Green)**

在 `Index.html` 的 `<script>` 區塊中定義 `formatTrackSchedule`、`composeDualTrackSchedule`、`parseScheduleString`：

```javascript
// --- 排程時間標準語法核心引擎 ---
function formatTrackSchedule(unit, time, options = {}) {
  const u = String(unit || '天').trim();
  const t = String(time || '00:00').trim();
  if (u === '版本') return '每次版本發布';
  if (u === '天') return `每日 ${t}`;
  if (u === '週') {
    const w = options.weekday || '日';
    return `每週${w} ${t}`;
  }
  if (u === '個月') {
    const d = options.day || '1';
    return d === '月底最後一日' ? `每月月底最後一日 ${t}` : `每月 ${d} 日 ${t}`;
  }
  if (u === '年') {
    const m = options.month || '12';
    const d = options.day || '31';
    return d === '月底最後一日' ? `每年 ${m} 月月底最後一日 ${t}` : `每年 ${m} 月 ${d} 日 ${t}`;
  }
  return `每日 ${t}`;
}

function composeDualTrackSchedule(options = {}) {
  const hasDiff = Boolean(options.hasDiff);
  const fullStr = String(options.fullSchedule || '').trim();
  const diffStr = hasDiff ? String(options.diffSchedule || '').trim() : '';

  if (hasDiff && fullStr && diffStr) {
    return `完整:${fullStr} / 差異:${diffStr}`;
  }
  if (fullStr) return fullStr;
  if (diffStr) return diffStr;
  return '每日 00:00';
}

function parseScheduleString(scheduleStr, defaultUnit = '天') {
  const raw = String(scheduleStr || '').trim();
  const result = {
    isDual: false,
    full: { time: '00:00', weekday: '日', month: '12', day: '1' },
    diff: { time: '00:00', weekday: '日', month: '12', day: '1' }
  };

  if (!raw) return result;

  const dualMatch = raw.match(/^完整:(.*?)\s*\/\s*差異:(.*)$/);
  if (dualMatch) {
    result.isDual = true;
    result.full = parseSingleScheduleSegment_(dualMatch[1].trim(), defaultUnit);
    result.diff = parseSingleScheduleSegment_(dualMatch[2].trim(), '天');
    return result;
  }

  result.full = parseSingleScheduleSegment_(raw, defaultUnit);
  return result;
}

function parseSingleScheduleSegment_(seg, unit) {
  const segment = String(seg || '').trim();
  const parsed = { time: '00:00', weekday: '日', month: '12', day: '1' };

  // 1. 提取時間 HH:mm
  const timeMatch = segment.match(/(?:^|[^\d])([01]?\d|2[0-3]):([0-5]\d)(?::\d{2})?(?:[^\d]|$)/);
  if (timeMatch) {
    const h = timeMatch[1].padStart(2, '0');
    const m = timeMatch[2].padStart(2, '0');
    parsed.time = `${h}:${m}`;
  }

  // 2. 提取星期
  const weekMatch = segment.match(/每?週([一二三四五六日])/);
  if (weekMatch) parsed.weekday = weekMatch[1];

  // 3. 提取月份與日期
  if (segment.includes('月底最後一日')) {
    parsed.day = '月底最後一日';
  } else {
    const dayMatch = segment.match(/每?月\s*([0-3]?\d)\s*日/);
    if (dayMatch) parsed.day = String(parseInt(dayMatch[1], 10));
  }

  const yearMatch = segment.match(/每?年\s*([0-1]?\d)\s*月/);
  if (yearMatch) parsed.month = String(parseInt(yearMatch[1], 10));

  return parsed;
}
```

- [x] **Step 4: 重新執行測試驗證通過**

Run: `node --test tests/schedule-picker.test.cjs`
Expected: PASS (All 3 tests passing)

- [x] **Step 5: Commit**

```bash
git add tests/schedule-picker.test.cjs Index.html
git commit -m "feat: implement core schedule syntax engine and dual-track composer"
```

---

### Task 2: 前端複合選擇器 HTML 結構與動態切換控制器 (UI & Dynamic Controller)

**Files:**
- Modify: `Index.html` (更新 `renderFormTargets`、增加 `renderSchedulePickerTrack` 與動態事件監聽)

**Interfaces:**
- Consumes: Task 1 的 `formatTrackSchedule` 與 `composeDualTrackSchedule`
- Produces:
  - `renderSchedulePickerHtml(targetId, track)`: 產生單一排程軌之複合 DOM 結構
  - `updateSchedulePickerUI(targetId, track, unit)`: 依單位切換欄位顯示/隱藏
  - `syncTargetSchedule(targetId)`: 自動更新隱藏 input 與 UI 預覽文字

- [x] **Step 1: 編寫 UI 控制器單元測試**

在 `tests/schedule-picker.test.cjs` 新增 UI 狀態連動測試：

```javascript
test('UI 控制器產生之預設輸出與單位切換對齊規範', () => {
  // 模擬天切換為週
  const weekOutput = formatTrackSchedule('週', '02:00', { weekday: '日' });
  assert.equal(weekOutput, '每週日 02:00');

  // 模擬切換為月 (月底)
  const monthOutput = formatTrackSchedule('個月', '00:00', { day: '月底最後一日' });
  assert.equal(monthOutput, '每月月底最後一日 00:00');
});
```

- [x] **Step 2: 在 `Index.html` 更新 `renderFormTargets` 渲染複合選擇器**

1. 將原本的單一文字框 `<input type="text" id="localScheduleTime_${target.id}">` 替換為：
   - 完整備份軌選擇器容器
   - 差異備份軌選擇器容器（有差異備份時顯示）
   - 隱藏的 input: `<input type="hidden" id="localScheduleTime_${target.id}">`
   - 即時預覽標籤: `<div id="localSchedulePreview_${target.id}">`
2. 異地備份區塊內替換為異地專屬選擇器容器與隱藏 input。
3. 日誌檔標的渲染固定的天週期時間選擇器。

- [x] **Step 3: 實作週期單位監聽與即時同步控制器**

1. 在 `renderFormTargets` 渲染後綁定監聽器：
   - `fullUnit_${target.id}` 變更時呼叫 `updateSchedulePickerUI(target.id, 'full', this.value)`
   - `diffUnit_${target.id}` 與 `diffValue_${target.id}` 變更時連動差異軌可見度與單位
   - `offsiteFreqUnit_${target.id}` 變更時呼叫 `updateSchedulePickerUI(target.id, 'offsite', this.value)`
2. 任何時間、星期、日期或月份下拉變更時，觸發 `syncTargetSchedule(target.id)`。

- [x] **Step 4: 語法檢查與測試驗證**

Run:
```bash
node --test tests/schedule-picker.test.cjs
node -e '
const fs = require("fs");
const html = fs.readFileSync("Index.html", "utf8");
const match = html.match(/<script>([\s\S]*?)<\/script>/i);
fs.writeFileSync("/tmp/syntax_check.js", match[1]);
' && node --check /tmp/syntax_check.js
```
Expected: PASS with 0 errors.

- [x] **Step 5: Commit**

```bash
git add Index.html tests/schedule-picker.test.cjs
git commit -m "feat: add composite schedule picker UI and dynamic unit controller"
```

---

### Task 3: 表單生命週期整合（編輯載入 Hydration、表單重置 Reset、送出驗證）與全系統驗證

**Files:**
- Modify: `Index.html` (更新 `editSystem`、`resetForm`、`handleFormSubmit`)

**Interfaces:**
- Consumes: Task 1 的 `parseScheduleString` 與 Task 2 的 `syncTargetSchedule`
- Produces: 完整的端到端表單生命週期與歷史資料無痕洗淨

- [x] **Step 1: 編寫生命週期整合測試**

在 `tests/schedule-picker.test.cjs` 驗證端到端情境：

```javascript
test('表單編輯時非標準舊資料正確反向解析並正規化', () => {
  // 測試舊資料 "每日 02:00" 載入至週週期時，提取 02:00 並補齊每週日
  const parsed = parseScheduleString('每日 02:00', '週');
  const normalized = formatTrackSchedule('週', parsed.full.time, { weekday: parsed.full.weekday });
  assert.equal(normalized, '每週日 02:00');
});
```

- [x] **Step 2: 更新 `editSystem(id)` 回填邏輯**

在 `editSystem(id)` 中：
1. 讀取既有 `tData.localScheduleTime`，呼叫 `parseScheduleString(tData.localScheduleTime, tData.localFullUnit)`。
2. 將解出的時間、星期、月份、日期設定回完整軌及差異軌之 DOM 元件。
3. 呼叫 `syncTargetSchedule(tid)` 重組並更新預覽。
4. 異地排程時間以相同邏輯解構回填。

- [x] **Step 3: 更新 `resetForm()` 重置邏輯**

在 `resetForm()` 中：
1. 重置各選擇器為標準預設值：
   - 星期：`日`
   - 日期：`1`（若是月底則為 `1`）
   - 月份：`12`
   - 時間：`00:00`
2. 本地排程隱藏值恢復為 `每日 00:00`。
3. 異地排程隱藏值恢復為 `每日 00:00`。

- [x] **Step 4: 全流程自動化測試與驗證**

Run:
```bash
node --test tests/schedule-picker.test.cjs
node --check code.js
```
Expected: PASS (All test suites pass, 0 syntax errors).

- [x] **Step 5: Commit**

```bash
git add Index.html tests/schedule-picker.test.cjs
git commit -m "feat: complete schedule picker lifecycle integration with hydration and reset"
```

---

## Verification Plan

### Automated Tests
1. **單元測試**：
   ```bash
   node --test tests/schedule-picker.test.cjs
   ```
   驗證各週期單位輸出格式、雙軌合成與非標準舊值正規化。
2. **語法合規測試**：
   ```bash
   node --check code.js
   ```

### Manual Verification
1. **動態切換驗證**：進入表單展開標的，將本地完整週期由「天」切換至「週」，確認介面即時出現「星期」下拉選單；切換至「個月」確認出現「1~31 日與月底最後一日」選單；切換至「版本」確認出現「每次版本發布」標籤。
2. **雙軌合成驗證**：差異備份設定 1 天 02:00，完整備份設定 1 週 03:00（每週日），確認即時預覽顯示 `完整:每週日 03:00 / 差異:每日 02:00`。
3. **單軌降級驗證**：將差異備份改為 0（或無），確認即時預覽自動變更為純單軌 `每週日 03:00`。
4. **送出與儀表板驗證**：送出表單，確認儀表板卡片詳細資訊正確顯示新合成之排程格式；檢視 Google Sheet 確認第 11 欄值完全吻合標準格式。
5. **編輯回填與洗淨驗證**：點擊卡片編輯，確認星期、日期、時間等控制項精確回填；重新送出確認資料完整無損。
