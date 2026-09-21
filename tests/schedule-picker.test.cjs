const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const vm = require('vm');

// 載入 Index.html 內部的核心純函式
const html = fs.readFileSync('Index.html', 'utf8');
const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/i);
if (!scriptMatch) throw new Error('Index.html script tag not found');

const sandbox = {
  localStorage: { getItem: () => null, setItem: () => {} },
  window: {},
  document: { addEventListener: () => {}, getElementById: () => null },
  console
};
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
