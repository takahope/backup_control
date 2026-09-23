const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const vm = require('vm');

// 載入 Index.html 內部的核心函式與控制器
const html = fs.readFileSync('Index.html', 'utf8');
const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/i);
if (!scriptMatch) throw new Error('Index.html script tag not found');

let currentMockDom = null;

function createMockDom() {
  const elements = new Map();

  function getOrCreate(id) {
    if (!elements.has(id)) {
      const element = {
        id,
        value: '',
        textContent: '',
        innerText: '',
        style: {},
        checked: false,
        disabled: false,
        required: false,
        reset() {
          this.value = '';
          this.checked = false;
        },
        options: [],
        appendChild(option) {
          option.parentNode = this;
          option.remove = function() {
            if (!this.parentNode) return;
            const index = this.parentNode.options.indexOf(this);
            if (index !== -1) this.parentNode.options.splice(index, 1);
          };
          this.options.push(option);
          return option;
        },
        querySelector: (sel) => getOrCreate('mock_child_' + sel),
        querySelectorAll(selector) {
          if (selector === 'option[data-legacy-option="true"]') {
            return this.options.filter(option => option.dataset && option.dataset.legacyOption === 'true');
          }
          return [];
        },
        classList: {
          _classes: new Set(),
          add(c) { this._classes.add(c); },
          remove(c) { this._classes.delete(c); },
          toggle(c, force) {
            if (force !== undefined) {
              if (force) this.add(c); else this.remove(c);
            } else {
              if (this.contains(c)) this.remove(c); else this.add(c);
            }
          },
          contains(c) { return this._classes.has(c); }
        }
      };
      elements.set(id, element);
    }
    return elements.get(id);
  }

  return {
    elements,
    getElementById: (id) => getOrCreate(id),
    querySelector: (selector) => {
      const checkedMatch = selector.match(/input\[name="([^"]+)"\]:checked/);
      if (checkedMatch) {
        const radioName = checkedMatch[1];
        const yesEl = getOrCreate(`${radioName}_val_是`);
        const noEl = getOrCreate(`${radioName}_val_否`);
        if (yesEl.checked) return yesEl;
        if (noEl.checked) return noEl;
        return noEl;
      }
      const valMatch = selector.match(/input\[name="([^"]+)"\]\[value="([^"]+)"\]/);
      if (valMatch) {
        const el = getOrCreate(`${valMatch[1]}_val_${valMatch[2]}`);
        el.value = valMatch[2];
        return el;
      }
      return getOrCreate(`mock_qs_${selector}`);
    },
    querySelectorAll: () => [],
    register: (id, initial = {}) => {
      const el = getOrCreate(id);
      Object.assign(el, initial);
      if (initial.classes) {
        initial.classes.forEach(c => el.classList.add(c));
      }
      return el;
    }
  };
}

const sandbox = {
  localStorage: { getItem: () => null, setItem: () => {} },
  window: { scrollTo: () => {} },
  document: {
    addEventListener: () => {},
    createElement: () => ({ value: '', textContent: '', dataset: {}, parentNode: null }),
    getElementById: (id) => currentMockDom ? currentMockDom.getElementById(id) : null,
    querySelector: (s) => currentMockDom ? currentMockDom.querySelector(s) : null,
    querySelectorAll: (s) => currentMockDom ? currentMockDom.querySelectorAll(s) : []
  },
  alert: () => {},
  setTimeout: (fn, ms) => setTimeout(fn, Math.min(ms || 0, 5)),
  clearTimeout,
  console
};
vm.createContext(sandbox);
vm.runInContext(scriptMatch[1], sandbox);

const {
  formatTrackSchedule,
  composeDualTrackSchedule,
  parseScheduleString,
  renderSchedulePickerTrackHtml,
  updateSchedulePickerUI,
  updateDiffPickerVisibility,
  syncTargetSchedule,
  renderTargetSetting,
  renderFormTargets,
  setChoiceSelectValue_,
  handleChoiceSelectChange_,
  resetChoiceSelect_,
  editSystem,
  resetForm,
  handleFormSubmit
} = sandbox;

// --- Task 1 核心語法測試 ---

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

// --- Task 2 前端複合選擇器 HTML 與動態控制器測試 ---

test('renderSchedulePickerTrackHtml 產出符合規範的 DOM 元件結構', () => {
  assert.equal(typeof renderSchedulePickerTrackHtml, 'function');

  // 測試完整軌
  const fullHtml = renderSchedulePickerTrackHtml('db', 'full', '完整排程:');
  assert.ok(fullHtml.includes('id="fullSchedulePicker_db"'), '應包含 track 容器 ID');
  assert.ok(fullHtml.includes('id="fullSchedPrefix_db"'), '應包含前綴 ID');
  assert.ok(fullHtml.includes('id="fullSchedMonth_db"'), '應包含月份下拉選單 ID');
  assert.ok(fullHtml.includes('id="fullSchedDay_db"'), '應包含日期下拉選單 ID');
  assert.ok(fullHtml.includes('id="fullSchedWeekday_db"'), '應包含星期下拉選單 ID');
  assert.ok(fullHtml.includes('id="fullSchedTime_db"'), '應包含時間選單 ID');
  assert.ok(fullHtml.includes('id="fullSchedVersionTag_db"'), '應包含版本標籤 ID');
  assert.ok(fullHtml.includes('月底最後一日'), '日期選單應包含月底最後一日');
  assert.ok(!fullHtml.includes('schedule-picker-track flex flex-wrap items-center gap-2 hidden'), '完整軌預設不應隱藏');

  // 測試差異軌
  const diffHtml = renderSchedulePickerTrackHtml('db', 'diff', '差異排程:');
  assert.ok(diffHtml.includes('id="diffSchedulePicker_db"'), '應包含差異軌容器 ID');
  assert.ok(diffHtml.includes('hidden'), '差異軌預設容器應包含 hidden');

  // 測試異地軌
  const offsiteHtml = renderSchedulePickerTrackHtml('db', 'offsite', '異地排程:');
  assert.ok(offsiteHtml.includes('id="offsiteSchedulePicker_db"'), '應包含異地軌容器 ID');
  assert.ok(offsiteHtml.includes('text-purple-700'), '異地軌應有紫色樣式');
});

test('renderFormTargets 將本地與異地處置及佐證渲染為相同的固定下拉選單', () => {
  currentMockDom = createMockDom();
  renderFormTargets();
  const rendered = currentMockDom.getElementById('targetsContainer').innerHTML;

  function getSelectHtml(id) {
    const match = rendered.match(new RegExp(`<select[^>]*id="${id}"[^>]*>([\\s\\S]*?)<\\/select>`));
    assert.ok(match, `${id} 必須使用 select 元件`);
    return match[1];
  }

  const localDisposition = getSelectHtml('localDispositionMethod_db');
  const offsiteDisposition = getSelectHtml('offsiteDispositionMethod_db');
  const localEvidence = getSelectHtml('localEvidenceRequirements_db');
  const offsiteEvidence = getSelectHtml('offsiteEvidenceRequirements_db');

  [localDisposition, offsiteDisposition].forEach(selectHtml => {
    assert.deepEqual(
      Array.from(selectHtml.matchAll(/<option value="([^"]+)"/g), match => match[1]),
      ['磁碟自動覆寫', '人工手動覆蓋']
    );
    assert.match(selectHtml, /^<option value="磁碟自動覆寫" selected>/);
  });
  [localEvidence, offsiteEvidence].forEach(selectHtml => {
    assert.deepEqual(
      Array.from(selectHtml.matchAll(/<option value="([^"]+)"/g), match => match[1]),
      ['備份系統執行Log', '備份檔名確認']
    );
    assert.match(selectHtml, /^<option value="備份系統執行Log" selected>/);
  });
  assert.match(rendered, /onchange="handleChoiceSelectChange_\(this\)"/);
});

test('選擇欄位可暫存歷史舊值並在重設時移除', () => {
  assert.equal(typeof setChoiceSelectValue_, 'function');
  assert.equal(typeof resetChoiceSelect_, 'function');

  currentMockDom = createMockDom();
  const select = currentMockDom.getElementById('offsiteEvidenceRequirements_db');
  select.appendChild({ value: '備份系統執行Log', textContent: '備份系統執行Log', dataset: {} });
  select.appendChild({ value: '備份檔名確認', textContent: '備份檔名確認', dataset: {} });

  setChoiceSelectValue_(select, '磁帶消磁簽收單 / 雲端過期報表', '備份系統執行Log');
  assert.equal(select.value, '磁帶消磁簽收單 / 雲端過期報表');
  assert.equal(select.options.length, 3);
  assert.equal(select.options[2].dataset.legacyOption, 'true');

  handleChoiceSelectChange_(select);
  assert.equal(select.options.length, 3, '仍選取歷史值時不得移除舊值選項');

  select.value = '備份檔名確認';
  handleChoiceSelectChange_(select);
  assert.equal(select.value, '備份檔名確認');
  assert.deepEqual(select.options.map(option => option.value), ['備份系統執行Log', '備份檔名確認']);

  setChoiceSelectValue_(select, '磁帶消磁簽收單 / 雲端過期報表', '備份系統執行Log');

  resetChoiceSelect_(select, '備份系統執行Log');
  assert.equal(select.value, '備份系統執行Log');
  assert.deepEqual(select.options.map(option => option.value), ['備份系統執行Log', '備份檔名確認']);
});

test('異地佐證缺值時與本地共用備份系統執行Log預設值', () => {
  const rendered = renderTargetSetting({
    targetKey: 'db',
    targetName: '資料庫',
    isOffsite: '是',
    offsiteEvidenceRequirements: ''
  });

  assert.match(rendered, /異地佐證：<\/span>備份系統執行Log/);
  assert.equal(rendered.includes('磁帶消磁簽收單 / 雲端過期報表'), false);
});

test('updateSchedulePickerUI 依週期單位動態切換各欄位顯示與隱藏', () => {
  assert.equal(typeof updateSchedulePickerUI, 'function');

  currentMockDom = createMockDom();
  const prefixEl = currentMockDom.register('fullSchedPrefix_db', { textContent: '' });
  const monthEl = currentMockDom.register('fullSchedMonth_db', { classes: ['hidden'] });
  const dayEl = currentMockDom.register('fullSchedDay_db', { classes: ['hidden'] });
  const weekdayEl = currentMockDom.register('fullSchedWeekday_db', { classes: ['hidden'] });
  const timeEl = currentMockDom.register('fullSchedTime_db');
  const versionTagEl = currentMockDom.register('fullSchedVersionTag_db', { classes: ['hidden'] });

  // 1. 切換為「天」
  updateSchedulePickerUI('db', 'full', '天');
  assert.equal(prefixEl.textContent, '每日');
  assert.equal(prefixEl.classList.contains('hidden'), false);
  assert.equal(timeEl.classList.contains('hidden'), false);
  assert.equal(monthEl.classList.contains('hidden'), true);
  assert.equal(dayEl.classList.contains('hidden'), true);
  assert.equal(weekdayEl.classList.contains('hidden'), true);
  assert.equal(versionTagEl.classList.contains('hidden'), true);

  // 2. 切換為「週」
  updateSchedulePickerUI('db', 'full', '週');
  assert.equal(prefixEl.classList.contains('hidden'), true);
  assert.equal(weekdayEl.classList.contains('hidden'), false);
  assert.equal(timeEl.classList.contains('hidden'), false);
  assert.equal(monthEl.classList.contains('hidden'), true);
  assert.equal(dayEl.classList.contains('hidden'), true);
  assert.equal(versionTagEl.classList.contains('hidden'), true);

  // 3. 切換為「個月」
  updateSchedulePickerUI('db', 'full', '個月');
  assert.equal(prefixEl.textContent, '每月');
  assert.equal(prefixEl.classList.contains('hidden'), false);
  assert.equal(dayEl.classList.contains('hidden'), false);
  assert.equal(timeEl.classList.contains('hidden'), false);
  assert.equal(monthEl.classList.contains('hidden'), true);
  assert.equal(weekdayEl.classList.contains('hidden'), true);
  assert.equal(versionTagEl.classList.contains('hidden'), true);

  // 4. 切換為「年」
  updateSchedulePickerUI('db', 'full', '年');
  assert.equal(prefixEl.textContent, '每年');
  assert.equal(prefixEl.classList.contains('hidden'), false);
  assert.equal(monthEl.classList.contains('hidden'), false);
  assert.equal(dayEl.classList.contains('hidden'), false);
  assert.equal(timeEl.classList.contains('hidden'), false);
  assert.equal(weekdayEl.classList.contains('hidden'), true);
  assert.equal(versionTagEl.classList.contains('hidden'), true);

  // 5. 切換為「版本」
  updateSchedulePickerUI('db', 'full', '版本');
  assert.equal(prefixEl.classList.contains('hidden'), true);
  assert.equal(versionTagEl.classList.contains('hidden'), false);
  assert.equal(timeEl.classList.contains('hidden'), true);
  assert.equal(monthEl.classList.contains('hidden'), true);
  assert.equal(dayEl.classList.contains('hidden'), true);
  assert.equal(weekdayEl.classList.contains('hidden'), true);
});

test('updateDiffPickerVisibility 依差異備份數值（>0）動態切換差異軌可見度', () => {
  assert.equal(typeof updateDiffPickerVisibility, 'function');

  currentMockDom = createMockDom();
  const diffValEl = currentMockDom.register('diffValue_db', { value: '0' });
  const diffPicker = currentMockDom.register('diffSchedulePicker_db', { classes: ['hidden'] });
  // 為避免 syncTargetSchedule 存取不存在元素，一併建立對應元素
  currentMockDom.register('fullUnit_db', { value: '天' });
  currentMockDom.register('fullSchedTime_db', { value: '00:00' });
  currentMockDom.register('localScheduleTime_db', { value: '' });
  currentMockDom.register('localSchedulePreview_db', { textContent: '' });

  // 數值為 0 時，差異軌維持 hidden
  updateDiffPickerVisibility('db');
  assert.equal(diffPicker.classList.contains('hidden'), true);

  // 數值 > 0 時，差異軌移除 hidden
  diffValEl.value = '1';
  updateDiffPickerVisibility('db');
  assert.equal(diffPicker.classList.contains('hidden'), false);

  // 數值改為空值或 <=0 時，重新加上 hidden
  diffValEl.value = '';
  updateDiffPickerVisibility('db');
  assert.equal(diffPicker.classList.contains('hidden'), true);
});

test('syncTargetSchedule 正確合成排程字串並更新隱藏 input 與預覽標籤', () => {
  assert.equal(typeof syncTargetSchedule, 'function');

  currentMockDom = createMockDom();
  // 設置一般標的 db
  currentMockDom.register('fullUnit_db', { value: '週' });
  currentMockDom.register('fullSchedWeekday_db', { value: '日' });
  currentMockDom.register('fullSchedTime_db', { value: '02:00' });

  currentMockDom.register('diffValue_db', { value: '0' });
  currentMockDom.register('diffUnit_db', { value: '天' });
  currentMockDom.register('diffSchedTime_db', { value: '00:00' });

  const localInput = currentMockDom.register('localScheduleTime_db', { value: '' });
  const localPreview = currentMockDom.register('localSchedulePreview_db', { textContent: '' });

  currentMockDom.register('offsiteFreqUnit_db', { value: '天' });
  currentMockDom.register('offsiteSchedTime_db', { value: '03:00' });
  const offsiteInput = currentMockDom.register('offsiteScheduleTime_db', { value: '' });
  const offsitePreview = currentMockDom.register('offsiteSchedulePreview_db', { textContent: '' });

  // 1. 單軌情境 (無差異備份)
  syncTargetSchedule('db');
  assert.equal(localInput.value, '每週日 02:00');
  assert.equal(localPreview.textContent, '排程設定值: 每週日 02:00');
  assert.equal(offsiteInput.value, '每日 03:00');
  assert.equal(offsitePreview.textContent, '排程設定值: 每日 03:00');

  // 2. 雙軌情境 (差異備份 > 0)
  currentMockDom.register('diffValue_db', { value: '1' });
  currentMockDom.register('diffUnit_db', { value: '天' });
  currentMockDom.register('diffSchedTime_db', { value: '01:00' });
  syncTargetSchedule('db');
  assert.equal(localInput.value, '完整:每週日 02:00 / 差異:每日 01:00');
  assert.equal(localPreview.textContent, '排程設定值: 完整:每週日 02:00 / 差異:每日 01:00');

  // 3. 系統日誌檔 (log) 固定天週期
  currentMockDom.register('fullSchedTime_log', { value: '04:15' });
  const logLocalInput = currentMockDom.register('localScheduleTime_log', { value: '' });
  const logLocalPreview = currentMockDom.register('localSchedulePreview_log', { textContent: '' });
  currentMockDom.register('offsiteFreqUnit_log', { value: '個月' });
  currentMockDom.register('offsiteSchedDay_log', { value: '15' });
  currentMockDom.register('offsiteSchedTime_log', { value: '05:00' });
  const logOffsiteInput = currentMockDom.register('offsiteScheduleTime_log', { value: '' });
  const logOffsitePreview = currentMockDom.register('offsiteSchedulePreview_log', { textContent: '' });

  syncTargetSchedule('log');
  assert.equal(logLocalInput.value, '每日 04:15');
  assert.equal(logLocalPreview.textContent, '排程設定值: 每日 04:15');
  assert.equal(logOffsiteInput.value, '每月 15 日 05:00');
  assert.equal(logOffsitePreview.textContent, '排程設定值: 每月 15 日 05:00');
});

// --- Task 3 表單生命週期整合測試 ---

test('表單編輯回填歷史非標準字串之自動解析與正規化', () => {
  currentMockDom = createMockDom();

  // 模擬載入含有非標準舊字串之歷史系統資料
  vm.runInContext(`appData = [{
    id: "sys-legacy",
    systemName: "歷史舊系統",
    systemLevel: "關鍵核心系統",
    managerName: "王小明",
    targets: [
      {
        targetKey: "db",
        targetName: "資料庫",
        localDiffValue: "0",
        localDiffUnit: "天",
        localFullValue: "1",
        localFullUnit: "週",
        localRetention: "3",
        localLocation: "機房 A 伺服器",
        localScheduleTime: "凌晨 03:00 執行備份",
        localDispositionMethod: "磁碟自動覆寫",
        localEvidenceRequirements: "備份系統執行Log",
        isOffsite: "是",
        offsiteFreqValue: "1",
        offsiteFreqUnit: "天",
        offsiteLocation: "雲端備份中心",
        offsiteRetention: "7",
        offsiteScheduleTime: "2026/01/01 02:30",
        offsiteDispositionMethod: "磁碟自動覆寫",
        offsiteEvidenceRequirements: "磁帶消磁簽收單 / 雲端過期報表"
      },
      {
        targetKey: "log",
        targetName: "系統日誌檔",
        logRetention: "保存 6 個月",
        localLocation: "Log 伺服器",
        localScheduleTime: "每日凌晨 04:15 執行",
        localDispositionMethod: "磁碟自動覆寫",
        localEvidenceRequirements: "備份系統執行Log",
        isOffsite: "否",
        offsiteFreqValue: "",
        offsiteFreqUnit: "",
        offsiteLocation: "",
        offsiteRetention: "",
        offsiteScheduleTime: "無",
        offsiteDispositionMethod: "無",
        offsiteEvidenceRequirements: "無"
      }
    ]
  }];`, sandbox);

  editSystem('sys-legacy');

  // 1. 一般標的 (db - 週週期) 回填與正規化檢驗
  assert.equal(currentMockDom.getElementById('fullSchedWeekday_db').value, '日');
  assert.equal(currentMockDom.getElementById('fullSchedTime_db').value, '03:00');
  assert.equal(currentMockDom.getElementById('localScheduleTime_db').value, '每週日 03:00');
  assert.equal(currentMockDom.getElementById('localSchedulePreview_db').textContent, '排程設定值: 每週日 03:00');

  // 異地 (db - 天週期)
  assert.equal(currentMockDom.getElementById('offsiteSchedTime_db').value, '02:30');
  assert.equal(currentMockDom.getElementById('offsiteScheduleTime_db').value, '每日 02:30');
  assert.equal(currentMockDom.getElementById('offsiteSchedulePreview_db').textContent, '排程設定值: 每日 02:30');

  // 2. 日誌檔標的 (log - 固定天週期) 回填與正規化檢驗
  assert.equal(currentMockDom.getElementById('fullSchedTime_log').value, '04:15');
  assert.equal(currentMockDom.getElementById('localScheduleTime_log').value, '每日 04:15');
  assert.equal(currentMockDom.getElementById('localSchedulePreview_log').textContent, '排程設定值: 每日 04:15');
  assert.equal(currentMockDom.getElementById('offsiteScheduleTime_log').value, '無');
  assert.equal(currentMockDom.getElementById('offsiteSchedulePreview_log').textContent, '排程設定值: 無');
});

test('表單編輯回填雙軌排程完整回填與狀態同步', () => {
  currentMockDom = createMockDom();

  vm.runInContext(`appData = [{
    id: "sys-dual",
    systemName: "雙軌系統",
    systemLevel: "關鍵核心系統",
    managerName: "李大同",
    targets: [
      {
        targetKey: "db",
        targetName: "資料庫",
        localDiffValue: "1",
        localDiffUnit: "天",
        localFullValue: "1",
        localFullUnit: "週",
        localRetention: "4",
        localLocation: "正式機房",
        localScheduleTime: "完整:每週日 02:00 / 差異:每日 01:30",
        localDispositionMethod: "磁碟自動覆寫",
        localEvidenceRequirements: "備份系統執行Log",
        isOffsite: "否",
        offsiteScheduleTime: "無"
      }
    ]
  }];`, sandbox);

  editSystem('sys-dual');

  // 驗證完整軌
  assert.equal(currentMockDom.getElementById('fullSchedWeekday_db').value, '日');
  assert.equal(currentMockDom.getElementById('fullSchedTime_db').value, '02:00');

  // 驗證差異軌
  assert.equal(currentMockDom.getElementById('diffSchedWeekday_db').value, '日');
  assert.equal(currentMockDom.getElementById('diffSchedTime_db').value, '01:30');
  assert.equal(currentMockDom.getElementById('diffSchedulePicker_db').classList.contains('hidden'), false);

  // 驗證合成字串
  assert.equal(currentMockDom.getElementById('localScheduleTime_db').value, '完整:每週日 02:00 / 差異:每日 01:30');
  assert.equal(currentMockDom.getElementById('localSchedulePreview_db').textContent, '排程設定值: 完整:每週日 02:00 / 差異:每日 01:30');
  assert.equal(currentMockDom.getElementById('offsiteScheduleTime_db').value, '無');
});

test('表單重置 resetForm 後所有標的與軌道恢復初始預設值', () => {
  currentMockDom = createMockDom();

  // 先手動填寫污染各元件狀態
  currentMockDom.getElementById('fullSchedTime_db').value = '15:00';
  currentMockDom.getElementById('diffSchedTime_db').value = '16:00';
  currentMockDom.getElementById('diffValue_db').value = '2';
  currentMockDom.getElementById('localScheduleTime_db').value = '非預設排程';
  currentMockDom.getElementById('offsiteScheduleTime_db').value = '每週五 03:00';
  currentMockDom.getElementById('fullSchedTime_log').value = '22:00';

  resetForm();

  // 檢驗 db 控制項恢復預設
  assert.equal(currentMockDom.getElementById('fullSchedTime_db').value, '00:00');
  assert.equal(currentMockDom.getElementById('fullSchedWeekday_db').value, '日');
  assert.equal(currentMockDom.getElementById('fullSchedMonth_db').value, '12');
  assert.equal(currentMockDom.getElementById('fullSchedDay_db').value, '1');
  assert.equal(currentMockDom.getElementById('diffSchedTime_db').value, '00:00');
  assert.equal(currentMockDom.getElementById('diffSchedulePicker_db').classList.contains('hidden'), true);
  assert.equal(currentMockDom.getElementById('localScheduleTime_db').value, '每日 00:00');
  assert.equal(currentMockDom.getElementById('offsiteScheduleTime_db').value, '無');
  assert.equal(currentMockDom.getElementById('offsiteSchedulePreview_db').textContent, '排程設定值: 無');

  // 檢驗 log 控制項恢復預設
  assert.equal(currentMockDom.getElementById('fullSchedTime_log').value, '00:00');
  assert.equal(currentMockDom.getElementById('localScheduleTime_log').value, '每日 00:00');
  assert.equal(currentMockDom.getElementById('offsiteScheduleTime_log').value, '無');
});

test('handleFormSubmit 送出前強制執行 syncTargetSchedule 確保所有標的排程字串更新', () => {
  currentMockDom = createMockDom();

  // 勾選 db 標的
  currentMockDom.getElementById('check_db').checked = true;
  currentMockDom.getElementById('fullUnit_db').value = '天';
  // 模擬使用者直接變更 time input 尚未 trigger 任何 event
  currentMockDom.getElementById('fullSchedTime_db').value = '05:45';
  currentMockDom.getElementById('localScheduleTime_db').value = '舊排程未同步';
  currentMockDom.getElementById('offsite_db_val_否').checked = true;
  currentMockDom.getElementById('offsite_db_val_否').value = '否';

  let submittedFormData = null;
  const run = {
    withSuccessHandler() { return this; },
    withFailureHandler() { return this; },
    submitData(formData) { submittedFormData = formData; }
  };
  sandbox.google = { script: { run } };

  let prevented = false;
  handleFormSubmit({
    preventDefault: () => { prevented = true; }
  });
  delete sandbox.google;

  assert.equal(prevented, true);
  // handleFormSubmit 內部應呼叫 syncTargetSchedule('db')，使 localScheduleTime_db 更新為最新格式
  assert.equal(currentMockDom.getElementById('localScheduleTime_db').value, '每日 05:45');
  assert.equal(submittedFormData.targets[0].offsiteDispositionMethod, '無');
  assert.equal(submittedFormData.targets[0].offsiteEvidenceRequirements, '無');
});
