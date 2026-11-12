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
        _innerHTML: '',
        get innerHTML() {
          return this._innerHTML;
        },
        set innerHTML(value) {
          this._innerHTML = value;
          if (value === '') this.options = [];
        },
        textContent: '',
        innerText: '',
        style: {},
        checked: false,
        disabled: false,
        required: false,
        setAttribute(name, value) {
          this[name] = value;
        },
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
  updateFullPickerVisibility,
  toggleLocalRetentionMode,
  toggleTarget,
  validateLocalBackupCycles,
  isCycleValueEnabled_,
  syncTargetSchedule,
  renderTargetSetting,
  renderFormTargets,
  setChoiceSelectValue_,
  handleChoiceSelectChange_,
  resetChoiceSelect_,
  handleSystemNameInput,
  renderSourceOptionsDatalists,
  renderStorageAssetOptionsDatalist,
  validateLocalLocationSelection,
  editSystem,
  resetForm,
  handleFormSubmit
} = sandbox;

// --- 資料來源與備份儲存設備選單責任分離測試 ---

test('來源資訊資產依系統屬性篩選，但備份儲存設備永遠包含全部資訊資產', () => {
  assert.equal(typeof renderSourceOptionsDatalists, 'function');
  assert.equal(typeof renderStorageAssetOptionsDatalist, 'function');

  currentMockDom = createMockDom();
  currentMockDom.getElementById('systemLevel').value = '一般系統';
  currentMockDom.querySelector = () => ({ value: 'assets' });
  vm.runInContext(`
    systemOptionsData = [];
    assetOptionsData = [
      { name: "AST-LOW - 一般儲存設備", level: "一般系統" },
      { name: "AST-HIGH - 核心儲存設備", level: "關鍵核心系統" }
    ];
  `, sandbox);

  renderSourceOptionsDatalists();
  renderStorageAssetOptionsDatalist();

  assert.deepEqual(
    currentMockDom.getElementById('sourceAssetOptionsList').options.map(option => option.value),
    ['AST-LOW - 一般儲存設備']
  );
  assert.deepEqual(
    currentMockDom.getElementById('storageAssetOptionsList').options.map(option => option.value),
    ['AST-LOW - 一般儲存設備', 'AST-HIGH - 核心儲存設備']
  );
  assert.equal(currentMockDom.getElementById('systemName').list, 'sourceAssetOptionsList');

  currentMockDom.getElementById('systemLevel').value = '關鍵核心系統';
  renderSourceOptionsDatalists();
  assert.deepEqual(
    currentMockDom.getElementById('sourceAssetOptionsList').options.map(option => option.value),
    ['AST-HIGH - 核心儲存設備']
  );
  assert.deepEqual(
    currentMockDom.getElementById('storageAssetOptionsList').options.map(option => option.value),
    ['AST-LOW - 一般儲存設備', 'AST-HIGH - 核心儲存設備'],
    '切換系統屬性不得縮限備份儲存設備清單'
  );
});

test('一般系統可選擇資產價值較高的資訊資產作為本地備份設備', () => {
  currentMockDom = createMockDom();
  currentMockDom.getElementById('localLoc_db').value = 'AST-HIGH - 核心儲存設備';
  vm.runInContext(`
    assetOptionsLoadState = "ready";
    assetOptionsData = [
      { name: "AST-HIGH - 核心儲存設備", level: "關鍵核心系統" }
    ];
  `, sandbox);

  assert.equal(validateLocalLocationSelection([{ id: 'db', name: '資料庫' }]), true);
});

test('資訊資產清單載入失敗或沒有資料時必須阻止送出', () => {
  currentMockDom = createMockDom();
  const alerts = [];
  sandbox.alert = message => alerts.push(message);

  vm.runInContext(`assetOptionsLoadState = "error"; assetOptionsData = [];`, sandbox);
  assert.equal(validateLocalLocationSelection([{ id: 'db', name: '資料庫' }]), false);
  assert.match(alerts.pop(), /資訊資產清單載入失敗/);

  vm.runInContext(`assetOptionsLoadState = "ready"; assetOptionsData = [];`, sandbox);
  assert.equal(validateLocalLocationSelection([{ id: 'db', name: '資料庫' }]), false);
  assert.match(alerts.pop(), /沒有可選資料/);
});

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

  // 僅差異軌：必須帶「差異:」前綴，避免被誤讀為完整軌
  const diffOnly = composeDualTrackSchedule({
    hasDiff: true,
    fullSchedule: '',
    diffSchedule: '每日 02:00'
  });
  assert.equal(diffOnly, '差異:每日 02:00');

  // 明確標示無完整備份（即使完整軌仍有殘留字串）
  const noFullExplicit = composeDualTrackSchedule({
    hasFull: false,
    hasDiff: true,
    fullSchedule: '每週日 02:00',
    diffSchedule: '每日 01:00'
  });
  assert.equal(noFullExplicit, '差異:每日 01:00');

  // 兩軌皆未啟用時維持預設值
  const neither = composeDualTrackSchedule({ hasFull: false, hasDiff: false });
  assert.equal(neither, '每日 00:00');
});

test('isCycleValueEnabled_ 僅數值 > 0 視為啟用，空白、0、N/A、無 皆為未啟用', () => {
  assert.equal(isCycleValueEnabled_('1'), true);
  assert.equal(isCycleValueEnabled_(2), true);
  assert.equal(isCycleValueEnabled_('0'), false);
  assert.equal(isCycleValueEnabled_(''), false);
  assert.equal(isCycleValueEnabled_(undefined), false);
  assert.equal(isCycleValueEnabled_(null), false);
  assert.equal(isCycleValueEnabled_('N/A'), false);
  assert.equal(isCycleValueEnabled_('無'), false);
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

  // 5. 僅差異備份格式：回填到差異軌，且不被誤判為雙軌
  const diffOnlyParsed = parseScheduleString('差異:每週三 01:30', '週');
  assert.equal(diffOnlyParsed.isDual, false);
  assert.equal(diffOnlyParsed.isDiffOnly, true);
  assert.equal(diffOnlyParsed.diff.weekday, '三');
  assert.equal(diffOnlyParsed.diff.time, '01:30');

  // 6. 舊資料不帶前綴的單軌字串仍視為完整軌
  const plainParsed = parseScheduleString('每日 02:00', '天');
  assert.equal(plainParsed.isDiffOnly, undefined);
  assert.equal(plainParsed.full.time, '02:00');
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
  assert.ok(fullHtml.includes('schedule-picker-track flex flex-wrap items-center gap-2 hidden'), '完整備份為選填，完整軌預設應隱藏（週期 > 0 才顯示）');

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
  assert.ok(rendered.includes('到期銷毀方法'), '表單應包含「到期銷毀方法」標籤');
  assert.ok(rendered.includes('銷毀確認方式'), '表單應包含「銷毀確認方式」標籤');
  assert.equal(rendered.includes('本地處置方法'), false, '不應再含有舊名詞「本地處置方法」');
  assert.equal(rendered.includes('異地佐證種類'), false, '不應再含有舊名詞「異地佐證種類」');
  assert.match(rendered, /onchange="handleChoiceSelectChange_\(this\)"/);
});

test('renderTargetSetting 採用雙軌子面板視覺設計與去前綴化標籤', () => {
  // 1. 啟用異地備份情境
  const renderedOffsiteYes = renderTargetSetting({
    targetKey: 'db',
    targetName: '資料庫',
    localFullValue: '1',
    localFullUnit: '週',
    localDiffValue: '1',
    localDiffUnit: '天',
    localRetention: '4',
    localLocation: 'NAS-01',
    localScheduleTime: '每日 00:00',
    localDispositionMethod: '磁碟自動覆寫',
    localEvidenceRequirements: '備份系統執行Log',
    isOffsite: '是',
    offsiteFreqValue: '1',
    offsiteFreqUnit: '個月',
    offsiteLocation: '國網中心',
    offsiteRetention: '2',
    offsiteScheduleTime: '每週日 03:00',
    offsiteDispositionMethod: '磁碟自動覆寫',
    offsiteEvidenceRequirements: '備份系統執行Log'
  });

  assert.ok(renderedOffsiteYes.includes('本地備份'), '應有本地備份軌道');
  assert.ok(renderedOffsiteYes.includes('異地備份'), '應有異地備份軌道');
  assert.ok(renderedOffsiteYes.includes('到期銷毀方法：'), '應顯示「到期銷毀方法」');
  assert.ok(renderedOffsiteYes.includes('銷毀確認方式：'), '應顯示「銷毀確認方式」');
  assert.equal(renderedOffsiteYes.includes('本地處置：'), false, '去前綴化：不應有「本地處置：」');
  assert.equal(renderedOffsiteYes.includes('異地處置：'), false, '去前綴化：不應有「異地處置：」');
  assert.equal(renderedOffsiteYes.includes('本地完整：'), false, '去前綴化：不應有「本地完整：」');

  // 2. 未啟用異地備份情境
  const renderedOffsiteNo = renderTargetSetting({
    targetKey: 'source',
    targetName: '原始碼',
    isOffsite: '否'
  });

  assert.ok(renderedOffsiteNo.includes('本地備份'));
  assert.ok(renderedOffsiteNo.includes('未啟用異地備份'), '未啟用異地時應呈現未啟用徽章');
  assert.equal(renderedOffsiteNo.includes('異地週期：'), false);
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

  assert.match(rendered, /銷毀確認方式：<\/span>備份系統執行Log/);
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
  currentMockDom.register('fullValue_db', { value: '1' });
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

  // 2.1 僅差異備份情境 (完整備份 0 / 空白)：完整軌排程不得被寫入
  currentMockDom.register('fullValue_db', { value: '' });
  syncTargetSchedule('db');
  assert.equal(localInput.value, '差異:每日 01:00');
  assert.equal(localPreview.textContent, '排程設定值: 差異:每日 01:00');
  currentMockDom.register('fullValue_db', { value: '0' });
  syncTargetSchedule('db');
  assert.equal(localInput.value, '差異:每日 01:00');

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
    systemId: "SYS-LEGACY-01",
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
  assert.equal(currentMockDom.getElementById('submissionId').value, 'sys-legacy');
  assert.equal(currentMockDom.getElementById('sourceSystemId').value, 'SYS-LEGACY-01');
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

test('編輯時改選另一系統會同步更新業務系統 ID', () => {
  currentMockDom = createMockDom();
  currentMockDom.querySelector = () => ({ value: 'systems' });
  currentMockDom.getElementById('systemName').value = 'SYS-NEW-02 - 新系統';
  currentMockDom.getElementById('sourceSystemId').value = 'SYS-OLD-01';
  vm.runInContext(`systemOptionsData = [{ name: "SYS-NEW-02 - 新系統", level: "一般系統" }];`, sandbox);

  handleSystemNameInput();

  assert.equal(currentMockDom.getElementById('sourceSystemId').value, 'SYS-NEW-02');

  currentMockDom.getElementById('systemName').value = '自訂新系統';
  handleSystemNameInput();
  assert.equal(currentMockDom.getElementById('sourceSystemId').value, '', '輸入不在來源清單的新系統時不得沿用舊 ID');
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
  currentMockDom.getElementById('sourceSystemId').value = 'SYS-RESET-01';

  resetForm();

  // 檢驗 db 控制項恢復預設
  assert.equal(currentMockDom.getElementById('fullSchedTime_db').value, '00:00');
  assert.equal(currentMockDom.getElementById('fullSchedWeekday_db').value, '日');
  assert.equal(currentMockDom.getElementById('fullSchedMonth_db').value, '12');
  assert.equal(currentMockDom.getElementById('fullSchedDay_db').value, '1');
  assert.equal(currentMockDom.getElementById('diffSchedTime_db').value, '00:00');
  assert.equal(currentMockDom.getElementById('diffSchedulePicker_db').classList.contains('hidden'), true);
  assert.equal(currentMockDom.getElementById('fullValue_db').value, '');
  assert.equal(currentMockDom.getElementById('fullSchedulePicker_db').classList.contains('hidden'), true);
  assert.equal(currentMockDom.getElementById('localScheduleTime_db').value, '每日 00:00');
  assert.equal(currentMockDom.getElementById('offsiteScheduleTime_db').value, '無');
  assert.equal(currentMockDom.getElementById('offsiteSchedulePreview_db').textContent, '排程設定值: 無');

  // 檢驗 log 控制項恢復預設
  assert.equal(currentMockDom.getElementById('fullSchedTime_log').value, '00:00');
  assert.equal(currentMockDom.getElementById('localScheduleTime_log').value, '每日 00:00');
  assert.equal(currentMockDom.getElementById('offsiteScheduleTime_log').value, '無');
  assert.equal(currentMockDom.getElementById('sourceSystemId').value, '');
});

test('handleFormSubmit 送出前強制執行 syncTargetSchedule 確保所有標的排程字串更新', () => {
  currentMockDom = createMockDom();

  // 勾選 db 標的
  currentMockDom.getElementById('check_db').checked = true;
  currentMockDom.getElementById('fullValue_db').value = '1';
  currentMockDom.getElementById('fullUnit_db').value = '天';
  // 模擬使用者直接變更 time input 尚未 trigger 任何 event
  currentMockDom.getElementById('fullSchedTime_db').value = '05:45';
  currentMockDom.getElementById('localScheduleTime_db').value = '舊排程未同步';
  currentMockDom.getElementById('offsite_db_val_否').checked = true;
  currentMockDom.getElementById('offsite_db_val_否').value = '否';
  currentMockDom.getElementById('submissionId').value = 'sub-001';
  currentMockDom.getElementById('sourceSystemId').value = 'SYS-001';
  currentMockDom.getElementById('localLoc_db').value = 'AST-001 - 備份設備';
  vm.runInContext(`
    assetOptionsLoadState = "ready";
    assetOptionsData = [{ name: "AST-001 - 備份設備", level: "一般系統" }];
  `, sandbox);

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
  assert.equal(submittedFormData.id, 'sub-001');
  assert.equal(submittedFormData.systemId, 'SYS-001');
  assert.equal(submittedFormData.targets[0].offsiteDispositionMethod, '無');
  assert.equal(submittedFormData.targets[0].offsiteEvidenceRequirements, '無');
});


// --- 完整備份改為選填（完整、差異二擇一必填）---

test('updateFullPickerVisibility 依完整備份數值（>0）動態切換完整軌可見度', () => {
  currentMockDom = createMockDom();
  const fullValEl = currentMockDom.register('fullValue_db', { value: '' });
  const fullPicker = currentMockDom.register('fullSchedulePicker_db', { classes: ['hidden'] });
  currentMockDom.register('fullUnit_db', { value: '天' });
  currentMockDom.register('fullSchedTime_db', { value: '00:00' });
  currentMockDom.register('localScheduleTime_db', { value: '' });
  currentMockDom.register('localSchedulePreview_db', { textContent: '' });

  updateFullPickerVisibility('db');
  assert.equal(fullPicker.classList.contains('hidden'), true);

  fullValEl.value = '1';
  updateFullPickerVisibility('db');
  assert.equal(fullPicker.classList.contains('hidden'), false);

  fullValEl.value = '0';
  updateFullPickerVisibility('db');
  assert.equal(fullPicker.classList.contains('hidden'), true);
});

test('完整備份欄位不再必填：標籤無紅色星號、min 為 0，勾選標的也不設定 required', () => {
  currentMockDom = createMockDom();
  renderFormTargets();
  const targetHtml = currentMockDom.getElementById('targetsContainer').innerHTML;
  const fullLabelMatch = targetHtml.match(/每 ___ 執行「完整備份」[^<]*<\/label>/);
  assert.ok(fullLabelMatch, '完整備份標籤內不得含必填星號 span（label 應直接結束）');
  assert.ok(/每 ___ 執行「差異備份」[^<]*<\/label>/.test(targetHtml), '差異備份標籤同樣無星號');
  const fullInput = targetHtml.match(/<input[^>]*id="fullValue_db"[^>]*>/)[0];
  assert.ok(fullInput.includes('min="0"'), '完整備份數值 min 應為 0');
  assert.ok(fullInput.includes('0代表無'), '應提示 0 代表無');
  assert.ok(fullInput.includes("updateFullPickerVisibility('db')"), '輸入時應連動完整軌可見度');

  currentMockDom = createMockDom();
  currentMockDom.getElementById('check_db').checked = true;
  vm.runInContext(`toggleTarget('db')`, sandbox);
  assert.equal(currentMockDom.getElementById('fullValue_db').required, false, '勾選標的後完整備份數值不得為 required');
});

test('validateLocalBackupCycles 完整、差異二擇一必填，系統日誌檔不適用', () => {
  const alerts = [];
  sandbox.alert = (msg) => alerts.push(msg);

  currentMockDom = createMockDom();
  currentMockDom.getElementById('fullValue_db').value = '';
  currentMockDom.getElementById('diffValue_db').value = '';
  assert.equal(validateLocalBackupCycles([{ id: 'db', name: '資料庫' }]), false, '兩者皆空必須擋下');
  assert.equal(alerts.length, 1);
  assert.ok(alerts[0].includes('資料庫'));

  currentMockDom.getElementById('fullValue_db').value = '0';
  currentMockDom.getElementById('diffValue_db').value = '0';
  assert.equal(validateLocalBackupCycles([{ id: 'db', name: '資料庫' }]), false, '兩者皆 0 必須擋下');

  currentMockDom.getElementById('fullValue_db').value = '';
  currentMockDom.getElementById('diffValue_db').value = '1';
  assert.equal(validateLocalBackupCycles([{ id: 'db', name: '資料庫' }]), true, '僅差異備份應通過');

  currentMockDom.getElementById('fullValue_db').value = '1';
  currentMockDom.getElementById('diffValue_db').value = '';
  assert.equal(validateLocalBackupCycles([{ id: 'db', name: '資料庫' }]), true, '僅完整備份應通過');

  assert.equal(validateLocalBackupCycles([{ id: 'log', name: '系統日誌檔' }]), true, '系統日誌檔不適用');

  sandbox.alert = () => {};
});

test('handleFormSubmit 僅差異備份可送出：完整週期送 0、排程為「差異:」格式；兩者皆空則阻止送出', () => {
  function runSubmit(fullValue, diffValue) {
    currentMockDom = createMockDom();
    currentMockDom.getElementById('check_db').checked = true;
    currentMockDom.getElementById('fullValue_db').value = fullValue;
    currentMockDom.getElementById('fullUnit_db').value = '週';
    currentMockDom.getElementById('diffValue_db').value = diffValue;
    currentMockDom.getElementById('diffUnit_db').value = '天';
    currentMockDom.getElementById('diffSchedTime_db').value = '02:00';
    currentMockDom.getElementById('offsite_db_val_否').checked = true;
    currentMockDom.getElementById('offsite_db_val_否').value = '否';
    currentMockDom.getElementById('localLoc_db').value = 'AST-001 - 備份設備';
    vm.runInContext(`
      assetOptionsLoadState = "ready";
      assetOptionsData = [{ name: "AST-001 - 備份設備", level: "一般系統" }];
    `, sandbox);

    let submitted = null;
    const run = {
      withSuccessHandler() { return this; },
      withFailureHandler() { return this; },
      submitData(formData) { submitted = formData; }
    };
    sandbox.google = { script: { run } };
    handleFormSubmit({ preventDefault: () => {} });
    delete sandbox.google;
    return submitted;
  }

  const diffOnly = runSubmit('', '1');
  assert.ok(diffOnly, '僅差異備份應可送出');
  assert.equal(diffOnly.targets[0].localFullValue, '0');
  assert.equal(diffOnly.targets[0].localDiffValue, '1');
  assert.equal(diffOnly.targets[0].localScheduleTime, '差異:每日 02:00');

  const neither = runSubmit('', '');
  assert.equal(neither, null, '完整與差異皆未設定時不得送出');
});

test('表單編輯回填僅差異備份：完整備份數值留空、完整軌隱藏、差異軌顯示並保留排程', () => {
  currentMockDom = createMockDom();
  currentMockDom.register('fullSchedulePicker_db', { classes: ['hidden'] });
  currentMockDom.register('diffSchedulePicker_db', { classes: ['hidden'] });

  vm.runInContext(`appData = [{
    id: "sys-diff-only",
    systemName: "僅差異系統",
    systemLevel: "一般系統",
    managerName: "陳小華",
    targets: [
      {
        targetKey: "db",
        targetName: "資料庫",
        localDiffValue: "1",
        localDiffUnit: "天",
        localFullValue: "",
        localFullUnit: "週",
        localRetention: "7",
        localLocation: "正式機房",
        localScheduleTime: "差異:每日 01:30",
        localDispositionMethod: "磁碟自動覆寫",
        localEvidenceRequirements: "備份系統執行Log",
        isOffsite: "否",
        offsiteScheduleTime: "無"
      }
    ]
  }];`, sandbox);

  editSystem('sys-diff-only');

  assert.equal(currentMockDom.getElementById('fullValue_db').value, '');
  assert.equal(currentMockDom.getElementById('fullSchedulePicker_db').classList.contains('hidden'), true);
  assert.equal(currentMockDom.getElementById('diffSchedulePicker_db').classList.contains('hidden'), false);
  assert.equal(currentMockDom.getElementById('diffSchedTime_db').value, '01:30');
  assert.equal(currentMockDom.getElementById('localScheduleTime_db').value, '差異:每日 01:30');
});


// --- 保留所有版本：停用到期銷毀方法與銷毀確認方式（全部保留不會銷毀，顯示並送出「無」）---

function registerDispositionSelects_(dom, targetId, disp, evid) {
  const dispEl = dom.register(`localDispositionMethod_${targetId}`, { value: disp || '磁碟自動覆寫' });
  dispEl.options.push({ value: '磁碟自動覆寫' }, { value: '人工手動覆蓋' });
  const evidEl = dom.register(`localEvidenceRequirements_${targetId}`, { value: evid || '備份系統執行Log' });
  evidEl.options.push({ value: '備份系統執行Log' }, { value: '備份檔名確認' });
  dom.register(`check_${targetId}`, { checked: true });
  dom.register(`localRet_${targetId}`, { value: '4' });
  return { dispEl, evidEl, retAllEl: dom.register(`localRetAll_${targetId}`, { checked: false }) };
}

test('勾選「保留所有版本」會停用兩個銷毀欄位並顯示「無」，取消勾選恢復原本的選擇', () => {
  currentMockDom = createMockDom();
  const { dispEl, evidEl, retAllEl } = registerDispositionSelects_(currentMockDom, 'db', '人工手動覆蓋', '備份檔名確認');

  retAllEl.checked = true;
  toggleLocalRetentionMode('db');
  assert.equal(dispEl.disabled, true);
  assert.equal(evidEl.disabled, true);
  assert.equal(dispEl.value, '無');
  assert.equal(evidEl.value, '無');
  assert.equal(dispEl.options.filter(o => o.value === '無').length, 1, '只應有一個暫時的「無」選項');

  // 重複呼叫不得重複新增選項，也不得覆蓋原本記住的選擇
  toggleLocalRetentionMode('db');
  assert.equal(dispEl.options.filter(o => o.value === '無').length, 1);

  retAllEl.checked = false;
  toggleLocalRetentionMode('db');
  assert.equal(dispEl.disabled, false);
  assert.equal(evidEl.disabled, false);
  assert.equal(dispEl.value, '人工手動覆蓋', '應恢復勾選前的選擇');
  assert.equal(evidEl.value, '備份檔名確認', '應恢復勾選前的選擇');
  assert.equal(dispEl.options.some(o => o.value === '無'), false, '暫時的「無」選項應被移除');
});

test('未曾更動選擇就勾選、取消「保留所有版本」，銷毀欄位回到預設值', () => {
  currentMockDom = createMockDom();
  const { dispEl, evidEl, retAllEl } = registerDispositionSelects_(currentMockDom, 'source');

  retAllEl.checked = true;
  toggleLocalRetentionMode('source');
  retAllEl.checked = false;
  toggleLocalRetentionMode('source');
  assert.equal(dispEl.value, '磁碟自動覆寫');
  assert.equal(evidEl.value, '備份系統執行Log');
  assert.equal(dispEl.disabled, false);
});

test('未勾選「保留所有版本」時，toggleLocalRetentionMode 不得更動使用者目前的銷毀欄位選擇', () => {
  currentMockDom = createMockDom();
  const { dispEl, evidEl } = registerDispositionSelects_(currentMockDom, 'db', '人工手動覆蓋', '備份檔名確認');

  toggleLocalRetentionMode('db');
  assert.equal(dispEl.value, '人工手動覆蓋');
  assert.equal(evidEl.value, '備份檔名確認');
  assert.equal(dispEl.disabled, false);
});

test('編輯「全部保留」的標的：勾選保留、銷毀欄位顯示「無」並停用；取消勾選後回到已存的選擇', () => {
  currentMockDom = createMockDom();
  const dispEl = currentMockDom.register('localDispositionMethod_db', { value: '磁碟自動覆寫' });
  dispEl.options.push({ value: '磁碟自動覆寫' }, { value: '人工手動覆蓋' });
  const evidEl = currentMockDom.register('localEvidenceRequirements_db', { value: '備份系統執行Log' });
  evidEl.options.push({ value: '備份系統執行Log' }, { value: '備份檔名確認' });

  vm.runInContext(`appData = [{
    id: "sys-all",
    systemName: "全部保留系統",
    systemLevel: "一般系統",
    managerName: "林小美",
    targets: [
      {
        targetKey: "db",
        targetName: "資料庫",
        localDiffValue: "1",
        localDiffUnit: "天",
        localFullValue: "1",
        localFullUnit: "週",
        localRetention: "全部保留",
        localLocation: "正式機房",
        localScheduleTime: "完整:每週日 02:00 / 差異:每日 01:30",
        localDispositionMethod: "人工手動覆蓋",
        localEvidenceRequirements: "備份檔名確認",
        isOffsite: "否",
        offsiteScheduleTime: "無"
      }
    ]
  }];`, sandbox);

  editSystem('sys-all');

  assert.equal(currentMockDom.getElementById('localRetAll_db').checked, true);
  assert.equal(dispEl.disabled, true);
  assert.equal(evidEl.disabled, true);
  assert.equal(dispEl.value, '無');
  assert.equal(evidEl.value, '無');

  // 使用者取消「保留所有版本」後，回到資料庫已存的選擇而不是預設值
  currentMockDom.getElementById('localRetAll_db').checked = false;
  toggleLocalRetentionMode('db');
  assert.equal(dispEl.disabled, false);
  assert.equal(dispEl.value, '人工手動覆蓋');
  assert.equal(evidEl.value, '備份檔名確認');
});

test('編輯已存為「無」的全部保留資料：取消勾選後回到預設值', () => {
  currentMockDom = createMockDom();
  const dispEl = currentMockDom.register('localDispositionMethod_db', { value: '磁碟自動覆寫' });
  dispEl.options.push({ value: '磁碟自動覆寫' }, { value: '人工手動覆蓋' });
  const evidEl = currentMockDom.register('localEvidenceRequirements_db', { value: '備份系統執行Log' });
  evidEl.options.push({ value: '備份系統執行Log' }, { value: '備份檔名確認' });

  vm.runInContext(`appData = [{
    id: "sys-all-none",
    systemName: "全部保留系統",
    systemLevel: "一般系統",
    managerName: "林小美",
    targets: [
      {
        targetKey: "db",
        targetName: "資料庫",
        localDiffValue: "0",
        localDiffUnit: "天",
        localFullValue: "1",
        localFullUnit: "週",
        localRetention: "ALL",
        localLocation: "正式機房",
        localScheduleTime: "每週日 02:00",
        localDispositionMethod: "無",
        localEvidenceRequirements: "無",
        isOffsite: "否",
        offsiteScheduleTime: "無"
      }
    ]
  }];`, sandbox);

  editSystem('sys-all-none');
  assert.equal(dispEl.disabled, true);
  assert.equal(dispEl.value, '無');

  currentMockDom.getElementById('localRetAll_db').checked = false;
  toggleLocalRetentionMode('db');
  assert.equal(dispEl.value, '磁碟自動覆寫');
  assert.equal(evidEl.value, '備份系統執行Log');
});

test('handleFormSubmit：全部保留時銷毀方法與確認方式送「無」，一般保留代數則送使用者選擇', () => {
  function runSubmit(isAll) {
    currentMockDom = createMockDom();
    const { dispEl, evidEl, retAllEl } = registerDispositionSelects_(currentMockDom, 'db', '人工手動覆蓋', '備份檔名確認');
    currentMockDom.getElementById('fullValue_db').value = '1';
    currentMockDom.getElementById('fullUnit_db').value = '天';
    currentMockDom.getElementById('offsite_db_val_否').checked = true;
    currentMockDom.getElementById('offsite_db_val_否').value = '否';
    currentMockDom.getElementById('localLoc_db').value = 'AST-001 - 備份設備';
    vm.runInContext(`
      assetOptionsLoadState = "ready";
      assetOptionsData = [{ name: "AST-001 - 備份設備", level: "一般系統" }];
    `, sandbox);

    retAllEl.checked = isAll;
    toggleLocalRetentionMode('db');

    let submitted = null;
    sandbox.google = { script: { run: {
      withSuccessHandler() { return this; },
      withFailureHandler() { return this; },
      submitData(formData) { submitted = formData; }
    } } };
    handleFormSubmit({ preventDefault: () => {} });
    delete sandbox.google;
    return submitted.targets[0];
  }

  const all = runSubmit(true);
  assert.equal(all.localRetention, 'ALL');
  assert.equal(all.localDispositionMethod, '無');
  assert.equal(all.localEvidenceRequirements, '無');

  const normal = runSubmit(false);
  assert.equal(normal.localRetention, '4');
  assert.equal(normal.localDispositionMethod, '人工手動覆蓋');
  assert.equal(normal.localEvidenceRequirements, '備份檔名確認');
});

test('resetForm 會解除「保留所有版本」造成的停用並回到預設值', () => {
  currentMockDom = createMockDom();
  const { dispEl, evidEl, retAllEl } = registerDispositionSelects_(currentMockDom, 'db', '人工手動覆蓋', '備份檔名確認');
  currentMockDom.getElementById('check_db').checked = false;

  retAllEl.checked = true;
  toggleLocalRetentionMode('db');
  assert.equal(dispEl.disabled, true);

  resetForm();
  assert.equal(retAllEl.checked, false);
  assert.equal(dispEl.disabled, false);
  assert.equal(evidEl.disabled, false);
  assert.equal(dispEl.value, '磁碟自動覆寫');
  assert.equal(evidEl.value, '備份系統執行Log');
});

test('取消勾選整個標的會解除「保留所有版本」的停用狀態', () => {
  currentMockDom = createMockDom();
  const { dispEl, retAllEl } = registerDispositionSelects_(currentMockDom, 'db', '人工手動覆蓋', '備份檔名確認');
  currentMockDom.register('details_db', {});
  currentMockDom.register('badge_db', {});

  retAllEl.checked = true;
  toggleLocalRetentionMode('db');
  assert.equal(dispEl.disabled, true);

  currentMockDom.getElementById('check_db').checked = false;
  toggleTarget('db');
  assert.equal(retAllEl.checked, false);
  assert.equal(dispEl.disabled, false);
  assert.equal(dispEl.value, '磁碟自動覆寫');
});
