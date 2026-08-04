# 雙來源下拉選單功能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在新增設定表單加入來源切換功能，並動態自「系統清冊」及「資訊資產清單」載入對應選項。

**Architecture:** 
後端 `getAssetOptions` 函式將同時查詢 `SECURITY_BASELINE_SPREADSHEET_ID` (系統清冊) 與 `ASSETS_SPREADSHEET_ID` (資訊資產清單)，回傳包含 `systems` 和 `assets` 兩組資料的物件。前端 `Index.html` 新增一組 Radio Buttons 供使用者切換來源，並根據所選來源動態替換下拉選單 (datalist) 的內容。因為送出的格式皆為 `ID - 名稱`，現有後端的 `submitData` 字串切割邏輯維持不變。

**Tech Stack:** Google Apps Script, HTML/JS, TailwindCSS

## Global Constraints

- 系統需維持與現有 `submitData` 的相容性（輸出格式必須是 `ID - 名稱`）
- 不破壞原本權限驗證 (`checkAppPermission_`)

---

### Task 1: 擴充後端選單資料擷取 (code.js)

**Files:**
- Modify: `code.js` (函式 `getAssetOptions` 的實作)

**Interfaces:**
- Consumes: `SECURITY_BASELINE_SPREADSHEET_ID`, `ASSETS_SPREADSHEET_ID`
- Produces: 回傳 `{ success: true, data: { systems: [...], assets: [...] } }`

- [ ] **Step 1: 修改 getAssetOptions 函式**
替換 `getAssetOptions` 函式內容，加入雙來源讀取：

```javascript
// 4. 讀取防護基準系統清冊與外部資訊資產清單，供前端動態切換來源
function getAssetOptions() {
  try {
    const permission = checkAppPermission_();
    if (!permission.allowed) {
      return { success: false, message: permission.message || '無權限讀取來源清冊。', data: { systems: [], assets: [] } };
    }

    const result = { systems: [], assets: [] };

    // 1. 讀取防護基準系統清冊
    if (typeof SECURITY_BASELINE_SPREADSHEET_ID !== 'undefined' && SECURITY_BASELINE_SPREADSHEET_ID !== 'PLEASE_SET_SECURITY_BASELINE_SPREADSHEET_ID') {
      try {
        const baselineSpreadsheet = SpreadsheetApp.openById(String(SECURITY_BASELINE_SPREADSHEET_ID).trim());
        const systemSheet = baselineSpreadsheet.getSheetByName('系統清冊');
        if (systemSheet) {
          const lastRow = systemSheet.getLastRow();
          if (lastRow > 1) {
            const values = systemSheet.getRange(2, 1, lastRow - 1, 3).getValues();
            const seen = {};
            values.forEach(function(row) {
              const colA = String(row[0] || '').trim();
              const colC = String(row[2] || '').trim();
              if (!colA && !colC) return;
              const name = colC ? (colA ? (colA + ' - ' + colC) : colC) : colA;
              if (!seen[name]) { seen[name] = true; result.systems.push(name); }
            });
            result.systems.sort();
          }
        }
      } catch (e) {
        console.warn('讀取系統清冊失敗', e);
      }
    }

    // 2. 讀取資訊資產清單
    if (typeof ASSETS_SPREADSHEET_ID !== 'undefined' && ASSETS_SPREADSHEET_ID !== 'PLEASE_SET_ASSETS_SPREADSHEET_ID') {
      try {
        const assetsSpreadsheet = SpreadsheetApp.openById(String(ASSETS_SPREADSHEET_ID).trim());
        // 嘗試抓取「資訊資產清單」，若無則退回抓取「資訊資產」
        let assetsSheet = assetsSpreadsheet.getSheetByName('資訊資產清單');
        if (!assetsSheet) assetsSheet = assetsSpreadsheet.getSheetByName('資訊資產');
        
        if (assetsSheet) {
          const lastRow = assetsSheet.getLastRow();
          if (lastRow > 1) {
            const values = assetsSheet.getRange(2, 1, lastRow - 1, 3).getValues();
            const seen = {};
            values.forEach(function(row) {
              const colA = String(row[0] || '').trim();
              const colC = String(row[2] || '').trim();
              if (!colA && !colC) return;
              const name = colC ? (colA ? (colA + ' - ' + colC) : colC) : colA;
              if (!seen[name]) { seen[name] = true; result.assets.push(name); }
            });
            result.assets.sort();
          }
        }
      } catch (e) {
        console.warn('讀取資訊資產失敗', e);
      }
    }

    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.toString(), data: { systems: [], assets: [] } };
  }
}
```

- [ ] **Step 2: 執行語法檢查**

Run: `node --check code.js`
Expected: PASS 且無語法錯誤

- [ ] **Step 3: Commit changes**

```bash
git add code.js
git commit -m "feat: read dual sources for system and asset dropdown options"
```

### Task 2: 調整前端 HTML 介面與事件綁定 (Index.html)

**Files:**
- Modify: `Index.html` (表單區塊與 JS)

**Interfaces:**
- Consumes: 後端回傳之 `{ systems: [...], assets: [...] }` 格式資料

- [ ] **Step 1: 新增 Radio 按鈕區塊與調整標籤**
在 `Index.html` 約第 149 行（`<label class="block text-sm font-medium text-gray-700">系統/設備名稱</label>`）前後，修改為包含 Radio 的結構：

```html
              <div>
                <div class="mb-2 flex items-center space-x-4 text-sm font-medium text-gray-700">
                  <span>來源：</span>
                  <label class="inline-flex items-center">
                    <input type="radio" name="sourceType" value="systems" class="form-radio text-blue-600" checked onchange="handleSourceChange()">
                    <span class="ml-1">系統清冊</span>
                  </label>
                  <label class="inline-flex items-center">
                    <input type="radio" name="sourceType" value="assets" class="form-radio text-blue-600" onchange="handleSourceChange()">
                    <span class="ml-1">資訊資產清單</span>
                  </label>
                </div>
                <label class="block text-sm font-medium text-gray-700">系統 / 資訊資產</label>
                <input type="text" id="systemName" list="assetOptionsList" required class="mt-1 block w-full rounded-md border-gray-300 shadow-sm p-2 border focus:ring-blue-500 focus:border-blue-500">
              </div>
```
*(注意：要記得幫原本的 input 加上 `list="assetOptionsList"`，目前可能沒有綁好 `list` 屬性)*

- [ ] **Step 2: 調整全域變數與 render 邏輯 (JS 區塊)**
在 `Index.html` 的 JS 區塊，修改 `assetOptions` 為兩個變數，並實作 `handleSourceChange` 與改寫 `loadAssetOptions`。

找到 JS 全域變數區段：
```javascript
    let globalSubmissions = [];
    let assetOptions = [];
```
修改為：
```javascript
    let globalSubmissions = [];
    let systemOptionsData = [];
    let assetOptionsData = [];
```

接著新增 `handleSourceChange` 並修改 `renderAssetOptionsDatalist` 與 `loadAssetOptions`：
```javascript
    function handleSourceChange() {
      const sourceType = document.querySelector('input[name="sourceType"]:checked').value;
      const list = document.getElementById('assetOptionsList');
      if (!list) return;

      list.innerHTML = '';
      const targetData = sourceType === 'systems' ? systemOptionsData : assetOptionsData;
      
      targetData.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        list.appendChild(option);
      });
      
      // 切換來源時清空現有輸入
      document.getElementById('systemName').value = '';
    }

    function renderAssetOptionsDatalist() {
      handleSourceChange(); // 依據當前選取的 Radio 渲染
    }

    function loadAssetOptions() {
      if (typeof google !== 'undefined' && google.script) {
        google.script.run
          .withSuccessHandler(function(response) {
            if (response && response.success && response.data) {
              systemOptionsData = response.data.systems || [];
              assetOptionsData = response.data.assets || [];
            } else {
              systemOptionsData = [];
              assetOptionsData = [];
            }
            renderAssetOptionsDatalist();
          })
          .withFailureHandler(function(error) {
            console.error("載入來源選項失敗: " + error);
          })
          .getAssetOptions();
      }
    }
```

- [ ] **Step 3: 確認 `systemName` 初始化寫入邏輯**
在原本 `editSubmission` 函式中，當帶入舊資料時，需要能正確設定來源 Radio（但因為目前 `submitData` 送出時並無紀錄是來自系統還是資產，前端可以預設保留在「系統清冊」並將值塞進去即可，`<input list="assetOptionsList">` 原本就是自由輸入的 text，不用特地去強制切換 radio）。請確認不用額外動 `editSubmission`，除非有報錯。

- [ ] **Step 4: Commit changes**

```bash
git add Index.html
git commit -m "feat: add radio toggle for dual source system and asset datalist"
```
