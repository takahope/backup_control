# 系統屬性自動判定與帶入 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 實作表單中「系統屬性」的自動判定與帶入功能。前端根據使用者在下拉選單選擇的「系統/資訊資產」，自動改變系統屬性（關鍵核心系統或一般系統），並保留手動修改彈性。

**Architecture:** 
1. `code.js` 中的 `getAssetOptions` 擴充讀取範圍，並在後端處理出 `{ name, level }` 結構。
2. `Index.html` 接收並存儲物件陣列，修正現有提取字串陣列的相關程式碼，並在 `#systemName` 綁定 `input` 事件以達成連動。

**Tech Stack:** Google Apps Script, HTML/JS

## Global Constraints

- 系統需維持原本 `ID - 名稱` 格式供 datalist 顯示與後續寫入表單使用。
- 必須保留 `validateLocalLocationSelection` 的驗證邏輯，且必須從新的 `{name, level}` 物件陣列中提取 `name` 來比對。

---

### Task 1: 擴充後端資料擷取與判定邏輯 (code.js)

**Files:**
- Modify: `code.js`

**Interfaces:**
- Consumes: Google Sheet (系統清冊 A-G 欄, 資訊資產清單 A-R 欄)
- Produces: 回傳 `{ success: true, data: { systems: [{name, level}...], assets: [{name, level}...] } }`

- [ ] **Step 1: 修改 getAssetOptions 函式**
替換 `getAssetOptions` 內容。主要改變是 `getRange` 的欄位數，以及新增 `level` 判定邏輯。

```javascript
    // 1. 讀取防護基準系統清冊 (A-G 欄, 共 7 欄)
    if (typeof SECURITY_BASELINE_SPREADSHEET_ID !== 'undefined' && SECURITY_BASELINE_SPREADSHEET_ID !== 'PLEASE_SET_SECURITY_BASELINE_SPREADSHEET_ID') {
      try {
        const baselineSpreadsheet = SpreadsheetApp.openById(String(SECURITY_BASELINE_SPREADSHEET_ID).trim());
        const systemSheet = baselineSpreadsheet.getSheetByName('系統清冊');
        if (systemSheet) {
          const lastRow = systemSheet.getLastRow();
          if (lastRow > 1) {
            const values = systemSheet.getRange(2, 1, lastRow - 1, 7).getValues();
            const seen = {};
            values.forEach(function(row) {
              const colA = String(row[0] || '').trim();
              const colC = String(row[2] || '').trim();
              const colG = String(row[6] || '').trim(); // 系統等級
              if (!colA && !colC) return;
              
              const name = colC ? (colA ? (colA + ' - ' + colC) : colC) : colA;
              const level = (colG === '中級' || colG === '高級') ? '關鍵核心系統' : '一般系統';
              
              if (!seen[name]) { 
                seen[name] = true; 
                result.systems.push({ name: name, level: level }); 
              }
            });
            result.systems.sort((a, b) => a.name.localeCompare(b.name));
          }
        }
      } catch (e) {
        console.warn('讀取系統清冊失敗', e);
      }
    }

    // 2. 讀取資訊資產清單 (A-R 欄, 共 18 欄)
    if (typeof ASSETS_SPREADSHEET_ID !== 'undefined' && ASSETS_SPREADSHEET_ID !== 'PLEASE_SET_ASSETS_SPREADSHEET_ID') {
      try {
        const assetsSpreadsheet = SpreadsheetApp.openById(String(ASSETS_SPREADSHEET_ID).trim());
        let assetsSheet = assetsSpreadsheet.getSheetByName(SHEET_ASSETS_NAME);
        if (!assetsSheet) assetsSheet = assetsSpreadsheet.getSheetByName('資訊資產');
        
        if (assetsSheet) {
          const lastRow = assetsSheet.getLastRow();
          if (lastRow > 1) {
            const values = assetsSheet.getRange(2, 1, lastRow - 1, 18).getValues();
            const seen = {};
            values.forEach(function(row) {
              const colA = String(row[0] || '').trim();
              const colC = String(row[2] || '').trim();
              const colR = row[17]; // 資產價值
              if (!colA && !colC) return;
              
              const name = colC ? (colA ? (colA + ' - ' + colC) : colC) : colA;
              
              let level = '一般系統';
              const rValue = parseFloat(colR);
              if (!isNaN(rValue) && rValue >= 7) {
                level = '關鍵核心系統';
              }

              if (!seen[name]) { 
                seen[name] = true; 
                result.assets.push({ name: name, level: level }); 
              }
            });
            result.assets.sort((a, b) => a.name.localeCompare(b.name));
          }
        }
      } catch (e) {
        console.warn('讀取資訊資產失敗', e);
      }
    }
```

- [ ] **Step 2: 執行語法檢查**

Run: `node --check code.js`
Expected: PASS 且無語法錯誤

- [ ] **Step 3: Commit changes**

```bash
git add code.js
git commit -m "feat: backend support for system level auto calculation"
```

### Task 2: 前端介面更新與事件綁定 (Index.html)

**Files:**
- Modify: `Index.html`

**Interfaces:**
- Consumes: `{ systems: [{name, level}...], assets: [{name, level}...] }` 格式資料

- [ ] **Step 1: 更新 `renderAssetOptionsDatalist` 邏輯**
原本是直接遍歷 `systemOptionsData` 與 `assetOptionsData` 字串陣列，現在陣列裡是物件，所以要用 `obj.name`。
```javascript
    function renderAssetOptionsDatalist() {
      const sysList = document.getElementById('systemOptionsList');
      const astList = document.getElementById('assetOptionsList');
      
      if (sysList) {
        sysList.innerHTML = '';
        systemOptionsData.forEach(item => {
          const option = document.createElement('option');
          option.value = item.name;
          sysList.appendChild(option);
        });
      }
      
      if (astList) {
        astList.innerHTML = '';
        assetOptionsData.forEach(item => {
          const option = document.createElement('option');
          option.value = item.name;
          astList.appendChild(option);
        });
      }
      
      handleSourceChange(false);
    }
```

- [ ] **Step 2: 更新 `validateLocalLocationSelection` 邏輯**
驗證時現在必須將 `assetOptionsData` map 成純字串陣列來做 `includes` 比對。
修改約 387 行：
```javascript
    function validateLocalLocationSelection(activeTargets) {
      if (!Array.isArray(assetOptionsData) || assetOptionsData.length === 0) return true;
      const validNames = assetOptionsData.map(item => item.name); // 提取字串陣列

      for (const t of activeTargets) {
        // ... (中間不變)
        if (!validNames.includes(value)) {
          alert(`「${t.name}」的本地存放地點/設備請從下拉選單選擇有效的資訊資產。`);
          input.focus();
          return false;
        }
      }
      return true;
    }
```

- [ ] **Step 3: 新增自動帶入事件 (`handleSystemNameInput`)**
在 `Index.html` 內的全域函式區塊新增 `handleSystemNameInput` 函式，並在 `DOMContentLoaded` (或直接加在 HTML tag 上) 綁定到 `#systemName` 欄位。

找到 `<input type="text" id="systemName" ... >`，修改 HTML 加上 `oninput="handleSystemNameInput()"` 或 `onchange="handleSystemNameInput()"`。為了更好的體驗，建議使用 `onchange`，因為輸入過程中尚未輸入完整名稱時不需要不斷尋找：
```html
<input type="text" id="systemName" list="systemOptionsList" required class="mt-1 block w-full rounded-md border-gray-300 shadow-sm p-2 border focus:ring-blue-500 focus:border-blue-500" onchange="handleSystemNameInput()">
```

在 JS 區塊加入函式：
```javascript
    function handleSystemNameInput() {
      const inputVal = document.getElementById('systemName').value.trim();
      if (!inputVal) return;

      const sourceRadio = document.querySelector('input[name="sourceType"]:checked');
      if (!sourceRadio) return;
      
      const targetData = sourceRadio.value === 'systems' ? systemOptionsData : assetOptionsData;
      
      // 在目標資料中尋找是否有吻合的名稱
      const matched = targetData.find(item => item.name === inputVal);
      if (matched && matched.level) {
        const levelSelect = document.getElementById('systemLevel');
        if (levelSelect) {
          levelSelect.value = matched.level;
        }
      }
    }
```

- [ ] **Step 4: Commit changes**

```bash
git add Index.html
git commit -m "feat: frontend auto-fill system level on systemName selection"
```
