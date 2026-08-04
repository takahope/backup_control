# 系統屬性自動判定與帶入 (System Level Auto-fill) Design Spec

## 1. 背景與目標 (Background & Goal)
目前新增備份管制表單中的「系統屬性」欄位（關鍵核心系統 / 一般系統）是靜態寫死的 HTML 選單，需由使用者手動判斷與選擇。
為降低人工判定錯誤率，本功能旨在透過後端資料擷取，自動判斷使用者所選的「系統/資訊資產」屬於何種等級，並在前端動態帶入對應的「系統屬性」，同時保留讓使用者可以手動覆寫的彈性。

## 2. 架構設計 (Architecture Design)

### 2.1 後端邏輯 (`code.js` -> `getAssetOptions`)
擴充現有的雙來源讀取邏輯，以提取等級判定所需的欄位。
回傳的資料結構將從 `Array<string>` 變更為 `Array<{name: string, level: string}>`。

**來源 A：系統清冊**
- **讀取範圍**：擴增至 G 欄（第 7 欄）。
- **判定邏輯**：
  - 當 G 欄 (`row[6]`) 值為「中級」或「高級」：`level = "關鍵核心系統"`
  - 當 G 欄 (`row[6]`) 值為「普級」或「套裝軟體」：`level = "一般系統"`

**來源 B：資訊資產清單**
- **讀取範圍**：擴增至 R 欄（第 18 欄）。
- **判定邏輯**：
  - 將 R 欄 (`row[17]`) 轉換為數值。
  - 當數值 `>= 7`：`level = "關鍵核心系統"`
  - 當數值 `< 7` 或無法解析：`level = "一般系統"`

### 2.2 前端邏輯 (`Index.html`)
- **資料儲存**：`systemOptionsData` 與 `assetOptionsData` 變數將儲存完整的物件陣列 (`{name, level}`)。
- **渲染邏輯**：`renderAssetOptionsDatalist()` 需改為讀取物件的 `name` 屬性來產生 `<option>`。
- **自動帶入事件**：
  - 在 `#systemName` 欄位綁定 `input`（或 `change`）事件 (`handleSystemNameInput`)。
  - 觸發時，根據當前選擇的 Radio (`sourceType`) 決定要在哪個陣列進行精準比對 (Exact match)。
  - 若 `systemName.value` 吻合陣列中某個物件的 `name`，則執行 `document.getElementById('systemLevel').value = matchedObject.level`。
- **保留修改彈性**：此機制僅為事件觸發時賦值，不會鎖定 (disable) `#systemLevel` 下拉選單，確保使用者隨時可手動修正。

## 3. 例外處理與邊界情況 (Edge Cases)
1. **找不到對應項目**：若使用者自行輸入了不存在於清單中的文字，前端不做任何連動改變（保留預設值或前一次的選擇）。
2. **編輯模式的競態條件 (Race Condition)**：若使用者透過「編輯」按鈕載入舊有資料，表單會先填入舊值。當 `getAssetOptions` 回傳後，只要使用者沒有實際去更動 `#systemName` 觸發 `input` 事件，該自動帶入機制就不會介入洗掉使用者原本儲存好的系統屬性值。
3. **空值處理**：對於來源表單缺少 G 欄或 R 欄資料的列，預設均退回判定為 `一般系統`，避免造成前端例外錯誤。

## 4. 實作範圍 (Scope)
- 修改 `code.js` 中的 `getAssetOptions` 函式。
- 修改 `Index.html` 中的資料存儲、`renderAssetOptionsDatalist` 邏輯，與驗證邏輯 `validateLocalLocationSelection` 提取 `name` 的部分。
- 於 `Index.html` 內新增並綁定 `#systemName` 的 `oninput` 監聽器。
