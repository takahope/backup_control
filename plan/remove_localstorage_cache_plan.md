# 徹底移除 localStorage 快取機制實作計畫

## 一、 目標說明 (Goal Description)

### 背景問題
在原設計中，[`Index.html`](file:///home/itsutaka/biobank_manager/backup_control/Index.html) 使用了瀏覽器 `localStorage`（鍵值 `isms_backup_records`）進行狀態暫存與失敗回退（Fallback）。
這造成了嚴重的資料不同步問題：
1. **載入即吞舊資料**：頁面開啟第一毫秒就先讀取 `localStorage` 舊資料，營造出快取畫面。
2. **靜默失敗隱蔽**：當後端 RPC 發生錯誤或回傳失敗時，前端不跳錯而是默默載入舊快取，導致使用者誤以為「後端改了／清空了但前端卡住沒反應」。
3. **幽靈資料殘留**：在表單送出時提前寫入 `localStorage`，即使後端已清空或刪除，使用者重新整理依然被舊快取覆蓋。

### 改造目標
1. **全面移除快取讀寫**：徹底刪除前端 `Index.html` 中的所有 `localStorage.getItem` 與 `localStorage.setItem`。
2. **主動洗淨歷史快取**：在頁面載入時主動呼叫 `localStorage.removeItem('isms_backup_records')`，瞬間清除使用者電腦中殘留的舊混亂快取。
3. **建立 Single Source of Truth**：儀表板 100% 依據後端 `getDashboardData()` 回傳之 Google Sheet 真實資料渲染；若讀取失敗則明確顯示錯誤資訊與重試入口，絕不擅自拿過期資料頂替。
4. **健全後端防禦**：同步在 [`code.js`](file:///home/itsutaka/biobank_manager/backup_control/code.js) 加入 `ensureSheetColumnsExist_`（格線不足 22 欄時自動擴充，避免座標無效例外）與 `runAutoBackfillNow()`（供 GAS 編輯器直跑驗證）。

---

## 二、 系統資料流對比

```mermaid
flowchart TD
    subgraph 改造前（受快取污染）
        A1["重新整理頁面"] --> B1["先讀取 localStorage 舊快取<br/>(直接顯示混亂資料)"]
        B1 --> C1["發送 getDashboardData()"]
        C1 -->|若後端拋錯或延遲| D1["靜默吞噬錯誤<br/>繼續停留在舊快取畫面"]
    end

    subgraph 改造後（100% 後端真實資料）
        A2["重新整理頁面"] --> B2["主動清除舊快取<br/>localStorage.removeItem()"]
        B2 --> C2["顯示進度載入動畫<br/>發送 getDashboardData()"]
        C2 -->|成功| D2["以 Google Sheet 真實資料渲染卡片<br/>(不存任何 localStorage)"]
        C2 -->|失敗| E2["清空卡片並於 Console/畫面輸出明確錯誤<br/>不再拿舊資料矇騙使用者"]
    end
```

---

## 三、 User Review Required

> [!IMPORTANT]
> **離線／純本地瀏覽器開啟行為改變**：
> 移除 `localStorage` 後，若在「未部署於 GAS 容器」的純本機環境直接用瀏覽器雙擊打開 `Index.html` 檔案（此時無 `google.script.run` 環境），畫面將預設顯示「目前尚無備份設定」空狀態，而不再顯示本地模擬假資料。在正式 Google Apps Script 網頁環境下則 100% 正確連線讀取 Google Sheet。

---

## 四、 預計變更檔案與細節 (Proposed Changes)

### 1. 前端視圖 [`Index.html`](file:///home/itsutaka/biobank_manager/backup_control/Index.html)

#### [MODIFY] [`Index.html`](file:///home/itsutaka/biobank_manager/backup_control/Index.html)

1. **移除全域 `STORAGE_KEY` 與快取預載**：
   - 刪除 `const STORAGE_KEY = 'isms_backup_records';`
   - 將 `let appData = JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];` 改為純乾淨陣列 `let appData = [];`
   - 在 `DOMContentLoaded` 初始化時加入主動清除：`try { localStorage.removeItem('isms_backup_records'); } catch(e) {}`

2. **重構 [`loadDashboardData()`](file:///home/itsutaka/biobank_manager/backup_control/Index.html#L790-L823)**：
   - 移除所有 `localStorage.setItem` 與 `localStorage.getItem`。
   - 成功時：`appData = response.data;` ➔ `renderDashboard();`
   - 失敗時：`appData = [];` ➔ 清楚輸出 `console.warn` / `console.error` ➔ `renderDashboard();`

3. **重構 [`deleteSystem()`](file:///home/itsutaka/biobank_manager/backup_control/Index.html#L1004-L1049)**：
   - 移除成功回呼中的 `localStorage.setItem(STORAGE_KEY, ...)`。
   - 移除本地模擬分支中的 `localStorage.setItem(STORAGE_KEY, ...)`。

4. **重構 [`handleFormSubmit()`](file:///home/itsutaka/biobank_manager/backup_control/Index.html#L1690-L1730)**：
   - 移除送出前的 `localStorage.setItem(STORAGE_KEY, ...)`。
   - 移除成功回呼中更新 ID 時的 `localStorage.setItem(STORAGE_KEY, ...)`。
   - 送出成功後維持呼叫 `loadDashboardData({ showProgress: false })`，由後端 Google Sheet 重新同步最新真實資料。

---

### 2. 後端腳本 [`code.js`](file:///home/itsutaka/biobank_manager/backup_control/code.js)

#### [MODIFY] [`code.js`](file:///home/itsutaka/biobank_manager/backup_control/code.js)

1. **新增 `ensureSheetColumnsExist_` 邊界保護**：
   - 避免試算表格線總欄數小於 22 欄時呼叫 `getRange` 拋錯。
2. **新增 `runAutoBackfillNow()` 管理員免部署驗收工具**：
   - 可直接在 GAS 編輯器點擊「執行」，即時補填並在執行記錄輸出診斷日誌。

---

## 五、 驗證計畫 (Verification Plan)

### 自動化測試
執行專案全量測試套件：
```bash
node tests/backend-header-and-migration.test.cjs
node tests/schedule-picker.test.cjs
```
**預期結果**：20 項測試全數 PASS。

### 手動驗收步驟
1. **驗證本地存儲完全清空**：
   - 開啟網頁後按 `F12` ➔ 切換至「Application（應用程式）」➔「Local Storage（本機存放區）」。
   - 確認 `isms_backup_records` 已被完全刪除，且任何點擊、新增、刪除都不會再寫入此 Key。
2. **驗證 Google Sheet 為唯一真實來源**：
   - 在 Google Sheet 手動變更某系統名稱或清空欄位。
   - 重新整理網頁，確認畫面呈現的資料 100% 與 Google Sheet 實時一致，無任何舊資料殘留。
