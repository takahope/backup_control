# ISMS 備份管制系統：原地編輯更新與稽核歷程封存雙軌架構實作計畫

本計畫旨在修復編輯備份設定時誤新增資料列的問題，並建立符合 ISMS 合規要求之「備份管制_歷程」雙軌封存機制。

## Goal Description
- **問題**：使用者在前端點擊「編輯內容」修改既有系統備份設定並儲存後，後端未能原地更新既有資料列，而是將其當作全新資料追加至表尾，導致試算表資料膨脹且同一系統有多組重複設定。
- **根因**：[`code.js`](file:///home/itsutaka/biobank_manager/backup_control/code.js) 中的比對函式 `isRowBelongsToSubmission_` 誤將 Index 13（管理人員）當作 Index 14（`submission_id`），且 Legacy Key 比對參數索引錯位。
- **目標**：
  1. 修復欄位比對索引，實現主表「原地覆寫更新（In-place Update）」。
  2. 建立「雙軌制」稽核歷程機制：在修改覆寫或刪除既有列前，將舊列完整快照自動備份至「備份管制_歷程」工作表（標註操作者、時間與動作類型）。
  3. 升級前端「刪除」功能，完整串接後端 `deleteSubmission` API，實現刪除前封存與主表同步清除。

---

## User Review Required

> [!IMPORTANT]
> - 本次變更將在試算表中自動建立新工作表分頁：`備份管制_歷程`（若不存在會自動生成 19 欄表頭並凍結第 1 列）。
> - 主工作表 `備份管制` 的欄位結構（16 欄）維持不變，完全相容既有資料。
> - 依據跨平台總綱 [`~/.agents/PLAYBOOK.md`](file:///home/itsutaka/.agents/PLAYBOOK.md) 鐵則 P1，本任務**絕不主動執行 `clasp push`**，修改完成後將提供手動部署指引與驗收清單由您親自部署測試。

---

## Open Questions
- 目前無未決問題。先前討論已確認：
  1. 採用雙軌制（主表維護現況、歷程表純追加）。
  2. 採用「完整封存」策略（更新前封存、刪除前封存）。
  3. 前端 Web App 暫不需變更歷程彈窗，維持簡潔。
  4. 一併補齊後端刪除 API。

---

## Proposed Changes

### 後端模組 (Backend: Google Apps Script)

#### [MODIFY] [code.js](file:///home/itsutaka/biobank_manager/backup_control/code.js)

- **新增常數**：
  - `SHEET_HISTORY_NAME = '備份管制_歷程'`
  - `EXPECTED_HISTORY_HEADERS`（19 欄，前置 3 個稽核欄位：異動時間戳記、操作人員、異動動作，後接 16 欄原始快照）
- **新增歷程管理模組**：
  - `checkAndSetHistoryHeaders(sheet)`：自動檢查並格式化歷程工作表表頭。
  - `archiveSubmissionRows_(ss, rawRowsData, actionType, operatorEmail)`：將原始資料附加稽核資訊後批次寫入歷程表。
- **修復比對邏輯**：
  - 修正 `isRowBelongsToSubmission_`：讀取 `row[14]` 作為 `submission_id`，並對齊 Legacy Key 傳參（`row[2]`, `row[3]`, `row[13]`）。
- **升級 `submitData(formData)`**：
  - 在原地更新與刪除過期標的之前，若比對到既有列，先擷取舊列資料呼叫 `archiveSubmissionRows_(...)` 封存（動作標註：`更新前封存`）。
  - 維持原有的標的覆寫、移除未勾選標的、追加新標的邏輯。
- **新增 `deleteSubmission(submissionId)`**：
  - 驗證權限後，找出該 `submissionId` 的所有列。
  - 先呼叫 `archiveSubmissionRows_(...)` 封存（動作標註：`刪除前封存`）。
  - 將列號由大至小排序，自主表逐列刪除。

```javascript
// 主要程式碼邏輯示意
function isRowBelongsToSubmission_(row, submissionId) {
  const id = String(submissionId || '').trim();
  if (!id) return false;

  const rowSubmissionId = String(row[14] || '').trim(); // 修正為 Index 14
  if (rowSubmissionId && rowSubmissionId === id) {
    return true;
  }

  if (!rowSubmissionId) {
    const timestamp = row[0] ? new Date(row[0]) : new Date(0);
    const legacyId = buildLegacySubmissionKey_(timestamp, row[2], row[3], row[13]); // 修正欄位索引
    return legacyId === id;
  }

  return false;
}
```

---

### 前端介面與互動 (Frontend: Web App UI)

#### [MODIFY] [Index.html](file:///home/itsutaka/biobank_manager/backup_control/Index.html)

- **升級 `deleteSystem(id)`**：
  - 加入明確確認提示（提示將進行歷史封存）。
  - 呼叫 `google.script.run.deleteSubmission(id)`。
  - 配對 `withSuccessHandler` 與 `withFailureHandler`：成功時更新本地 `appData` 並重繪畫面；失敗時顯示錯誤訊息並觸發 `loadDashboardData()` 重新校正畫面。
- **表單儲存後連動**：
  - 在 `handleFormSubmit` 成功回呼時，自動呼叫 `loadDashboardData({ showProgress: false })` 背景同步後端主表最新資料，確保無伺服端快取差異。

---

## Verification Plan

### Automated Tests
1. **語法與靜態檢查**：
   ```bash
   node --check code.js
   ```
   確保伺服端無任何 JavaScript 語法錯誤或保留字衝突。

### Manual Verification
部署後執行以下 5 步情境驗證：

1. **全新新增測試**：
   - 新增一個測試系統（勾選 2 個標的）送出。
   - 檢查：主表「備份管制」新增 2 列；歷程表無新增；儀表板正確顯示。
2. **編輯更新測試（核心）**：
   - 編輯剛建立的測試系統，修改標的週期，並取消 1 個標的、勾選 1 個新標的送出。
   - 檢查：
     - 歷程表「備份管制_歷程」自動產生 2 列舊快照（標記 `更新前封存`，附帶時間與操作者 Email）。
     - 主表「備份管制」維持總列數 2 列（1 列更新、1 列移除、1 列新增），無多餘冗餘列。
3. **刪除測試**：
   - 點擊該系統的「刪除」按鈕並確認。
   - 檢查：
     - 歷程表「備份管制_歷程」產生 2 列舊快照（標記 `刪除前封存`）。
     - 主表「備份管制」該系統列完全清除。
     - 前端儀表板該卡片即時移除，重新整理後不復現。
4. **一鍵生成表單相容性**：
   - 點擊「一鍵生成表單」，確認生成的 Google Doc 僅讀取主表之有效設定。
