# 備份管制系統：20 欄跨專案政策標準化暨自動遷移與前端升級實作計畫 [COMPLETED]

本計畫已全部實作完成並通過自動化測試驗證。將備份管制系統升級為標準 20 欄業務欄位（加表尾 2 欄技術欄位，共 22 欄），對齊 `data-termination-backup-termination` 專案之政策標準，並實現試算表零停機平滑自動遷移與前端 Smart Defaults 優化。

---

## User Review Required

> [!IMPORTANT]
> - **工作表結構升級**：主表 `備份管制` 將由舊版 16 欄平滑升級為 22 欄；歷程表 `備份管制_歷程` 由 19 欄升級為 25 欄。
> - **既有資料安全性**：後端內建自動遷移演算法（`migrateSheetIfNeeded_`），系統首次載入時會自動將舊列資料依欄名對齊新位置並填補預設值，**現存資料 100% 完整保留，無需手動搬移**。
> - **預設值統一**：依使用者指示，排程時間統一預設為 `每日 00:00`，處置方法統一預設為 `磁碟自動覆寫`。
> - **官方 Doc 報表格式保持**：依使用者指示，「一鍵生成表單」之 Google Doc 表格維持原始核定範本格式，不注入排程與處置文字。
> - **部署規範**：嚴格遵守 [`~/.agents/PLAYBOOK.md`](file:///home/itsutaka/.agents/PLAYBOOK.md) 鐵則 P1，變更完成後**絕不主動執行 `clasp push`**，由您確認並手動部署測試。

---

## Proposed Changes

### 後端模組 (Backend: Google Apps Script)

#### [MODIFY] [code.js](file:///home/itsutaka/biobank_manager/backup_control/code.js)

1. **更新常數定義**：
   - `EXPECTED_HEADERS`：擴充為 22 欄（第 6 欄為管理人員、新增 6 個新欄位、第 21/22 欄為 `submission_id` 與 `target_key`）。
   - `EXPECTED_HISTORY_HEADERS`：擴充為 25 欄（前置 3 個稽核欄位 ＋ 22 個快照欄位）。
2. **表頭動態索引解析器 (`resolveHeaderIndices_`)**：
   - 透過表頭文字動態解析出各欄位置（支援別名如 `系統ID`/`sourceSystemId`、`管理人員`/`managerName`、`本地_排程時間` 等），徹底擺脫硬編碼索引。
3. **零停機平滑自動遷移演算法 (`migrateSheetIfNeeded_`)**：
   - 當 `checkAndSetHeaders` 或 `checkAndSetHistoryHeaders` 偵測到現存工作表與預期表頭不完全一致時：
     - 若為空表，直接寫入新表頭並格式化。
     - 若有既有資料列，以舊表頭建立欄位字典對照，逐列重新映射為 22 欄（或歷程表 25 欄）格式。
     - 缺漏之新欄位自動填入智慧預設值（排程時間 `每日 00:00`、處置方法 `磁碟自動覆寫`、佐證種類 `備份系統執行Log` 等）。
     - 一次性整批覆寫回工作表，保全所有歷史紀錄。
4. **重構讀取與寫入模組**：
   - `getDashboardData()`：由動態索引讀取各欄值，標的物件新增 6 個新欄位屬性回傳前端。
   - `buildSheetRowData_(options)`：依 22 欄順序組裝陣列。
   - `isRowBelongsToSubmission_`：改用動態索引 `indices.submissionId` 比對。
   - `submitData()` 與 `deleteSubmission()`：適配 22 欄主表寫入與 25 欄歷程表封存。
   - `generateBackupControlDocument()`：動態對齊存放地點與保留代數欄位，保持原始匯出文字格式，不額外附加排程與處置文字。

---

### 前端介面與互動 (Frontend: Web App UI)

#### [MODIFY] [Index.html](file:///home/itsutaka/biobank_manager/backup_control/Index.html)

1. **標的卡片表單擴充 (`renderFormTargets`)**：
   - **本地備份區塊**：
     - 新增「本地排程時間」文字輸入框，預設值 `每日 00:00`。
     - 新增「本地處置方法」輸入選單，預設值 `磁碟自動覆寫`。
     - 新增「本地佐證種類」輸入選單，預設值 `備份系統執行Log`。
   - **異地存放區塊（勾選「是」展開）**：
     - 新增「異地排程時間」文字輸入框，預設值 `每日 00:00`。
     - 新增「異地處置方法」輸入選單，預設值 `磁碟自動覆寫`。
     - 新增「異地佐證種類」輸入選單，預設值 `磁帶消磁簽收單 / 雲端過期報表`。
2. **編輯模式載入 (`editSystem`) 與送出 (`handleFormSubmit`)**：
   - 載入時將這 6 個欄位正確回填表單（舊紀錄若無值自動帶入預設值）。
   - 送出時收集 6 個欄位封裝進 `formData.targets` 傳給後端。
3. **儀表板卡片呈現 (`renderTargetSetting`)**：
   - 在已啟用標的明細中，清晰展示排程時段、處置方法與佐證種類。

---

## Verification Plan

### Automated Tests
```bash
node --check code.js
```
確保語法完全正確、無語法衝突。

### Manual Verification（5 步驗證）

1. **試算表平滑遷移驗證**：
   - 重新整理網頁，後端自動將工作表升級為 22 欄，第 6 欄為管理人員，第 21/22 欄為技術鍵值，舊資料無遺失。
2. **新增表單預設值驗證**：
   - 打開「新增設定」，檢查排程時間皆為 `每日 00:00`、處置方法皆為 `磁碟自動覆寫`。
3. **編輯更新與歷程封存驗證**：
   - 編輯既有系統微調時段後送出，主表原地更新，歷程表記錄 25 欄「更新前封存」快照。
4. **一鍵生成表單格式驗證**：
   - 點擊「一鍵生成表單」，確認生成的 Doc 維持原始範本樣式。
5. **外部專案讀取相容性驗證**：
   - 確認主表前 20 欄完全符合 `data-termination-backup-termination` 之標準規格。
