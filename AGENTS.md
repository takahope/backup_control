# Repository Guidelines

## Project Structure & Module Organization
This repository is a Google Apps Script web app for backup control management.

- `code.js`: Server-side Apps Script logic (Web App entrypoint, sheet I/O, permissions, asset lookup).
- `Index.html`: Frontend UI (dashboard, form flow, rendering, validation, loading/progress UX).
- `env.js`: Environment constants (for example `ASSETS_SPREADSHEET_ID` for external sheets).
- `appsscript.json`: Apps Script manifest.
- `.clasp.json`: Local clasp project mapping.

Keep backend concerns in `code.js` and UI behavior in `Index.html`; avoid mixing sheet access logic into HTML.

## Build, Test, and Development Commands
This project is deployed with `clasp` (Google Apps Script CLI).

- `clasp status`: Show local changes vs remote project.
- `clasp push`: Upload local files to Apps Script project.
- `clasp pull`: Sync remote project changes to local.
- `clasp open`: Open the script project in browser.
- `clasp deployments`: List deployment versions.

There is no automated build pipeline in this repo. Validate by running the deployed Web App and checking sheet read/write flows.

## Coding Style & Naming Conventions
- Use 2-space indentation and semicolons in JavaScript.
- Prefer `const`/`let`; avoid `var`.
- Use descriptive camelCase for functions/variables (e.g., `getDashboardData`, `loadAssetOptions`).
- Constants use upper snake case (e.g., `SHEET_NAME`, `SHEET_PERMISSIONS_NAME`).
- For frontend safety, do not inject raw user content into `innerHTML`; escape content or use DOM APIs.

## Testing Guidelines
No formal test framework is configured. Use scenario-based manual tests:

1. Authorized user can load dashboard, submit form, and edit/delete records.
2. Unauthorized user is blocked (permission sheet `權限` column B).
3. Asset dropdown loads from external sheet (`資訊資產`) and filters allowed types.
4. Loading bar behavior: stalls at 90% until data load completion, then reaches 100%.

## Commit & Pull Request Guidelines
Follow Conventional Commit style used in history:
- `feat: ...`
- `fix: ...`
- `refactor: ...`

PRs should include:
- Clear summary of user-visible behavior changes.
- Affected files and sheet dependencies.
- Manual test evidence (steps + result, screenshots for UI changes).
- Notes for config changes (especially `env.js` and spreadsheet permissions).

## Security & Configuration Tips
- Never commit real spreadsheet IDs or secrets to shared branches.
- Ensure `ASSETS_SPREADSHEET_ID` points to the correct external file and access is granted.
- Keep permission checks server-side for every data API, not only `doGet`.

## Incident & Change Log
- 2026-11-12: 本地備份勾選「保留所有版本（全部保留）」時，停用「到期銷毀方法」與「銷毀確認方式」兩個下拉選單並顯示「無」（全部保留代表不會銷毀），寫入試算表的值為「無」，比照異地備份未啟用時的做法。前端新增 `syncDispositionLockWithRetention_`（鎖定時以暫時的「無」選項顯示並記住原選擇，取消勾選恢復原選擇，未曾選擇則回預設值），由 `toggleLocalRetentionMode` 觸發；`editSystem` 回填銷毀欄位後會重新套用鎖定；`handleFormSubmit` 全部保留時兩欄送「無」；取消勾選標的與 `resetForm` 皆會解除鎖定。後端新增 `isRetentionAll_`，`buildSheetRowData_` 於伺服器端強制寫入「無」（不信任前端；同時修正 `全部保留` 字串被誤判為一般代數的問題）。異地備份不受影響；表頭與資料契約不變，無需執行 `runDataMigration`；已存在的舊「全部保留」資料維持原值，待下次編輯儲存才會變成「無」；`GLOSSARY.md` 新增「全部保留」；全套 67 項測試通過。
- 2026-11-09: 本地備份「完整備份」由必填改為選填，規則為「完整備份、差異備份二擇一必填」，以支援僅執行差異備份的標的。前端移除完整備份標籤星號與 `required`，數值欄 `min` 改 0 並提示「0代表無」，新增 `updateFullPickerVisibility`（週期 > 0 才顯示完整軌排程選擇器，與差異軌對稱），送出前以 `validateLocalBackupCycles` 阻擋兩者皆空；排程字串新增僅差異格式 `差異:每日 02:00`（`composeDualTrackSchedule`／`parseScheduleString` 同步支援，舊資料無前綴字串仍視為完整軌，免遷移）。後端新增 `isCycleEnabled_`，`buildSheetRowData_` 將完整週期空白或 0 寫為「無」（不再產生殘缺的「 週」），`submitData` 以 `assertLocalBackupCyclePresent_` 於寫入前拒絕兩者皆空（系統日誌檔除外）。新增 `GLOSSARY.md`；表頭與資料契約不變，無需執行 `runDataMigration`；全套 54 項測試通過。
- 2026-09-23: 修復名詞升級後 assertSheetSchema_ 因嚴格字串比對拒絕未遷移試算表（處置/佐證舊表頭）導致前端顯示空狀態的缺陷。在後端引入 isHeaderCompatible_ 允許新舊表頭別名向下相容通過，確保舊版試算表免執行 runDataMigration 即可零停機平滑讀取；前端 loadDashboardData 補強 dashboardLoadError 狀態與重試入口，消除錯誤被靜默吞沒的隱患；全套 39 項測試全數通過。
- 2026-09-23: 完成備份管制表「到期銷毀方法」與「銷毀確認方式」名詞升級與雙軌子面板視覺重構。後端升級 EXPECTED_HEADERS（12、13、19、20 欄）與歷程表頭，resolveHeaderIndices_ 與 parseTargetFromRow_ 完整支援新舊欄位別名相容，migrateSheetIfNeeded_ 平滑升級資料表；前端儀表板標的卡片重構為「本地（藍灰）／異地（紫羅蘭）」雙軌獨立子面板，內部標籤徹底去前綴化，未啟用異地時呈現優雅收合徽章；表單彈窗同步更新為直覺的到期銷毀方法與銷毀確認方式標籤；後端與生命週期測試全套 38 項全數通過。
- 2026-09-23: 徹底移除前端 Index.html 的 localStorage 快取機制（isms_backup_records），載入時主動清除歷史殘留快取，建立 100% 以 Google Sheet 後端為唯一真實來源（Single Source of Truth）的無快取架構；後端加入 ensureSheetColumnsExist_ 格線邊界保護與 runAutoBackfillNow 免部署直驗工具，全套 21 項測試全數通過。
- 2026-09-23: 實作第一層聚合前自動補填技術欄位 (submission_id, target_key) 機制。在 getDashboardData 建立 autoBackfillTechnicalKeysIfNeeded_ 輕量偵測旁路（Fast-path 0 寫入），當發現空白列時自動依同系統關聯性補齊 UUID 並推導 target_key，精準批次持久化至主表第 21、22 欄（不碰觸 1~20 欄業務資料），徹底杜絕同一系統標的被拆散或覆蓋的聚合錯誤；後端單元測試擴充至 9 項全數通過。
- 2026-09-22: 修復每次載入頁面時觸發破壞性遷移導致後端資料錯位與資料洗成「無」的嚴重缺陷。徹底解耦 CQRS，從 `getDashboardData` 與 `deleteSubmission` 中移除 `checkAndSetHeaders`，主表與歷程表改為僅在全空時以 `ensureHeadersExist_` 寫入表頭；修正 `resolveHeaderIndices_` 缺漏欄位改為回傳 `-1` 徹底杜絕索引錯位；修復 `migrateSheetIfNeeded_` 在 26 欄工作表中的表頭比對判定；加強 `projectBusinessRow_` 異地資料與布林值保護；新增專用技術欄位安全補齊工具 `backfillMissingTechnicalKeys`；完成後端 5 項單元測試。
- 2026-09-21: 完成智慧排程選擇器表單生命週期整合（編輯回填 Hydration、表單重置 Reset、送出驗證）。在 editSystem 實作反向解析與歷史字串洗淨正規化，resetForm 正確重設所有標的與各軌控制項預設值（本地「每日 00:00」、異地「無」），handleFormSubmit 送出前強制執行 syncTargetSchedule；完成全系統 11 項單元與生命週期整合測試。
- 2026-09-20: 實作跨專案資料契約「備份管制表 20 欄（含 2 欄技術欄位共 22 欄）」升級；後端建立 Header-Aware 動態表頭解析與零停機平滑資料遷移機制，並新增管理員專用一鍵安全遷移工具 `runDataMigration`（含事前自動快照備份）；前端表單與儀表板擴充本地／異地排程時間（預設 每日 00:00）、處置方法（預設 磁碟自動覆寫）及佐證種類；文件生成維護原有核定格式不改動。
- 2026-09-19: 修復編輯時 `isRowBelongsToSubmission_` 因欄位索引錯位（Index 13 誤判為 submission_id）導致誤將更新判為新增的問題；實作雙軌制歷程封存，在主表原地更新與刪除前先將舊資料快照寫入「備份管制_歷程」工作表，並補齊後端 `deleteSubmission` API。
