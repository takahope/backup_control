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
- 2026-09-21: 完成智慧排程選擇器表單生命週期整合（編輯回填 Hydration、表單重置 Reset、送出驗證）。在 editSystem 實作反向解析與歷史字串洗淨正規化，resetForm 正確重設所有標的與各軌控制項預設值（本地「每日 00:00」、異地「無」），handleFormSubmit 送出前強制執行 syncTargetSchedule；完成全系統 11 項單元與生命週期整合測試。
- 2026-09-20: 實作跨專案資料契約「備份管制表 20 欄（含 2 欄技術欄位共 22 欄）」升級；後端建立 Header-Aware 動態表頭解析與零停機平滑資料遷移機制，並新增管理員專用一鍵安全遷移工具 `runDataMigration`（含事前自動快照備份）；前端表單與儀表板擴充本地／異地排程時間（預設 每日 00:00）、處置方法（預設 磁碟自動覆寫）及佐證種類；文件生成維護原有核定格式不改動。
- 2026-09-19: 修復編輯時 `isRowBelongsToSubmission_` 因欄位索引錯位（Index 13 誤判為 submission_id）導致誤將更新判為新增的問題；實作雙軌制歷程封存，在主表原地更新與刪除前先將舊資料快照寫入「備份管制_歷程」工作表，並補齊後端 `deleteSubmission` API。
