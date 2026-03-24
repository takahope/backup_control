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
