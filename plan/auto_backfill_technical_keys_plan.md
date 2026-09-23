# 自動補填技術欄位 (submission_id / target_key) 實作計畫

## 問題分析與目標說明

### 背景與根本原因
在目前備份管制系統架構中，Google Sheet 主表（`備份管制`）的每一列記錄代表「單一系統的單一備份標的」（例如同一系統啟用資料庫與原始碼時會記錄為 2 列）。

當主表中存在部分歷史資料、手動匯入或手動插入之資料列，其第 21 欄（`submission_id`）或第 22 欄（`target_key`）為空白時，後端在執行 [`getDashboardData()`](file:///home/itsutaka/biobank_manager/backup_control/code.js#L591) 時會發生以下連鎖判斷錯誤：

1. **`submission_id` 空白導致第一層聚合被拆散**：
   - 舊邏輯退回使用 `legacyKey`（由系統名稱、等級、管理人員與時間戳記毫秒值組成）。
   - 若手動填寫或資料庫微小差異導致時間戳記不一致或為空，同一系統的多個備份標的會被切成多個獨立的 `submissionKey`。
2. **第二層聚合導致標的被洗掉**：
   - 第二層聚合在比對最新版本時，僅保留時間戳記最新的一個 submission，導致同一系統的其他已設定標的（如「資料庫」被「日誌檔」覆蓋）徹底消失。
3. **`target_key` 空白導致編輯與狀態對齊失敗**：
   - 無法精準與前端預設之 6 大標的常數對齊，造成編輯回填或保存時無法正向關聯。

### 解決目標
在後端 `getDashboardData()` 進入第一層聚合前，建立一組**輕量、安全、冪等（Idempotent）**的自動補填機制：
1. **即時檢測**：先在記憶體中掃描是否有空白的技術欄位；若無空白則 0 毫秒穿透（Fast-path，不執行任何寫入）。
2. **關係關聯**：
   - 同一系統（`系統名稱 + 系統等級 + 管理人員`）中，若已有列具備 `submission_id`，其餘空白列自動繼承該 UUID。
   - 若該系統全數皆無 `submission_id`，則由系統自動產生一組全新 UUID，並賦予該系統的所有標的列，確保所有標的牢牢綁定於同一提交批次。
3. **標的推導**：依據 `備份標的` 名稱自動正規化為標準代碼（`db`、`source`、`config`、`log`、`image`、`other`）。
4. **實體持久化**：自動批次寫回 Google Sheet（僅精準更新第 21、22 欄，絕不碰觸第 1~20 欄業務資料），使工作表實體資料庫永久固化。
5. **記憶體無縫接軌**：直接將補齊結果同步至目前運算中之資料陣列，使後續的第一層與第二層聚合 100% 正常運作。

---

## 系統架構與資料流

```mermaid
flowchart TD
    A["進入 getDashboardData()"] --> B["讀取 Google Sheet 主表 rows"]
    B --> C{"檢測是否有空白<br/>submission_id 或 target_key？"}
    C -->|全部齊全 (Fast-path)| F["直接進入第一層聚合 (Map by submissionKey)"]
    C -->|存在空白列| D["執行 autoBackfillTechnicalKeysIfNeeded_()"]
    
    subgraph 自動補填與固化流程
        D --> D1["同系統既有 UUID 收集<br/>(以 systemKey 建立映射)"]
        D1 --> D2["空白列補齊<br/>• submission_id: 繼承或分配新 UUID<br/>• target_key: 依標的名稱推導代碼"]
        D2 --> D3["精準批次寫入 Sheet 21/22 欄<br/>(業務資料 1~20 欄完全不動)"]
        D3 --> D4["同步更新記憶體中的 values 陣列"]
    end
    
    D4 --> F
    F --> G["第二層聚合 (最新版本去重)"]
    G --> H["回傳完整、無缺漏標的之系統卡片資料至前端"]
```

---

## User Review Required

> [!IMPORTANT]
> **寫入安全性保障**：
> 1. 自動補填僅會針對主表第 21 欄（`submission_id`）與第 22 欄（`target_key`）進行精準範圍寫入（`Range.setValues`），完全不會更動第 1 至 20 欄的使用者業務設定。
> 2. 若工作表全部資料均已具備技術欄位（正常運作狀態），會命中快速旁路（Fast-path），不產生任何 Google Sheets API 寫入請求，無額外延遲。

---

## Open Questions

目前設計已針對現有系統的資料結構與 CQRS 架構進行最佳化，暫無阻礙執行的未決問題。若您對於補填時的分組邏輯（以「系統名稱 + 系統等級 + 管理人員」為同一系統）有特殊需求，請隨時提出。

---

## Proposed Changes

### 後端模組（Google Apps Script）

#### [MODIFY] [`code.js`](file:///home/itsutaka/biobank_manager/backup_control/code.js)

1. **新增 `autoBackfillTechnicalKeysIfNeeded_` 函式**：
   - 負責技術欄位空值偵測。
   - 針對同一系統的所有標的賦予相同 UUID。
   - 根據標的名稱推導正規化 `target_key`。
   - 批次寫回 Google Sheet 第 21 與 22 欄，並即時更新記憶體 `values`。

2. **新增 `ensureTechnicalHeadersExist_` 輔助函式**：
   - 防禦性檢查：若工作表表頭尚未擴展第 21、22 欄表頭名稱，自動在第一列安全補上 `submission_id` 與 `target_key`，避免寫入無效區域。

3. **在 [`getDashboardData()`](file:///home/itsutaka/biobank_manager/backup_control/code.js#L609) 串接自動補填機制**：
   - 在讀取 `values` 之後、建立 `bySubmission` Map 之前，呼叫 `autoBackfillTechnicalKeysIfNeeded_(sheet, values, indices)`。

```javascript
// 程式碼異動預覽 (code.js)

// 1. 在 getDashboardData() 中串接
function getDashboardData() {
  ...
  const values = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();

  // 自動補填空白技術欄位（若有缺失則先補填寫回，再執行聚合）
  autoBackfillTechnicalKeysIfNeeded_(sheet, values, indices);

  // 將每一列（單一標的）轉換為 submission group，再取每個系統最新一次提交
  const bySubmission = new Map();
  ...
}

// 2. 自動補齊核心函式實作
function autoBackfillTechnicalKeysIfNeeded_(sheet, values, indices) {
  if (!values || values.length === 0) return values;

  // 1. 唯讀快篩：若無任何空白技術欄位，直接跳出（Fast-path）
  let hasMissing = false;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const sysName = (indices.systemName !== -1 && row[indices.systemName]) ? String(row[indices.systemName]).trim() : '';
    const tgtName = (indices.targetName !== -1 && row[indices.targetName]) ? String(row[indices.targetName]).trim() : '';
    if (!sysName || !tgtName) continue;

    const subId = (indices.submissionId !== -1 && row[indices.submissionId] !== undefined) ? String(row[indices.submissionId]).trim() : '';
    const tgtKey = (indices.targetKey !== -1 && row[indices.targetKey] !== undefined) ? normalizeTargetKey_(row[indices.targetKey]) : '';
    if (!subId || !tgtKey) {
      hasMissing = true;
      break;
    }
  }

  if (!hasMissing) return values;

  // 2. 表頭防護：確認表頭索引有效
  if (sheet && (indices.submissionId === -1 || indices.targetKey === -1)) {
    ensureTechnicalHeadersExist_(sheet, indices);
  }

  // 3. 第一階段：收集同系統已有之 UUID
  const systemToExistingUuid = new Map();
  values.forEach(function(row) {
    const sysName = (indices.systemName !== -1 && row[indices.systemName]) ? String(row[indices.systemName]).trim() : '';
    const sysLevel = (indices.systemLevel !== -1 && row[indices.systemLevel]) ? String(row[indices.systemLevel]).trim() : '';
    const mgrName = (indices.managerName !== -1 && row[indices.managerName]) ? String(row[indices.managerName]).trim() : '';
    const subId = (indices.submissionId !== -1 && row[indices.submissionId]) ? String(row[indices.submissionId]).trim() : '';

    if (sysName && subId) {
      const sysKey = [sysName, sysLevel, mgrName].join('||');
      if (!systemToExistingUuid.has(sysKey)) {
        systemToExistingUuid.set(sysKey, subId);
      }
    }
  });

  // 4. 第二階段：為空白列分配同一組 UUID 並推導 target_key
  const systemToAssignedUuid = new Map();
  let updatedCount = 0;

  values.forEach(function(row) {
    const sysName = (indices.systemName !== -1 && row[indices.systemName]) ? String(row[indices.systemName]).trim() : '';
    const sysLevel = (indices.systemLevel !== -1 && row[indices.systemLevel]) ? String(row[indices.systemLevel]).trim() : '';
    const mgrName = (indices.managerName !== -1 && row[indices.managerName]) ? String(row[indices.managerName]).trim() : '';
    const tgtName = (indices.targetName !== -1 && row[indices.targetName]) ? String(row[indices.targetName]).trim() : '';

    if (!sysName || !tgtName) return;

    const sysKey = [sysName, sysLevel, mgrName].join('||');
    let currentSubId = (indices.submissionId !== -1 && row[indices.submissionId] !== undefined) ? String(row[indices.submissionId]).trim() : '';
    let currentTargetKey = (indices.targetKey !== -1 && row[indices.targetKey] !== undefined) ? normalizeTargetKey_(row[indices.targetKey]) : '';
    let rowChanged = false;

    if (!currentSubId) {
      let targetUuid = systemToExistingUuid.get(sysKey);
      if (!targetUuid) {
        if (!systemToAssignedUuid.has(sysKey)) {
          systemToAssignedUuid.set(sysKey, Utilities.getUuid());
        }
        targetUuid = systemToAssignedUuid.get(sysKey);
      }
      if (indices.submissionId !== -1) {
        row[indices.submissionId] = targetUuid;
        rowChanged = true;
      }
    }

    if (!currentTargetKey) {
      const derivedKey = targetNameToKey_(tgtName) || 'other';
      if (indices.targetKey !== -1) {
        row[indices.targetKey] = derivedKey;
        rowChanged = true;
      }
    }

    if (rowChanged) updatedCount++;
  });

  // 5. 第三階段：精準批次寫回 Google Sheet 第 21 與 22 欄
  if (sheet && updatedCount > 0 && indices.submissionId !== -1 && indices.targetKey !== -1) {
    try {
      const numRows = values.length;
      const techUpdates = values.map(function(row) {
        return [row[indices.submissionId] || '', row[indices.targetKey] || ''];
      });

      if (indices.targetKey === indices.submissionId + 1) {
        sheet.getRange(2, indices.submissionId + 1, numRows, 2).setValues(techUpdates);
      } else {
        sheet.getRange(2, indices.submissionId + 1, numRows, 1).setValues(techUpdates.map(p => [p[0]]));
        sheet.getRange(2, indices.targetKey + 1, numRows, 1).setValues(techUpdates.map(p => [p[1]]));
      }
      console.log(`[autoBackfill] 自動補齊完成：更新了 ${updatedCount} 筆資料列之技術欄位`);
    } catch (err) {
      console.warn(`[autoBackfill] 寫回試算表異常（仍已在記憶體中補齊以維持聚合正確）：${err}`);
    }
  }

  return values;
}
```

---

### 自動化測試模組

#### [MODIFY] [`tests/backend-header-and-migration.test.cjs`](file:///home/itsutaka/biobank_manager/backup_control/tests/backend-header-and-migration.test.cjs)

新增以下自動化測試案例，確保新機制 100% 正確可靠：
1. **全滿資料不觸發寫入測試（Fast-path 驗收）**：模擬所有列皆有 `submission_id` 與 `target_key`，驗證 `setValues` 不會被呼叫。
2. **多標的系統補齊相同 UUID 測試**：同一個系統下有「資料庫」與「系統日誌檔」兩列均無 `submission_id`，驗證兩者補齊後分配完全相同的 UUID，且 `target_key` 分別正確為 `db` 與 `log`。
3. **部分列缺失繼承既有 UUID 測試**：同系統第一列已有 UUID，第二列為空白，驗證第二列正確繼承第一列之 UUID。
4. **第一層聚合完整性測試**：模擬前述缺失資料經過 `getDashboardData()`，驗證回傳的系統卡片物件中，其 `targets` 陣列包含完整的 2 個標的，無任何遺失。

---

## Verification Plan

### Automated Tests
於終端機執行現有與新增的後端測試套件：
```bash
node tests/backend-header-and-migration.test.cjs
```
**預期結果**：所有測試項目（原 5 項 + 新增測試）全數 Pass 通過。

### Manual Verification
1. **空白技術欄位測試**：
   - 在測試試算表的主表隨意挑選一個系統，手動將其 `submission_id` 與 `target_key` 刪除清空。
   - 重新整理網頁儀表板，確認該系統卡片依然完整呈現所有備份標的（未被拆散或縮水）。
   - 回到 Google Sheet 查看，確認原本被清空的欄位已自動補齊標準 UUID 與對應標的代碼。
2. **多標的正常編輯/儲存流程**：
   - 點擊該卡片的「編輯內容」，確認所有標的完整載入並成功打勾。
   - 點擊「更新儲存紀錄」，確認能正常在工作表中原地更新，不產生多餘重複列。
