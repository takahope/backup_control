// Code.gs - 處理網頁請求與寫入 Google Sheets
const SHEET_NAME = '備份管制'; // 請確認您的 Google Sheet 分頁名稱是否為此
const SHEET_HISTORY_NAME = '備份管制_歷程';
const SHEET_ASSETS_NAME = '資訊資產清單';
const SHEET_PERMISSIONS_NAME = '權限';

// 定義正確的表頭陣列（20 欄業務標準 + 2 欄技術識別欄位，共 22 欄）
const EXPECTED_HEADERS = [
  '時間戳記', '系統ID', '系統名稱', '系統等級', '備份標的', '管理人員',
  '本地_完整週期', '本地_差異週期', '本地_保留代數', '本地_存放地點',
  '本地_排程時間', '本地_處置方法', '本地_佐證種類',
  '是否異地', '異地_週期', '異地_地點', '異地_保留代數',
  '異地_排程時間', '異地_處置方法', '異地_佐證種類',
  'submission_id', 'target_key'
];

// 定義歷程表頭陣列（前置 3 個稽核欄位 + 原始 22 欄快照，共 25 欄）
const EXPECTED_HISTORY_HEADERS = [
  '異動時間戳記', '操作人員', '異動動作',
  '時間戳記', '系統ID', '系統名稱', '系統等級', '備份標的', '管理人員',
  '本地_完整週期', '本地_差異週期', '本地_保留代數', '本地_存放地點',
  '本地_排程時間', '本地_處置方法', '本地_佐證種類',
  '是否異地', '異地_週期', '異地_地點', '異地_保留代數',
  '異地_排程時間', '異地_處置方法', '異地_佐證種類',
  'submission_id', 'target_key'
];

// 1. 輸出前端 HTML 網頁
function doGet(e) {
  const permission = checkAppPermission_();
  if (!permission.allowed) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>無權限存取</title></head><body style="font-family:sans-serif;background:#f3f4f6;padding:24px;"><div style="max-width:680px;margin:0 auto;background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:20px;"><h2 style="margin:0 0 12px 0;color:#b91c1c;">無權限存取</h2><p style="color:#374151;line-height:1.7;">您目前沒有此系統的使用權限。請聯繫系統管理員，將您的帳號加入「權限」工作表 B 欄白名單。</p><p style="color:#6b7280;font-size:13px;margin-top:12px;">' + sanitizeHtml_(permission.message || '') + '</p></div></body></html>'
    ).setTitle('無權限存取');
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ISMS 備份管制表單系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 動態解析表頭名稱索引（Header-Aware Resolver），支援別名與安全預設值
function resolveHeaderIndices_(headers) {
  const hList = (headers || []).map(function(h) { return String(h || '').trim(); });
  function find_(aliases) {
    for (let i = 0; i < aliases.length; i++) {
      const idx = hList.indexOf(aliases[i]);
      if (idx !== -1) return idx;
    }
    return -1;
  }
  return {
    timestamp: find_(['時間戳記', 'timestamp']),
    systemId: find_(['系統ID', '系統代碼', 'sourceSystemId']),
    systemName: find_(['系統名稱', 'systemName']),
    systemLevel: find_(['系統等級', 'systemLevel']),
    targetName: find_(['備份標的', '標的名稱', 'targetName']),
    managerName: find_(['管理人員', '系統管理人員', '備份負責人', 'managerName']),
    localFull: find_(['本地_完整週期', '本地完整週期', '完整週期']),
    localDiff: find_(['本地_差異週期', '本地差異週期', '差異週期', '差異備份週期', '差異']),
    localRetention: find_(['本地_保留代數', '本地保留代數', '保留代數']),
    localLocation: find_(['本地_存放地點', '本地存放地點', '存放地點']),
    localScheduleTime: find_(['本地_排程時間', '本地排程時間']),
    localDispositionMethod: find_(['本地_處置方法', '本地處置方法']),
    localEvidenceRequirements: find_(['本地_佐證種類', '本地佐證種類']),
    isOffsite: find_(['是否異地', '異地備份', '是否異地備份']),
    offsiteCycle: find_(['異地_週期', '異地週期', '異地備份週期']),
    offsiteLocation: find_(['異地_地點', '異地存放地點', '異地地點']),
    offsiteRetention: find_(['異地_保留代數', '異地保留代數']),
    offsiteScheduleTime: find_(['異地_排程時間', '異地排程時間']),
    offsiteDispositionMethod: find_(['異地_處置方法', '異地處置方法']),
    offsiteEvidenceRequirements: find_(['異地_佐證種類', '異地佐證種類']),
    submissionId: find_(['submission_id', 'submissionId']),
    targetKey: find_(['target_key', 'targetKey'])
  };
}

// 格式化表頭樣式
function formatHeaderRow_(sheet, colCount, isHistory) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setFontWeight("bold");
  range.setBackground(isHistory ? "#fef3c7" : "#e5e7eb");
  sheet.setFrozenRows(1);
}

// 零停機平滑資料遷移演算法：偵測並自動重排既有資料列、補齊預設值（僅在 runDataMigration 主動執行）
function migrateSheetIfNeeded_(sheet, expectedHeaders, isHistory) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    formatHeaderRow_(sheet, expectedHeaders.length, isHistory);
    return;
  }

  // 嘗試取得第一列表頭比對：只要前 expectedHeaders.length 欄完全相符，即視為表頭正確，絕不重複遷移
  const checkWidth = Math.max(lastCol, expectedHeaders.length);
  const currentHeaders = sheet.getRange(1, 1, 1, checkWidth).getValues()[0];
  let isMatch = true;
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (String(currentHeaders[i] || '').trim() !== expectedHeaders[i]) {
      isMatch = false;
      break;
    }
  }

  // 若表頭已完全一致，直接返回
  if (isMatch) return;

  // 需要資料遷移：讀取現有所有資料列
  const existingValues = sheet.getRange(1, 1, lastRow, Math.max(lastCol, 1)).getValues();
  const oldHeaders = existingValues[0].map(function(h) { return String(h || '').trim(); });

  function getOldVal(row, aliases) {
    for (let i = 0; i < aliases.length; i++) {
      const idx = oldHeaders.indexOf(aliases[i]);
      if (idx !== -1 && row[idx] !== undefined && row[idx] !== null && String(row[idx]).trim() !== '') {
        return row[idx];
      }
    }
    return '';
  }

  const newValues = [];
  newValues.push(expectedHeaders);

  for (let r = 1; r < existingValues.length; r++) {
    const row = existingValues[r];
    const isRowEmpty = row.every(function(cell) { return cell === '' || cell === null || cell === undefined; });
    if (isRowEmpty) continue;

    if (isHistory) {
      // 歷程表保留前置 3 個稽核欄位
      const actionTime = getOldVal(row, ['異動時間戳記']) || row[0] || new Date();
      const operator = getOldVal(row, ['操作人員']) || (oldHeaders.indexOf('操作人員') !== -1 ? row[oldHeaders.indexOf('操作人員')] : '');
      const action = getOldVal(row, ['異動動作']) || (oldHeaders.indexOf('異動動作') !== -1 ? row[oldHeaders.indexOf('異動動作')] : '歷史封存');
      const baseRow = projectBusinessRow_(row, getOldVal);
      newValues.push([actionTime, operator, action].concat(baseRow));
    } else {
      const baseRow = projectBusinessRow_(row, getOldVal);
      newValues.push(baseRow);
    }
  }

  // 清空現有內容並寫入投影後之新資料
  sheet.clearContents();
  sheet.getRange(1, 1, newValues.length, expectedHeaders.length).setValues(newValues);
  formatHeaderRow_(sheet, expectedHeaders.length, isHistory);
}

// 將舊資料列重新投影映射為 22 欄標準業務格式（安全保護，不盲目覆寫為「無」）
function projectBusinessRow_(row, getOldVal) {
  const timestamp = getOldVal(row, ['時間戳記', 'timestamp']) || new Date();
  const systemId = getOldVal(row, ['系統ID', '系統代碼', 'sourceSystemId']);
  const systemName = getOldVal(row, ['系統名稱', 'systemName']);
  const systemLevel = getOldVal(row, ['系統等級', 'systemLevel']) || '一般系統';
  const targetName = getOldVal(row, ['備份標的', '標的名稱', 'targetName']);
  const managerName = getOldVal(row, ['管理人員', '系統管理人員', '備份負責人', 'managerName']);

  const localFull = getOldVal(row, ['本地_完整週期', '本地完整週期', '完整週期']) || '1 週';
  const localDiff = getOldVal(row, ['本地_差異週期', '本地差異週期', '差異週期', '差異備份週期', '差異']) || '無';
  const localRetention = getOldVal(row, ['本地_保留代數', '本地保留代數', '保留代數']) || '4 代';
  const localLocation = getOldVal(row, ['本地_存放地點', '本地存放地點', '存放地點']);
  const localScheduleTime = getOldVal(row, ['本地_排程時間', '本地排程時間']) || '每日 00:00';
  const localDispositionMethod = getOldVal(row, ['本地_處置方法', '本地處置方法']) || '磁碟自動覆寫';
  const localEvidenceRequirements = getOldVal(row, ['本地_佐證種類', '本地佐證種類']) || '備份系統執行Log';

  const rawIsOffsite = String(getOldVal(row, ['是否異地', '異地備份', '是否異地備份']) || '').trim();
  const rawOffsiteCycle = getOldVal(row, ['異地_週期', '異地週期', '異地備份週期']);
  const rawOffsiteLocation = getOldVal(row, ['異地_地點', '異地存放地點', '異地地點']);
  const rawOffsiteRetention = getOldVal(row, ['異地_保留代數', '異地保留代數']);
  const rawOffsiteScheduleTime = getOldVal(row, ['異地_排程時間', '異地排程時間']);
  const rawOffsiteDispositionMethod = getOldVal(row, ['異地_處置方法', '異地處置方法']);
  const rawOffsiteEvidenceRequirements = getOldVal(row, ['異地_佐證種類', '異地佐證種類']);

  const hasOffsiteData = (rawOffsiteLocation && rawOffsiteLocation !== '無' && rawOffsiteLocation !== 'N/A') ||
                         (rawOffsiteCycle && rawOffsiteCycle !== '無' && rawOffsiteCycle !== 'N/A');
  const isOffsite = (rawIsOffsite === '是' || rawIsOffsite.toLowerCase() === 'true' || rawIsOffsite === '1' || rawIsOffsite.toUpperCase() === 'Y' || hasOffsiteData)
    ? '是'
    : (rawIsOffsite || '否');

  const offsiteCycle = isOffsite === '是' ? (rawOffsiteCycle || '1 週') : (rawOffsiteCycle || '無');
  const offsiteLocation = isOffsite === '是' ? (rawOffsiteLocation || '') : (rawOffsiteLocation || '無');
  const offsiteRetention = isOffsite === '是' ? (rawOffsiteRetention || '4 代') : (rawOffsiteRetention || '無');
  const offsiteScheduleTime = isOffsite === '是' ? (rawOffsiteScheduleTime || '每日 00:00') : (rawOffsiteScheduleTime || '無');
  const offsiteDispositionMethod = isOffsite === '是' ? (rawOffsiteDispositionMethod || '磁碟自動覆寫') : (rawOffsiteDispositionMethod || '無');
  const offsiteEvidenceRequirements = isOffsite === '是' ? (rawOffsiteEvidenceRequirements || '磁帶消磁簽收單 / 雲端過期報表') : (rawOffsiteEvidenceRequirements || '無');

  const submissionId = getOldVal(row, ['submission_id', 'submissionId']) || Utilities.getUuid();
  const targetKey = getOldVal(row, ['target_key', 'targetKey']) || targetNameToKey_(targetName) || 'other';

  return [
    timestamp, systemId, systemName, systemLevel, targetName, managerName,
    localFull, localDiff, localRetention, localLocation,
    localScheduleTime, localDispositionMethod, localEvidenceRequirements,
    isOffsite, offsiteCycle, offsiteLocation, offsiteRetention,
    offsiteScheduleTime, offsiteDispositionMethod, offsiteEvidenceRequirements,
    submissionId, targetKey
  ];
}

// 僅在工作表完全為空時初始化表頭，絕不破壞既有資料
function ensureHeadersExist_(sheet, expectedHeaders, isHistory) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    formatHeaderRow_(sheet, expectedHeaders.length, isHistory);
  }
}

// 檢測並自動填寫主工作表表頭（提供 runDataMigration 主動遷移使用）
function checkAndSetHeaders(sheet) {
  migrateSheetIfNeeded_(sheet, EXPECTED_HEADERS, false);
}

// 檢測並自動填寫歷程工作表表頭（提供 runDataMigration 主動遷移使用）
function checkAndSetHistoryHeaders(sheet) {
  migrateSheetIfNeeded_(sheet, EXPECTED_HISTORY_HEADERS, true);
}

/**
 * 管理員主動安全遷移工具（可直接在 Apps Script 編輯器選單點選執行）
 * 1. 自動建立原工作表快照分頁（如 備份管制_backup_yyyyMMdd_HHmmss）
 * 2. 執行主表平滑資料遷移（升級至 22 欄）
 * 3. 執行歷程表平滑資料遷移（升級至 25 欄）
 * 4. 輸出詳細報告日誌與 UI 提示
 */
function runDataMigration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const report = [];

  // 1. 遷移主工作表（備份管制）
  const mainSheet = ss.getSheetByName(SHEET_NAME);
  if (mainSheet) {
    const backupName = `${SHEET_NAME}_backup_${timeStr}`;
    mainSheet.copyTo(ss).setName(backupName);
    report.push(`【主表備份】已建立快照分頁：${backupName}`);

    const beforeLastCol = mainSheet.getLastColumn();
    const rowCount = Math.max(mainSheet.getLastRow() - 1, 0);
    checkAndSetHeaders(mainSheet);
    report.push(`【主表遷移】「${SHEET_NAME}」由 ${beforeLastCol} 欄升級為 ${EXPECTED_HEADERS.length} 欄（共處理 ${rowCount} 筆資料列）。`);
  } else {
    report.push(`【主表略過】未找到工作表「${SHEET_NAME}」。`);
  }

  // 2. 遷移歷程工作表（備份管制_歷程）
  const historySheet = ss.getSheetByName(SHEET_HISTORY_NAME);
  if (historySheet) {
    const backupHistoryName = `${SHEET_HISTORY_NAME}_backup_${timeStr}`;
    historySheet.copyTo(ss).setName(backupHistoryName);
    report.push(`【歷程備份】已建立快照分頁：${backupHistoryName}`);

    const beforeHistCol = historySheet.getLastColumn();
    const histRowCount = Math.max(historySheet.getLastRow() - 1, 0);
    checkAndSetHistoryHeaders(historySheet);
    report.push(`【歷程遷移】「${SHEET_HISTORY_NAME}」由 ${beforeHistCol} 欄升級為 ${EXPECTED_HISTORY_HEADERS.length} 欄（共處理 ${histRowCount} 筆資料列）。`);
  } else {
    report.push(`【歷程略過】未找到歷程工作表「${SHEET_HISTORY_NAME}」。`);
  }

  const summary = report.join('\n');
  console.log(summary);
  try {
    SpreadsheetApp.getUi().alert('資料庫遷移完成', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // 若非在試算表容器介面中執行，略過 UI 彈窗
  }
  return summary;
}

// 將異動前列快照寫入歷程表
function archiveSubmissionRows_(ss, rawRowsData, actionType, operatorEmail) {
  if (!rawRowsData || rawRowsData.length === 0) return;

  let historySheet = ss.getSheetByName(SHEET_HISTORY_NAME);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_HISTORY_NAME);
  }
  ensureHeadersExist_(historySheet, EXPECTED_HISTORY_HEADERS, true);

  const now = new Date();
  const operator = String(operatorEmail || '').trim();
  const action = String(actionType || '異動封存').trim();

  const historyRows = rawRowsData.map(function(row) {
    return [now, operator, action].concat(row);
  });

  let startRow = historySheet.getLastRow() + 1;
  if (startRow < 2) startRow = 2;
  historySheet.getRange(startRow, 1, historyRows.length, historyRows[0].length).setValues(historyRows);
}

// 2. 接收前端資料並寫入 Sheet
function submitData(formData) {
  try {
    const permission = checkAppPermission_();
    if (!permission.allowed) {
      return { success: false, message: permission.message || '無權限執行此操作。' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("找不到指定的工作表，請確認名稱是否為 '" + SHEET_NAME + "'。");

    // 僅在全新工作表時寫入表頭，保護既有資料列
    ensureHeadersExist_(sheet, EXPECTED_HEADERS, false);

    // 檢查是否有傳入標的陣列
    if (!formData.targets || formData.targets.length === 0) {
      throw new Error("沒有收到任何備份標的資料！");
    }

    const timestamp = new Date();
    let systemId = '';
    let parsedSystemName = String(formData.systemName || '').trim();
    if (parsedSystemName.indexOf(' - ') !== -1) {
      const parts = parsedSystemName.split(' - ');
      systemId = parts[0].trim();
      parsedSystemName = parts.slice(1).join(' - ').trim();
    }
    const systemName = parsedSystemName;
    const systemLevel = String(formData.systemLevel || '').trim();
    const managerName = String(formData.managerName || '').trim();
    const requestedId = String(formData.id || '').trim();
    const submissionId = requestedId || Utilities.getUuid();

    const payloadRowsByTargetKey = {};
    formData.targets.forEach(function(target) {
      const targetName = String(target.targetName || '').trim();
      const targetKey = normalizeTargetKey_(target.targetKey) || targetNameToKey_(targetName);
      if (!targetKey) {
        throw new Error('備份標的缺少可識別 target_key。');
      }
      payloadRowsByTargetKey[targetKey] = buildSheetRowData_({
        timestamp: timestamp,
        systemId: systemId,
        systemName: systemName,
        systemLevel: systemLevel,
        managerName: managerName,
        submissionId: submissionId,
        targetKey: targetKey,
        target: target
      });
    });

    const payloadTargetKeys = Object.keys(payloadRowsByTargetKey);
    const existingRowsByTargetKey = {};
    const existingRawRows = [];
    const rowsToDelete = [];
    let updatedCount = 0;
    let insertedCount = 0;
    let deletedCount = 0;

    const lastRow = sheet.getLastRow();
    const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const indices = resolveHeaderIndices_(headerValues);

    if (lastRow > 1) {
      const colCount = Math.max(sheet.getLastColumn(), EXPECTED_HEADERS.length);
      const values = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
      values.forEach(function(row, idx) {
        const rowNumber = idx + 2;
        if (!isRowBelongsToSubmission_(row, submissionId, indices)) return;

        existingRawRows.push(row);

        const targetKey = normalizeTargetKey_(row[indices.targetKey]) || targetNameToKey_(row[indices.targetName]);
        if (!targetKey) return;
        if (!existingRowsByTargetKey[targetKey]) {
          existingRowsByTargetKey[targetKey] = [];
        }
        existingRowsByTargetKey[targetKey].push(rowNumber);
      });
    }

    // 若有匹配到既有列（編輯更新模式），在執行覆寫/刪除前先將舊列完整封存至歷程表
    if (existingRawRows.length > 0) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      archiveSubmissionRows_(ss, existingRawRows, '更新前封存', permission.email || '');
    }

    // 更新既有列或標記為待新增
    const rowsToAppend = [];
    payloadTargetKeys.forEach(function(targetKey) {
      const existingRows = existingRowsByTargetKey[targetKey] || [];
      const rowData = payloadRowsByTargetKey[targetKey];

      if (existingRows.length > 0) {
        const keepRow = existingRows.shift();
        sheet.getRange(keepRow, 1, 1, rowData.length).setValues([rowData]);
        updatedCount++;
        // 同一 submission/target 若有重複列，保留第一筆其餘刪除
        existingRows.forEach(function(rowNum) { rowsToDelete.push(rowNum); });
      } else {
        rowsToAppend.push(rowData);
      }

      delete existingRowsByTargetKey[targetKey];
    });

    // 編輯時取消勾選的標的：刪除舊列
    Object.keys(existingRowsByTargetKey).forEach(function(targetKey) {
      const staleRows = existingRowsByTargetKey[targetKey] || [];
      staleRows.forEach(function(rowNum) { rowsToDelete.push(rowNum); });
    });

    if (rowsToDelete.length > 0) {
      const uniqueRows = Array.from(new Set(rowsToDelete)).sort(function(a, b) { return b - a; });
      uniqueRows.forEach(function(rowNum) {
        sheet.deleteRow(rowNum);
        deletedCount++;
      });
    }

    if (rowsToAppend.length > 0) {
      let startRow = sheet.getLastRow() + 1;
      if (startRow < 2) startRow = 2;
      sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      insertedCount += rowsToAppend.length;
    }

    // 回傳成功訊息與處理筆數給前端
    const resultMsg = existingRawRows.length > 0
      ? `同步完成：更新 ${updatedCount} 筆、新增 ${insertedCount} 筆、移除 ${deletedCount} 筆，已封存舊歷程。`
      : `新增完成：共新增 ${insertedCount} 筆備份標的設定。`;

    return { 
      success: true, 
      submissionId: submissionId,
      message: resultMsg
    };
    
  } catch (error) {
    // 發生錯誤時將詳細資訊回傳給前端
    return { success: false, message: error.toString() };
  }
}

// 2.1 接收前端請求刪除特定系統之備份設定（先封存後刪除）
function deleteSubmission(submissionId) {
  try {
    const permission = checkAppPermission_();
    if (!permission.allowed) {
      return { success: false, message: permission.message || '無權限執行此操作。' };
    }

    const id = String(submissionId || '').trim();
    if (!id) {
      return { success: false, message: '未指定要刪除的系統編號 (submission_id)。' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("找不到指定的工作表，請確認名稱是否為 '" + SHEET_NAME + "'。");

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, message: '資料表中無任何資料可刪除。' };
    }

    const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const indices = resolveHeaderIndices_(headerValues);
    const colCount = Math.max(sheet.getLastColumn(), EXPECTED_HEADERS.length);
    const values = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
    const rowsToDelete = [];
    const rawRowsToArchive = [];

    values.forEach(function(row, idx) {
      if (isRowBelongsToSubmission_(row, id, indices)) {
        rowsToDelete.push(idx + 2);
        rawRowsToArchive.push(row);
      }
    });

    if (rowsToDelete.length === 0) {
      return { success: false, message: '在工作表中找不到對應的系統設定紀錄（可能已被刪除）。' };
    }

    // 先將被刪除的完整列資料封存至歷程表
    archiveSubmissionRows_(ss, rawRowsToArchive, '刪除前封存', permission.email || '');

    // 由大到小降序刪除，避免列索引位移
    rowsToDelete.sort(function(a, b) { return b - a; }).forEach(function(rowNum) {
      sheet.deleteRow(rowNum);
    });

    return {
      success: true,
      submissionId: id,
      deletedCount: rowsToDelete.length,
      message: `已成功刪除 ${rowsToDelete.length} 筆備份標的設定，並已封存至「${SHEET_HISTORY_NAME}」工作表。`
    };
  } catch (error) {
    return { success: false, message: '刪除失敗：' + error.toString() };
  }
}

// 3. 讀取儀表板資料（由 Sheet 還原成前端 appData 結構）
function getDashboardData() {
  try {
    const permission = checkAppPermission_();
    if (!permission.allowed) {
      return { success: false, message: permission.message || '無權限讀取資料。', data: [] };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("找不到指定的工作表，請確認名稱是否為 '" + SHEET_NAME + "'。");

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }

    const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const indices = resolveHeaderIndices_(headerValues);
    const colCount = Math.max(sheet.getLastColumn(), EXPECTED_HEADERS.length);
    const values = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();

    // 將每一列（單一標的）轉換為 submission group，再取每個系統最新一次提交
    const bySubmission = new Map();
    values.forEach(function(row) {
      const timestamp = row[indices.timestamp] ? new Date(row[indices.timestamp]) : new Date(0);
      const systemId = String(row[indices.systemId] || '').trim();
      const systemName = String(row[indices.systemName] || '').trim();
      const systemLevel = String(row[indices.systemLevel] || '').trim();
      const targetName = String(row[indices.targetName] || '').trim();
      const managerName = String(row[indices.managerName] || '').trim();
      const submissionId = String(row[indices.submissionId] || '').trim();
      const targetKey = normalizeTargetKey_(row[indices.targetKey]) || targetNameToKey_(targetName);

      if (!systemName || !targetName) return;

      const ts = timestamp.getTime() || 0;
      const legacyKey = buildLegacySubmissionKey_(timestamp, systemName, systemLevel, managerName);
      const submissionKey = submissionId || legacyKey;

      if (!bySubmission.has(submissionKey)) {
        bySubmission.set(submissionKey, {
          id: submissionKey,
          timestamp: ts,
          systemName: systemName,
          systemLevel: systemLevel || '一般系統',
          managerName: managerName,
          targets: []
        });
      }

      bySubmission.get(submissionKey).targets.push(parseTargetFromRow(row, indices, targetKey));
    });

    const latestBySystem = new Map();
    bySubmission.forEach(function(record) {
      const systemKey = [record.systemName, record.systemLevel, record.managerName].join('||');
      const existing = latestBySystem.get(systemKey);
      if (!existing || record.timestamp >= existing.timestamp) {
        latestBySystem.set(systemKey, record);
      }
    });

    const data = Array.from(latestBySystem.values())
      .sort(function(a, b) { return b.timestamp - a.timestamp; })
      .map(function(item) {
        return {
          id: item.id,
          systemName: item.systemName,
          systemLevel: item.systemLevel,
          managerName: item.managerName,
          targets: item.targets
        };
      });

    return { success: true, data: data };
  } catch (error) {
    return { success: false, message: error.toString(), data: [] };
  }
}

function parseTargetFromRow(row, indices, targetKeyFromRow) {
  const targetName = (indices.targetName !== -1 && row[indices.targetName]) ? String(row[indices.targetName]).trim() : '';
  const localFull = splitCycleValue((indices.localFull !== -1 && row[indices.localFull] !== undefined) ? row[indices.localFull] : '', '週');
  const localDiff = splitCycleValue((indices.localDiff !== -1 && row[indices.localDiff] !== undefined) ? row[indices.localDiff] : '', '天');
  const localRetention = parseRetentionCount((indices.localRetention !== -1 && row[indices.localRetention] !== undefined) ? row[indices.localRetention] : '');
  const localLocation = (indices.localLocation !== -1 && row[indices.localLocation] !== undefined) ? String(row[indices.localLocation]).trim() : '';

  const rawIsOffsite = (indices.isOffsite !== -1 && row[indices.isOffsite] !== undefined) ? String(row[indices.isOffsite]).trim() : '否';
  const isOffsite = (rawIsOffsite === '是' || rawIsOffsite.toLowerCase() === 'true' || rawIsOffsite === '1' || rawIsOffsite.toUpperCase() === 'Y') ? '是' : (rawIsOffsite || '否');

  const offsiteFreq = splitCycleValue((indices.offsiteCycle !== -1 && row[indices.offsiteCycle] !== undefined) ? row[indices.offsiteCycle] : '', '週');
  const offsiteRetention = parseRetentionCount((indices.offsiteRetention !== -1 && row[indices.offsiteRetention] !== undefined) ? row[indices.offsiteRetention] : '');

  const localScheduleTime = (indices.localScheduleTime !== -1 && row[indices.localScheduleTime] !== undefined && String(row[indices.localScheduleTime]).trim() !== '')
    ? String(row[indices.localScheduleTime]).trim()
    : '每日 00:00';
  const localDispositionMethod = (indices.localDispositionMethod !== -1 && row[indices.localDispositionMethod] !== undefined && String(row[indices.localDispositionMethod]).trim() !== '')
    ? String(row[indices.localDispositionMethod]).trim()
    : '磁碟自動覆寫';
  const localEvidenceRequirements = (indices.localEvidenceRequirements !== -1 && row[indices.localEvidenceRequirements] !== undefined && String(row[indices.localEvidenceRequirements]).trim() !== '')
    ? String(row[indices.localEvidenceRequirements]).trim()
    : '備份系統執行Log';

  const offsiteLocationVal = (indices.offsiteLocation !== -1 && row[indices.offsiteLocation] !== undefined) ? String(row[indices.offsiteLocation]).trim() : '';
  const offsiteScheduleTimeVal = (indices.offsiteScheduleTime !== -1 && row[indices.offsiteScheduleTime] !== undefined && String(row[indices.offsiteScheduleTime]).trim() !== '')
    ? String(row[indices.offsiteScheduleTime]).trim()
    : '每日 00:00';
  const offsiteDispositionMethodVal = (indices.offsiteDispositionMethod !== -1 && row[indices.offsiteDispositionMethod] !== undefined && String(row[indices.offsiteDispositionMethod]).trim() !== '')
    ? String(row[indices.offsiteDispositionMethod]).trim()
    : '磁碟自動覆寫';
  const offsiteEvidenceRequirementsVal = (indices.offsiteEvidenceRequirements !== -1 && row[indices.offsiteEvidenceRequirements] !== undefined && String(row[indices.offsiteEvidenceRequirements]).trim() !== '')
    ? String(row[indices.offsiteEvidenceRequirements]).trim()
    : '磁帶消磁簽收單 / 雲端過期報表';

  const offsiteScheduleTime = isOffsite === '是' ? offsiteScheduleTimeVal : '無';
  const offsiteDispositionMethod = isOffsite === '是' ? offsiteDispositionMethodVal : '無';
  const offsiteEvidenceRequirements = isOffsite === '是' ? offsiteEvidenceRequirementsVal : '無';

  const target = {
    targetKey: normalizeTargetKey_(targetKeyFromRow) || targetNameToKey_(targetName),
    targetName: targetName,
    localDiffValue: localDiff.value,
    localDiffUnit: localDiff.unit,
    localFullValue: localFull.value,
    localFullUnit: localFull.unit,
    localRetention: localRetention,
    localLocation: localLocation,
    localScheduleTime: localScheduleTime,
    localDispositionMethod: localDispositionMethod,
    localEvidenceRequirements: localEvidenceRequirements,
    logRetention: '',
    isOffsite: isOffsite,
    offsiteFreqValue: offsiteFreq.value,
    offsiteFreqUnit: offsiteFreq.unit,
    offsiteLocation: isOffsite === '是' ? offsiteLocationVal : '',
    offsiteRetention: isOffsite === '是' ? offsiteRetention : '',
    offsiteScheduleTime: offsiteScheduleTime,
    offsiteDispositionMethod: offsiteDispositionMethod,
    offsiteEvidenceRequirements: offsiteEvidenceRequirements
  };

  // 日誌檔在資料表中以「本地_保留代數」欄位保存「保存 x 月」字串
  if (targetName === '系統日誌檔') {
    target.localDiffValue = 'N/A';
    target.localDiffUnit = '';
    target.localFullValue = 'N/A';
    target.localFullUnit = '';
    target.localRetention = 'N/A';
    target.logRetention = (indices.localRetention !== -1 && row[indices.localRetention] !== undefined) ? String(row[indices.localRetention]).trim() : '';
  }

  return target;
}

function buildSheetRowData_(options) {
  const target = options.target || {};
  const targetName = String(target.targetName || '').trim();
  const targetKey = normalizeTargetKey_(options.targetKey) || targetNameToKey_(targetName);
  const isLogTarget = (targetName === '系統日誌檔' || targetKey === 'log');

  let diffDisplay;
  let fullDisplay;
  let retentionDisplay;
  let offsiteCycleDisplay;
  let offsiteRetentionDisplay;

  if (isLogTarget) {
    diffDisplay = '不適用';
    fullDisplay = '不適用';
    retentionDisplay = String(target.logRetention || '').trim();
    offsiteCycleDisplay = '不適用';
    offsiteRetentionDisplay = '不適用';
  } else {
    diffDisplay = (target.localDiffValue && target.localDiffValue !== '0' && target.localDiffValue !== 'N/A')
      ? (target.localDiffValue + ' ' + target.localDiffUnit)
      : '無';
    fullDisplay = target.localFullValue + ' ' + target.localFullUnit;
    retentionDisplay = target.localRetention === 'ALL' ? '全部保留' : (target.localRetention + ' 代');
    offsiteCycleDisplay = target.isOffsite === '是' ? (target.offsiteFreqValue + ' ' + target.offsiteFreqUnit) : '無';
    offsiteRetentionDisplay = target.isOffsite === '是' ? (target.offsiteRetention + ' 代') : '無';
  }

  const localScheduleTime = String(target.localScheduleTime || '每日 00:00').trim() || '每日 00:00';
  const localDispositionMethod = String(target.localDispositionMethod || '磁碟自動覆寫').trim() || '磁碟自動覆寫';
  const localEvidenceRequirements = String(target.localEvidenceRequirements || '備份系統執行Log').trim() || '備份系統執行Log';

  let offsiteScheduleTime = '無';
  let offsiteDispositionMethod = '無';
  let offsiteEvidenceRequirements = '無';

  if (target.isOffsite === '是') {
    offsiteScheduleTime = String(target.offsiteScheduleTime || '每日 00:00').trim() || '每日 00:00';
    offsiteDispositionMethod = String(target.offsiteDispositionMethod || '磁碟自動覆寫').trim() || '磁碟自動覆寫';
    offsiteEvidenceRequirements = String(target.offsiteEvidenceRequirements || '磁帶消磁簽收單 / 雲端過期報表').trim() || '磁帶消磁簽收單 / 雲端過期報表';
  }

  return [
    options.timestamp,                                                           // 1. 時間戳記
    options.systemId,                                                            // 2. 系統ID
    options.systemName,                                                          // 3. 系統名稱
    options.systemLevel,                                                         // 4. 系統等級
    targetName,                                                                  // 5. 備份標的
    options.managerName,                                                         // 6. 管理人員
    fullDisplay,                                                                 // 7. 本地_完整週期
    diffDisplay,                                                                 // 8. 本地_差異週期
    retentionDisplay,                                                            // 9. 本地_保留代數
    target.localLocation || '',                                                  // 10. 本地_存放地點
    localScheduleTime,                                                           // 11. 本地_排程時間
    localDispositionMethod,                                                      // 12. 本地_處置方法
    localEvidenceRequirements,                                                   // 13. 本地_佐證種類
    target.isOffsite,                                                            // 14. 是否異地
    offsiteCycleDisplay,                                                         // 15. 異地_週期
    target.isOffsite === '是' ? (target.offsiteLocation || '') : '無',           // 16. 異地_地點
    offsiteRetentionDisplay,                                                     // 17. 異地_保留代數
    offsiteScheduleTime,                                                         // 18. 異地_排程時間
    offsiteDispositionMethod,                                                    // 19. 異地_處置方法
    offsiteEvidenceRequirements,                                                 // 20. 異地_佐證種類
    options.submissionId,                                                        // 21. submission_id
    targetKey                                                                    // 22. target_key
  ];
}

function normalizeTargetKey_(value) {
  const key = String(value || '').trim().toLowerCase();
  if (key === 'config' || key === 'source' || key === 'db' || key === 'log' || key === 'image' || key === 'other') {
    return key;
  }
  return '';
}

function targetNameToKey_(name) {
  const text = String(name || '').trim();
  if (text === '設定檔') return 'config';
  if (text === '原始碼') return 'source';
  if (text === '資料庫') return 'db';
  if (text === '系統日誌檔') return 'log';
  if (text === '整體影像備份') return 'image';
  if (text === '其他') return 'other';
  return '';
}

function buildLegacySubmissionKey_(timestamp, systemName, systemLevel, managerName) {
  const ts = timestamp && timestamp.getTime ? (timestamp.getTime() || 0) : 0;
  return [String(systemName || '').trim(), String(systemLevel || '').trim(), String(managerName || '').trim(), ts].join('||');
}

function isRowBelongsToSubmission_(row, submissionId, indices) {
  const id = String(submissionId || '').trim();
  if (!id) return false;

  const subCol = (indices && indices.submissionId !== undefined && indices.submissionId !== -1) ? indices.submissionId : -1;
  const rowSubmissionId = subCol !== -1 ? String(row[subCol] || '').trim() : '';
  if (rowSubmissionId && rowSubmissionId === id) {
    return true;
  }

  if (!rowSubmissionId) {
    const tsCol = (indices && indices.timestamp !== undefined && indices.timestamp !== -1) ? indices.timestamp : 0;
    const nameCol = (indices && indices.systemName !== undefined && indices.systemName !== -1) ? indices.systemName : 2;
    const levelCol = (indices && indices.systemLevel !== undefined && indices.systemLevel !== -1) ? indices.systemLevel : 3;
    const mgrCol = (indices && indices.managerName !== undefined && indices.managerName !== -1) ? indices.managerName : 5;

    const timestamp = (tsCol !== -1 && row[tsCol]) ? new Date(row[tsCol]) : new Date(0);
    const legacyId = buildLegacySubmissionKey_(
      timestamp,
      nameCol !== -1 ? row[nameCol] : '',
      levelCol !== -1 ? row[levelCol] : '',
      mgrCol !== -1 ? row[mgrCol] : ''
    );
    return legacyId === id;
  }

  return false;
}

function splitCycleValue(rawValue, fallbackUnit) {
  const text = String(rawValue || '').trim();
  if (!text || text === 'N/A' || text === '無') {
    return { value: '', unit: fallbackUnit };
  }

  const match = text.match(/^(\d+)\s*(\S+)$/);
  if (!match) {
    return { value: text, unit: fallbackUnit };
  }

  return { value: match[1], unit: match[2] || fallbackUnit };
}

function parseRetentionCount(rawValue) {
  const text = String(rawValue || '').trim();
  if (text === '全部保留') return 'ALL';
  const match = text.match(/(\d+)/);
  return match ? match[1] : '';
}

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
            const values = systemSheet.getRange(2, 1, lastRow - 1, 17).getValues();
            const seen = {};
            values.forEach(function(row) {
              const colA = String(row[0] || '').trim();
              const colC = String(row[2] || '').trim();
              const colQ = String(row[16] || '').trim(); // 系統等級
              if (!colA && !colC) return;
              
              const name = colC ? (colA ? (colA + ' - ' + colC) : colC) : colA;
              // 支援「中級」、「高級」或是單純的「中」、「高」
              const level = (colQ === '中級' || colQ === '中' || colQ === '高級' || colQ === '高') ? '關鍵核心系統' : '一般系統';
              
              if (!seen[name]) { 
                seen[name] = true; 
                result.systems.push({ name: name, level: level, rawQ: colQ }); 
              }
            });
            result.systems.sort((a, b) => a.name.localeCompare(b.name));
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
                result.assets.push({ name: name, level: level, rawR: colR }); 
              }
            });
            result.assets.sort((a, b) => a.name.localeCompare(b.name));
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

// 5. 一鍵生成備份管制表單（Google Doc 範本 + Sheet 資料表格）
function generateBackupControlDocument() {
  try {
    const permission = checkAppPermission_();
    if (!permission.allowed) {
      return { success: false, message: permission.message || '無權限產生文件。' };
    }

    if (typeof FORM_TEMPLATE_DOC_ID === 'undefined' || !String(FORM_TEMPLATE_DOC_ID).trim() || FORM_TEMPLATE_DOC_ID === 'PLEASE_SET_FORM_TEMPLATE_DOC_ID') {
      return { success: false, message: '尚未設定 FORM_TEMPLATE_DOC_ID。' };
    }
    if (typeof FORM_OUTPUT_FOLDER_ID === 'undefined' || !String(FORM_OUTPUT_FOLDER_ID).trim() || FORM_OUTPUT_FOLDER_ID === 'PLEASE_SET_FORM_OUTPUT_FOLDER_ID') {
      return { success: false, message: '尚未設定 FORM_OUTPUT_FOLDER_ID。' };
    }

    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sourceSheet) {
      return { success: false, message: "找不到資料來源工作表（備份管制）。" };
    }

    const templateFile = DriveApp.getFileById(String(FORM_TEMPLATE_DOC_ID).trim());
    const outputFolderId = String(FORM_OUTPUT_FOLDER_ID).trim();
    const outputFolder = DriveApp.getFolderById(outputFolderId);

    const now = new Date();
    const dateKey = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
    const year = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy');
    const month = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM');
    const day = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd');
    const recordNo = createRecordNoFromFolder_(outputFolder, dateKey);
    const newFileName = '備份管制表單_' + recordNo;
    const copiedFile = templateFile.makeCopy(newFileName, outputFolder);

    const doc = DocumentApp.openById(copiedFile.getId());
    const body = doc.getBody();

    replaceTemplateTokens_(doc, {
      '年': year,
      '月': month,
      '日': day,
      '紀錄編號': recordNo
    });

    const lastRow = sourceSheet.getLastRow();
    if (lastRow <= 1) {
      body.appendParagraph('目前沒有可匯出的備份資料。').setForegroundColor('#6b7280');
      doc.saveAndClose();
      return { success: true, message: '文件已生成（目前無資料列）。', url: doc.getUrl(), fileId: copiedFile.getId(), recordNo: recordNo };
    }

    checkAndSetHeaders(sourceSheet);
    const sourceHeaders = sourceSheet.getRange(1, 1, 1, Math.max(sourceSheet.getLastColumn(), 1)).getValues()[0];
    const indices = resolveHeaderIndices_(sourceHeaders);
    const values = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
    const tableData = [];
    tableData.push(['項次', '系統/設備名稱', '是否備份', '備份標的及週期', '備份方式', '系統管理人員']);

    values.forEach(function(row, idx) {
      const timestamp = formatCellValue_(row[indices.timestamp]);
      const systemId = formatCellValue_(row[indices.systemId]);
      const systemName = formatCellValue_(row[indices.systemName]);
      const systemLevel = formatCellValue_(row[indices.systemLevel]);
      const backupTarget = formatCellValue_(row[indices.targetName]);
      const localDiff = formatCellValue_(row[indices.localDiff]);
      const localFull = formatCellValue_(row[indices.localFull]);
      const localRetention = formatCellValue_(row[indices.localRetention]);
      const localLocation = formatCellValue_(row[indices.localLocation]);
      const isOffsite = formatCellValue_(row[indices.isOffsite]);
      const offsiteCycle = formatCellValue_(row[indices.offsiteCycle]);
      const offsiteLocation = formatCellValue_(row[indices.offsiteLocation]);
      const offsiteRetention = formatCellValue_(row[indices.offsiteRetention]);
      const manager = formatCellValue_(row[indices.managerName]);

      const col2 = systemName + (systemLevel ? ('（' + systemLevel + '）') : '');
      const col3 = backupTarget && backupTarget !== '無' ? '是' : '否';
      const col4 = [
        '備份標的：' + backupTarget,
        '差異週期：' + localDiff,
        '完整週期：' + localFull,
        '異地週期：' + offsiteCycle
      ].join('\n');
      const col5 = [
        '本地存放：' + localLocation,
        '本地保留：' + localRetention,
        '異地備份：' + isOffsite,
        '異地地點：' + offsiteLocation,
        '異地保留：' + offsiteRetention,
        '時間戳記：' + timestamp
      ].join('\n');

      tableData.push([
        String(idx + 1),
        col2,
        col3,
        col4,
        col5,
        manager
      ]);
    });

    const table = insertTableAtPlaceholder_(body, tableData);
    styleGeneratedTable_(table);

    doc.saveAndClose();
    return {
      success: true,
      message: '文件已成功生成。',
      url: doc.getUrl(),
      fileId: copiedFile.getId(),
      recordNo: recordNo
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function styleGeneratedTable_(table) {
  if (!table) return;

  const rowCount = table.getNumRows();
  for (let r = 0; r < rowCount; r++) {
    const row = table.getRow(r);
    const cellCount = row.getNumCells();
    for (let c = 0; c < cellCount; c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(6).setPaddingRight(6);
      if (r === 0) {
        cell.setBackgroundColor('#e5e7eb');
        cell.editAsText().setBold(true).setForegroundColor('#111827');
      } else {
        cell.setBackgroundColor('#ffffff');
        cell.editAsText().setBold(false).setForegroundColor('#1f2937');
      }
    }
  }
}

function insertTableAtPlaceholder_(body, tableData) {
  const placeholderPattern = '\\{\\{表格\\}\\}';
  const found = body.findText(placeholderPattern);

  if (!found) {
    body.appendParagraph('');
    body.appendParagraph('備份資訊明細表').setBold(true).setFontSize(12);
    return body.appendTable(tableData);
  }

  const textElement = found.getElement().asText();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();
  textElement.deleteText(start, end);

  let paragraph = textElement.getParent();
  while (paragraph && paragraph.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    paragraph = paragraph.getParent();
  }

  if (!paragraph) {
    return body.appendTable(tableData);
  }

  const paragraphText = paragraph.asParagraph().getText().trim();
  const index = body.getChildIndex(paragraph);
  const table = body.insertTable(index + 1, tableData);

  if (!paragraphText) {
    body.removeChild(paragraph);
  }

  return table;
}

function replaceTemplateTokens_(doc, tokenMap) {
  const sections = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(Boolean);
  const keys = Object.keys(tokenMap || {});

  sections.forEach(function(section) {
    keys.forEach(function(key) {
      // 支援 {{年}} 與 {{ 年 }} 這種帶空白寫法
      const pattern = '\\{\\{\\s*' + escapeRegExp_(key) + '\\s*\\}\\}';
      section.replaceText(pattern, String(tokenMap[key]));
    });
  });
}

function createRecordNoFromFolder_(folder, dateKey) {
  const prefix = (typeof RECORD_NUMBER_PREFIX !== 'undefined' && String(RECORD_NUMBER_PREFIX).trim())
    ? String(RECORD_NUMBER_PREFIX).trim()
    : 'IS-R-032';
  const escapedPrefix = escapeRegExp_(prefix);
  const pattern = new RegExp(escapedPrefix + '-' + dateKey + '-(\\d+)');

  let maxSerial = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const name = String(file.getName() || '').trim();
    const match = name.match(pattern);
    if (!match) continue;

    const serial = parseInt(match[1], 10);
    if (!isNaN(serial) && serial > maxSerial) {
      maxSerial = serial;
    }
  }

  const nextSerial = maxSerial + 1;
  const serialText = nextSerial < 100 ? ('0' + nextSerial).slice(-2) : String(nextSerial);
  return prefix + '-' + dateKey + '-' + serialText;
}

function escapeRegExp_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function formatCellValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  }
  const text = String(value == null ? '' : value).trim();
  return text || 'N/A';
}

function checkAppPermission_() {
  try {
    const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    if (!email) {
      return { allowed: false, message: '無法辨識目前使用者帳號，請使用可識別的 Google 帳號登入後再試。' };
    }

    if (typeof ASSETS_SPREADSHEET_ID === 'undefined' || !String(ASSETS_SPREADSHEET_ID).trim() || ASSETS_SPREADSHEET_ID === 'PLEASE_SET_ASSETS_SPREADSHEET_ID') {
      return { allowed: false, message: '尚未設定外部權限來源試算表 ID。' };
    }

    const assetsSpreadsheet = SpreadsheetApp.openById(String(ASSETS_SPREADSHEET_ID).trim());
    const permissionSheet = assetsSpreadsheet.getSheetByName(SHEET_PERMISSIONS_NAME);
    if (!permissionSheet) {
      return { allowed: false, message: '外部試算表中找不到「權限」工作表。' };
    }

    const lastRow = permissionSheet.getLastRow();
    if (lastRow <= 1) {
      return { allowed: false, message: '「權限」工作表尚未設定任何可用帳號。' };
    }

    const values = permissionSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B 欄
    const allowed = values.some(function(row) {
      return String(row[0] || '').trim().toLowerCase() === email;
    });

    if (!allowed) {
      return { allowed: false, message: '您的帳號不在白名單中。' };
    }

    return { allowed: true, email: email };
  } catch (error) {
    return { allowed: false, message: '權限檢查失敗：' + error.toString() };
  }
}

function sanitizeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
