// Code.gs - 處理網頁請求與寫入 Google Sheets
const SHEET_NAME = '工作表1'; // 請確認您的 Google Sheet 分頁名稱是否為此
const SHEET_ASSETS_NAME = '資訊資產';
const SHEET_PERMISSIONS_NAME = '權限';

// 定義正確的表頭陣列（含技術欄位）
const EXPECTED_HEADERS = [
  '時間戳記', '系統名稱', '系統等級', '備份標的',
  '本地_差異週期', '本地_完整週期', '本地_保留代數', '本地_存放地點',
  '是否異地', '異地_週期', '異地_地點', '異地_保留代數', '管理人員',
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

// 新增功能：檢測並自動填寫正確的表頭
function checkAndSetHeaders(sheet) {
  const lastCol = EXPECTED_HEADERS.length;
  // 嘗試取得第一列的資料
  const range = sheet.getRange(1, 1, 1, lastCol);
  const currentHeaders = range.getValues()[0];

  let headersMatch = true;
  for (let i = 0; i < EXPECTED_HEADERS.length; i++) {
    // 將現有內容轉換為字串比對，避免 undefined 或其他型態造成誤判
    if (String(currentHeaders[i]).trim() !== EXPECTED_HEADERS[i]) {
      headersMatch = false;
      break;
    }
  }

  // 如果不匹配或為空白，覆寫第一列為正確表頭
  if (!headersMatch) {
    range.setValues([EXPECTED_HEADERS]);
    range.setFontWeight("bold");       // 字體加粗
    range.setBackground("#e5e7eb");    // 加上淺灰色背景以利辨識
    sheet.setFrozenRows(1);            // 凍結第一列，往下捲動時表頭會固定在上方
  }
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

    // 執行表頭檢測與自動填寫功能
    checkAndSetHeaders(sheet);

    // 檢查是否有傳入標的陣列
    if (!formData.targets || formData.targets.length === 0) {
      throw new Error("沒有收到任何備份標的資料！");
    }

    const timestamp = new Date();
    const systemName = String(formData.systemName || '').trim();
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
    const rowsToDelete = [];
    let updatedCount = 0;
    let insertedCount = 0;
    let deletedCount = 0;

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const values = sheet.getRange(2, 1, lastRow - 1, EXPECTED_HEADERS.length).getValues();
      values.forEach(function(row, idx) {
        const rowNumber = idx + 2;
        if (!isRowBelongsToSubmission_(row, submissionId)) return;

        const targetKey = normalizeTargetKey_(row[14]) || targetNameToKey_(row[3]);
        if (!targetKey) return;
        if (!existingRowsByTargetKey[targetKey]) {
          existingRowsByTargetKey[targetKey] = [];
        }
        existingRowsByTargetKey[targetKey].push(rowNumber);
      });
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
    return { 
      success: true, 
      submissionId: submissionId,
      message: `同步完成：更新 ${updatedCount} 筆、新增 ${insertedCount} 筆、移除 ${deletedCount} 筆。`
    };
    
  } catch (error) {
    // 發生錯誤時將詳細資訊回傳給前端
    return { success: false, message: error.toString() };
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

    checkAndSetHeaders(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }

    const values = sheet.getRange(2, 1, lastRow - 1, EXPECTED_HEADERS.length).getValues();

    // 將每一列（單一標的）轉換為 submission group，再取每個系統最新一次提交
    const bySubmission = new Map();
    values.forEach(function(row) {
      const timestamp = row[0] ? new Date(row[0]) : new Date(0);
      const systemName = String(row[1] || '').trim();
      const systemLevel = String(row[2] || '').trim();
      const targetName = String(row[3] || '').trim();
      const managerName = String(row[12] || '').trim();
      const submissionId = String(row[13] || '').trim();
      const targetKey = normalizeTargetKey_(row[14]) || targetNameToKey_(targetName);

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

      bySubmission.get(submissionKey).targets.push(parseTargetFromRow(row, targetKey));
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

function parseTargetFromRow(row, targetKeyFromRow) {
  const targetName = String(row[3] || '').trim();
  const localDiff = splitCycleValue(row[4], '天');
  const localFull = splitCycleValue(row[5], '週');
  const localRetention = parseRetentionCount(row[6]);
  const localLocation = String(row[7] || '').trim();
  const isOffsite = String(row[8] || '否').trim() || '否';
  const offsiteFreq = splitCycleValue(row[9], '週');
  const offsiteRetention = parseRetentionCount(row[11]);

  const target = {
    targetKey: normalizeTargetKey_(targetKeyFromRow) || targetNameToKey_(targetName),
    targetName: targetName,
    localDiffValue: localDiff.value,
    localDiffUnit: localDiff.unit,
    localFullValue: localFull.value,
    localFullUnit: localFull.unit,
    localRetention: localRetention,
    localLocation: localLocation,
    logRetention: '',
    isOffsite: isOffsite,
    offsiteFreqValue: offsiteFreq.value,
    offsiteFreqUnit: offsiteFreq.unit,
    offsiteLocation: isOffsite === '是' ? String(row[10] || '').trim() : '',
    offsiteRetention: isOffsite === '是' ? offsiteRetention : ''
  };

  // 日誌檔在資料表中以「本地_保留代數」欄位保存「保存 x 月」字串
  if (targetName === '系統日誌檔') {
    target.localDiffValue = 'N/A';
    target.localDiffUnit = '';
    target.localFullValue = 'N/A';
    target.localFullUnit = '';
    target.localRetention = 'N/A';
    target.logRetention = String(row[6] || '').trim();
  }

  return target;
}

function buildSheetRowData_(options) {
  const target = options.target || {};
  const targetName = String(target.targetName || '').trim();
  const targetKey = normalizeTargetKey_(options.targetKey) || targetNameToKey_(targetName);

  let diffDisplay;
  let fullDisplay;
  let retentionDisplay;
  if (targetName === '系統日誌檔' || targetKey === 'log') {
    diffDisplay = '無';
    fullDisplay = '無';
    retentionDisplay = String(target.logRetention || '').trim();
  } else {
    diffDisplay = (target.localDiffValue && target.localDiffValue !== '0' && target.localDiffValue !== 'N/A')
      ? (target.localDiffValue + ' ' + target.localDiffUnit)
      : '無';
    fullDisplay = target.localFullValue + ' ' + target.localFullUnit;
    retentionDisplay = target.localRetention === 'ALL' ? '全部保留' : (target.localRetention + ' 代');
  }

  return [
    options.timestamp,                                                           // A欄: 時間戳記
    options.systemName,                                                          // B欄: 系統名稱
    options.systemLevel,                                                         // C欄: 系統等級
    targetName,                                                                  // D欄: 備份標的
    diffDisplay,                                                                 // E欄: 本地_差異週期
    fullDisplay,                                                                 // F欄: 本地_完整週期
    retentionDisplay,                                                            // G欄: 本地_保留代數
    target.localLocation,                                                        // H欄: 本地_存放地點
    target.isOffsite,                                                            // I欄: 是否異地
    target.isOffsite === '是' ? (target.offsiteFreqValue + ' ' + target.offsiteFreqUnit) : '無', // J欄: 異地_週期
    target.isOffsite === '是' ? target.offsiteLocation : '無',                 // K欄: 異地_地點
    target.isOffsite === '是' ? target.offsiteRetention + ' 代' : '無',        // L欄: 異地_保留代數
    options.managerName,                                                         // M欄: 管理人員
    options.submissionId,                                                        // N欄: submission_id
    targetKey                                                                    // O欄: target_key
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

function isRowBelongsToSubmission_(row, submissionId) {
  const id = String(submissionId || '').trim();
  if (!id) return false;

  const rowSubmissionId = String(row[13] || '').trim();
  if (rowSubmissionId && rowSubmissionId === id) {
    return true;
  }

  if (!rowSubmissionId) {
    const timestamp = row[0] ? new Date(row[0]) : new Date(0);
    const legacyId = buildLegacySubmissionKey_(timestamp, row[1], row[2], row[12]);
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

// 4. 讀取資訊資產（僅 B 欄為 HW），回傳「A欄 + C欄」供前端下拉搜尋
function getAssetOptions() {
  try {
    const permission = checkAppPermission_();
    if (!permission.allowed) {
      return { success: false, message: permission.message || '無權限讀取資訊資產。', data: [] };
    }

    if (typeof ASSETS_SPREADSHEET_ID === 'undefined' || !String(ASSETS_SPREADSHEET_ID).trim() || ASSETS_SPREADSHEET_ID === 'PLEASE_SET_ASSETS_SPREADSHEET_ID') {
      return { success: false, message: "尚未設定外部資訊資產 Google Sheet ID（ASSETS_SPREADSHEET_ID）。", data: [] };
    }

    const assetsSpreadsheet = SpreadsheetApp.openById(String(ASSETS_SPREADSHEET_ID).trim());
    const sheet = assetsSpreadsheet.getSheetByName(SHEET_ASSETS_NAME);
    if (!sheet) {
      return { success: false, message: "在外部試算表中找不到「資訊資產」工作表。", data: [] };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }

    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A,B,C
    const options = [];
    const seen = {};

    values.forEach(function(row) {
      const colA = String(row[0] || '').trim();
      const colB = String(row[1] || '').trim();
      const colC = String(row[2] || '').trim();
      const normalizedType = colB
        .replace(/Ｓ/g, 'S')
        .replace(/Ｈ/g, 'H')
        .replace(/Ｗ/g, 'W')
        .toUpperCase();

      if (normalizedType !== 'HW' && normalizedType !== 'SW') return;
      if (!colA && !colC) return;

      const name = colC ? (colA ? (colA + ' - ' + colC) : colC) : colA;
      if (!seen[name]) {
        seen[name] = true;
        options.push(name);
      }
    });

    options.sort();
    return { success: true, data: options };
  } catch (error) {
    return { success: false, message: error.toString(), data: [] };
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
      return { success: false, message: "找不到資料來源工作表（工作表1）。" };
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

    const values = sourceSheet.getRange(2, 1, lastRow - 1, EXPECTED_HEADERS.length).getValues();
    const tableData = [];
    tableData.push(['項次', '系統/設備名稱', '是否備份', '備份標的及週期', '備份方式', '系統管理人員']);

    values.forEach(function(row, idx) {
      const timestamp = formatCellValue_(row[0]);
      const systemName = formatCellValue_(row[1]);
      const systemLevel = formatCellValue_(row[2]);
      const backupTarget = formatCellValue_(row[3]);
      const localDiff = formatCellValue_(row[4]);
      const localFull = formatCellValue_(row[5]);
      const localRetention = formatCellValue_(row[6]);
      const localLocation = formatCellValue_(row[7]);
      const isOffsite = formatCellValue_(row[8]);
      const offsiteCycle = formatCellValue_(row[9]);
      const offsiteLocation = formatCellValue_(row[10]);
      const offsiteRetention = formatCellValue_(row[11]);
      const manager = formatCellValue_(row[12]);

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
