// Code.gs - 處理網頁請求與寫入 Google Sheets
const SHEET_NAME = '工作表1'; // 請確認您的 Google Sheet 分頁名稱是否為此
const SHEET_ASSETS_NAME = '資訊資產';
const SHEET_PERMISSIONS_NAME = '權限';

// 定義正確的表頭陣列 (新增 '本地_存放地點')
const EXPECTED_HEADERS = [
  '時間戳記', '系統名稱', '系統等級', '備份標的',
  '本地_差異週期', '本地_完整週期', '本地_保留代數', '本地_存放地點',
  '是否異地', '異地_週期', '異地_地點', '異地_保留代數', '管理人員'
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

    // 取得表單共用的基本資訊
    const timestamp = new Date();
    const systemName = formData.systemName;
    const systemLevel = formData.systemLevel;
    const managerName = formData.managerName;
    
    // 準備一個二維陣列存放要寫入的所有資料列
    const rowsToWrite = [];
    
    // 檢查是否有傳入標的陣列
    if (!formData.targets || formData.targets.length === 0) {
      throw new Error("沒有收到任何備份標的資料！");
    }

    // 將前端傳來的所有被勾選的備份標的，逐一轉換成試算表的「列 (Row)」
    formData.targets.forEach(function(target) {
      
      let diffDisplay, fullDisplay, retentionDisplay;

      // 針對「系統日誌檔」的特殊處理：直接顯示保存期間，無備份週期
      if (target.targetName === '系統日誌檔') {
        diffDisplay = '無';
        fullDisplay = '無';
        retentionDisplay = target.logRetention; // 例如："保存 6 個月"
      } else {
        // 一般備份標的處理
        diffDisplay = (target.localDiffValue && target.localDiffValue !== '0' && target.localDiffValue !== 'N/A') 
                            ? target.localDiffValue + ' ' + target.localDiffUnit 
                            : '無';
        fullDisplay = target.localFullValue + ' ' + target.localFullUnit;
        retentionDisplay = target.localRetention === 'ALL' ? '全部保留' : (target.localRetention + ' 代');
      }

      // 組合單列資料 (需對應您的試算表表頭順序)
      const rowData = [
        timestamp,                                           // A欄: 時間戳記
        systemName,                                          // B欄: 系統名稱
        systemLevel,                                         // C欄: 系統等級
        target.targetName,                                   // D欄: 備份標的
        diffDisplay,                                         // E欄: 本地_差異週期
        fullDisplay,                                         // F欄: 本地_完整週期
        retentionDisplay,                                    // G欄: 本地_保留代數 (或日誌保存期間)
        target.localLocation,                                // H欄: 本地_存放地點 (新增)
        target.isOffsite,                                    // I欄: 是否異地
        target.isOffsite === '是' ? (target.offsiteFreqValue + ' ' + target.offsiteFreqUnit) : 'N/A', // J欄: 異地_週期
        target.isOffsite === '是' ? target.offsiteLocation : 'N/A',                                   // K欄: 異地_地點
        target.isOffsite === '是' ? target.offsiteRetention + ' 代' : 'N/A',                          // L欄: 異地_保留代數
        managerName                                          // M欄: 管理人員
      ];
      
      rowsToWrite.push(rowData); // 將這列加入準備寫入的陣列中
    });

    // 將資料批次寫入 Google Sheets
    if (rowsToWrite.length > 0) {
      // 取得目前最後一列。若試算表原本為空，寫完表頭後 getLastRow() 會是 1
      let startRow = sheet.getLastRow() + 1; 
      
      // 強制最少從第 2 列開始寫，保護第 1 列的表頭不被覆寫
      if (startRow < 2) {
        startRow = 2; 
      }

      const startCol = 1;                      // 從第 A 欄開始寫
      const numRows = rowsToWrite.length;      // 要寫入幾列
      const numCols = rowsToWrite[0].length;   // 每列有幾欄

      // 一次性寫入多行資料
      sheet.getRange(startRow, startCol, numRows, numCols).setValues(rowsToWrite);
    }

    // 回傳成功訊息與處理筆數給前端
    return { 
      success: true, 
      message: `成功寫入 ${rowsToWrite.length} 筆標的紀錄！` 
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

      if (!systemName || !targetName) return;

      const ts = timestamp.getTime() || 0;
      const submissionKey = [systemName, systemLevel, managerName, ts].join('||');

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

      bySubmission.get(submissionKey).targets.push(parseTargetFromRow(row));
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

function parseTargetFromRow(row) {
  const targetName = String(row[3] || '').trim();
  const localDiff = splitCycleValue(row[4], '天');
  const localFull = splitCycleValue(row[5], '週');
  const localRetention = parseRetentionCount(row[6]);
  const localLocation = String(row[7] || '').trim();
  const isOffsite = String(row[8] || '否').trim() || '否';
  const offsiteFreq = splitCycleValue(row[9], '週');
  const offsiteRetention = parseRetentionCount(row[11]);

  const target = {
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

    const recordNo = createRecordNo_();
    const now = new Date();
    const year = String(now.getFullYear());
    const month = String(now.getMonth() + 1);
    const day = String(now.getDate());

    const templateFile = DriveApp.getFileById(String(FORM_TEMPLATE_DOC_ID).trim());
    const outputFolder = DriveApp.getFolderById(String(FORM_OUTPUT_FOLDER_ID).trim());
    const newFileName = '備份管制表單_' + recordNo;
    const copiedFile = templateFile.makeCopy(newFileName, outputFolder);

    const doc = DocumentApp.openById(copiedFile.getId());
    const body = doc.getBody();

    body.replaceText('\\{\\{年\\}\\}', year);
    body.replaceText('\\{\\{月\\}\\}', month);
    body.replaceText('\\{\\{日\\}\\}', day);
    body.replaceText('\\{\\{紀錄編號\\}\\}', recordNo);

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

function createRecordNo_() {
  const now = new Date();
  const y = now.getFullYear();
  const m = ('0' + (now.getMonth() + 1)).slice(-2);
  const d = ('0' + now.getDate()).slice(-2);
  const hh = ('0' + now.getHours()).slice(-2);
  const mm = ('0' + now.getMinutes()).slice(-2);
  const ss = ('0' + now.getSeconds()).slice(-2);
  return 'REC-' + y + m + d + '-' + hh + mm + ss;
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
