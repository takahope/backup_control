const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

// 從 code.js 載入相關常數與純函式
const codeJsContent = fs.readFileSync(path.join(__dirname, '..', 'code.js'), 'utf8');

// 建立模擬環境執行純函式
const sandbox = {
  Utilities: {
    getUuid: () => 'mock-uuid-123'
  },
  console: console
};

const fnDefinitions = `
${codeJsContent}
exports.EXPECTED_HEADERS = EXPECTED_HEADERS;
exports.EXPECTED_HISTORY_HEADERS = EXPECTED_HISTORY_HEADERS;
exports.resolveHeaderIndices_ = resolveHeaderIndices_;
exports.parseTargetFromRow = parseTargetFromRow;
exports.splitCycleValue = splitCycleValue;
exports.parseRetentionCount = parseRetentionCount;
exports.normalizeTargetKey_ = normalizeTargetKey_;
exports.targetNameToKey_ = targetNameToKey_;
exports.projectBusinessRow_ = projectBusinessRow_;
exports.migrateSheetIfNeeded_ = migrateSheetIfNeeded_;
exports.backfillMissingTechnicalKeys = backfillMissingTechnicalKeys;
exports.ensureTechnicalHeadersExist_ = ensureTechnicalHeadersExist_;
exports.autoBackfillTechnicalKeysIfNeeded_ = autoBackfillTechnicalKeysIfNeeded_;
exports.ensureSheetColumnsExist_ = ensureSheetColumnsExist_;
exports.runAutoBackfillNow = runAutoBackfillNow;
`;

const mod = {};
const fn = new Function('exports', 'Utilities', 'console', fnDefinitions);
fn(mod, sandbox.Utilities, sandbox.console);

test('resolveHeaderIndices_ 在舊 16 欄工作表上不應發生欄位錯位或將排程時間指向異地週期', () => {
  const OLD_16_HEADERS = [
    '時間戳記', '系統ID', '系統名稱', '系統等級', '備份標的',
    '本地_差異週期', '本地_完整週期', '本地_保留代數', '本地_存放地點',
    '是否異地', '異地_週期', '異地_地點', '異地_保留代數', '管理人員',
    'submission_id', 'target_key'
  ];

  const indices = mod.resolveHeaderIndices_(OLD_16_HEADERS);

  // 既有欄位必須正確解析
  assert.equal(indices.timestamp, 0);
  assert.equal(indices.systemId, 1);
  assert.equal(indices.systemName, 2);
  assert.equal(indices.systemLevel, 3);
  assert.equal(indices.targetName, 4);
  assert.equal(indices.localDiff, 5);
  assert.equal(indices.localFull, 6);
  assert.equal(indices.localRetention, 7);
  assert.equal(indices.localLocation, 8);
  assert.equal(indices.isOffsite, 9);
  assert.equal(indices.offsiteCycle, 10);
  assert.equal(indices.offsiteLocation, 11);
  assert.equal(indices.offsiteRetention, 12);
  assert.equal(indices.managerName, 13);
  assert.equal(indices.submissionId, 14);
  assert.equal(indices.targetKey, 15);

  // 舊表不存在之新欄位，索引不可指向既有有效欄位（必須為 -1）
  assert.equal(indices.localScheduleTime, -1, 'localScheduleTime 不可錯位指向異地欄位');
  assert.equal(indices.localDispositionMethod, -1, 'localDispositionMethod 不可錯位指向異地地點');
  assert.equal(indices.localEvidenceRequirements, -1, 'localEvidenceRequirements 不可錯位');
  assert.equal(indices.offsiteScheduleTime, -1, 'offsiteScheduleTime 不可錯位');
});

test('parseTargetFromRow 在讀取舊 16 欄資料列時，新欄位應正確套用預設值且既有資料不遺失', () => {
  const OLD_16_HEADERS = [
    '時間戳記', '系統ID', '系統名稱', '系統等級', '備份標的',
    '本地_差異週期', '本地_完整週期', '本地_保留代數', '本地_存放地點',
    '是否異地', '異地_週期', '異地_地點', '異地_保留代數', '管理人員',
    'submission_id', 'target_key'
  ];

  const indices = mod.resolveHeaderIndices_(OLD_16_HEADERS);

  const sampleOldRow = [
    '2026/9/19 10:00:00', 'SYS-01', '測試系統', '關鍵核心系統', '資料庫檔案',
    '1 天', '1 週', '4 代', 'NAS-01',
    '是', '2 週', '異地機房', '8 代', '張三',
    'sub-001', 'db'
  ];

  const target = mod.parseTargetFromRow(sampleOldRow, indices, 'db');

  // 驗證既有資料
  assert.equal(target.targetName, '資料庫檔案');
  assert.equal(target.localDiffValue, '1');
  assert.equal(target.localDiffUnit, '天');
  assert.equal(target.localFullValue, '1');
  assert.equal(target.localFullUnit, '週');
  assert.equal(target.localRetention, '4');
  assert.equal(target.localLocation, 'NAS-01');
  assert.equal(target.isOffsite, '是');
  assert.equal(target.offsiteFreqValue, '2');
  assert.equal(target.offsiteFreqUnit, '週');
  assert.equal(target.offsiteLocation, '異地機房');
  assert.equal(target.offsiteRetention, '8');

  // 驗證新欄位未包含在舊表時，應正確回退為標準預設值，而不是讀取到異地欄位的內容
  assert.equal(target.localScheduleTime, '每日 00:00');
  assert.equal(target.localDispositionMethod, '磁碟自動覆寫');
  assert.equal(target.localEvidenceRequirements, '備份系統執行Log');
  assert.equal(target.offsiteScheduleTime, '每日 00:00');
  assert.equal(target.offsiteDispositionMethod, '磁碟自動覆寫');
  assert.equal(target.offsiteEvidenceRequirements, '磁帶消磁簽收單 / 雲端過期報表');
});

test('projectBusinessRow_ 安全投影：異地地點有資料或布林值時絕不可被強制洗成「無」', () => {
  const rowWithOffsite = [
    '2026/9/19 10:00:00', 'SYS-01', '測試系統', '關鍵核心系統', '資料庫檔案',
    '1 天', '1 週', '4 代', 'NAS-01',
    'true', '2 週', '台北備份中心', '4 代', '李四',
    'sub-002', 'db'
  ];
  const oldHeaders = [
    '時間戳記', '系統ID', '系統名稱', '系統等級', '備份標的',
    '本地_差異週期', '本地_完整週期', '本地_保留代數', '本地_存放地點',
    '是否異地', '異地_週期', '異地_地點', '異地_保留代數', '管理人員',
    'submission_id', 'target_key'
  ];

  function getOldVal(row, aliases) {
    for (let i = 0; i < aliases.length; i++) {
      const idx = oldHeaders.indexOf(aliases[i]);
      if (idx !== -1 && row[idx] !== undefined && row[idx] !== null && String(row[idx]).trim() !== '') {
        return row[idx];
      }
    }
    return '';
  }

  const projected = mod.projectBusinessRow_(rowWithOffsite, getOldVal);

  // 新 22 欄索引：
  // 13: 是否異地, 14: 異地_週期, 15: 異地_地點, 16: 異地_保留代數
  assert.equal(projected[13], '是', '布林值 true 或有異地地點應被辨識為「是」');
  assert.equal(projected[14], '2 週');
  assert.equal(projected[15], '台北備份中心', '異地地點不可被洗成「無」');
  assert.equal(projected[16], '4 代');
});

test('migrateSheetIfNeeded_ 在 26 欄工作表（A~Z）但前 22 欄表頭正確時，絕不可清空工作表', () => {
  let clearCalled = false;
  const mockSheet = {
    getLastRow: () => 10,
    getLastColumn: () => 26, // 預設 26 欄
    getRange: (r, c, numR, numC) => ({
      getValues: () => {
        // 第一列：前 22 欄正確，後 4 欄為空
        const row1 = [...mod.EXPECTED_HEADERS, '', '', '', ''];
        return [row1];
      },
      setValues: () => {},
      setFontWeight: () => {},
      setBackground: () => {}
    }),
    clearContents: () => {
      clearCalled = true;
    },
    setFrozenRows: () => {}
  };

  mod.migrateSheetIfNeeded_(mockSheet, mod.EXPECTED_HEADERS, false);
  assert.equal(clearCalled, false, '表頭吻合時絕不可呼叫 clearContents');
});

test('backfillMissingTechnicalKeys 正確為同一系統補齊相同 UUID 並依標的名稱推導 target_key', () => {
  let writtenValues = null;
  let copiedBackupName = null;

  const mockRows = [
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '資料庫', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '系統日誌檔', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['2026/9/20 12:00:00', 'SYS-02', 'LIMS系統', '一般系統', '原始碼', '李四', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'existing-uuid-002', 'source']
  ];

  const mockSheet = {
    getLastRow: () => 4,
    getLastColumn: () => 22,
    getRange: (r, c, numR, numC) => {
      if (r === 1) {
        return { getValues: () => [mod.EXPECTED_HEADERS] };
      }
      return {
        getValues: () => mockRows,
        setValues: (vals) => { writtenValues = vals; }
      };
    },
    copyTo: () => ({
      setName: (name) => { copiedBackupName = name; }
    })
  };

  const mockSs = {
    getSheetByName: (name) => (name === '備份管制' ? mockSheet : null)
  };

  const SpreadsheetApp = {
    getActiveSpreadsheet: () => mockSs,
    getUi: () => ({ alert: () => {} })
  };

  const Session = {
    getScriptTimeZone: () => 'GMT+8'
  };

  const Utilities = {
    getUuid: (() => {
      let count = 1;
      return () => 'mock-uuid-' + (count++);
    })(),
    formatDate: () => '20260922_120000'
  };

  const fnSrc = `
    ${codeJsContent}
    return backfillMissingTechnicalKeys();
  `;
  const runner = new Function('SpreadsheetApp', 'Session', 'Utilities', 'console', fnSrc);
  const result = runner(SpreadsheetApp, Session, Utilities, console);

  assert.ok(copiedBackupName.includes('備份管制_backup_'), '必須在寫入前建立備份快照分頁');
  assert.ok(writtenValues !== null, '必須寫回補齊後的數值');
  assert.equal(writtenValues[0][20], writtenValues[1][20], '同一系統的多個標的必須分配同一個 UUID');
  assert.equal(writtenValues[0][21], 'db', '資料庫標的應補齊為 db');
  assert.equal(writtenValues[1][21], 'log', '系統日誌檔標的應補齊為 log');
  assert.equal(writtenValues[2][20], 'existing-uuid-002', '既有 UUID 必須保留');
  assert.equal(writtenValues[2][21], 'source', '既有 target_key 必須保留');
});

test('autoBackfillTechnicalKeysIfNeeded_ 命中 Fast-path 時絕不執行任何寫入操作', () => {
  let writeCalled = false;
  const mockSheet = {
    getRange: () => ({
      setValues: () => { writeCalled = true; },
      setValue: () => { writeCalled = true; }
    })
  };

  const completeRows = [
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '資料庫', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'uuid-1', 'db'],
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '系統日誌檔', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'uuid-1', 'log']
  ];

  const indices = mod.resolveHeaderIndices_(mod.EXPECTED_HEADERS);
  const result = mod.autoBackfillTechnicalKeysIfNeeded_(mockSheet, completeRows, indices);

  assert.equal(writeCalled, false, '所有欄位齊全時絕不可觸發 setValues 寫入');
  assert.equal(result.length, 2);
  assert.equal(result[0][20], 'uuid-1');
});

test('autoBackfillTechnicalKeysIfNeeded_ 正確為同一系統的多標的補齊相同 UUID 與 target_key 並批次寫回', () => {
  let writtenRange = null;
  let writtenData = null;

  const mockSheet = {
    getRange: (r, c, numR, numC) => {
      writtenRange = { r, c, numR, numC };
      return {
        setValues: (vals) => { writtenData = vals; }
      };
    }
  };

  const missingRows = [
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '資料庫', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '系統日誌檔', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
  ];

  const indices = mod.resolveHeaderIndices_(mod.EXPECTED_HEADERS);
  const result = mod.autoBackfillTechnicalKeysIfNeeded_(mockSheet, missingRows, indices);

  assert.ok(writtenData !== null, '必須批次寫回補齊之數值');
  assert.equal(writtenRange.c, 21, '必須寫回第 21 欄起算');
  assert.equal(writtenRange.numC, 2, '必須寫回 2 欄寬度 (submission_id, target_key)');
  assert.ok(result[0][20].length > 0, 'submission_id 必須已補齊');
  assert.equal(result[0][20], result[1][20], '同一系統之多標的必須分配相同 UUID');
  assert.equal(result[0][21], 'db', '資料庫標的應補齊為 db');
  assert.equal(result[1][21], 'log', '系統日誌檔標的應補齊為 log');
});

test('autoBackfillTechnicalKeysIfNeeded_ 部分缺失時正確繼承同系統既有 UUID', () => {
  let writtenData = null;
  const mockSheet = {
    getRange: () => ({
      setValues: (vals) => { writtenData = vals; }
    })
  };

  const partialRows = [
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '資料庫', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'existing-uuid-999', 'db'],
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '原始碼', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
  ];

  const indices = mod.resolveHeaderIndices_(mod.EXPECTED_HEADERS);
  const result = mod.autoBackfillTechnicalKeysIfNeeded_(mockSheet, partialRows, indices);

  assert.equal(result[1][20], 'existing-uuid-999', '空白列必須繼承同系統既有的 UUID');
  assert.equal(result[1][21], 'source', '原始碼標的應補齊為 source');
});

test('整合驗證：缺失技術欄位時 getDashboardData 能自動補齊並完整保留同系統所有標的', () => {
  let sheetWrites = [];
  const mockRows = [
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '資料庫', '張三', '1週', '1天', '4', 'NAS-01', '每日 00:00', '磁碟自動覆寫', 'Log', '否', '', '', '', '無', '無', '無', '', ''],
    ['2026/9/19 10:00:01', 'SYS-01', 'HIS系統', '關鍵核心系統', '系統日誌檔', '張三', '', '', '保存 6 月', 'NAS-01', '每日 00:00', '磁碟自動覆寫', 'Log', '否', '', '', '', '無', '無', '無', '', '']
  ];

  const mockSheet = {
    getLastRow: () => 3,
    getLastColumn: () => 22,
    getMaxColumns: () => 22,
    getRange: (r, c, numR, numC) => {
      if (r === 1) {
        return { getValues: () => [mod.EXPECTED_HEADERS] };
      }
      return {
        getValues: () => mockRows,
        setValues: (vals) => { sheetWrites.push(vals); }
      };
    }
  };

  const mockSs = {
    getSheetByName: (name) => (name === '備份管制' ? mockSheet : null)
  };

  const SpreadsheetApp = {
    getActiveSpreadsheet: () => mockSs
  };

  const Session = {
    getScriptTimeZone: () => 'GMT+8',
    getActiveUser: () => ({ getEmail: () => 'admin@example.com' })
  };

  let uuidCounter = 1;
  const Utilities = {
    getUuid: () => 'auto-uuid-' + (uuidCounter++)
  };

  // 讀取最新 codeJsContent
  const currentCodeJs = fs.readFileSync(path.join(__dirname, '..', 'code.js'), 'utf8');
  const fnSrc = `
    ${currentCodeJs}
    // 覆寫權限檢查以利單元測試
    checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
    return getDashboardData();
  `;
  const runner = new Function('SpreadsheetApp', 'Session', 'Utilities', 'console', fnSrc);
  const resp = runner(SpreadsheetApp, Session, Utilities, console);

  assert.equal(resp.success, true, 'getDashboardData 必須成功');
  assert.equal(resp.data.length, 1, '應聚合為 1 個系統卡片');
  assert.equal(resp.data[0].systemName, 'HIS系統');
  assert.equal(resp.data[0].targets.length, 2, '該系統卡片底下的 targets 必須包含全部 2 個標的（資料庫與系統日誌檔），絕不可遺漏');
  assert.equal(resp.data[0].targets[0].targetName, '資料庫');
  assert.equal(resp.data[0].targets[1].targetName, '系統日誌檔');
});

test('ensureSheetColumnsExist_ 在格線不足 22 欄時正確擴展欄位，超過時不重複插入', () => {
  let inserted = null;
  const mockSheetSmall = {
    getMaxColumns: () => 20,
    insertColumnsAfter: (pos, count) => {
      inserted = { pos, count };
    }
  };

  mod.ensureSheetColumnsExist_(mockSheetSmall, 22);
  assert.deepEqual(inserted, { pos: 20, count: 2 }, '必須在第 20 欄後插入 2 欄');

  let insertedBig = null;
  const mockSheetBig = {
    getMaxColumns: () => 26,
    insertColumnsAfter: (pos, count) => {
      insertedBig = { pos, count };
    }
  };

  mod.ensureSheetColumnsExist_(mockSheetBig, 22);
  assert.equal(insertedBig, null, '欄數已達 26 欄時絕不可重複插入欄位');
});

