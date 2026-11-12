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
exports.archiveSubmissionRows_ = archiveSubmissionRows_;
exports.buildSheetRowData_ = buildSheetRowData_;
exports.isCycleEnabled_ = isCycleEnabled_;
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
  const sheetRows = [
    [...mod.EXPECTED_HEADERS, '', '', '', ''],
    Array(26).fill('')
  ];
  const mockSheet = {
    getLastRow: () => 10,
    getLastColumn: () => 26, // 預設 26 欄
    getRange: (r, c, numR, numC) => ({
      getValues: () => sheetRows
        .slice(r - 1, r - 1 + Math.min(numR, sheetRows.length))
        .map(row => row.slice(c - 1, c - 1 + numC)),
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

test('migrateSheetIfNeeded_ 必須移除 V 欄後的有效殘留資料並重建 22 欄契約', () => {
  let clearCalled = false;
  let written = null;
  const dataRow = Array.from({ length: 22 }, (_, i) => `value-${i + 1}`).concat('W-殘留');
  const sheetRows = [[...mod.EXPECTED_HEADERS, '舊欄位'], dataRow];
  const mockSheet = {
    getLastRow: () => 2,
    getLastColumn: () => 23,
    getRange: (r, c, numR, numC) => ({
      getValues: () => sheetRows
        .slice(r - 1, r - 1 + numR)
        .map(row => row.slice(c - 1, c - 1 + numC)),
      setValues: (values) => { written = { r, c, numR, numC, values }; },
      setFontWeight: () => {},
      setBackground: () => {}
    }),
    clearContents: () => { clearCalled = true; },
    setFrozenRows: () => {}
  };

  mod.migrateSheetIfNeeded_(mockSheet, mod.EXPECTED_HEADERS, false);

  assert.equal(clearCalled, true);
  assert.equal(written.numC, 22);
  assert.deepEqual(written.values[0], mod.EXPECTED_HEADERS);
  assert.equal(written.values[1].length, 22);
  assert.equal(written.values[1].includes('W-殘留'), false);
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
  assert.equal(writtenValues[0][0], writtenValues[1][0], '同一系統的多個標的必須分配同一個 UUID');
  assert.equal(writtenValues[0][1], 'db', '資料庫標的應補齊為 db');
  assert.equal(writtenValues[1][1], 'log', '系統日誌檔標的應補齊為 log');
  assert.equal(writtenValues[2][0], 'existing-uuid-002', '既有 UUID 必須保留');
  assert.equal(writtenValues[2][1], 'source', '既有 target_key 必須保留');
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

test('autoBackfillTechnicalKeysIfNeeded_ 系統 ID 部分缺失時正確繼承唯一同系統既有 UUID', () => {
  let writtenData = null;
  const mockSheet = {
    getRange: () => ({
      setValues: (vals) => { writtenData = vals; }
    })
  };

  const partialRows = [
    ['2026/9/19 10:00:00', 'SYS-01', 'HIS系統', '關鍵核心系統', '資料庫', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'existing-uuid-999', 'db'],
    ['2026/9/19 10:00:00', '', 'HIS系統', '關鍵核心系統', '原始碼', '張三', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
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
  assert.equal(resp.data[0].systemId, 'SYS-01', '讀取結果必須保留 B 欄系統 ID 供編輯回填');
  assert.equal(resp.data[0].systemName, 'HIS系統');
  assert.equal(resp.data[0].targets.length, 2, '該系統卡片底下的 targets 必須包含全部 2 個標的（資料庫與系統日誌檔），絕不可遺漏');
  assert.equal(resp.data[0].targets[0].targetName, '資料庫');
  assert.equal(resp.data[0].targets[1].targetName, '系統日誌檔');
});

test('getDashboardData 不得合併系統 ID 不同但名稱與管理人員相同的紀錄', () => {
  const mockRows = [
    ['2026/9/22 10:00:00', 'SYS-01', '同名系統', '一般系統', '資料庫', '王小明', '1 週', '無', '4 代', 'NAS-01', '每日 00:00', '磁碟自動覆寫', 'Log', '否', '', '', '', '無', '無', '無', '', ''],
    ['2026/9/23 10:00:00', 'SYS-02', '同名系統', '一般系統', '資料庫', '王小明', '1 週', '無', '4 代', 'NAS-02', '每日 00:00', '磁碟自動覆寫', 'Log', '否', '', '', '', '無', '無', '無', '', '']
  ];
  const mockSheet = {
    getLastRow: () => 3,
    getLastColumn: () => 22,
    getMaxColumns: () => 22,
    getRange: row => ({
      getValues: () => row === 1 ? [mod.EXPECTED_HEADERS] : mockRows,
      setValues: () => {}
    })
  };
  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: name => name === '備份管制' ? mockSheet : null })
  };
  const runner = new Function('SpreadsheetApp', 'Utilities', 'console', `
    ${codeJsContent}
    checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
    return getDashboardData();
  `);

  let uuidCounter = 0;
  const resp = runner(SpreadsheetApp, { getUuid: () => 'generated-' + (++uuidCounter) }, console);
  assert.equal(resp.success, true);
  assert.deepEqual(resp.data.map(item => item.systemId).sort(), ['SYS-01', 'SYS-02']);
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

test('submitData 遇到非標準 A:V 表頭時必須拒絕寫入並提示執行遷移', () => {
  let writeCalled = false;
  const malformedHeaders = [...mod.EXPECTED_HEADERS];
  [malformedHeaders[13], malformedHeaders[14]] = [malformedHeaders[14], malformedHeaders[13]];

  const mockSheet = {
    getLastRow: () => 2,
    getLastColumn: () => 22,
    getRange: (row) => ({
      getValues: () => row === 1 ? [malformedHeaders] : [[]],
      setValues: () => { writeCalled = true; }
    }),
    deleteRow: () => { writeCalled = true; }
  };
  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => name === '備份管制' ? mockSheet : null
    })
  };
  const Utilities = { getUuid: () => 'new-uuid' };
  const Session = { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) };
  const runner = new Function('SpreadsheetApp', 'Utilities', 'Session', 'console', `
    ${codeJsContent}
    checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
    return submitData({
      systemName: 'SYS-01 - 測試系統',
      systemLevel: '一般系統',
      managerName: '王小明',
      targets: [{
        targetKey: 'db', targetName: '資料庫',
        localFullValue: '1', localFullUnit: '週',
        localDiffValue: '1', localDiffUnit: '天',
        localRetention: '4', localLocation: 'NAS',
        isOffsite: '否'
      }]
    });
  `);

  const result = runner(SpreadsheetApp, Utilities, Session, console);
  assert.equal(result.success, false);
  assert.match(result.message, /runDataMigration/);
  assert.equal(writeCalled, false, 'schema 不符時不得覆寫、追加或刪除任何資料列');
});

test('submitData 編輯舊版前端 payload 時必須依標的保留各列既有 B 欄系統 ID', () => {
  const existingDbRow = [
    new Date('2026-09-23T01:00:00Z'), 'SYS-001', '測試系統', '一般系統', '資料庫', '王小明',
    '1 週', '無', '4 代', 'NAS', '每日 00:00', '磁碟自動覆寫', '備份系統執行Log',
    '否', '無', '無', '無', '無', '無', '無', 'sub-001', 'db'
  ];
  const existingSourceRow = [
    new Date('2026-09-23T01:00:00Z'), 'SYS-002', '測試系統', '一般系統', '原始碼', '王小明',
    '1 週', '無', '4 代', 'Git', '每日 00:00', '磁碟自動覆寫', '備份系統執行Log',
    '否', '無', '無', '無', '無', '無', '無', 'sub-001', 'source'
  ];
  const updatedRows = {};

  function makeSheet(rows, headers, captureUpdate) {
    return {
      getLastRow: () => rows.length,
      getLastColumn: () => headers.length,
      getRange: (row, col, numRows, numCols) => ({
        getValues: () => rows
          .slice(row - 1, row - 1 + numRows)
          .map(sourceRow => sourceRow.slice(col - 1, col - 1 + numCols)),
        setValues: values => {
          if (captureUpdate && row >= 2 && col === 1) updatedRows[row] = values[0];
          while (rows.length < row - 1 + values.length) rows.push([]);
          values.forEach((valueRow, index) => { rows[row - 1 + index] = valueRow.slice(); });
        },
        setFontWeight: () => {},
        setBackground: () => {}
      }),
      setFrozenRows: () => {},
      deleteRow: rowNumber => rows.splice(rowNumber - 1, 1)
    };
  }

  const mainSheet = makeSheet([mod.EXPECTED_HEADERS, existingDbRow, existingSourceRow], mod.EXPECTED_HEADERS, true);
  const historySheet = makeSheet([mod.EXPECTED_HISTORY_HEADERS], mod.EXPECTED_HISTORY_HEADERS, false);
  const spreadsheet = {
    getSheetByName: name => name === '備份管制' ? mainSheet : (name === '備份管制_歷程' ? historySheet : null),
    insertSheet: () => historySheet
  };
  const SpreadsheetApp = { getActiveSpreadsheet: () => spreadsheet };
  const Utilities = { getUuid: () => 'new-uuid' };
  const Session = { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) };
  const runner = new Function('SpreadsheetApp', 'Utilities', 'Session', 'console', `
    ${codeJsContent}
    checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
    return submitData({
      id: 'sub-001',
      systemName: '測試系統',
      systemLevel: '一般系統',
      managerName: '王小明',
      targets: [
        {
          targetKey: 'db', targetName: '資料庫',
          localFullValue: '1', localFullUnit: '週', localDiffValue: '0', localDiffUnit: '天',
          localRetention: '4', localLocation: 'NAS', isOffsite: '否'
        },
        {
          targetKey: 'source', targetName: '原始碼',
          localFullValue: '1', localFullUnit: '週', localDiffValue: '0', localDiffUnit: '天',
          localRetention: '4', localLocation: 'Git', isOffsite: '否'
        }
      ]
    });
  `);

  const result = runner(SpreadsheetApp, Utilities, Session, console);
  assert.equal(result.success, true);
  assert.equal(updatedRows[2][1], 'SYS-001', '資料庫列應保留自己的 B 欄');
  assert.equal(updatedRows[3][1], 'SYS-002', '原始碼列應保留自己的 B 欄');
  assert.equal(updatedRows[2][20], 'sub-001');
  assert.equal(updatedRows[3][20], 'sub-001');
});

test('backfillMissingTechnicalKeys 只能寫入 U:V，不得重寫 A:T 業務欄位', () => {
  const writes = [];
  const sourceRows = [
    ['2026/9/23 10:00:00', 'SYS-01', '測試系統', '一般系統', '資料庫', '王小明', '=FORMULA()', '', '', '', '', '', '', '否', '', '', '', '', '', '', '', '']
  ];
  const mockSheet = {
    getLastRow: () => 2,
    getLastColumn: () => 22,
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => row === 1 ? [mod.EXPECTED_HEADERS] : sourceRows,
      setValues: (values) => writes.push({ row, col, numRows, numCols, values })
    }),
    copyTo: () => ({ setName: () => {} })
  };
  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => mockSheet }),
    getUi: () => ({ alert: () => {} })
  };
  const Session = { getScriptTimeZone: () => 'Asia/Taipei' };
  const Utilities = {
    getUuid: () => 'generated-uuid',
    formatDate: () => '20260923_100000'
  };
  const runner = new Function('SpreadsheetApp', 'Session', 'Utilities', 'console', `
    ${codeJsContent}
    return backfillMissingTechnicalKeys();
  `);

  runner(SpreadsheetApp, Session, Utilities, console);
  assert.equal(writes.length, 1);
  assert.deepEqual(
    { row: writes[0].row, col: writes[0].col, numRows: writes[0].numRows, numCols: writes[0].numCols },
    { row: 2, col: 21, numRows: 1, numCols: 2 }
  );
  assert.deepEqual(writes[0].values, [['generated-uuid', 'db']]);
});

test('archiveSubmissionRows_ 永遠只封存 3 個稽核欄位加 A:V 快照', () => {
  let archived = null;
  const historySheet = {
    getLastRow: () => 1,
    getLastColumn: () => 25,
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => [mod.EXPECTED_HISTORY_HEADERS],
      setValues: (values) => { archived = { row, col, numRows, numCols, values }; }
    })
  };
  const ss = { getSheetByName: () => historySheet };
  const sourceRowWithTrailingData = [...Array.from({ length: 22 }, (_, i) => `col-${i + 1}`), 'W-殘留', 'X-殘留'];

  mod.archiveSubmissionRows_(ss, [sourceRowWithTrailingData], '更新前封存', 'admin@example.com');

  assert.equal(archived.numCols, 25);
  assert.equal(archived.values[0].length, 25);
  assert.deepEqual(archived.values[0].slice(3), sourceRowWithTrailingData.slice(0, 22));
});

test('getDashboardData 即使只有表頭也必須拒絕非標準 A:V schema', () => {
  const malformedHeaders = [...mod.EXPECTED_HEADERS];
  malformedHeaders[21] = '錯誤技術欄';
  const mockSheet = {
    getLastRow: () => 1,
    getLastColumn: () => 22,
    getRange: () => ({ getValues: () => [malformedHeaders] })
  };
  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => mockSheet })
  };
  const runner = new Function('SpreadsheetApp', 'console', `
    ${codeJsContent}
    checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
    return getDashboardData();
  `);

  const result = runner(SpreadsheetApp, console);
  assert.equal(result.success, false);
  assert.match(result.message, /runDataMigration/);
});

test('deleteSubmission 遇到非標準 A:V schema 時不得刪除資料', () => {
  let deleteCalled = false;
  const malformedHeaders = [...mod.EXPECTED_HEADERS];
  [malformedHeaders[20], malformedHeaders[21]] = [malformedHeaders[21], malformedHeaders[20]];
  const mockSheet = {
    getLastRow: () => 2,
    getLastColumn: () => 22,
    getRange: (row) => ({ getValues: () => row === 1 ? [malformedHeaders] : [[]] }),
    deleteRow: () => { deleteCalled = true; }
  };
  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => mockSheet })
  };
  const runner = new Function('SpreadsheetApp', 'console', `
    ${codeJsContent}
    checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
    return deleteSubmission('submission-1');
  `);

  const result = runner(SpreadsheetApp, console);
  assert.equal(result.success, false);
  assert.match(result.message, /runDataMigration/);
  assert.equal(deleteCalled, false);
});

test('generateBackupControlDocument 必須在建立 Drive 文件前拒絕非標準 schema', () => {
  let driveCalled = false;
  const malformedHeaders = [...mod.EXPECTED_HEADERS];
  malformedHeaders[13] = '異地欄位錯誤';
  const mockSheet = {
    getLastRow: () => 2,
    getLastColumn: () => 22,
    getRange: () => ({ getValues: () => [malformedHeaders] })
  };
  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => mockSheet })
  };
  const DriveApp = {
    getFileById: () => {
      driveCalled = true;
      throw new Error('Drive 不應被呼叫');
    },
    getFolderById: () => {
      driveCalled = true;
      throw new Error('Drive 不應被呼叫');
    }
  };
  const runner = new Function(
    'SpreadsheetApp', 'DriveApp', 'FORM_TEMPLATE_DOC_ID', 'FORM_OUTPUT_FOLDER_ID', 'console',
    `
      ${codeJsContent}
      checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
      return generateBackupControlDocument();
    `
  );

  const result = runner(SpreadsheetApp, DriveApp, 'template-id', 'folder-id', console);
  assert.equal(result.success, false);
  assert.match(result.message, /runDataMigration/);
  assert.equal(driveCalled, false, 'schema 不符時不可建立或複製任何 Drive 文件');
});

test('EXPECTED_HEADERS 包含新版銷毀方法與確認方式表頭', () => {
  assert.equal(mod.EXPECTED_HEADERS[11], '本地_到期銷毀方法');
  assert.equal(mod.EXPECTED_HEADERS[12], '本地_銷毀確認方式');
  assert.equal(mod.EXPECTED_HEADERS[18], '異地_到期銷毀方法');
  assert.equal(mod.EXPECTED_HEADERS[19], '異地_銷毀確認方式');
  assert.equal(mod.EXPECTED_HEADERS.length, 22);

  assert.equal(mod.EXPECTED_HISTORY_HEADERS[14], '本地_到期銷毀方法');
  assert.equal(mod.EXPECTED_HISTORY_HEADERS[15], '本地_銷毀確認方式');
  assert.equal(mod.EXPECTED_HISTORY_HEADERS[21], '異地_到期銷毀方法');
  assert.equal(mod.EXPECTED_HISTORY_HEADERS[22], '異地_銷毀確認方式');
  assert.equal(mod.EXPECTED_HISTORY_HEADERS.length, 25);
});

test('resolveHeaderIndices_ 正確相容解析新版表頭與歷史處置/佐證別名', () => {
  // 1. 新版標準表頭解析
  const newIndices = mod.resolveHeaderIndices_(mod.EXPECTED_HEADERS);
  assert.equal(newIndices.localDispositionMethod, 11);
  assert.equal(newIndices.localEvidenceRequirements, 12);
  assert.equal(newIndices.offsiteDispositionMethod, 18);
  assert.equal(newIndices.offsiteEvidenceRequirements, 19);

  // 2. 舊版歷史 22 欄表頭解析（向下相容）
  const LEGACY_22_HEADERS = [
    '時間戳記', '系統ID', '系統名稱', '系統等級', '備份標的', '管理人員',
    '本地_完整週期', '本地_差異週期', '本地_保留代數', '本地_存放地點',
    '本地_排程時間', '本地_處置方法', '本地_佐證種類',
    '是否異地', '異地_週期', '異地_地點', '異地_保留代數',
    '異地_排程時間', '異地_處置方法', '異地_佐證種類',
    'submission_id', 'target_key'
  ];
  const legacyIndices = mod.resolveHeaderIndices_(LEGACY_22_HEADERS);
  assert.equal(legacyIndices.localDispositionMethod, 11, '歷史表頭本地_處置方法應正確解析');
  assert.equal(legacyIndices.localEvidenceRequirements, 12, '歷史表頭本地_佐證種類應正確解析');
  assert.equal(legacyIndices.offsiteDispositionMethod, 18, '歷史表頭異地_處置方法應正確解析');
  assert.equal(legacyIndices.offsiteEvidenceRequirements, 19, '歷史表頭異地_佐證種類應正確解析');
});

test('migrateSheetIfNeeded_ 將歷史舊版處置/佐證表頭平滑升級為新版表頭且資料無損', () => {
  const LEGACY_22_HEADERS = [
    '時間戳記', '系統ID', '系統名稱', '系統等級', '備份標的', '管理人員',
    '本地_完整週期', '本地_差異週期', '本地_保留代數', '本地_存放地點',
    '本地_排程時間', '本地_處置方法', '本地_佐證種類',
    '是否異地', '異地_週期', '異地_地點', '異地_保留代數',
    '異地_排程時間', '異地_處置方法', '異地_佐證種類',
    'submission_id', 'target_key'
  ];
  const legacyRow = [
    '2026/09/20 12:00:00', 'SYS-001', '人資系統', '關鍵核心系統', '資料庫', '王小明',
    '1 週', '1 天', '4 代', 'NAS-01',
    '每日 00:00', '磁碟自動覆寫', '備份系統執行Log',
    '是', '1 週', '雲端空間', '4 代',
    '每日 00:00', '磁碟自動覆寫', '備份系統執行Log',
    'sub-001', 'db'
  ];

  const sheetData = [LEGACY_22_HEADERS.slice(), legacyRow.slice()];
  let writtenValues = null;

  const mockSheet = {
    getLastRow: () => sheetData.length,
    getLastColumn: () => 22,
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => sheetData.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols)),
      setValues: (values) => {
        writtenValues = values;
        values.forEach((r, idx) => {
          sheetData[row - 1 + idx] = r.slice();
        });
      },
      setFontWeight: () => {},
      setBackground: () => {}
    }),
    clearContents: () => {
      sheetData.length = 0;
    },
    setFrozenRows: () => {}
  };

  mod.migrateSheetIfNeeded_(mockSheet, mod.EXPECTED_HEADERS, false);

  assert.deepEqual(sheetData[0], mod.EXPECTED_HEADERS, '表頭應成功更新為新版表頭');
  assert.equal(sheetData[1][11], '磁碟自動覆寫', '第 12 欄到期銷毀方法資料應完整保留');
  assert.equal(sheetData[1][12], '備份系統執行Log', '第 13 欄銷毀確認方式資料應完整保留');
  assert.equal(sheetData[1][18], '磁碟自動覆寫', '第 19 欄異地到期銷毀方法資料應完整保留');
  assert.equal(sheetData[1][19], '備份系統執行Log', '第 20 欄異地銷毀確認方式資料應完整保留');
  assert.equal(sheetData[1][20], 'sub-001');
  assert.equal(sheetData[1][21], 'db');
});

test('assertSheetSchema_ 與 getDashboardData 在使用歷史 22 欄表頭（本地_處置方法等）時必須相容通過', () => {
  const LEGACY_22_HEADERS = [
    '時間戳記', '系統ID', '系統名稱', '系統等級', '備份標的', '管理人員',
    '本地_完整週期', '本地_差異週期', '本地_保留代數', '本地_存放地點',
    '本地_排程時間', '本地_處置方法', '本地_佐證種類',
    '是否異地', '異地_週期', '異地_地點', '異地_保留代數',
    '異地_排程時間', '異地_處置方法', '異地_佐證種類',
    'submission_id', 'target_key'
  ];
  const legacyRow = [
    '2026/09/20 12:00:00', 'SYS-001', '人資系統', '關鍵核心系統', '資料庫', '王小明',
    '1 週', '1 天', '4 代', 'NAS-01',
    '每日 00:00', '磁碟自動覆寫', '備份系統執行Log',
    '是', '1 週', '雲端空間', '4 代',
    '每日 00:00', '磁碟自動覆寫', '備份系統執行Log',
    'sub-001', 'db'
  ];

  const mockSheet = {
    getLastRow: () => 2,
    getLastColumn: () => 22,
    getRange: (r, c, nr, nc) => {
      if (r === 1) return { getValues: () => [LEGACY_22_HEADERS] };
      return { getValues: () => [legacyRow], setValues: () => {} };
    }
  };

  const SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: () => mockSheet })
  };

  const runner = new Function('SpreadsheetApp', 'console', `
    ${codeJsContent}
    checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
    return getDashboardData();
  `);

  const result = runner(SpreadsheetApp, console);
  assert.equal(result.success, true, '歷史 22 欄表頭工作表必須相容通過，不可回傳失敗');
  assert.equal(result.data.length, 1);
  assert.equal(result.data[0].systemName, '人資系統');
  assert.equal(result.data[0].targets[0].localDispositionMethod, '磁碟自動覆寫');
});


// --- 完整備份改為選填（完整、差異二擇一必填）---

function buildRowFor_(targetOverrides) {
  return mod.buildSheetRowData_({
    timestamp: new Date('2026-11-09T00:00:00Z'),
    systemId: 'SYS-001',
    systemName: '測試系統',
    systemLevel: '一般系統',
    managerName: '王小明',
    submissionId: 'sub-001',
    targetKey: 'db',
    target: Object.assign({
      targetName: '資料庫',
      localFullValue: '1', localFullUnit: '週',
      localDiffValue: '0', localDiffUnit: '天',
      localRetention: '4', localLocation: 'NAS',
      isOffsite: '否'
    }, targetOverrides)
  });
}

test('isCycleEnabled_ 僅數值 > 0 視為啟用', () => {
  assert.equal(mod.isCycleEnabled_('1'), true);
  assert.equal(mod.isCycleEnabled_('0'), false);
  assert.equal(mod.isCycleEnabled_(''), false);
  assert.equal(mod.isCycleEnabled_(undefined), false);
  assert.equal(mod.isCycleEnabled_('N/A'), false);
  assert.equal(mod.isCycleEnabled_('無'), false);
});

test('buildSheetRowData_ 完整週期為空白或 0 時寫入「無」，不得產生殘缺字串', () => {
  const diffOnly = buildRowFor_({ localFullValue: '0', localFullUnit: '週', localDiffValue: '1', localDiffUnit: '天' });
  assert.equal(diffOnly[6], '無', '本地_完整週期應為「無」');
  assert.equal(diffOnly[7], '1 天');

  const blankFull = buildRowFor_({ localFullValue: '', localFullUnit: '週', localDiffValue: '1', localDiffUnit: '天' });
  assert.equal(blankFull[6], '無');
  assert.ok(!/^\s*週/.test(String(blankFull[6])), '不得寫成「 週」');

  const fullOnly = buildRowFor_({});
  assert.equal(fullOnly[6], '1 週');
  assert.equal(fullOnly[7], '無');

  const both = buildRowFor_({ localDiffValue: '1', localDiffUnit: '天' });
  assert.equal(both[6], '1 週');
  assert.equal(both[7], '1 天');
});

test('僅差異備份寫入「無」後，parseTargetFromRow 讀回完整週期為空值且差異週期保留', () => {
  const row = buildRowFor_({
    localFullValue: '0', localFullUnit: '週', localDiffValue: '1', localDiffUnit: '天',
    localScheduleTime: '差異:每日 01:30'
  });
  const indices = mod.resolveHeaderIndices_(mod.EXPECTED_HEADERS);
  const parsed = mod.parseTargetFromRow(row, indices, 'db');
  assert.equal(parsed.localFullValue, '');
  assert.equal(parsed.localDiffValue, '1');
  assert.equal(parsed.localDiffUnit, '天');
  assert.equal(parsed.localScheduleTime, '差異:每日 01:30');
});

function runSubmitData_(targets) {
  const rows = [mod.EXPECTED_HEADERS];
  const written = [];
  const sheet = {
    getLastRow: () => rows.length,
    getLastColumn: () => mod.EXPECTED_HEADERS.length,
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => rows.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols)),
      setValues: values => { values.forEach(v => written.push(v.slice())); },
      setFontWeight: () => {},
      setBackground: () => {}
    }),
    setFrozenRows: () => {},
    deleteRow: () => {}
  };
  const historySheet = Object.assign({}, sheet, { getLastRow: () => 1 });
  const spreadsheet = {
    getSheetByName: name => name === '備份管制' ? sheet : (name === '備份管制_歷程' ? historySheet : null),
    insertSheet: () => historySheet
  };
  const runner = new Function('SpreadsheetApp', 'Utilities', 'Session', 'console', 'targets', `
    ${codeJsContent}
    checkAppPermission_ = () => ({ allowed: true, email: 'admin@example.com' });
    return submitData({
      systemId: 'SYS-001', systemName: '測試系統', systemLevel: '一般系統', managerName: '王小明',
      targets: targets
    });
  `);
  const result = runner(
    { getActiveSpreadsheet: () => spreadsheet },
    { getUuid: () => 'new-uuid' },
    { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
    console,
    targets
  );
  return { result, written };
}

test('submitData 完整與差異備份皆未設定時必須拒絕寫入', () => {
  const { result, written } = runSubmitData_([{
    targetKey: 'db', targetName: '資料庫',
    localFullValue: '0', localFullUnit: '週', localDiffValue: '0', localDiffUnit: '天',
    localRetention: '4', localLocation: 'NAS', isOffsite: '否'
  }]);
  assert.equal(result.success, false);
  assert.ok(String(result.message).includes('至少須設定其中一種'));
  assert.equal(written.length, 0, '驗證失敗時不得寫入任何資料列');
});

test('submitData 僅差異備份可成功寫入，且完整週期欄為「無」', () => {
  const { result, written } = runSubmitData_([{
    targetKey: 'db', targetName: '資料庫',
    localFullValue: '0', localFullUnit: '週', localDiffValue: '1', localDiffUnit: '天',
    localScheduleTime: '差異:每日 02:00',
    localRetention: '4', localLocation: 'NAS', isOffsite: '否'
  }]);
  assert.equal(result.success, true, result.message);
  assert.equal(written.length, 1);
  assert.equal(written[0][6], '無');
  assert.equal(written[0][7], '1 天');
  assert.equal(written[0][10], '差異:每日 02:00');
});

test('submitData 系統日誌檔不受完整、差異二擇一規則限制', () => {
  const { result, written } = runSubmitData_([{
    targetKey: 'log', targetName: '系統日誌檔',
    localFullValue: '不適用', localDiffValue: '不適用',
    logRetention: '保存 6 個月', localLocation: 'NAS', isOffsite: '否'
  }]);
  assert.equal(result.success, true, result.message);
  assert.equal(written.length, 1);
});


// --- 保留所有版本：到期銷毀方法與銷毀確認方式一律寫入「無」---

test('buildSheetRowData_ 本地保留代數為全部保留時，銷毀方法與確認方式寫入「無」', () => {
  ['ALL', '全部保留'].forEach(function(retention) {
    const row = buildRowFor_({
      localRetention: retention,
      localDispositionMethod: '人工手動覆蓋',
      localEvidenceRequirements: '備份檔名確認'
    });
    assert.equal(row[8], '全部保留', '本地_保留代數');
    assert.equal(row[11], '無', '本地_到期銷毀方法：' + retention);
    assert.equal(row[12], '無', '本地_銷毀確認方式：' + retention);
  });
});

test('buildSheetRowData_ 一般保留代數時維持使用者選擇與預設值', () => {
  const chosen = buildRowFor_({
    localRetention: '4',
    localDispositionMethod: '人工手動覆蓋',
    localEvidenceRequirements: '備份檔名確認'
  });
  assert.equal(chosen[11], '人工手動覆蓋');
  assert.equal(chosen[12], '備份檔名確認');

  const defaults = buildRowFor_({ localRetention: '4' });
  assert.equal(defaults[11], '磁碟自動覆寫');
  assert.equal(defaults[12], '備份系統執行Log');
});

test('buildSheetRowData_ 全部保留不影響異地備份的銷毀設定', () => {
  const row = buildRowFor_({
    localRetention: 'ALL',
    isOffsite: '是',
    offsiteFreqValue: '1', offsiteFreqUnit: '個月',
    offsiteLocation: '國網中心', offsiteRetention: '2',
    offsiteDispositionMethod: '人工手動覆蓋',
    offsiteEvidenceRequirements: '備份檔名確認'
  });
  assert.equal(row[11], '無');
  assert.equal(row[12], '無');
  assert.equal(row[18], '人工手動覆蓋', '異地_到期銷毀方法不受影響');
  assert.equal(row[19], '備份檔名確認', '異地_銷毀確認方式不受影響');
});

test('submitData 即使前端（或舊版前端）送來預設銷毀值，全部保留仍強制寫「無」', () => {
  const { result, written } = runSubmitData_([{
    targetKey: 'db', targetName: '資料庫',
    localFullValue: '1', localFullUnit: '週', localDiffValue: '0', localDiffUnit: '天',
    localRetention: 'ALL', localLocation: 'NAS', isOffsite: '否',
    localDispositionMethod: '磁碟自動覆寫',
    localEvidenceRequirements: '備份系統執行Log'
  }]);
  assert.equal(result.success, true, result.message);
  assert.equal(written[0][8], '全部保留');
  assert.equal(written[0][11], '無');
  assert.equal(written[0][12], '無');
});

test('parseTargetFromRow 讀取已存為「無」的銷毀欄位時不得被預設值蓋回', () => {
  const row = buildRowFor_({ localRetention: 'ALL' });
  const indices = mod.resolveHeaderIndices_(mod.EXPECTED_HEADERS);
  const parsed = mod.parseTargetFromRow(row, indices, 'db');
  assert.equal(parsed.localRetention, 'ALL');
  assert.equal(parsed.localDispositionMethod, '無');
  assert.equal(parsed.localEvidenceRequirements, '無');
});
