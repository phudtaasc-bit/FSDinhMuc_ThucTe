const FSFIX_CFG = Object.freeze({
  SUMMARY_NN: '00. Tổng hợp',
  SUMMARY_TT: '00.Tổng hợp_thực tế',
  TECH_NN: '01A. Kỹ thuật',
  TECH_TT: '01A. Kỹ thuật_thực tế',
  MASTER_SUMMARY_NN: '__FSFIX_MASTER_SUMMARY_NN__',
  MASTER_TECH_NN: '__FSFIX_MASTER_TECH_NN__'
});

/**
 * Chạy và chốt kịch bản Định mức (NN).
 * Sau khi hoàn thành, lưu một bản master đã đóng băng để TT không thể làm thay đổi kết quả NN.
 */
function FS_FIX_chayNN() {
  const ss = SpreadsheetApp.getActive();

  if (typeof FSNT_deleteLegacyTriggers_ === 'function') FSNT_deleteLegacyTriggers_();
  FSNT_runScenario_('NN');
  FSNT_assertNoFail_();
  FSNT_markSummary_(FSFIX_CFG.SUMMARY_NN, 'NN');
  SpreadsheetApp.flush();

  FSFIX_saveMaster_(ss, FSFIX_CFG.SUMMARY_NN, FSFIX_CFG.MASTER_SUMMARY_NN);
  FSFIX_saveMaster_(ss, FSFIX_CFG.TECH_NN, FSFIX_CFG.MASTER_TECH_NN);

  if (typeof FSNT_sapXepSheetNN_TT_ === 'function') FSNT_sapXepSheetNN_TT_();
  SpreadsheetApp.getUi().alert(
    'Đã chạy xong kịch bản Định mức (NN).\n' +
    'Kết quả NN đã được chốt làm bản gốc để phục hồi sau khi chạy TT.'
  );
}

/**
 * Chạy kịch bản Thực tế (TT), chốt kết quả riêng, sau đó phục hồi NN từ bản master.
 * Bắt buộc phải chạy menu NN ít nhất một lần trước menu này.
 */
function FS_FIX_chayTT() {
  const ss = SpreadsheetApp.getActive();

  if (typeof FSNT_deleteLegacyTriggers_ === 'function') FSNT_deleteLegacyTriggers_();
  if (!ss.getSheetByName(FSFIX_CFG.MASTER_SUMMARY_NN) || !ss.getSheetByName(FSFIX_CFG.MASTER_TECH_NN)) {
    throw new Error(
      'Chưa có bản gốc Định mức. Hãy chạy menu "Chạy toàn bộ mô hình - Định mức (NN)" trước.'
    );
  }

  try {
    FS_taoKyThuatTuDauVao('TT');
    SpreadsheetApp.flush();
    FSNT_assertActiveScenario_('TT');

    FSFIX_copyLiveSheet_(ss, FSFIX_CFG.TECH_NN, FSFIX_CFG.TECH_TT, true);

    FSNT_runModelFromActiveTech_('TT');
    FSNT_assertNoFail_();
    FSFIX_copyLiveSheet_(ss, FSFIX_CFG.SUMMARY_NN, FSFIX_CFG.SUMMARY_TT, true);
    FSNT_markSummary_(FSFIX_CFG.SUMMARY_TT, 'TT');
  } finally {
    FSFIX_restoreFromMaster_(ss, FSFIX_CFG.MASTER_TECH_NN, FSFIX_CFG.TECH_NN);
    FSFIX_restoreFromMaster_(ss, FSFIX_CFG.MASTER_SUMMARY_NN, FSFIX_CFG.SUMMARY_NN);
    FSNT_markSummary_(FSFIX_CFG.SUMMARY_NN, 'NN');
    if (typeof FSNT_sapXepSheetNN_TT_ === 'function') FSNT_sapXepSheetNN_TT_();
    SpreadsheetApp.flush();
  }

  FSFIX_assertOutputsDiffer_(ss);
  SpreadsheetApp.getUi().alert(
    'Đã chạy xong kịch bản Thực tế (TT).\n' +
    '- 00. Tổng hợp: Định mức (NN)\n' +
    '- 00.Tổng hợp_thực tế: Thực tế (TT)'
  );
}

function FSFIX_saveMaster_(ss, sourceName, masterName) {
  const old = ss.getSheetByName(masterName);
  if (old) ss.deleteSheet(old);

  const source = ss.getSheetByName(sourceName);
  if (!source) throw new Error('Không tìm thấy sheet nguồn "' + sourceName + '".');

  const master = source.copyTo(ss).setName(masterName);
  SpreadsheetApp.flush();
  const range = master.getDataRange();
  range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  master.hideSheet();
  return master;
}

function FSFIX_copyLiveSheet_(ss, sourceName, targetName, freezeValues) {
  const source = ss.getSheetByName(sourceName);
  if (!source) throw new Error('Không tìm thấy sheet nguồn "' + sourceName + '".');

  const old = ss.getSheetByName(targetName);
  if (old) ss.deleteSheet(old);

  const target = source.copyTo(ss).setName(targetName);
  SpreadsheetApp.flush();
  if (freezeValues) {
    const range = target.getDataRange();
    range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }
  return target;
}

function FSFIX_restoreFromMaster_(ss, masterName, targetName) {
  const master = ss.getSheetByName(masterName);
  if (!master) throw new Error('Không tìm thấy bản master "' + masterName + '".');

  let target = ss.getSheetByName(targetName);
  if (!target) target = ss.insertSheet(targetName);

  const rows = Math.max(1, master.getMaxRows());
  const cols = Math.max(1, master.getMaxColumns());
  if (target.getMaxRows() < rows) target.insertRowsAfter(target.getMaxRows(), rows - target.getMaxRows());
  if (target.getMaxColumns() < cols) target.insertColumnsAfter(target.getMaxColumns(), cols - target.getMaxColumns());

  target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).breakApart();
  target.clear();
  master.getRange(1, 1, rows, cols).copyTo(target.getRange(1, 1, rows, cols));
  target.setFrozenRows(master.getFrozenRows());
  target.setFrozenColumns(master.getFrozenColumns());
  for (let c = 1; c <= cols; c++) target.setColumnWidth(c, master.getColumnWidth(c));
}

function FSFIX_assertOutputsDiffer_(ss) {
  const nn = ss.getSheetByName(FSFIX_CFG.SUMMARY_NN);
  const tt = ss.getSheetByName(FSFIX_CFG.SUMMARY_TT);
  if (!nn || !tt) return;

  const nnValue = Number(nn.getRange('D31').getValue()) || 0;
  const ttValue = Number(tt.getRange('D31').getValue()) || 0;
  const techNN = FSFIX_readTechCost_(ss.getSheetByName(FSFIX_CFG.TECH_NN), 'Chi phí XD/TB/khác');
  const techTT = FSFIX_readTechCost_(ss.getSheetByName(FSFIX_CFG.TECH_TT), 'Chi phí XD/TB/khác');

  if (Math.abs(techNN - techTT) > 1 && Math.abs(nnValue - ttValue) < 0.000001) {
    throw new Error(
      'Hai nguồn chi phí NN/TT khác nhau nhưng kết quả Tổng vốn đầu tư tại hai sheet 00 vẫn trùng. ' +
      'NN CPXD=' + techNN + '; TT CPXD=' + techTT + '; D31=' + nnValue + '. '
    );
  }
}

function FSFIX_readTechCost_(sheet, itemName) {
  if (!sheet) return 0;
  const values = sheet.getDataRange().getValues();
  const target = String(itemName || '').trim().toLowerCase();
  for (let r = 0; r < values.length; r++) {
    if (String(values[r][0] || '').trim().toLowerCase() === target) {
      return Number(values[r][1]) || 0;
    }
  }
  return 0;
}
