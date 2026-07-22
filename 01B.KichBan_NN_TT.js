const FSNT_CFG = Object.freeze({
  SUMMARY_NN: '00. Tổng hợp',
  SUMMARY_TT: '00.Tổng hợp_thực tế',
  CHECKS: '99. Kiểm tra',
  TECH_NN: '01A. Kỹ thuật',
  TECH_TT: '01A. Kỹ thuật_thực tế',
  BACKUP_SUMMARY_NN: '__FSNT_BACKUP_SUMMARY_NN__',
  BACKUP_TECH_NN: '__FSNT_BACKUP_TECH_NN__',
  LEGACY_TRIGGER_HANDLER: 'FSNT_tiepTucHaiKichBan_'
});

function FSNT_tiepTucHaiKichBan_() {
  FSNT_deleteLegacyTriggers_();
}

function FS_chayHaiKichBanNN_TT() {
  FSNT_deleteLegacyTriggers_();
  SpreadsheetApp.getUi().alert(
    'Chức năng chạy nối tiếp NN/TT đã dừng.\n' +
    'Hãy chạy riêng hai menu:\n' +
    '- Chạy toàn bộ mô hình - Định mức (NN)\n' +
    '- Chạy toàn bộ mô hình - Thực tế (TT)'
  );
}

function FS_huyChayHaiKichBan() {
  FSNT_deleteLegacyTriggers_();
  SpreadsheetApp.getUi().alert('Đã xóa các trigger NN/TT cũ.');
}

/**
 * Chạy mô hình Định mức.
 * 01A. Kỹ thuật luôn được trả về trạng thái NN.
 */
function FS_chayKichBanDinhMuc() {
  FSNT_deleteLegacyTriggers_();
  FSNT_runScenario_('NN');
  FSNT_assertNoFail_();
  FSNT_markSummary_(FSNT_CFG.SUMMARY_NN, 'NN');
  FSNT_sapXepSheetNN_TT_();
  SpreadsheetApp.flush();
}

/**
 * Chạy mô hình Thực tế.
 * - Lưu riêng dữ liệu kỹ thuật TT tại 01A. Kỹ thuật_thực tế.
 * - Chốt kết quả TT tại 00.Tổng hợp_thực tế.
 * - Khôi phục 01A. Kỹ thuật và 00. Tổng hợp về trạng thái NN sau khi chạy.
 */
function FS_chayKichBanThucTe() {
  FSNT_deleteLegacyTriggers_();
  const ss = SpreadsheetApp.getActive();

  // Luôn tái tạo NN trước khi sao lưu để tránh lấy nhầm trạng thái TT của lần chạy trước.
  FS_taoKyThuatTuDauVao('NN');
  SpreadsheetApp.flush();
  FSNT_assertActiveScenario_('NN');

  FSNT_backupSheet_(ss, FSNT_CFG.TECH_NN, FSNT_CFG.BACKUP_TECH_NN);
  FSNT_backupSheet_(ss, FSNT_CFG.SUMMARY_NN, FSNT_CFG.BACKUP_SUMMARY_NN);

  try {
    // Dựng bộ kỹ thuật TT trên sheet làm việc 01A rồi lưu bản riêng để kiểm tra/audit.
    FS_taoKyThuatTuDauVao('TT');
    SpreadsheetApp.flush();
    FSNT_assertActiveScenario_('TT');
    FSNT_snapshotSheet_(ss, FSNT_CFG.TECH_NN, FSNT_CFG.TECH_TT, false);

    // Chạy các sheet lõi từ đúng bộ kỹ thuật TT đang hiện hành.
    FSNT_runModelFromActiveTech_('TT');
    FSNT_assertNoFail_();
    FSNT_snapshotSummary_(FSNT_CFG.SUMMARY_TT, 'TT');
  } finally {
    // Khôi phục trạng thái NN, kể cả khi quá trình TT phát sinh lỗi.
    FSNT_restoreSheetInPlace_(ss, FSNT_CFG.BACKUP_TECH_NN, FSNT_CFG.TECH_NN);
    FSNT_restoreSummary_(ss);
    FSNT_sapXepSheetNN_TT_();
  }

  SpreadsheetApp.flush();
}

function FSNT_runScenario_(scenario) {
  FS_taoKyThuatTuDauVao(scenario);
  SpreadsheetApp.flush();
  FSNT_assertActiveScenario_(scenario);
  FSNT_runModelFromActiveTech_(scenario);
}

function FSNT_runModelFromActiveTech_(scenario) {
  FSNT_assertActiveScenario_(scenario);

  FSNT_callRequired_('FS_lapSheet02');
  FSNT_callRequired_('FS_lapSheet03');

  if (typeof globalThis.FS_hoiTuTaiTro === 'function') {
    globalThis.FS_hoiTuTaiTro();
  } else {
    FSNT_callRequired_('FS_lapSheet03A');
    FSNT_callRequired_('FS_lapSheet04');
  }

  FSNT_callOptional_('FS_lapSheet04A');

  if (typeof globalThis.FS_lapSheet00_TheoDanhMuc === 'function') {
    globalThis.FS_lapSheet00_TheoDanhMuc();
  } else {
    FSNT_callRequired_('FS_lapSheet00');
  }

  if (typeof globalThis.FS00I_phanTichHieuQuaTheoSanPham === 'function') {
    globalThis.FS00I_phanTichHieuQuaTheoSanPham();
  } else if (typeof globalThis.FS00F_phanTichHieuQuaTheoSanPham === 'function') {
    globalThis.FS00F_phanTichHieuQuaTheoSanPham();
  }

  FSNT_callOptional_('FS_lapSheet99');
  SpreadsheetApp.flush();
}

function FSNT_assertActiveScenario_(expected) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(FSNT_CFG.TECH_NN);
  if (!sheet) throw new Error('Không tìm thấy sheet 01A. Kỹ thuật.');

  const values = sheet.getDataRange().getValues();
  let actual = '';
  for (let r = 0; r < values.length; r++) {
    if (String(values[r][0] || '').trim() === 'Kịch bản chi phí') {
      actual = String(values[r][1] || '').trim().toUpperCase();
      break;
    }
  }

  if (actual !== expected) {
    throw new Error('Sai kịch bản tại 01A. Kỹ thuật. Yêu cầu ' + expected + ', thực tế ' + actual + '.');
  }
}

function FSNT_backupSheet_(ss, sourceName, backupName) {
  const oldBackup = ss.getSheetByName(backupName);
  if (oldBackup) ss.deleteSheet(oldBackup);

  const source = ss.getSheetByName(sourceName);
  if (!source) return null;

  return source.copyTo(ss).setName(backupName).hideSheet();
}

/**
 * Khôi phục nội dung vào đúng sheet đích để không làm thay đổi sheetId và các tham chiếu.
 */
function FSNT_restoreSheetInPlace_(ss, backupName, targetName) {
  const backup = ss.getSheetByName(backupName);
  if (!backup) return;

  let target = ss.getSheetByName(targetName);
  if (!target) target = ss.insertSheet(targetName);

  const sourceRows = Math.max(1, backup.getMaxRows());
  const sourceCols = Math.max(1, backup.getMaxColumns());
  if (target.getMaxRows() < sourceRows) {
    target.insertRowsAfter(target.getMaxRows(), sourceRows - target.getMaxRows());
  }
  if (target.getMaxColumns() < sourceCols) {
    target.insertColumnsAfter(target.getMaxColumns(), sourceCols - target.getMaxColumns());
  }

  target.getRange(1, 1, target.getMaxRows(), target.getMaxColumns()).breakApart();
  target.clear();
  backup.getRange(1, 1, sourceRows, sourceCols).copyTo(target.getRange(1, 1, sourceRows, sourceCols));

  target.setFrozenRows(backup.getFrozenRows());
  target.setFrozenColumns(backup.getFrozenColumns());
  for (let c = 1; c <= sourceCols; c++) target.setColumnWidth(c, backup.getColumnWidth(c));

  ss.deleteSheet(backup);
}

function FSNT_restoreSummary_(ss) {
  const backup = ss.getSheetByName(FSNT_CFG.BACKUP_SUMMARY_NN);
  if (!backup) return;

  const current = ss.getSheetByName(FSNT_CFG.SUMMARY_NN);
  if (current) ss.deleteSheet(current);

  backup.showSheet().setName(FSNT_CFG.SUMMARY_NN);
  FSNT_markSummary_(FSNT_CFG.SUMMARY_NN, 'NN');
}

/**
 * Sao chép một sheet. freezeValues=true dùng cho báo cáo kết quả cần chốt số.
 */
function FSNT_snapshotSheet_(ss, sourceName, targetName, freezeValues) {
  const source = ss.getSheetByName(sourceName);
  if (!source) throw new Error('Không tìm thấy sheet nguồn "' + sourceName + '".');

  const existing = ss.getSheetByName(targetName);
  if (existing) ss.deleteSheet(existing);

  const snapshot = source.copyTo(ss).setName(targetName);
  SpreadsheetApp.flush();

  if (freezeValues) {
    const range = snapshot.getDataRange();
    range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }

  return snapshot;
}

function FSNT_snapshotSummary_(targetName, scenario) {
  const ss = SpreadsheetApp.getActive();
  FSNT_snapshotSheet_(ss, FSNT_CFG.SUMMARY_NN, targetName, true);
  FSNT_markSummary_(targetName, scenario);
}

function FSNT_markSummary_(sheetName, scenario) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) return;

  const label = scenario === 'TT' ? 'THỰC TẾ (TT)' : 'ĐỊNH MỨC (NN)';
  sheet.getRange('A1').setNote('Nguồn chi phí: ' + label);

  const title = sheet.getRange('A2').getDisplayValue();
  if (title && !title.includes('[' + label + ']')) {
    const cleanTitle = title.replace(/\s*\[(?:ĐỊNH MỨC \(NN\)|THỰC TẾ \(TT\))\]\s*$/i, '');
    sheet.getRange('A2').setValue(cleanTitle + ' [' + label + ']');
  }
}

/**
 * Sắp xếp các sheet chính để dễ theo dõi:
 * 1. 00. Tổng hợp
 * 2. 00.Tổng hợp_thực tế
 * 01A. Kỹ thuật_thực tế nằm ngay sau 01A. Kỹ thuật.
 */
function FSNT_sapXepSheetNN_TT_() {
  const ss = SpreadsheetApp.getActive();

  FSNT_moveSheetTo_(ss, FSNT_CFG.SUMMARY_NN, 1);
  FSNT_moveSheetTo_(ss, FSNT_CFG.SUMMARY_TT, 2);

  const techNN = ss.getSheetByName(FSNT_CFG.TECH_NN);
  const techTT = ss.getSheetByName(FSNT_CFG.TECH_TT);
  if (techNN && techTT) {
    const techNNIndex = ss.getSheets().findIndex(sheet => sheet.getSheetId() === techNN.getSheetId());
    if (techNNIndex >= 0) {
      FSNT_moveSheetTo_(ss, FSNT_CFG.TECH_TT, techNNIndex + 2);
    }
  }
}

function FSNT_moveSheetTo_(ss, sheetName, position) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const maxPosition = ss.getSheets().length;
  const targetPosition = Math.max(1, Math.min(position, maxPosition));
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(targetPosition);
}

function FSNT_assertNoFail_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(FSNT_CFG.CHECKS);
  if (!sheet) return;

  const values = sheet.getDataRange().getDisplayValues();
  const failures = [];

  values.forEach((row, r) => {
    row.forEach((value, c) => {
      if (String(value || '').trim().toUpperCase() === 'FAIL') {
        failures.push(sheet.getRange(r + 1, c + 1).getA1Notation());
      }
    });
  });

  if (failures.length) {
    throw new Error(
      'Sheet 99 còn FAIL tại: ' + failures.slice(0, 20).join(', ') +
      (failures.length > 20 ? '...' : '')
    );
  }
}

function FSNT_deleteLegacyTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === FSNT_CFG.LEGACY_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function FSNT_callRequired_(functionName) {
  const fn = globalThis[functionName];
  if (typeof fn !== 'function') {
    throw new Error('Không tìm thấy hàm bắt buộc: ' + functionName);
  }
  return fn();
}

function FSNT_callOptional_(functionName) {
  const fn = globalThis[functionName];
  if (typeof fn === 'function') return fn();
  return null;
}
