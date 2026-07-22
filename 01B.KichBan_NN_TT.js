const FSNT_CFG = Object.freeze({
  SUMMARY_NN: '00. Tổng hợp',
  SUMMARY_TT: '00.Tổng hợp_thực tế',
  CHECKS: '99. Kiểm tra',
  TECH: '01A. Kỹ thuật',
  BACKUP_NN: '__FSNT_BACKUP_NN__',
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

function FS_chayKichBanDinhMuc() {
  FSNT_deleteLegacyTriggers_();
  FSNT_runScenario_('NN');
  FSNT_assertNoFail_();
  FSNT_markSummary_(FSNT_CFG.SUMMARY_NN, 'NN');
  SpreadsheetApp.flush();
}

function FS_chayKichBanThucTe() {
  FSNT_deleteLegacyTriggers_();
  const ss = SpreadsheetApp.getActive();
  FSNT_backupCurrentNN_(ss);

  try {
    FSNT_runScenario_('TT');
    FSNT_assertNoFail_();
    FSNT_snapshotSummary_(FSNT_CFG.SUMMARY_TT, 'TT');
  } finally {
    FSNT_restoreNN_(ss);
  }

  SpreadsheetApp.flush();
}

function FSNT_runScenario_(scenario) {
  FS_taoKyThuatTuDauVao(scenario);
  SpreadsheetApp.flush();
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
  const sheet = SpreadsheetApp.getActive().getSheetByName(FSNT_CFG.TECH);
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

function FSNT_backupCurrentNN_(ss) {
  const oldBackup = ss.getSheetByName(FSNT_CFG.BACKUP_NN);
  if (oldBackup) ss.deleteSheet(oldBackup);

  const source = ss.getSheetByName(FSNT_CFG.SUMMARY_NN);
  if (!source) return;

  source.copyTo(ss).setName(FSNT_CFG.BACKUP_NN).hideSheet();
}

function FSNT_restoreNN_(ss) {
  const backup = ss.getSheetByName(FSNT_CFG.BACKUP_NN);
  if (!backup) return;

  const current = ss.getSheetByName(FSNT_CFG.SUMMARY_NN);
  if (current) ss.deleteSheet(current);

  backup.showSheet().setName(FSNT_CFG.SUMMARY_NN);
  FSNT_markSummary_(FSNT_CFG.SUMMARY_NN, 'NN');
}

function FSNT_snapshotSummary_(targetName, scenario) {
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName(FSNT_CFG.SUMMARY_NN);
  if (!source) throw new Error('Không tìm thấy sheet nguồn "' + FSNT_CFG.SUMMARY_NN + '".');

  const existing = ss.getSheetByName(targetName);
  if (existing) ss.deleteSheet(existing);

  const snapshot = source.copyTo(ss).setName(targetName);
  SpreadsheetApp.flush();
  const range = snapshot.getDataRange();
  range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
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
