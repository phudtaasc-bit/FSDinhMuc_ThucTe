const FSNT_CFG = Object.freeze({
  SUMMARY_NN: '00. Tổng hợp',
  SUMMARY_TT: '00.Tổng hợp_thực tế',
  CHECKS: '99. Kiểm tra',
  LEGACY_TRIGGER_HANDLER: 'FSNT_tiepTucHaiKichBan_'
});

/**
 * Hàm cũ được giữ để không làm lỗi các trigger đã tồn tại.
 * Khi chạy, hàm chỉ xóa trigger cũ và kết thúc.
 */
function FSNT_tiepTucHaiKichBan_() {
  FSNT_deleteLegacyTriggers_();
}

/**
 * Hàm cũ không còn chạy nối tiếp hai kịch bản.
 * Người dùng chạy riêng từng menu NN hoặc TT để tránh vượt thời gian thực thi.
 */
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

/** Chạy toàn bộ mô hình theo nguồn Định mức và giữ kết quả tại 00. Tổng hợp. */
function FS_chayKichBanDinhMuc() {
  FSNT_runScenario_('NN');
  FSNT_assertNoFail_();
  FSNT_markSummary_(FSNT_CFG.SUMMARY_NN, 'NN');
  SpreadsheetApp.flush();
}

/** Chạy toàn bộ mô hình theo nguồn Thực tế và chốt kết quả tại 00.Tổng hợp_thực tế. */
function FS_chayKichBanThucTe() {
  FSNT_deleteLegacyTriggers_();
  FSNT_runScenario_('TT');
  FSNT_assertNoFail_();
  FSNT_snapshotSummary_(FSNT_CFG.SUMMARY_TT, 'TT');
  SpreadsheetApp.flush();
}

function FSNT_runScenario_(scenario) {
  FS_taoKyThuatTuDauVao(scenario);
  SpreadsheetApp.flush();

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

function FSNT_snapshotSummary_(targetName, scenario) {
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName(FSNT_CFG.SUMMARY_NN);
  if (!source) throw new Error('Không tìm thấy sheet nguồn "' + FSNT_CFG.SUMMARY_NN + '".');

  const existing = ss.getSheetByName(targetName);
  if (existing) ss.deleteSheet(existing);

  const snapshot = source.copyTo(ss).setName(targetName);
  const range = snapshot.getDataRange();
  range.setValues(range.getValues());
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
