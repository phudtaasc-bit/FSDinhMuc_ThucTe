const FSNT_CFG = Object.freeze({
  SUMMARY_NN: '00. Tổng hợp',
  SUMMARY_TT: '00.Tổng hợp_thực tế',
  CHECKS: '99. Kiểm tra'
});

/**
 * Chạy toàn bộ mô hình theo hai nguồn chi phí trên cùng một bộ sheet trung gian.
 * - Chạy TT trước, chốt kết quả sang 00.Tổng hợp_thực tế.
 * - Chạy NN sau cùng, giữ trạng thái mô hình hiện hành tại 00. Tổng hợp.
 *
 * Không giữ DocumentLock ở hàm bao ngoài vì một số hàm lõi tự quản lý khóa.
 * Giữ khóa bao ngoài sẽ gây deadlock/timeout khi chạy chuỗi hai kịch bản.
 */
function FS_chayHaiKichBanNN_TT() {
  const ss = SpreadsheetApp.getActive();

  ss.toast('Đang chạy kịch bản Thực tế...', 'FS NN/TT', 5);
  FSNT_runScenario_('TT');
  FSNT_assertNoFail_();
  FSNT_snapshotSummary_(FSNT_CFG.SUMMARY_TT, 'TT');

  ss.toast('Đang chạy kịch bản Định mức...', 'FS NN/TT', 5);
  FSNT_runScenario_('NN');
  FSNT_assertNoFail_();
  FSNT_markSummary_(FSNT_CFG.SUMMARY_NN, 'NN');

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Đã hoàn thành hai kịch bản:\n' +
    '- 00. Tổng hợp: nguồn Định mức (NN)\n' +
    '- 00.Tổng hợp_thực tế: nguồn Thực tế (TT)\n' +
    '- Sheet 99 không có FAIL ở lần chạy cuối.'
  );
}

function FS_chayKichBanDinhMuc() {
  FSNT_runScenario_('NN');
  FSNT_assertNoFail_();
  FSNT_markSummary_(FSNT_CFG.SUMMARY_NN, 'NN');
}

function FS_chayKichBanThucTe() {
  FSNT_runScenario_('TT');
  FSNT_assertNoFail_();
  FSNT_snapshotSummary_(FSNT_CFG.SUMMARY_TT, 'TT');
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
