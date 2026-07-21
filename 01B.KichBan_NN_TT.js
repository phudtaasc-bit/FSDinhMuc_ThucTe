const FSNT_CFG = Object.freeze({
  SUMMARY_NN: '00. Tổng hợp',
  SUMMARY_TT: '00.Tổng hợp_thực tế',
  CHECKS: '99. Kiểm tra',
  STATE_KEY: 'FSNT_TWO_SCENARIO_STATE',
  TRIGGER_HANDLER: 'FSNT_tiepTucHaiKichBan_',
  NEXT_RUN_DELAY_MS: 15000
});

/**
 * Khởi động quy trình hai kịch bản theo cơ chế chia phiên thực thi.
 * Mỗi kịch bản chạy trong một execution riêng để không vượt giới hạn Apps Script.
 */
function FS_chayHaiKichBanNN_TT() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Đang có một tiến trình NN/TT khác hoạt động. Vui lòng chờ.');
  }

  try {
    FSNT_deleteContinuationTriggers_();
    FSNT_saveState_({
      step: 'TT',
      status: 'PENDING',
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      message: 'Đã xếp lịch chạy kịch bản Thực tế.'
    });
    FSNT_scheduleNext_();

    SpreadsheetApp.getActive().toast(
      'Đã khởi động. Hệ thống chạy TT trước, sau đó tự động chạy NN.',
      'FS NN/TT',
      8
    );
  } finally {
    lock.releaseLock();
  }
}

/** Hàm trigger tiếp tục quy trình. Không gọi trực tiếp từ menu. */
function FSNT_tiepTucHaiKichBan_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    const state = FSNT_loadState_();
    if (!state || !state.step) {
      FSNT_deleteContinuationTriggers_();
      return;
    }

    if (state.step === 'TT') {
      FSNT_saveState_(Object.assign({}, state, {
        status: 'RUNNING',
        updatedAt: new Date().toISOString(),
        message: 'Đang chạy kịch bản Thực tế.'
      }));

      FSNT_runScenario_('TT');
      FSNT_assertNoFail_();
      FSNT_snapshotSummary_(FSNT_CFG.SUMMARY_TT, 'TT');

      FSNT_saveState_({
        step: 'NN',
        status: 'PENDING',
        startedAt: state.startedAt,
        updatedAt: new Date().toISOString(),
        message: 'Kịch bản TT hoàn thành. Đã xếp lịch chạy kịch bản NN.'
      });
      FSNT_scheduleNext_();
      return;
    }

    if (state.step === 'NN') {
      FSNT_saveState_(Object.assign({}, state, {
        status: 'RUNNING',
        updatedAt: new Date().toISOString(),
        message: 'Đang chạy kịch bản Định mức.'
      }));

      FSNT_runScenario_('NN');
      FSNT_assertNoFail_();
      FSNT_markSummary_(FSNT_CFG.SUMMARY_NN, 'NN');

      FSNT_saveState_({
        step: 'DONE',
        status: 'DONE',
        startedAt: state.startedAt,
        updatedAt: new Date().toISOString(),
        message: 'Đã hoàn thành TT và NN. Sheet 99 không có FAIL ở lần chạy cuối.'
      });
      FSNT_deleteContinuationTriggers_();

      SpreadsheetApp.getActive().toast(
        'Đã hoàn thành: 00. Tổng hợp = NN; 00.Tổng hợp_thực tế = TT.',
        'FS NN/TT',
        10
      );
      return;
    }

    FSNT_deleteContinuationTriggers_();
  } catch (error) {
    const previous = FSNT_loadState_() || {};
    FSNT_saveState_({
      step: previous.step || 'UNKNOWN',
      status: 'ERROR',
      startedAt: previous.startedAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      message: error && error.message ? error.message : String(error)
    });
    FSNT_deleteContinuationTriggers_();
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function FS_xemTrangThaiHaiKichBan() {
  const state = FSNT_loadState_();
  const message = state
    ? [
        'Trạng thái: ' + (state.status || ''),
        'Bước: ' + (state.step || ''),
        'Cập nhật: ' + (state.updatedAt || ''),
        'Nội dung: ' + (state.message || '')
      ].join('\n')
    : 'Chưa có tiến trình NN/TT.';
  SpreadsheetApp.getUi().alert(message);
}

function FS_huyChayHaiKichBan() {
  FSNT_deleteContinuationTriggers_();
  FSNT_saveState_({
    step: 'CANCELLED',
    status: 'CANCELLED',
    startedAt: '',
    updatedAt: new Date().toISOString(),
    message: 'Tiến trình đã được hủy thủ công.'
  });
  SpreadsheetApp.getUi().alert('Đã hủy tiến trình NN/TT đang chờ.');
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

function FSNT_scheduleNext_() {
  FSNT_deleteContinuationTriggers_();
  ScriptApp.newTrigger(FSNT_CFG.TRIGGER_HANDLER)
    .timeBased()
    .after(FSNT_CFG.NEXT_RUN_DELAY_MS)
    .create();
}

function FSNT_deleteContinuationTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === FSNT_CFG.TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function FSNT_saveState_(state) {
  PropertiesService.getScriptProperties().setProperty(
    FSNT_CFG.STATE_KEY,
    JSON.stringify(state || {})
  );
}

function FSNT_loadState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(FSNT_CFG.STATE_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
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
