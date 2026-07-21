const FS01A_CFG = Object.freeze({
  INPUT: '01. Đầu vào',
  TECH: '01A. Kỹ thuật',
  TECH_LEGACY: '01. Kỹ thuật',
  MENU: 'FS - CẬP NHẬT MÔ HÌNH',
  SCENARIO_NN: 'NN',
  SCENARIO_TT: 'TT'
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(FS01A_CFG.MENU)
    .addItem('1. Tạo lại sheet 01A. Kỹ thuật', 'FS_taoKyThuatTuDauVao')
    .addSeparator()
    .addItem('2. Lập sheet 02 - Doanh thu', 'FS_lapSheet02')
    .addItem('3. Lập sheet 03 - Chi phí & Vốn', 'FS_lapSheet03')
    .addItem('4. Lập sheet 03A - Lợi nhuận & Thuế', 'FS_lapSheet03A')
    .addItem('5. Lập sheet 04 - Dòng tiền & Tài trợ', 'FS_lapSheet04')
    .addItem('6. Lập sheet 04A - Tổng hợp dòng tiền', 'FS_lapSheet04A')
    .addItem('7. Lập sheet 00 - Tổng hợp', 'FS_lapSheet00')
    .addItem('8. Lập sheet 99 - Kiểm tra', 'FS_lapSheet99')
    .addSeparator()
    .addItem('9. Chạy toàn bộ mô hình', 'FS_chayToanBoMoHinh')
    .addToUi();
}

function FS_chayToanBoMoHinh() {
  FS_taoKyThuatTuDauVao(FS01A_CFG.SCENARIO_NN);
  ['FS_lapSheet02', 'FS_lapSheet03', 'FS_lapSheet03A', 'FS_lapSheet04', 'FS_lapSheet04A', 'FS_lapSheet00', 'FS_lapSheet99']
    .forEach(FS01A_runIfExists_);
}

function FS01A_runIfExists_(functionName) {
  const fn = globalThis[functionName];
  if (typeof fn === 'function') fn();
}

function FS_taoKyThuatDinhMuc() {
  return FS_taoKyThuatTuDauVao(FS01A_CFG.SCENARIO_NN);
}

function FS_taoKyThuatThucTe() {
  return FS_taoKyThuatTuDauVao(FS01A_CFG.SCENARIO_TT);
}

function FS_taoKyThuatTuDauVao(kichBan) {
  const scenario = FS01A_scenario_(kichBan);
  const ss = SpreadsheetApp.getActive();
  const input = ss.getSheetByName(FS01A_CFG.INPUT);
  let tech = ss.getSheetByName(FS01A_CFG.TECH);
  const legacy = ss.getSheetByName(FS01A_CFG.TECH_LEGACY);

  if (!input) throw new Error('Không tìm thấy sheet "01. Đầu vào".');
  if (!tech && legacy) {
    legacy.setName(FS01A_CFG.TECH);
    tech = legacy;
  }
  if (!tech) tech = ss.insertSheet(FS01A_CFG.TECH);

  tech.clear();
  tech.clearFormats();

  let row = 1;
  row = FS01A_writeInfo_(input, tech, row, scenario) + 2;
  row = FS01A_writeCostsByScenario_(input, tech, row, scenario) + 2;

  const productIndex = FS01A_writeProducts_(input, tech, row, scenario);
  row = productIndex.nextRow + 2;

  row = FS01A_writePlans_(input, tech, row, productIndex.byCode, productIndex.codeByName) + 2;

  FS01A_writeTable_(input, tech, row, 'TIEN_DO_CHI_PHI', 'F. TIẾN ĐỘ CHI PHÍ', 'Khoản mục',
    ['Khoản mục', 'Tháng bắt đầu', 'Thời gian', 'Tỷ lệ', 'Loại']);

  FS01A_format_(tech);
  tech.getRange('A1').setNote('Kịch bản chi phí hiện hành: ' + scenario);
  SpreadsheetApp.flush();
  return { scenario, sheet: tech.getName() };
}

function FS01A_writeInfo_(input, tech, startRow, scenario) {
  const fields = [
    ['Tên dự án', 'A. THÔNG TIN CHUNG', 'text'],
    ['Ngày bắt đầu dự án', 'A. THÔNG TIN CHUNG', 'date'],
    ['Số tháng mô hình', 'A. THÔNG TIN CHUNG', 'integer'],
    ['Đơn vị tiền', 'A. THÔNG TIN CHUNG', 'text'],
    ['Tỷ suất chiết khấu', 'A. THÔNG TIN CHUNG', 'percent'],
    ['Tỷ lệ tăng giá bán/năm', 'A. THÔNG TIN CHUNG', 'percent'],
    ['Tỷ lệ tăng giá thuê/năm', 'A. THÔNG TIN CHUNG', 'percent'],
    ['Tỷ lệ trượt chi phí/năm', 'A. THÔNG TIN CHUNG', 'percent'],
    ['Diện tích đất', 'B. QUY HOẠCH', 'number'],
    ['Bắt đầu xây dựng', 'B. QUY HOẠCH', 'integer'],
    ['Thời gian xây dựng', 'B. QUY HOẠCH', 'integer'],
    ['Tỷ lệ vốn vay', 'B. QUY HOẠCH', 'percent'],
    ['Lãi suất vay năm', 'B. QUY HOẠCH', 'percent'],
    ['Tháng bắt đầu trả gốc', 'B. QUY HOẠCH', 'integer'],
    ['Thời gian trả gốc', 'B. QUY HOẠCH', 'integer']
  ];

  const out = [['THONG_TIN_CHUNG', 'GIÁ TRỊ', 'Ô NGUỒN', 'KIỂU DỮ LIỆU']];
  out.push(['Kịch bản chi phí', scenario, '', 'text']);
  fields.forEach(([label, section, type]) => {
    const cell = FS01A_findValueCell_(input, section, label);
    out.push([label, cell ? cell.getValue() : '', cell ? cell.getA1Notation() : '', type]);
  });

  tech.getRange(startRow, 1, out.length, 4).setValues(out);
  return startRow + out.length;
}

function FS01A_writeCostsByScenario_(input, tech, startRow, scenario) {
  const sectionRow = FS01A_findRow_(input, 'C. CHI PHÍ CHUNG');
  if (!sectionRow) throw new Error('Không tìm thấy mục C. CHI PHÍ CHUNG tại 01. Đầu vào.');

  const headerRow = FS01A_findHeaderRow_(input, 'Khoản mục', sectionRow);
  if (!headerRow) throw new Error('Không tìm thấy hàng tiêu đề Khoản mục tại mục C.');

  const lastCol = FS01A_lastCol_(input, headerRow);
  const headers = input.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const headerKeys = headers.map(FS01A_key_);
  const itemCol = headerKeys.indexOf(FS01A_key_('Khoản mục'));
  const codeCol = headerKeys.indexOf(FS01A_key_('Mã khoản mục'));
  const beforeCols = FS01A_allIndexes_(headerKeys, FS01A_key_('Trước VAT'));
  const vatCols = FS01A_allIndexes_(headerKeys, FS01A_key_('VAT đầu vào'));
  const afterCols = FS01A_allIndexes_(headerKeys, FS01A_key_('Sau VAT'));
  const noteCols = FS01A_allIndexes_(headerKeys, FS01A_key_('Ghi chú'));
  const rateCols = FS01A_allIndexes_(headerKeys, FS01A_key_('Tỷ lệ'));
  const scenarioIndex = scenario === FS01A_CFG.SCENARIO_TT ? 1 : 0;

  if (itemCol < 0 || beforeCols.length < 2 || vatCols.length < 2 || afterCols.length < 2 || noteCols.length < 2 || rateCols.length < 2) {
    throw new Error('Bảng C phải có đủ hai nhóm cột Theo Định mức và Thực tế.');
  }

  const out = [[
    'Khoản mục', 'Trước VAT', 'VAT đầu vào', 'Sau VAT', 'Ghi chú', 'Tỷ lệ',
    'Mã khoản mục', 'Kịch bản'
  ]];

  let blankCount = 0;
  for (let r = headerRow + 1; r <= input.getLastRow(); r++) {
    const display = input.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    if (display.some(v => /^[A-Z]\./.test(String(v).trim()))) break;
    if (!display.some(v => String(v).trim())) {
      if (++blankCount >= 3) break;
      continue;
    }
    blankCount = 0;

    const values = input.getRange(r, 1, 1, lastCol).getValues()[0];
    const item = values[itemCol];
    if (!String(item || '').trim()) continue;

    out.push([
      item,
      values[beforeCols[scenarioIndex]],
      values[vatCols[scenarioIndex]],
      values[afterCols[scenarioIndex]],
      values[noteCols[scenarioIndex]],
      values[rateCols[scenarioIndex]],
      codeCol >= 0 ? values[codeCol] : '',
      scenario
    ]);
  }

  tech.getRange(startRow, 1).setValue('CHI_PHI_CHUNG');
  tech.getRange(startRow + 1, 1, out.length, out[0].length).setValues(out);
  return startRow + 1 + out.length;
}

function FS01A_writeTable_(input, tech, startRow, title, section, header, headersOut) {
  const table = FS01A_readTable_(input, section, header);
  tech.getRange(startRow, 1).setValue(title);

  if (!table) {
    tech.getRange(startRow + 1, 1).setValue('Không tìm thấy dữ liệu');
    return startRow + 2;
  }

  const out = [headersOut];
  table.rows.forEach(row => out.push(headersOut.map(h => FS01A_get_(row, table.headers, [h]))));
  tech.getRange(startRow + 1, 1, out.length, headersOut.length).setValues(out);
  return startRow + 1 + out.length;
}

function FS01A_writeProducts_(input, tech, startRow, scenario) {
  const table = FS01A_readTable_(input, 'D. CHI TIẾT SẢN PHẨM', 'Loại sản phẩm');
  tech.getRange(startRow, 1).setValue('SAN_PHAM');
  if (!table) throw new Error('Không tìm thấy bảng D. CHI TIẾT SẢN PHẨM.');

  const headers = [
    'Mã SP', 'Loại sản phẩm', 'Nhóm', 'DTKD', 'Giá bán trước thuế/m²',
    'Giá thuê/m²/tháng', 'CPXD/m²', 'VAT đầu ra', 'Thuế TNDN', 'Lấp đầy',
    'CPVH/Doanh thu', 'CPBT/Doanh thu', 'Thời gian thuê (năm)', 'Diện tích đất (m²)'
  ];

  const out = [headers];
  const byCode = {};
  const codeByName = {};
  const cpxdCandidates = scenario === FS01A_CFG.SCENARIO_TT
    ? ['CPXD/m2_TT', 'CPXD/m²_TT', 'CPXD TT/m2', 'CPXD TT/m²']
    : ['CPXD/m2_ĐM', 'CPXD/m²_ĐM', 'CPXD/m2_DM', 'CPXD/m²_DM', 'CPXD NN/m2', 'CPXD NN/m²'];

  table.rows.forEach(row => {
    const code = String(FS01A_get_(row, table.headers, ['Mã SP', 'Mã sản phẩm']) || '').trim().toUpperCase();
    const name = String(FS01A_get_(row, table.headers, ['Loại sản phẩm', 'Loại SP', 'Sản phẩm']) || '').trim();
    if (!code && !name) return;
    if (!code) throw new Error('D. CHI TIẾT SẢN PHẨM còn thiếu Mã SP tại sản phẩm: ' + name);
    if (byCode[code]) throw new Error('Mã SP bị trùng: ' + code);

    const product = {
      code,
      name,
      group: FS01A_get_(row, table.headers, ['Nhóm', 'Nhóm sản phẩm'])
    };

    byCode[code] = product;
    const nameKey = FS01A_key_(name);
    if (nameKey) codeByName[nameKey] = code;

    out.push([
      code,
      name,
      product.group,
      FS01A_get_(row, table.headers, ['DTKD', 'Diện tích kinh doanh']),
      FS01A_get_(row, table.headers, ['Giá bán trước thuế/m2', 'Giá bán trước thuế/m²', 'Giá bán/m2', 'Giá bán']),
      FS01A_get_(row, table.headers, ['Giá thuê/m2/tháng', 'Giá thuê/m²/tháng', 'Giá thuê/m2/th', 'Giá thuê']),
      FS01A_get_(row, table.headers, cpxdCandidates.concat(['CPXD/m2', 'CPXD/m²', 'Chi phí XD/m2', 'Suất CPXD', 'CPXD'])),
      FS01A_get_(row, table.headers, ['VAT đầu ra', 'VAT']),
      FS01A_get_(row, table.headers, ['Thuế TNDN', 'TNDN']),
      FS01A_get_(row, table.headers, ['Lấp đầy', 'Lấp đầy thuê', 'Tỷ lệ lấp đầy']),
      FS01A_get_(row, table.headers, ['CPVH/doanh thu', 'CPVH/Doanh thu', 'Chi phí vận hành/doanh thu', 'CPVH']),
      FS01A_get_(row, table.headers, ['Chi phí bảo trì/doanh thu', 'CPBT/doanh thu', 'CPBT/Doanh thu', 'CPBT']),
      FS01A_get_(row, table.headers, ['Thời gian thuê (năm)', 'Thời gian thuê', 'Số năm thuê']),
      FS01A_get_(row, table.headers, ['Diện tích đất', 'Diện tích đất (m²)', 'DT đất'])
    ]);
  });

  if (out.length === 1) throw new Error('D. CHI TIẾT SẢN PHẨM không có sản phẩm hoạt động.');

  tech.getRange(startRow + 1, 1, out.length, headers.length).setValues(out);
  return {
    nextRow: startRow + 1 + out.length,
    byCode,
    codeByName
  };
}

function FS01A_writePlans_(input, tech, startRow, productsByCode, codeByName) {
  const table = FS01A_readTable_(input, 'E. KẾ HOẠCH BÁN HÀNG', 'Nhóm');
  tech.getRange(startRow, 1).setValue('KE_HOACH_BAN_THU_TIEN');

  const headers = ['Nhóm', 'Mã SP', 'Số đợt', 'Đợt', 'Tháng bắt đầu', 'Thời gian', 'Tỷ lệ', 'Ghi chú'];
  const out = [headers];

  if (table) {
    table.rows.forEach(row => {
      const sourceCode = String(FS01A_get_(row, table.headers, ['Mã SP', 'Mã sản phẩm']) || '').trim().toUpperCase();
      const sourceName = String(FS01A_get_(row, table.headers, ['Loại sản phẩm', 'Loại SP', 'Sản phẩm']) || '').trim();
      const code = productsByCode[sourceCode]
        ? sourceCode
        : (codeByName[FS01A_key_(sourceName)] || '');

      if (!code) return;

      out.push([
        FS01A_get_(row, table.headers, ['Nhóm']),
        code,
        FS01A_get_(row, table.headers, ['Số đợt']),
        FS01A_get_(row, table.headers, ['Đợt']),
        FS01A_get_(row, table.headers, ['Tháng bắt đầu']),
        FS01A_get_(row, table.headers, ['Thời gian']),
        FS01A_get_(row, table.headers, ['Tỷ lệ']),
        FS01A_get_(row, table.headers, ['Ghi chú'])
      ]);
    });
  }

  tech.getRange(startRow + 1, 1, out.length, headers.length).setValues(out);
  return startRow + 1 + out.length;
}

function FS01A_readTable_(sheet, sectionText, headerText) {
  const sectionRow = FS01A_findRow_(sheet, sectionText);
  if (!sectionRow) return null;
  const headerRow = FS01A_findHeaderRow_(sheet, headerText, sectionRow);
  if (!headerRow) return null;

  const lastCol = FS01A_lastCol_(sheet, headerRow);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const rows = [];
  let blanks = 0;

  for (let r = headerRow + 1; r <= sheet.getLastRow(); r++) {
    const display = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    if (display.some(v => /^[A-Z]\./.test(String(v).trim()))) break;
    if (!display.some(v => String(v).trim())) {
      if (++blanks >= 3) break;
      continue;
    }
    blanks = 0;
    rows.push(sheet.getRange(r, 1, 1, lastCol).getValues()[0]);
  }

  return { headers, rows };
}

function FS01A_findValueCell_(sheet, sectionText, label) {
  const data = sheet.getDataRange().getDisplayValues();
  const section = FS01A_norm_(sectionText);
  const target = FS01A_norm_(label);
  let start = -1;

  for (let r = 0; r < data.length; r++) {
    if (data[r].some(v => FS01A_norm_(v).includes(section))) {
      start = r;
      break;
    }
  }
  if (start < 0) return null;

  for (let r = start + 1; r <= Math.min(start + 15, data.length - 1); r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      const value = FS01A_norm_(data[r][c]);
      if (value === target || value.startsWith(target + '/')) return sheet.getRange(r + 1, c + 2);
    }
  }
  return null;
}

function FS01A_findRow_(sheet, text) {
  const target = FS01A_norm_(text);
  const data = sheet.getDataRange().getDisplayValues();
  for (let r = 0; r < data.length; r++) {
    if (data[r].some(v => FS01A_norm_(v).includes(target))) return r + 1;
  }
  return 0;
}

function FS01A_findHeaderRow_(sheet, headerText, startRow) {
  const target = FS01A_key_(headerText);
  for (let r = startRow; r <= Math.min(startRow + 25, sheet.getLastRow()); r++) {
    const row = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    if (row.some(v => FS01A_key_(v) === target)) return r;
  }
  return 0;
}

function FS01A_lastCol_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  for (let c = values.length - 1; c >= 0; c--) {
    if (String(values[c]).trim()) return c + 1;
  }
  return 1;
}

function FS01A_get_(row, headers, candidates) {
  for (const candidate of candidates) {
    const target = FS01A_key_(candidate);
    const index = headers.findIndex(h => FS01A_key_(h) === target);
    if (index >= 0) return row[index];
  }
  return '';
}

function FS01A_allIndexes_(values, target) {
  const indexes = [];
  values.forEach((value, index) => {
    if (value === target) indexes.push(index);
  });
  return indexes;
}

function FS01A_scenario_(value) {
  const key = String(value || FS01A_CFG.SCENARIO_NN).trim().toUpperCase();
  if (key !== FS01A_CFG.SCENARIO_NN && key !== FS01A_CFG.SCENARIO_TT) {
    throw new Error('Kịch bản chi phí chỉ nhận NN hoặc TT. Giá trị nhận được: ' + value);
  }
  return key;
}

function FS01A_format_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) return;

  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  const productRow = FS01A_findRow_(sheet, 'SAN_PHAM');
  const planRow = FS01A_findRow_(sheet, 'KE_HOACH_BAN_THU_TIEN');
  const costRow = FS01A_findRow_(sheet, 'CHI_PHI_CHUNG');
  const scheduleRow = FS01A_findRow_(sheet, 'TIEN_DO_CHI_PHI');

  [1, costRow, productRow, planRow, scheduleRow]
    .filter(Boolean)
    .forEach(r => sheet.getRange(r, 1, 1, Math.min(lastCol, 14)).setFontWeight('bold'));

  if (productRow) {
    sheet.getRange(productRow + 1, 1, 1, 14).setFontWeight('bold').setWrap(true);
    const rows = planRow ? planRow - productRow - 4 : 0;
    if (rows > 0) {
      sheet.getRange(productRow + 2, 4, rows, 4).setNumberFormat('#,##0');
      sheet.getRange(productRow + 2, 8, rows, 5).setNumberFormat('0.00%');
      sheet.getRange(productRow + 2, 13, rows, 1).setNumberFormat('0');
      sheet.getRange(productRow + 2, 14, rows, 1).setNumberFormat('#,##0');
    }
  }

  if (planRow) {
    sheet.getRange(planRow + 1, 1, 1, 8).setFontWeight('bold').setWrap(true);
  }

  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 95);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 145);
  sheet.setColumnWidth(6, 135);
  sheet.setColumnWidth(7, 105);
  sheet.setColumnWidths(8, 3, 95);
  sheet.setColumnWidth(11, 125);
  sheet.setColumnWidth(12, 125);
  sheet.setColumnWidth(13, 125);
  sheet.setColumnWidth(14, 120);
}

function FS01A_norm_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function FS01A_key_(value) {
  return FS01A_norm_(value)
    .replace(/²/g, '2')
    .replace(/\^2/g, '2')
    .replace(/m\s*2/g, 'm2')
    .replace(/[^a-z0-9]/g, '');
}
