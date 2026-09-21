const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

global.window = {};
global.XLSX = require('../xlsx.full.min.js');
global.JSZip = require('../jszip.min.js');
vm.runInThisContext(fs.readFileSync(require.resolve('../vtb_converter.js'), 'utf8'));

const { buildVTBMappingRows, splitOldVTBItems, applyVTBBulkTechnology } = window.VTBConverter;

function baseItem(overrides = {}) {
  return {
    row: 12,
    vendor: 'Huawei',
    name: 'Thiết bị thử nghiệm',
    dvt: 'Bộ',
    dg: 100,
    kl_uctt_27: 0,
    kl_uctt_28: 0,
    kl_bbk_27: 0,
    kl_bbk_28: 0,
    kl_vp_27: 0,
    kl_vp_28: 0,
    kl_gn_27: 0,
    kl_gn_28: 0,
    kl_tt_27: 0,
    kl_tt_28: 0,
    kl_cg_27: 0,
    kl_cg_28: 0,
    kl_hd_27: 0,
    kl_hd_28: 0,
    ...overrides
  };
}

async function test(name, run) {
  try {
    await run();
    console.log(`PASS ${name}`);
  } catch (error) {
    console.error(`FAIL ${name}`);
    throw error;
  }
}

(async () => {
  await test('gộp Đầu tư bắt buộc khác vào UCTT và dùng Mã DV ƯCTT', () => {
    const item = baseItem({
      kl_uctt_27: 2,
      kl_bbk_27: 3,
      kl_bbk_28: 4
    });

    const result = splitOldVTBItems([item], {});

    assert.deepStrictEqual(result.rows.map(row => ({
      branch: row.branch,
      maDV: row.maDV,
      kl27: row.kl27,
      kl28: row.kl28
    })), [{ branch: 'quality', maDV: 'ƯCTT', kl27: 5, kl28: 4 }]);
  });

  await test('tách Vùng phủ và Tăng trưởng lưu lượng thành hai dòng 3G', () => {
    const item = baseItem({
      kl_vp_27: 7,
      kl_vp_28: 8,
      kl_tt_27: 9,
      kl_tt_28: 10
    });

    const result = splitOldVTBItems([item], { '12': '3G' });

    assert.deepStrictEqual(result.rows.map(row => ({
      branch: row.branch,
      maDV: row.maDV,
      purpose: row.purpose,
      kl27: row.kl27,
      kl28: row.kl28
    })), [
      { branch: 'coverage_3g', maDV: 'VPDĐ', purpose: 'Vùng phủ', kl27: 7, kl28: 8 },
      { branch: 'capacity_3g', maDV: 'DLDĐ', purpose: 'Tăng trưởng lưu lượng', kl27: 9, kl28: 10 }
    ]);
  });

  await test('vẫn tách hai dòng khi Vùng phủ và Tăng trưởng cùng được gán 5G', () => {
    const item = baseItem({ kl_vp_27: 1, kl_tt_27: 2 });

    const result = splitOldVTBItems([item], { '12': '5G' });

    assert.deepStrictEqual(result.rows.map(row => ({ branch: row.branch, maDV: row.maDV, kl27: row.kl27 })), [
      { branch: 'network_5g', maDV: '5G', kl27: 1 },
      { branch: 'network_5g', maDV: '5G', kl27: 2 }
    ]);
  });

  await test('ánh xạ Hiện đại hóa sang nhánh Hiện đại hóa mạng lưới và Mã DV VHKT', () => {
    const item = baseItem({ kl_hd_27: 6, kl_hd_28: 7 });

    const result = splitOldVTBItems([item], {});

    assert.deepStrictEqual(result.rows.map(row => ({
      branch: row.branch,
      maDV: row.maDV,
      kl27: row.kl27,
      kl28: row.kl28
    })), [{ branch: 'modernization', maDV: 'VHKT', kl27: 6, kl28: 7 }]);
  });

  await test('yêu cầu gán công nghệ cho mọi vật tư có Vùng phủ hoặc Tăng trưởng', () => {
    const items = [
      baseItem({ row: 12, name: 'Thiết bị chưa rõ', kl_vp_27: 1 }),
      baseItem({ row: 13, name: 'Thiết bị 4G', kl_tt_27: 2 })
    ];

    const mappingRows = buildVTBMappingRows(items);
    const result = splitOldVTBItems(items, { '13': '4G' });

    assert.deepStrictEqual(mappingRows.map(row => ({ id: row.id, suggestion: row.suggestion })), [
      { id: '12', suggestion: '' },
      { id: '13', suggestion: '4G' }
    ]);
    assert.deepStrictEqual(result.missingAssignments, ['12']);
  });

  await test('phát hiện dữ liệu ngoài phạm vi ở Giải nghẽn và Củng cố', () => {
    const item = baseItem({ kl_gn_27: 3, kl_cg_28: 4 });

    const result = splitOldVTBItems([item], {});

    assert.deepStrictEqual(result.unmappedAmounts, [
      { row: 12, column: 'Giải nghẽn, nâng cấp', year: 2027, quantity: 3 },
      { row: 12, column: 'Củng cố bền vững', year: 2028, quantity: 4 }
    ]);
  });

  await test('tự động co gọn dòng, bảo toàn 100% định dạng và giữ đúng mã đầu ra', async () => {
    const rows = Array.from({ length: 16 }, () => Array(37).fill(''));
    for (let index = 0; index < 7; index += 1) {
      const excelRow = 9 + index;
      rows[excelRow - 1][0] = `VT-${index + 1}`;
      rows[excelRow - 1][1] = `Thiết bị dung lượng ${index + 1}`;
      rows[excelRow - 1][2] = 'Bộ';
      rows[excelRow - 1][7] = 1;
      rows[excelRow - 1][10] = 1;
      rows[excelRow - 1][19] = 100;
      rows[excelRow - 1][27] = 100;
    }
    const oldWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(oldWorkbook, XLSX.utils.aoa_to_sheet(rows), 'Sheet1');
    const assignments = Object.fromEntries(Array.from({ length: 7 }, (_, index) => [String(9 + index), '4G']));

    const templateBuffer = fs.readFileSync(require.resolve('../Masterlist 2027-2028_Mau_moi.xlsx'));
    const result = await window.VTBConverter.convertOldVTBToNewFormat(oldWorkbook, assignments, templateBuffer);
    const outputSheet = result.convertedWorkbook.Sheets[result.convertedWorkbook.SheetNames[0]];
    const outputRange = XLSX.utils.decode_range(outputSheet['!ref']);
    const capacityRows = [];
    let cap4GHeaderRow = 0;
    let threeGHeaderRow = 0;
    for (let row = 0; row <= outputRange.e.r; row += 1) {
      const name = outputSheet[XLSX.utils.encode_cell({ r: row, c: 1 })]?.v;
      const maDV = outputSheet[XLSX.utils.encode_cell({ r: row, c: 10 })]?.v;
      if (name === 'Thiết bị 4G' && row >= 15) cap4GHeaderRow = row + 1;
      if (name === 'Thiết bị 3G' && row >= 15) threeGHeaderRow = row + 1;
      if (maDV === 'DLDĐ') {
        capacityRows.push({
          maMang: outputSheet[XLSX.utils.encode_cell({ r: row, c: 9 })]?.v,
          maLoai: outputSheet[XLSX.utils.encode_cell({ r: row, c: 11 })]?.v
        });
      }
    }

    assert.strictEqual(result.mappedCount, 7);
    assert.strictEqual(capacityRows.length, 7);
    assert.ok(cap4GHeaderRow > 0, 'Phải tìm thấy tiêu đề Thiết bị 4G trong Dung lượng');
    assert.strictEqual(threeGHeaderRow, cap4GHeaderRow + 8, 'Tiêu đề Thiết bị 3G phải nằm ngay sau 7 vật tư 4G (không có dòng trống thừa)');
    assert.deepStrictEqual(capacityRows, Array.from({ length: 7 }, () => ({ maMang: 'VT', maLoai: 'VTTB' })));
    assert.strictEqual(outputSheet['!autofilter'].ref, `A7:Q${outputRange.e.r + 1}`);

    // Kiểm tra styles.xml được giữ nguyên vẹn từ phôi gốc (> 50KB)
    const zip = await global.JSZip.loadAsync(result.convertedBuffer);
    const stylesXml = await zip.file('xl/styles.xml').async('string');
    assert.ok(stylesXml.length > 50000, 'styles.xml phải được giữ nguyên từ phôi gốc (> 50KB)');
  });

  await test('gán một công nghệ cho nhiều vật tư đã quét chọn', () => {
    const assignments = { '12': '2G', '13': '5G', '14': '3G' };

    const updated = applyVTBBulkTechnology(assignments, ['12', '14'], '4G');

    assert.deepStrictEqual(updated, { '12': '4G', '13': '5G', '14': '4G' });
    assert.deepStrictEqual(assignments, { '12': '2G', '13': '5G', '14': '3G' });
  });
})();
