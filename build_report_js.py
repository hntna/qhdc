# -*- coding: utf-8 -*-
"""
Script tạo file report.js hoàn chỉnh với:
1. Bảng Tổng hợp theo Dịch vụ & Mảng (chuẩn Sheet "TH theo DV" - 28 hàng x 33 cột).
2. Bảng Tổng hợp theo Mảng (Toàn mạng).
3. 11 bảng trong sheet "TH theo Mảng".
4. 14 bảng trong sheet "TH theo Dịch vụ".
5. Bảng Đối chiếu Chiến lược 5 năm.
6. Chế độ xem kỳ (Tất cả, 2027, 2028, 2027-2028), tìm kiếm, xuất Excel, sao chép.
"""
import json

with open('scratch_template_schema.json', 'r', encoding='utf-8') as f:
    schema = json.load(f)

schema_json = json.dumps(schema, ensure_ascii=False)

part1 = """/**
 * report.js - Báo cáo Tổng Hợp & So Sánh Chiến Lược QHĐC 2027-2028
 * Tự động tính toán số liệu thời gian thực từ dữ liệu các mảng theo công thức chuẩn của file mẫu.
 */

// Schema cấu trúc bảng & công thức trích xuất chuẩn từ Masterlist Mau
const REPORT_TEMPLATE_SCHEMA = """ + schema_json + """;

// Cột ánh xạ trong sheet "TH theo DV"
// 2027: D=VT, E=ML, F=CĐBR, G=CNTT, H=TD, I=IP, J=CĐ, K=HT, L=TOTAL
// 2028: N=VT, O=ML, P=CĐBR, Q=CNTT, R=TD, S=IP, T=CĐ, U=HT, V=TOTAL
const MANG_COL_MAP_2027 = {
  'D': 'VT', 'E': 'ML', 'F': 'CDBR', 'G': 'CNTT', 'H': 'TD', 'I': 'IP', 'J': 'CD', 'K': 'HT', 'L': 'TOTAL'
};
const MANG_COL_MAP_2028 = {
  'N': 'VT', 'O': 'ML', 'P': 'CDBR', 'Q': 'CNTT', 'R': 'TD', 'S': 'IP', 'T': 'CD', 'U': 'HT', 'V': 'TOTAL'
};

// 8 Mảng kỹ thuật trong sheet "TH theo DV"
const TH_DV_SECTORS = ['VT', 'ML', 'CDBR', 'CNTT', 'TD', 'IP', 'CD', 'HT'];

// Cấu trúc 28 hàng chuẩn trong sheet "TH theo DV" (từ Row 4 đến Row 31 của Excel)
const TH_DV_ROW_DEFS = [
  // R04: Tổng cộng
  { id: 'R04', type: 'grand_total', stt: '*', nd: 'Tổng', madv: null },
  // R05: Tỷ trọng
  { id: 'R05', type: 'share_header', stt: '', nd: 'Tỷ trọng theo mảng (%)', madv: null },
  // R06: Nhóm I
  { id: 'R06', type: 'group', stt: 'I', nd: 'Công nghệ mới triển khai diện rộng cho kinh doanh', madv: null, children: ['R07', 'R08', 'R09', 'R10'] },
  { id: 'R07', type: 'leaf', stt: '1', nd: 'Mạng 5G', madv: '5G' },
  { id: 'R08', type: 'leaf', stt: '2', nd: 'XGSPON', madv: 'XGSPON' },
  { id: 'R09', type: 'leaf', stt: '3', nd: 'AI/GPU', madv: 'AI/GPU' },
  { id: 'R10', type: 'leaf', stt: '4', nd: 'Private Mobile Network', madv: 'PMN' },
  // R11: Nhóm II
  { id: 'R11', type: 'group', stt: 'II', nd: 'Mở rộng mạng lưới hiện tại cho kinh doanh', madv: null, children: ['R12', 'R15', 'R20', 'R21', 'R22'] },
  { id: 'R12', type: 'subgroup', stt: '1', nd: 'Mạng 2/3/4G', madv: null, children: ['R13', 'R14'] },
  { id: 'R13', type: 'leaf', stt: '-', nd: '2/3/4G vùng phủ', madv: 'VPDD' },
  { id: 'R14', type: 'leaf', stt: '-', nd: '2/3/4G dung lượng', madv: 'DLDD' },
  { id: 'R15', type: 'subgroup', stt: '2', nd: 'Mạng BRCĐ&TH', madv: null, children: ['R16', 'R17', 'R18', 'R19'] },
  { id: 'R16', type: 'leaf', stt: '-', nd: 'Vùng phủ GPON', madv: 'VPGPON' },
  { id: 'R17', type: 'leaf', stt: '-', nd: 'Dung lượng Internet', madv: 'DLGPON' },
  { id: 'R18', type: 'leaf', stt: '-', nd: 'Dung lượng Truyền hình', madv: 'TH' },
  { id: 'R19', type: 'leaf', stt: '-', nd: 'Kênh truyền', madv: 'KT' },
  { id: 'R20', type: 'leaf', stt: '3', nd: 'Phục vụ kinh doanh Cloud', madv: 'CLOUD_KD' },
  { id: 'R21', type: 'leaf', stt: '4', nd: 'Data Center', madv: 'DC' },
  { id: 'R22', type: 'subgroup', stt: '5', nd: 'Triển khai hạ tầng CNTT', madv: null, children: ['R23', 'R24', 'R25'] },
  { id: 'R23', type: 'leaf', stt: '-', nd: 'Trực tiếp kinh doanh', madv: 'TTKD' },
  { id: 'R24', type: 'leaf', stt: '-', nd: 'Hỗ trợ kinh doanh và quản trị', madv: 'HTKD' },
  { id: 'R25', type: 'leaf', stt: '-', nd: 'Phần mềm', madv: 'PM' },
  // R26: Nhóm III
  { id: 'R26', type: 'group', stt: 'III', nd: 'Đảm bảo VHKT, kiên cố, bền vững trong thiên tai', madv: null, children: ['R27', 'R28', 'R29', 'R30'] },
  { id: 'R27', type: 'leaf', stt: '1', nd: 'Đảm bảo dự phòng ƯCTT', madv: 'UCTT' },
  { id: 'R28', type: 'leaf', stt: '2', nd: 'Đảm bảo VHKT và nâng cao chất lượng mạng', madv: 'VHKT' },
  { id: 'R29', type: 'leaf', stt: '3', nd: 'Củng cố kiên cố thường trình', madv: 'KCTT' },
  { id: 'R30', type: 'leaf', stt: '4', nd: 'Bền vững mạng lưới trong thiên tai', madv: 'PCTT' },
  // R31: Nhóm IV
  { id: 'R31', type: 'leaf', stt: 'IV', nd: 'Ví điện tử', madv: 'VI' }
];

// Trạng thái báo cáo
const reportState = {
  rawResponse: null,
  activeSubtab: 'subtabTongHop',
  unitMultiplier: 1.0, // 1.0 = Triệu USD (M$), 1000 = Nghìn USD (K$), 1000000 = USD
  unitLabel: 'M$',
  searchKeyword: '',
  selectedSectorFilter: 'ALL',
  selectedServiceFilter: 'ALL',
  thdvPeriodView: 'ALL'
};

// Chuẩn hóa tên mảng
function normMangCode(m) {
  if (!m) return '';
  const s = String(m).trim().toUpperCase();
  if (s === 'VT' || s.includes('VÔ TUYẾN') || s.includes('VO TUYEN')) return 'VT';
  if (s === 'ML' || s.includes('MẠNG LÕI') || s.includes('MANG LOI')) return 'ML';
  if (s === 'CDBR' || s === 'CĐBR' || s.includes('CỐ ĐỊNH') || s.includes('CO DINH') || s.includes('BRCĐ')) return 'CDBR';
  if (s === 'CNTT' || s.includes('CÔNG NGHỆ THÔNG TIN') || s.includes('CONG NGHE THONG TIN')) return 'CNTT';
  if (s === 'IP') return 'IP';
  if (s === 'TD' || s.includes('TRUYỀN DẪN') || s.includes('TRUYEN DAN')) return 'TD';
  if (s === 'CD' || s === 'CĐ' || s.includes('CƠ ĐIỆN') || s.includes('CO DIEN')) return 'CD';
  if (s === 'HT' || s.includes('HẠ TẦNG') || s.includes('HA TANG')) return 'HT';
  return s;
}

// Chuẩn hóa mã dịch vụ
function normMaDVCode(d) {
  if (!d) return '';
  return String(d).trim().toUpperCase()
    .replace(/Ư/g, 'U')
    .replace(/Đ/g, 'D');
}

// Bản đồ Row -> Mã DV chuẩn hóa
const DV_ROW_MAP = {};
for (const [rStr, v] of Object.entries(REPORT_TEMPLATE_SCHEMA.dv_map || {})) {
  if (v && v.madv) {
    DV_ROW_MAP[parseInt(rStr, 10)] = normMaDVCode(v.madv);
  }
}

/**
 * TÍNH TOÁN DỮ LIỆU BÁO CÁO THỜI GIAN THỰC TỪ EXTRACTED DATA CỦA APP
 */
function updateReportFromExtractedData(extractedData, extractedByMang) {
  const items = extractedData || [];
  
  // 1. Khởi tạo ma trận (Mảng x Mã DV) và tổng mảng
  const matrix27 = {};
  const matrix28 = {};
  const mangTotals27 = { VT: 0, ML: 0, CDBR: 0, CNTT: 0, TD: 0, IP: 0, CD: 0, HT: 0 };
  const mangTotals28 = { VT: 0, ML: 0, CDBR: 0, CNTT: 0, TD: 0, IP: 0, CD: 0, HT: 0 };
  const serviceTotals27 = {};
  const serviceTotals28 = {};

  let countItemsWithData = 0;

  items.forEach(it => {
    // Bỏ qua dòng Header Cấp 1 của Mảng nếu không có đơn giá
    if (it.isMangHeader) return;
    // Bỏ qua các dòng nhóm thuần túy (không phải dòng vật tư có số liệu)
    if (it.isGroup && (!it.dg || it.dg <= 0) && (!it.kl27 && !it.kl28)) return;

    const val27 = (parseFloat(it.tt27) || 0) / 1000000.0;
    const val28 = (parseFloat(it.tt28) || 0) / 1000000.0;
    if (val27 === 0 && val28 === 0) return;

    countItemsWithData++;

    const m = normMangCode(it.ma_mang || it.mangCode || it.mang);
    const dv = normMaDVCode(it.ma_dv || it.madv);

    if (m) {
      mangTotals27[m] = (mangTotals27[m] || 0) + val27;
      mangTotals28[m] = (mangTotals28[m] || 0) + val28;
    }

    if (dv) {
      serviceTotals27[dv] = (serviceTotals27[dv] || 0) + val27;
      serviceTotals28[dv] = (serviceTotals28[dv] || 0) + val28;
    }

    if (m && dv) {
      const k = m + '_' + dv;
      matrix27[k] = (matrix27[k] || 0) + val27;
      matrix28[k] = (matrix28[k] || 0) + val28;
    }
  });

  // 2. Tính toán các bảng TH theo Mảng
  const tablesMang = computeTablesMang(matrix27, matrix28, mangTotals27, mangTotals28, serviceTotals27, serviceTotals28);
  
  // 3. Tính toán các bảng TH theo Dịch vụ
  const tablesDV = computeTablesDV(matrix27, matrix28, mangTotals27, mangTotals28, serviceTotals27, serviceTotals28);

  // 4. Tính toán bảng Tổng hợp TH theo DV & Mảng (33 cột x 28 hàng chuẩn Excel)
  const thTheoDV = computeTHTheoDV(matrix27, matrix28);

  // 5. Tổng hợp chung toàn mạng
  const summary = computeOverallSummary(tablesMang);

  // 6. Đối chiếu Chiến lược 5 năm
  const stratComp = buildStrategyComparison(tablesMang);

  reportState.rawResponse = {
    success: true,
    file_name: 'Dữ liệu đầu vào ứng dụng (' + countItemsWithData + ' dòng vật tư)',
    total_items_processed: countItemsWithData,
    summary: summary,
    tables_mang: tablesMang,
    tables_dv: tablesDV,
    th_theo_dv: thTheoDV,
    strategy_comparison: stratComp
  };

  renderAllTabs();
}

// Hàm giải công thức tham chiếu 'TH theo DV'!ColRow
function resolveFormula(form, colMap, matrix, serviceTotals) {
  if (!form || typeof form !== 'string') return 0.0;
  const regex = /'?TH theo DV'?!([A-Z]+)(\\d+)/gi;
  let match;
  let sum = 0.0;
  let found = false;

  while ((match = regex.exec(form)) !== null) {
    const col = match[1].toUpperCase();
    const rIdx = parseInt(match[2], 10);
    const dvCode = DV_ROW_MAP[rIdx] || '';

    if (colMap[col]) {
      const mang = colMap[col];
      if (mang === 'TOTAL') {
        sum += (serviceTotals[dvCode] || 0.0);
      } else {
        const k = mang + '_' + dvCode;
        sum += (matrix[k] || 0.0);
      }
      found = true;
    }
  }

  return found ? sum : 0.0;
}

// Tính toán 11 bảng trong sheet "TH theo Mảng"
function computeTablesMang(matrix27, matrix28, mangTotals27, mangTotals28, serviceTotals27, serviceTotals28) {
  const schemaTables = REPORT_TEMPLATE_SCHEMA.tables_mang || [];

  return schemaTables.map(t => {
    const computedRows = t.rows.map(r => {
      let y27 = 0.0;
      let y28 = 0.0;

      if (r.form_e) {
        y27 = resolveFormula(r.form_e, MANG_COL_MAP_2027, matrix27, serviceTotals27);
      }
      if (r.form_f) {
        y28 = resolveFormula(r.form_f, MANG_COL_MAP_2028, matrix28, serviceTotals28);
      }

      return {
        ...r,
        y2027: y27,
        y2028: y28,
        tong: y27 + y28
      };
    });

    // Tính lại dòng Tổng
    const totRow = computedRows.find(r => r.is_total);
    if (totRow) {
      const sub27 = computedRows.filter(r => !r.is_total && !r.is_sub).reduce((acc, r) => acc + r.y2027, 0);
      const sub28 = computedRows.filter(r => !r.is_total && !r.is_sub).reduce((acc, r) => acc + r.y2028, 0);

      if (sub27 > 0 || sub28 > 0) {
        totRow.y2027 = sub27;
        totRow.y2028 = sub28;
        totRow.tong = sub27 + sub28;
      } else {
        // Fallback theo tổng mảng trực tiếp nếu bảng không có công thức con chi tiết
        const titleU = t.title.toUpperCase();
        let fallbackKey = null;
        if (titleU.includes('VÔ TUYẾN')) fallbackKey = 'VT';
        else if (titleU.includes('BRCĐ')) fallbackKey = 'CDBR';
        else if (titleU.includes('MẠNG LÕI')) fallbackKey = 'ML';
        else if (titleU.includes('THÔNG TIN')) fallbackKey = 'CNTT';
        else if (titleU.includes('TRUYỀN DẪN QUANG')) fallbackKey = 'TD';
        else if (titleU.includes('TRUYỀN DẪN IP')) fallbackKey = 'IP';
        else if (titleU.includes('CƠ ĐIỆN TỔNG TRẠM')) fallbackKey = 'CD';
        else if (titleU.includes('HẠ TẦNG')) fallbackKey = 'HT';

        if (fallbackKey) {
          totRow.y2027 = mangTotals27[fallbackKey] || 0.0;
          totRow.y2028 = mangTotals28[fallbackKey] || 0.0;
          totRow.tong = totRow.y2027 + totRow.y2028;
        }
      }
    }

    return {
      title: t.title,
      rows: computedRows
    };
  });
}

// Tính toán 14 bảng trong sheet "TH theo Dịch vụ"
function computeTablesDV(matrix27, matrix28, mangTotals27, mangTotals28, serviceTotals27, serviceTotals28) {
  const schemaTables = REPORT_TEMPLATE_SCHEMA.tables_dv || [];

  return schemaTables.map(t => {
    const computedRows = t.rows.map(r => {
      let y27 = 0.0;
      let y28 = 0.0;

      if (r.form_f) {
        y27 = resolveFormula(r.form_f, MANG_COL_MAP_2027, matrix27, serviceTotals27);
      }
      if (r.form_g) {
        y28 = resolveFormula(r.form_g, MANG_COL_MAP_2028, matrix28, serviceTotals28);
      }

      return {
        ...r,
        y2027: y27,
        y2028: y28,
        tong: y27 + y28
      };
    });

    const totRow = computedRows.find(r => r.is_total);
    if (totRow) {
      const sub27 = computedRows.filter(r => !r.is_total && !r.is_sub).reduce((acc, r) => acc + r.y2027, 0);
      const sub28 = computedRows.filter(r => !r.is_total && !r.is_sub).reduce((acc, r) => acc + r.y2028, 0);
      if (sub27 > 0 || sub28 > 0) {
        totRow.y2027 = sub27;
        totRow.y2028 = sub28;
        totRow.tong = sub27 + sub28;
      }
    }

    return {
      title: t.title,
      rows: computedRows
    };
  });
}

// Tính toán ma trận TH theo DV (28 hàng x 33 cột chuẩn Excel)
function computeTHTheoDV(matrix27, matrix28) {
  const rowDataMap = {};

  // 1. Khởi tạo đối tượng cho từng hàng
  TH_DV_ROW_DEFS.forEach(r => {
    rowDataMap[r.id] = {
      id: r.id,
      type: r.type,
      stt: r.stt,
      nd: r.nd,
      madv: r.madv,
      y2027: { VT: 0, ML: 0, CDBR: 0, CNTT: 0, TD: 0, IP: 0, CD: 0, HT: 0, total: 0, share: 0 },
      y2028: { VT: 0, ML: 0, CDBR: 0, CNTT: 0, TD: 0, IP: 0, CD: 0, HT: 0, total: 0, share: 0 },
      yTotal: { VT: 0, ML: 0, CDBR: 0, CNTT: 0, TD: 0, IP: 0, CD: 0, HT: 0, total: 0, share: 0 }
    };
  });

  // 2. Điền giá trị cho các dòng leaf (có Mã DV)
  TH_DV_ROW_DEFS.filter(r => r.madv).forEach(r => {
    const d = rowDataMap[r.id];
    const code = normMaDVCode(r.madv);
    TH_DV_SECTORS.forEach(s => {
      const k = s + '_' + code;
      const v27 = matrix27[k] || 0.0;
      const v28 = matrix28[k] || 0.0;
      d.y2027[s] = v27;
      d.y2028[s] = v28;
      d.yTotal[s] = v27 + v28;
    });
    d.y2027.total = TH_DV_SECTORS.reduce((sum, s) => sum + d.y2027[s], 0);
    d.y2028.total = TH_DV_SECTORS.reduce((sum, s) => sum + d.y2028[s], 0);
    d.yTotal.total = TH_DV_SECTORS.reduce((sum, s) => sum + d.yTotal[s], 0);
  });

  // 3. Tính toán Subgroups (R12, R15, R22)
  ['R12', 'R15', 'R22'].forEach(subId => {
    const def = TH_DV_ROW_DEFS.find(r => r.id === subId);
    if (!def || !def.children) return;
    const subRow = rowDataMap[subId];
    TH_DV_SECTORS.forEach(s => {
      let sum27 = 0, sum28 = 0, sumTot = 0;
      def.children.forEach(cId => {
        const c = rowDataMap[cId];
        if (c) {
          sum27 += c.y2027[s];
          sum28 += c.y2028[s];
          sumTot += c.yTotal[s];
        }
      });
      subRow.y2027[s] = sum27;
      subRow.y2028[s] = sum28;
      subRow.yTotal[s] = sumTot;
    });
    subRow.y2027.total = TH_DV_SECTORS.reduce((sum, s) => sum + subRow.y2027[s], 0);
    subRow.y2028.total = TH_DV_SECTORS.reduce((sum, s) => sum + subRow.y2028[s], 0);
    subRow.yTotal.total = TH_DV_SECTORS.reduce((sum, s) => sum + subRow.yTotal[s], 0);
  });

  // 4. Tính toán Groups lớn (R06, R11, R26)
  ['R06', 'R11', 'R26'].forEach(grpId => {
    const def = TH_DV_ROW_DEFS.find(r => r.id === grpId);
    if (!def || !def.children) return;
    const grpRow = rowDataMap[grpId];
    TH_DV_SECTORS.forEach(s => {
      let sum27 = 0, sum28 = 0, sumTot = 0;
      def.children.forEach(cId => {
        const c = rowDataMap[cId];
        if (c) {
          sum27 += c.y2027[s];
          sum28 += c.y2028[s];
          sumTot += c.yTotal[s];
        }
      });
      grpRow.y2027[s] = sum27;
      grpRow.y2028[s] = sum28;
      grpRow.yTotal[s] = sumTot;
    });
    grpRow.y2027.total = TH_DV_SECTORS.reduce((sum, s) => sum + grpRow.y2027[s], 0);
    grpRow.y2028.total = TH_DV_SECTORS.reduce((sum, s) => sum + grpRow.y2028[s], 0);
    grpRow.yTotal.total = TH_DV_SECTORS.reduce((sum, s) => sum + grpRow.yTotal[s], 0);
  });

  // 5. Tính toán Hàng Tổng cộng R04 (* Tổng) = R06 + R11 + R26 + R31
  const r04 = rowDataMap['R04'];
  const grandChildren = ['R06', 'R11', 'R26', 'R31'];
  TH_DV_SECTORS.forEach(s => {
    let sum27 = 0, sum28 = 0, sumTot = 0;
    grandChildren.forEach(cId => {
      const c = rowDataMap[cId];
      if (c) {
        sum27 += c.y2027[s];
        sum28 += c.y2028[s];
        sumTot += c.yTotal[s];
      }
    });
    r04.y2027[s] = sum27;
    r04.y2028[s] = sum28;
    r04.yTotal[s] = sumTot;
  });
  r04.y2027.total = TH_DV_SECTORS.reduce((sum, s) => sum + r04.y2027[s], 0);
  r04.y2028.total = TH_DV_SECTORS.reduce((sum, s) => sum + r04.y2028[s], 0);
  r04.yTotal.total = TH_DV_SECTORS.reduce((sum, s) => sum + r04.yTotal[s], 0);
  r04.y2027.share = 100.0;
  r04.y2028.share = 100.0;
  r04.yTotal.share = 100.0;

  // 6. Tính toán Hàng R05 (Tỷ trọng theo mảng %)
  const r05 = rowDataMap['R05'];
  TH_DV_SECTORS.forEach(s => {
    r05.y2027[s] = r04.y2027.total > 0 ? (r04.y2027[s] / r04.y2027.total * 100) : 0;
    r05.y2028[s] = r04.y2028.total > 0 ? (r04.y2028[s] / r04.y2028.total * 100) : 0;
    r05.yTotal[s] = r04.yTotal.total > 0 ? (r04.yTotal[s] / r04.yTotal.total * 100) : 0;
  });
  r05.y2027.total = 100.0;
  r05.y2028.total = 100.0;
  r05.yTotal.total = 100.0;

  // 7. Tính Tỷ trọng % cho tất cả các hàng còn lại (R06 -> R31)
  TH_DV_ROW_DEFS.forEach(r => {
    if (r.id === 'R04' || r.id === 'R05') return;
    const d = rowDataMap[r.id];
    d.y2027.share = r04.y2027.total > 0 ? (d.y2027.total / r04.y2027.total * 100) : 0;
    d.y2028.share = r04.y2028.total > 0 ? (d.y2028.total / r04.y2028.total * 100) : 0;
    d.yTotal.share = r04.yTotal.total > 0 ? (d.yTotal.total / r04.yTotal.total * 100) : 0;
  });

  return TH_DV_ROW_DEFS.map(r => rowDataMap[r.id]);
}

// Tổng hợp chung toàn mạng
function computeOverallSummary(tablesMang) {
  let total27 = 0.0;
  let total28 = 0.0;
  let totalStrat = 0.0;
  const byMang = [];

  tablesMang.forEach(t => {
    const totRow = t.rows.find(r => r.is_total);
    const row27 = totRow ? totRow.y2027 : 0.0;
    const row28 = totRow ? totRow.y2028 : 0.0;
    const rowTot = totRow ? totRow.tong : (row27 + row28);
    const stratTot = totRow ? (totRow.strat_tot || 0.0) : 0.0;

    total27 += row27;
    total28 += row28;
    totalStrat += stratTot;

    byMang.push({
      name: t.title,
      tong: rowTot,
      y2027: row27,
      y2028: row28,
      strat_tot: stratTot
    });
  });

  const totalInv = total27 + total28;
  byMang.forEach(m => {
    m.share = totalInv > 0 ? (m.tong / totalInv * 100) : 0.0;
  });

  return {
    total_investment: totalInv,
    total_2027: total27,
    total_2028: total28,
    total_strategy_5y: totalStrat,
    share_2027: totalInv > 0 ? (total27 / totalInv * 100) : 0.0,
    share_2028: totalInv > 0 ? (total28 / totalInv * 100) : 0.0,
    by_mang: byMang
  };
}

// Xây dựng bảng đối chiếu Chiến lược 5 năm
function buildStrategyComparison(tablesMang) {
  const list = [];
  tablesMang.forEach(t => {
    const totRow = t.rows.find(r => r.is_total);
    if (!totRow) return;

    const qTot = totRow.tong;
    const q27 = totRow.y2027;
    const q28 = totRow.y2028;
    const s27 = totRow.strat_27 || 0.0;
    const s28 = totRow.strat_28 || 0.0;
    const s27_28 = s27 + s28;
    const s5y = totRow.strat_tot || 0.0;
    const delta = qTot - s27_28;

    let status = 'Vừa khớp';
    let statusColor = 'green';
    if (delta > 0.01) {
      status = `Vượt CL +${delta.toFixed(2)} M$`;
      statusColor = 'orange';
    } else if (delta < -0.01) {
      status = `Dưới CL ${delta.toFixed(2)} M$`;
      statusColor = 'blue';
    }

    list.push({
      title: t.title,
      qhdc_2027: q27,
      qhdc_2028: q28,
      qhdc_total: qTot,
      strat_2027: s27,
      strat_2028: s28,
      strat_27_28: s27_28,
      strat_5y: s5y,
      delta: delta,
      status: status,
      status_color: statusColor
    });
  });

  return list;
}

/* ==================== RENDERING UI ==================== */

// Định dạng số
function fmtVal(num) {
  if (num === null || num === undefined || isNaN(num) || num === 0) return '0.00';
  const scaled = num * reportState.unitMultiplier;
  return scaled.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function renderAllTabs() {
  if (!reportState.rawResponse) return;
  renderTabTongHop();
  renderTabMang();
  renderTabDichVu();
  renderTabStrategy();
  if (window.lucide) lucide.createIcons();
}

// TAB 1: TỔNG HỢP
function renderTabTongHop() {
  const summary = reportState.rawResponse?.summary;
  if (!summary) return;

  const elTot = document.getElementById('kpiTotalInv');
  if (elTot) elTot.textContent = fmtVal(summary.total_investment);

  const el27 = document.getElementById('kpi2027');
  if (el27) el27.textContent = fmtVal(summary.total_2027);

  const el28 = document.getElementById('kpi2028');
  if (el28) el28.textContent = fmtVal(summary.total_2028);

  const elStrat = document.getElementById('kpiStrat5Y');
  if (elStrat) elStrat.textContent = fmtVal(summary.total_strategy_5y);

  // 1. Render Bảng TH theo DV & Mảng (33 cột x 28 hàng chuẩn)
  renderTableTHTheoDV();

  // 2. Render Bảng tóm tắt theo mảng phía dưới
  renderTableSummaryMang(summary);
}

// Render Master Matrix Table: Sheet "TH theo DV" (33 cột x 28 hàng)
function renderTableTHTheoDV() {
  const tbody = document.getElementById('tableTHTheoDVBody');
  if (!tbody) return;

  const rows = reportState.rawResponse?.th_theo_dv || [];
  if (rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="33" style="text-align: center; padding: 2rem; color: #94a3b8;">Chưa có dữ liệu ma trận TH theo Dịch vụ & Mảng</td></tr>`;
    return;
  }

  const cellFmt = (val, isShare) => {
    if (isShare) {
      return (val !== null && val !== undefined && val > 0.0001) ? (val.toFixed(1) + '%') : '-';
    }
    if (val === null || val === undefined || isNaN(val) || Math.abs(val) < 0.000001) {
      return '<span style="color: #cbd5e1;">-</span>';
    }
    return fmtVal(val);
  };

  let html = '';
  rows.forEach(r => {
    let rowClass = 'row-leaf';
    if (r.type === 'grand_total') rowClass = 'row-grand-total';
    else if (r.type === 'share_header') rowClass = 'row-share-header';
    else if (r.type === 'group') rowClass = 'row-group';
    else if (r.type === 'subgroup') rowClass = 'row-subgroup';

    const isShareHeader = (r.type === 'share_header');
    const isSubItem = (r.stt === '-');
    const ndClass = isSubItem ? 'col-sticky-nd cell-sub-item' : 'col-sticky-nd';
    const ndStyle = isSubItem ? 'padding-left: 24px;' : (r.type === 'group' || r.type === 'grand_total' ? 'font-weight: 800;' : (r.type === 'subgroup' ? 'font-weight: 700;' : ''));

    html += `<tr class="${rowClass}" data-nd="${escapeHtml(r.nd).toLowerCase()}" data-madv="${escapeHtml(r.madv || '').toLowerCase()}">`;
    html += `<td class="col-sticky-stt">${escapeHtml(r.stt || '')}</td>`;
    html += `<td class="${ndClass}" style="${ndStyle}" title="${escapeHtml(r.nd)}">${escapeHtml(r.nd)}</td>`;
    html += `<td class="col-sticky-madv" style="font-weight: 600; color: #64748b;">${escapeHtml(r.madv || '')}</td>`;

    // 2027 (10 cột)
    TH_DV_SECTORS.forEach(s => {
      html += `<td class="col-num col-p-2027">${cellFmt(r.y2027[s], isShareHeader)}</td>`;
    });
    html += `<td class="col-num col-p-2027 col-tot" style="font-weight: 800; background: #e0e7ff; color: #3730a3;">${isShareHeader ? '100%' : cellFmt(r.y2027.total, false)}</td>`;
    html += `<td class="col-num col-p-2027 period-sep" style="font-weight: 700; color: #4338ca; background: #eef2ff;">${isShareHeader ? '-' : cellFmt(r.y2027.share, true)}</td>`;

    // 2028 (10 cột)
    TH_DV_SECTORS.forEach(s => {
      html += `<td class="col-num col-p-2028">${cellFmt(r.y2028[s], isShareHeader)}</td>`;
    });
    html += `<td class="col-num col-p-2028 col-tot" style="font-weight: 800; background: #d1fae5; color: #065f46;">${isShareHeader ? '100%' : cellFmt(r.y2028.total, false)}</td>`;
    html += `<td class="col-num col-p-2028 period-sep" style="font-weight: 700; color: #047857; background: #ecfdf5;">${isShareHeader ? '-' : cellFmt(r.y2028.share, true)}</td>`;

    // 2027-2028 (10 cột)
    TH_DV_SECTORS.forEach(s => {
      html += `<td class="col-num col-p-tot">${cellFmt(r.yTotal[s], isShareHeader)}</td>`;
    });
    html += `<td class="col-num col-p-tot col-tot" style="font-weight: 800; background: #ede9fe; color: #5b21b6;">${isShareHeader ? '100%' : cellFmt(r.yTotal.total, false)}</td>`;
    html += `<td class="col-num col-p-tot" style="font-weight: 800; color: #5b21b6; background: #f5f3ff;">${isShareHeader ? '-' : cellFmt(r.yTotal.share, true)}</td>`;

    html += `</tr>`;
  });

  tbody.innerHTML = html;
}

// Render Bảng tóm tắt theo mảng toàn mạng
function renderTableSummaryMang(summary) {
  const tbody = document.getElementById('tableSummaryMangBody');
  if (!tbody) return;

  const rows = summary.by_mang || [];
  let html = '';

  rows.forEach((m, idx) => {
    html += `
      <tr>
        <td class="col-mang" style="font-weight: 700; color: #0f172a;">
          <span style="display: inline-block; width: 22px; color: #64748b;">${idx + 1}.</span>
          ${escapeHtml(m.name)}
        </td>
        <td class="col-num col-tot">${fmtVal(m.tong)}</td>
        <td class="col-num col-27">${fmtVal(m.y2027)}</td>
        <td class="col-num col-28">${fmtVal(m.y2028)}</td>
        <td class="col-num" style="color: #475569; font-weight: 700; width: 100px;">
          ${m.share ? m.share.toFixed(1) + '%' : '-'}
        </td>
      </tr>
    `;
  });

  // Hàng tổng cộng
  html += `
    <tr class="row-total-table">
      <td class="col-mang" style="font-weight: 800; font-size: 0.9rem; color: #065f46;">
        TỔNG CỘNG TOÀN MẠNG
      </td>
      <td class="col-num col-tot" style="font-size: 0.95rem; font-weight: 800; color: #065f46;">${fmtVal(summary.total_investment)}</td>
      <td class="col-num col-27" style="font-weight: 800; color: #065f46;">${fmtVal(summary.total_2027)}</td>
      <td class="col-num col-28" style="font-weight: 800; color: #065f46;">${fmtVal(summary.total_2028)}</td>
      <td class="col-num" style="color: #065f46; font-weight: 800;">100%</td>
    </tr>
  `;

  tbody.innerHTML = html;
}

// TAB 2: TỔNG HỢP THEO MẢNG
function renderTabMang() {
  const container = document.getElementById('containerTablesMang');
  if (!container) return;

  const tables = reportState.rawResponse?.tables_mang || [];
  if (tables.length === 0) {
    container.innerHTML = `<div style="text-align: center; padding: 3rem; color: #94a3b8;">Chưa có dữ liệu bảng Tổng hợp theo Mảng</div>`;
    return;
  }

  const kw = reportState.searchKeyword;
  const sectorFilter = reportState.selectedSectorFilter;
  let html = '';

  tables.forEach((t, tIdx) => {
    if (sectorFilter !== 'ALL') {
      const match = t.title.toUpperCase().includes(sectorFilter);
      if (!match) return;
    }

    const filteredRows = t.rows.filter(r => {
      if (!kw) return true;
      return (r.name && r.name.toLowerCase().includes(kw)) ||
             (t.title && t.title.toLowerCase().includes(kw));
    });

    if (kw && filteredRows.length === 0) return;

    html += `
      <div class="table-card" style="margin-bottom: 1.5rem;" id="card_mang_${tIdx}">
        <div class="table-card-header">
          <h2 class="table-card-title">
            <i data-lucide="layers" class="w-4 h-4 text-indigo-600"></i>
            <span>${escapeHtml(t.title)}</span>
          </h2>
          <div style="display: flex; align-items: center; gap: 8px;">
            <span class="table-badge-count">${filteredRows.length} hạng mục</span>
            <button type="button" class="btn btn-outline btn-xs" onclick="copyTableHtml('table_mang_${tIdx}', '${escapeHtml(t.title)}')">
              <i data-lucide="copy" class="w-3.5 h-3.5 inline mr-1"></i> Sao chép
            </button>
          </div>
        </div>
        <div style="overflow-x: auto;">
          <table class="report-table" id="table_mang_${tIdx}">
            <thead>
              <tr>
                <th class="col-mang">Mảng</th>
                <th class="col-num col-tot">Tổng (<span class="unit-label-text">M$</span>)</th>
                <th class="col-num col-27">2027 (<span class="unit-label-text">M$</span>)</th>
                <th class="col-num col-28">2028 (<span class="unit-label-text">M$</span>)</th>
              </tr>
            </thead>
            <tbody>
    `;

    filteredRows.forEach(r => {
      const rowClass = r.is_total ? 'row-total-table' : (r.is_sub ? 'row-sub-item' : '');
      const namePrefix = r.stt && r.stt !== '-' ? `<span style="display:inline-block;width:20px;color:#64748b;">${escapeHtml(r.stt)}.</span>` : (r.is_sub ? '<span style="color:#94a3b8;margin-right:6px;">-</span>' : '');

      html += `
        <tr class="${rowClass}">
          <td class="col-mang" style="${r.is_total ? 'font-weight: 800; color: #065f46;' : ''}">
            ${namePrefix}${escapeHtml(r.name)}
          </td>
          <td class="col-num col-tot">${fmtVal(r.tong)}</td>
          <td class="col-num col-27">${fmtVal(r.y2027)}</td>
          <td class="col-num col-28">${fmtVal(r.y2028)}</td>
        </tr>
      `;
    });

    html += `
            </tbody>
          </table>
        </div>
      </div>
    `;
  });

  container.innerHTML = html || `<div style="text-align: center; padding: 3rem; color: #94a3b8;">Không tìm thấy kết quả phù hợp với bộ lọc</div>`;
}

// TAB 3: TỔNG HỢP THEO DỊCH VỤ
function renderTabDichVu() {
  const container = document.getElementById('containerTablesDichVu');
  if (!container) return;

  const tables = reportState.rawResponse?.tables_dv || [];
  if (tables.length === 0) {
    container.innerHTML = `<div style="text-align: center; padding: 3rem; color: #94a3b8;">Chưa có dữ liệu bảng Tổng hợp theo Dịch vụ</div>`;
    return;
  }

  const kw = reportState.searchKeyword;
  const srvFilter = reportState.selectedServiceFilter;
  let html = '';

  tables.forEach((t, tIdx) => {
    if (srvFilter !== 'ALL') {
      const match = t.title.toUpperCase().includes(srvFilter);
      if (!match) return;
    }

    const filteredRows = t.rows.filter(r => {
      if (!kw) return true;
      return (r.name && r.name.toLowerCase().includes(kw)) ||
             (r.service && r.service.toLowerCase().includes(kw)) ||
             (r.mang && r.mang.toLowerCase().includes(kw)) ||
             (t.title && t.title.toLowerCase().includes(kw));
    });

    if (kw && filteredRows.length === 0) return;

    html += `
      <div class="table-card" style="margin-bottom: 1.5rem;" id="card_dv_${tIdx}">
        <div class="table-card-header">
          <h2 class="table-card-title">
            <i data-lucide="network" class="w-4 h-4 text-emerald-600"></i>
            <span>${escapeHtml(t.title)}</span>
          </h2>
          <div style="display: flex; align-items: center; gap: 8px;">
            <span class="table-badge-count">${filteredRows.length} hạng mục</span>
            <button type="button" class="btn btn-outline btn-xs" onclick="copyTableHtml('table_dv_${tIdx}', '${escapeHtml(t.title)}')">
              <i data-lucide="copy" class="w-3.5 h-3.5 inline mr-1"></i> Sao chép
            </button>
          </div>
        </div>
        <div style="overflow-x: auto;">
          <table class="report-table" id="table_dv_${tIdx}">
            <thead>
              <tr>
                <th class="col-mang">Dịch vụ / Hạng mục</th>
                <th class="col-num col-tot">Tổng (<span class="unit-label-text">M$</span>)</th>
                <th class="col-num col-27">2027 (<span class="unit-label-text">M$</span>)</th>
                <th class="col-num col-28">2028 (<span class="unit-label-text">M$</span>)</th>
              </tr>
            </thead>
            <tbody>
    `;

    filteredRows.forEach(r => {
      const rowClass = r.is_total ? 'row-total-table' : (r.is_sub ? 'row-sub-item' : '');
      const displayName = r.service ? `${r.service} - ${r.mang || r.name}` : (r.mang || r.name);
      const namePrefix = r.stt && r.stt !== '-' ? `<span style="display:inline-block;width:20px;color:#64748b;">${escapeHtml(r.stt)}.</span>` : (r.is_sub ? '<span style="color:#94a3b8;margin-right:6px;">-</span>' : '');

      html += `
        <tr class="${rowClass}">
          <td class="col-mang" style="${r.is_total ? 'font-weight: 800; color: #065f46;' : ''}">
            ${namePrefix}${escapeHtml(displayName)}
          </td>
          <td class="col-num col-tot">${fmtVal(r.tong)}</td>
          <td class="col-num col-27">${fmtVal(r.y2027)}</td>
          <td class="col-num col-28">${fmtVal(r.y2028)}</td>
        </tr>
      `;
    });

    html += `
            </tbody>
          </table>
        </div>
      </div>
    `;
  });

  container.innerHTML = html || `<div style="text-align: center; padding: 3rem; color: #94a3b8;">Không tìm thấy kết quả phù hợp với bộ lọc</div>`;
}

// TAB 4: ĐỐI CHIẾU CHIẾN LƯỢC 5 NĂM
function renderTabStrategy() {
  const tbody = document.getElementById('tableStrategyBody');
  if (!tbody) return;

  const compList = reportState.rawResponse?.strategy_comparison || [];
  if (compList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align: center; padding: 2rem; color: #94a3b8;">Chưa có dữ liệu đối chiếu chiến lược</td></tr>`;
    return;
  }

  let html = '';
  let sumQ27 = 0, sumQ28 = 0, sumQTot = 0;
  let sumS27 = 0, sumS28 = 0, sumS27_28 = 0, sumS5Y = 0;

  compList.forEach((c, idx) => {
    sumQ27 += c.qhdc_2027;
    sumQ28 += c.qhdc_2028;
    sumQTot += c.qhdc_total;
    sumS27 += c.strat_2027;
    sumS28 += c.strat_2028;
    sumS27_28 += c.strat_27_28;
    sumS5Y += c.strat_5y;

    const deltaClass = c.delta > 0.01 ? 'text-amber-600' : (c.delta < -0.01 ? 'text-blue-600' : 'text-emerald-600');
    const badgeClass = c.status_color === 'orange' ? 'badge-orange' : (c.status_color === 'blue' ? 'badge-blue' : 'badge-green');

    html += `
      <tr>
        <td style="font-weight: 700; color: #0f172a; min-width: 240px;">
          <span style="display:inline-block;width:20px;color:#64748b;">${idx + 1}.</span>
          ${escapeHtml(c.title)}
        </td>
        <td class="col-num col-27">${fmtVal(c.qhdc_2027)}</td>
        <td class="col-num col-28">${fmtVal(c.qhdc_2028)}</td>
        <td class="col-num col-tot" style="border-right: 2px solid #cbd5e1;">${fmtVal(c.qhdc_total)}</td>

        <td class="col-num" style="color: #475569;">${fmtVal(c.strat_2027)}</td>
        <td class="col-num" style="color: #475569;">${fmtVal(c.strat_2028)}</td>
        <td class="col-num" style="font-weight: 800; color: #b45309;">${fmtVal(c.strat_27_28)}</td>
        <td class="col-num" style="font-weight: 800; color: #92400e; border-right: 2px solid #cbd5e1;">${fmtVal(c.strat_5y)}</td>

        <td class="col-num ${deltaClass}" style="font-weight: 800;">
          ${c.delta > 0 ? '+' : ''}${fmtVal(c.delta)}
        </td>
        <td style="text-align: center;">
          <span class="badge-status ${badgeClass}">${escapeHtml(c.status)}</span>
        </td>
      </tr>
    `;
  });

  // Hàng tổng cộng đối chiếu
  const totalDelta = sumQTot - sumS27_28;
  const totDeltaClass = totalDelta > 0.01 ? 'text-amber-600' : (totalDelta < -0.01 ? 'text-blue-600' : 'text-emerald-600');
  const totBadgeClass = totalDelta > 0.01 ? 'badge-orange' : (totalDelta < -0.01 ? 'badge-blue' : 'badge-green');
  let totStatus = 'Khớp chiến lược';
  if (totalDelta > 0.05) totStatus = `Vượt CL +${totalDelta.toFixed(2)} M$`;
  else if (totalDelta < -0.05) totStatus = `Dưới CL ${totalDelta.toFixed(2)} M$`;

  html += `
    <tr class="row-total-table" style="background: #f8fafc; border-top: 2.5px solid #065f46;">
      <td style="font-weight: 800; color: #065f46; font-size: 0.9rem;">TỔNG CỘNG TOÀN MẠNG</td>
      <td class="col-num col-27" style="font-weight: 800;">${fmtVal(sumQ27)}</td>
      <td class="col-num col-28" style="font-weight: 800;">${fmtVal(sumQ28)}</td>
      <td class="col-num col-tot" style="font-weight: 800; border-right: 2px solid #cbd5e1; font-size: 0.95rem;">${fmtVal(sumQTot)}</td>

      <td class="col-num" style="font-weight: 700; color: #475569;">${fmtVal(sumS27)}</td>
      <td class="col-num" style="font-weight: 700; color: #475569;">${fmtVal(sumS28)}</td>
      <td class="col-num" style="font-weight: 800; color: #b45309;">${fmtVal(sumS27_28)}</td>
      <td class="col-num" style="font-weight: 800; color: #92400e; border-right: 2px solid #cbd5e1;">${fmtVal(sumS5Y)}</td>

      <td class="col-num ${totDeltaClass}" style="font-weight: 800; font-size: 0.95rem;">
        ${totalDelta > 0 ? '+' : ''}${fmtVal(totalDelta)}
      </td>
      <td style="text-align: center;">
        <span class="badge-status ${totBadgeClass}">${totStatus}</span>
      </td>
    </tr>
  `;

  tbody.innerHTML = html;
}

/* ==================== CÁC HÀM TƯƠNG TÁC NGƯỜI DÙNG ==================== */

// Chuyển đổi Sub-tab
function switchReportSubtab(subtabId) {
  reportState.activeSubtab = subtabId;

  document.querySelectorAll('.report-subtab-btn').forEach(btn => {
    if (btn.getAttribute('data-subtab') === subtabId) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  document.querySelectorAll('.report-subtab-content').forEach(content => {
    if (content.id === subtabId) {
      content.classList.add('active');
    } else {
      content.classList.remove('active');
    }
  });

  renderAllTabs();
}

// Chuyển đổi đơn vị tiền tệ
function setReportUnit(unit) {
  if (unit === 'K') {
    reportState.unitMultiplier = 1000.0;
    reportState.unitLabel = 'K$ (Nghìn USD)';
  } else if (unit === 'USD') {
    reportState.unitMultiplier = 1000000.0;
    reportState.unitLabel = 'USD';
  } else {
    reportState.unitMultiplier = 1.0;
    reportState.unitLabel = 'M$ (Triệu USD)';
  }

  document.querySelectorAll('.unit-btn').forEach(btn => {
    if (btn.getAttribute('data-unit') === unit) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  document.querySelectorAll('.unit-label-text').forEach(el => {
    el.textContent = unit === 'USD' ? 'USD' : (unit === 'K' ? 'K$' : 'M$');
  });

  renderAllTabs();
}

// Chế độ xem kỳ của bảng TH theo DV
function setTHDVPeriod(period) {
  reportState.thdvPeriodView = period;
  const tbl = document.getElementById('tableTHTheoDV');
  const btns = document.querySelectorAll('#thdvPeriodButtons .period-control-btn');
  btns.forEach(b => {
    if (b.getAttribute('data-period') === period) b.classList.add('active');
    else b.classList.remove('active');
  });

  if (!tbl) return;
  tbl.classList.remove('view-p-2027', 'view-p-2028', 'view-p-tot', 'view-single-period');
  if (period === '2027') {
    tbl.classList.add('view-p-2027', 'view-single-period');
  } else if (period === '2028') {
    tbl.classList.add('view-p-2028', 'view-single-period');
  } else if (period === 'TOT') {
    tbl.classList.add('view-p-tot', 'view-single-period');
  }
}

// Lọc nhanh dòng trong bảng TH theo DV
function filterTHDVTable(kw) {
  const query = (kw || '').trim().toLowerCase();
  const rows = document.querySelectorAll('#tableTHTheoDVBody tr');
  rows.forEach(tr => {
    if (!query) {
      tr.style.display = '';
      return;
    }
    const nd = tr.getAttribute('data-nd') || '';
    const madv = tr.getAttribute('data-madv') || '';
    const isHeader = tr.classList.contains('row-grand-total') || tr.classList.contains('row-share-header');
    if (isHeader || nd.includes(query) || madv.includes(query)) {
      tr.style.display = '';
    } else {
      tr.style.display = 'none';
    }
  });
}

// Xuất riêng bảng TH theo DV ra file Excel (.xls)
function exportTHTheoDVExcel() {
  const tbl = document.getElementById('tableTHTheoDV');
  if (!tbl) return;

  const clone = tbl.cloneNode(true);
  clone.querySelectorAll('tr').forEach(tr => {
    if (tr.style.display === 'none') tr.remove();
  });

  const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="utf-8">
      <style>
        table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt; }
        th, td { border: 1px solid #999; padding: 4px 6px; }
        th { background-color: #e2e8f0; font-weight: bold; text-align: center; }
        .col-num { text-align: right; mso-number-format: "#,##0.00"; }
        .row-grand-total { background-color: #ecfdf5; font-weight: bold; }
        .row-group { background-color: #eef2ff; font-weight: bold; }
        .row-subgroup { background-color: #f1f5f9; font-weight: bold; }
      </style>
    </head>
    <body>
      ${clone.outerHTML}
    </body>
    </html>
  `;

  const blob = new Blob(['\\uFEFF' + html], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'QHDC_TH_theo_DV_33Cot_' + new Date().toISOString().slice(0, 10) + '.xls';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Lọc từ khóa toàn cục
function onReportSearch(keyword) {
  reportState.searchKeyword = (keyword || '').toLowerCase().trim();
  renderAllTabs();
}

// Sao chép HTML bảng vào Clipboard (hỗ trợ dán trực tiếp vào Excel)
function copyTableHtml(tableId, tableName) {
  const table = document.getElementById(tableId);
  if (!table) return;

  let text = '';
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r];
    if (row.style.display === 'none') continue;
    const cells = [];
    for (let c = 0; c < row.cells.length; c++) {
      if (row.cells[c].style.display === 'none') continue;
      cells.push(row.cells[c].innerText.trim().replace(/\\t/g, ' '));
    }
    text += cells.join('\\t') + '\\n';
  }

  navigator.clipboard.writeText(text).then(() => {
    if (typeof showToast === 'function') {
      showToast(`Đã sao chép bảng "${tableName}" vào Clipboard! Dán vào Excel bằng Ctrl+V.`, 'success');
    } else {
      alert(`Đã sao chép "${tableName}"! Bạn có thể dán trực tiếp vào Excel bằng Ctrl+V.`);
    }
  }).catch(err => {
    console.error('Lỗi khi sao chép:', err);
  });
}

// Sao chép toàn bộ các bảng của sub-tab hiện tại
function copyCurrentTabTables() {
  const subtab = reportState.activeSubtab;
  let targetId = 'tableTHTheoDV';
  let name = 'Bảng Tổng hợp theo Dịch vụ & Mảng (TH theo DV)';

  if (subtab === 'subtabTongHop') {
    targetId = 'tableTHTheoDV';
    name = 'Bảng Tổng hợp TH theo DV';
  } else if (subtab === 'subtabStratDetail') {
    targetId = 'tableStrategy';
    name = 'Bảng Đối chiếu Chiến lược 5 năm';
  } else if (subtab === 'subtabMang') {
    targetId = 'containerTablesMang';
    name = 'Các bảng TH theo Mảng';
  } else if (subtab === 'subtabDichVu') {
    targetId = 'containerTablesDichVu';
    name = 'Các bảng TH theo Dịch vụ';
  }

  const el = document.getElementById(targetId);
  if (!el) return;

  const tables = el.tagName === 'TABLE' ? [el] : el.querySelectorAll('table');
  let fullText = '';
  tables.forEach(table => {
    for (let r = 0; r < table.rows.length; r++) {
      const row = table.rows[r];
      if (row.style.display === 'none') continue;
      const cells = [];
      for (let c = 0; c < row.cells.length; c++) {
        if (row.cells[c].style.display === 'none') continue;
        cells.push(row.cells[c].innerText.trim().replace(/\\t/g, ' '));
      }
      fullText += cells.join('\\t') + '\\n';
    }
    fullText += '\\n';
  });

  if (fullText) {
    navigator.clipboard.writeText(fullText).then(() => {
      if (typeof showToast === 'function') {
        showToast(`Đã sao chép ${name} vào Clipboard! Dán vào Excel bằng Ctrl+V.`, 'success');
      }
    });
  }
}

// Xuất toàn bộ Báo cáo ra file Excel workbook
function exportReportExcel() {
  if (typeof XLSX === 'undefined') {
    alert('Thư viện Excel chưa được nạp!');
    return;
  }

  const wb = XLSX.utils.book_new();

  // Sheet 0: TH theo DV (Ma trận 33 cột x 28 hàng chuẩn)
  const tableTHDV = document.getElementById('tableTHTheoDV');
  if (tableTHDV) {
    const ws0 = XLSX.utils.table_to_sheet(tableTHDV);
    XLSX.utils.book_append_sheet(wb, ws0, 'TH theo DV');
  }

  // Sheet 1: TH Toàn Mạng
  const tableSummary = document.getElementById('tableSummaryMang');
  if (tableSummary) {
    const ws1 = XLSX.utils.table_to_sheet(tableSummary);
    XLSX.utils.book_append_sheet(wb, ws1, 'TH Toàn mạng');
  }

  // Sheet 2: Đối chiếu Chiến Lược
  const tableStrat = document.getElementById('tableStrategy');
  if (tableStrat) {
    const ws2 = XLSX.utils.table_to_sheet(tableStrat);
    XLSX.utils.book_append_sheet(wb, ws2, 'ĐC Chiến lược 5 năm');
  }

  XLSX.writeFile(wb, `Bao_cao_QHDC_Chien_luoc_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// Khởi tạo listeners khi tải trang
document.addEventListener('DOMContentLoaded', () => {
  // Khởi tạo sẵn cấu trúc bảng 28 hàng x 33 cột
  if (!reportState.rawResponse) {
    updateReportFromExtractedData([]);
  }
  // Lắng nghe click các chip lọc mảng
  document.querySelectorAll('#mangSectorChips .strat-sector-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('#mangSectorChips .strat-sector-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      reportState.selectedSectorFilter = chip.getAttribute('data-sector') || 'ALL';
      renderTabMang();
      if (window.lucide) lucide.createIcons();
    });
  });

  // Lắng nghe click các chip lọc dịch vụ
  document.querySelectorAll('#dvServiceChips .strat-sector-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('#dvServiceChips .strat-sector-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      reportState.selectedServiceFilter = chip.getAttribute('data-service') || 'ALL';
      renderTabDichVu();
      if (window.lucide) lucide.createIcons();
    });
  });
});

// Expose toàn cục
window.reportState = reportState;
window.updateReportFromExtractedData = updateReportFromExtractedData;
window.renderAllReportTabs = renderAllTabs;
window.setTHDVPeriod = setTHDVPeriod;
window.filterTHDVTable = filterTHDVTable;
window.exportTHTheoDVExcel = exportTHTheoDVExcel;
window.copyTableHtml = copyTableHtml;
window.copyCurrentTabTables = copyCurrentTabTables;
window.exportReportExcel = exportReportExcel;
"""

with open('report.js', 'w', encoding='utf-8') as f:
    f.write(part1)

print('Successfully written report.js! Size:', len(part1))
