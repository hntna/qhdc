# -*- coding: utf-8 -*-
"""
Module trích xuất & tính toán dữ liệu Báo cáo Tổng hợp từ file Masterlist Excel
Hỗ trợ cả file CNTT_Masterlist 2027-2028_Mau.xlsx và Masterlist 2027-2028_Mau_Tach.xlsx
Tự động giải các công thức SUMIFS từ sheet PL1.1 nếu file có dữ liệu vật tư.
"""
import os
import re
import openpyxl

DEFAULT_CNTT_PATH = r'd:\Haint28\CONG VIEC\20260824 QHĐC Thi truong\STL\QHDC 2027-2028\Final\CNTT\CNTT_Masterlist 2027-2028_Mau.xlsx'
WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
FALLBACK_TACH_PATH = os.path.join(WORKSPACE_DIR, 'Masterlist 2027-2028_Mau_Tach.xlsx')
FALLBACK_MAU_PATH = os.path.join(WORKSPACE_DIR, 'Masterlist 2027-2028_Mau.xlsx')

def find_best_report_file(custom_path=None):
    if custom_path and os.path.exists(custom_path):
        return custom_path
    if os.path.exists(DEFAULT_CNTT_PATH):
        return DEFAULT_CNTT_PATH
    if os.path.exists(FALLBACK_TACH_PATH):
        return FALLBACK_TACH_PATH
    if os.path.exists(FALLBACK_MAU_PATH):
        return FALLBACK_MAU_PATH
    return None

def safe_num(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        s = str(val).replace(',', '').strip()
        return float(s)
    except Exception:
        return 0.0

# Ánh xạ cột trong sheet "TH theo DV"
# D-J: 2027 (VT, ML, CĐBR, CNTT, TD, CĐ, HT)
# M-S: 2028 (VT, ML, CĐBR, CNTT, TD, CĐ, HT)
MANG_COL_MAP_2027 = {
    'D': 'VT', 'E': 'ML', 'F': 'CĐBR', 'G': 'CNTT', 'H': 'TD', 'I': 'CĐ', 'J': 'HT', 'K': 'TOTAL'
}
MANG_COL_MAP_2028 = {
    'M': 'VT', 'N': 'ML', 'O': 'CĐBR', 'P': 'CNTT', 'Q': 'TD', 'R': 'CĐ', 'S': 'HT', 'T': 'TOTAL'
}

def build_th_dv_matrix(wb_val):
    """Tính toán ma trận TH theo DV trực tiếp từ PL1.1 nếu có dữ liệu"""
    matrix_27 = {}
    matrix_28 = {}

    if 'PL1.1 ML2027-2028' not in wb_val.sheetnames:
        return matrix_27, matrix_28

    ws_pl = wb_val['PL1.1 ML2027-2028']
    if ws_pl.max_row <= 7:
        return matrix_27, matrix_28

    for r in range(7, ws_pl.max_row + 1):
        tt27 = ws_pl.cell(r, 7).value
        tt28 = ws_pl.cell(r, 8).value
        mang = ws_pl.cell(r, 10).value
        madv = ws_pl.cell(r, 11).value

        if not mang or (tt27 is None and tt28 is None):
            continue

        mang_key = str(mang).strip().upper()
        madv_key = str(madv or '').strip().upper()

        val27 = safe_num(tt27) / 1000000.0
        val28 = safe_num(tt28) / 1000000.0

        k = (mang_key, madv_key)
        matrix_27[k] = matrix_27.get(k, 0.0) + val27
        matrix_28[k] = matrix_28.get(k, 0.0) + val28

    return matrix_27, matrix_28

def parse_report_workbook(file_path):
    if not file_path or not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    wb_val = openpyxl.load_workbook(file_path, data_only=True)
    wb_form = openpyxl.load_workbook(file_path, data_only=False)

    sheet_names = wb_val.sheetnames

    name_m = None
    name_dv = None
    for s in sheet_names:
        s_lower = s.lower()
        if 'mảng' in s_lower or 'mang' in s_lower:
            name_m = s
        if 'dịch vụ' in s_lower or 'dich vu' in s_lower:
            name_dv = s

    if not name_m or not name_dv:
        if os.path.exists(FALLBACK_TACH_PATH):
            wb_val = openpyxl.load_workbook(FALLBACK_TACH_PATH, data_only=True)
            wb_form = openpyxl.load_workbook(FALLBACK_TACH_PATH, data_only=False)
            file_path = FALLBACK_TACH_PATH
            name_m = 'TH theo Mảng'
            name_dv = 'TH theo Dịch vụ'

    # Build dynamic matrix nếu PL1.1 có dữ liệu
    matrix_27, matrix_28 = build_th_dv_matrix(wb_val)

    # Đọc bản đồ Mã DV theo hàng trong sheet 'TH theo DV'
    dv_row_map = {}
    if 'TH theo DV' in wb_val.sheetnames:
        ws_dv_ref = wb_val['TH theo DV']
        for r in range(4, ws_dv_ref.max_row + 1):
            c_val = ws_dv_ref.cell(r, 3).value
            if c_val:
                dv_row_map[r] = str(c_val).strip().upper()

    tables_mang = parse_sheet_mang(wb_val, wb_form, name_m, matrix_27, matrix_28, dv_row_map) if name_m else []
    tables_dv = parse_sheet_dv(wb_val, wb_form, name_dv, matrix_27, matrix_28, dv_row_map) if name_dv else []
    summary_data = compute_overall_summary(wb_val, tables_mang, tables_dv)
    strategy_comp = build_strategy_comparison(tables_mang, tables_dv)

    return {
        'success': True,
        'file_path': file_path,
        'file_name': os.path.basename(file_path),
        'file_size': os.path.getsize(file_path),
        'sheets': wb_val.sheetnames,
        'summary': summary_data,
        'tables_mang': tables_mang,
        'tables_dv': tables_dv,
        'strategy_comparison': strategy_comp
    }

def resolve_cell_formula(form_val, val_val, matrix_27, matrix_28, dv_row_map):
    """Giải giá trị công thức tham chiếu 'TH theo DV'!<Col><Row> nếu có matrix từ PL1.1"""
    if val_val is not None and isinstance(val_val, (int, float)) and val_val > 0:
        return float(val_val)

    if form_val and isinstance(form_val, str) and matrix_27:
        m = re.search(r"'?TH theo DV'?!([A-Z]+)(\d+)", form_val, re.IGNORECASE)
        if m:
            col_letter = m.group(1).upper()
            row_idx = int(m.group(2))
            ma_dv = dv_row_map.get(row_idx, '')
            if col_letter in MANG_COL_MAP_2027:
                mang = MANG_COL_MAP_2027[col_letter]
                return matrix_27.get((mang, ma_dv), 0.0)
            elif col_letter in MANG_COL_MAP_2028:
                mang = MANG_COL_MAP_2028[col_letter]
                return matrix_28.get((mang, ma_dv), 0.0)

    return safe_num(val_val)

def parse_sheet_mang(wb_val, wb_form, sheet_name, matrix_27, matrix_28, dv_row_map):
    if sheet_name not in wb_val.sheetnames:
        return []

    ws_val = wb_val[sheet_name]
    ws_form = wb_form[sheet_name]

    tables = []
    current_table = None

    for r in range(1, ws_val.max_row + 1):
        b_val = ws_val.cell(r, 2).value
        c_val = ws_val.cell(r, 3).value
        d_val = ws_val.cell(r, 4).value
        e_val = ws_val.cell(r, 5).value
        f_val = ws_val.cell(r, 6).value

        b_form = ws_form.cell(r, 2).value
        c_form = ws_form.cell(r, 3).value
        d_form = ws_form.cell(r, 4).value
        e_form = ws_form.cell(r, 5).value
        f_form = ws_form.cell(r, 6).value

        b_str = str(b_val or '').strip()
        c_str = str(c_val or '').strip()

        is_sec = False
        for prefix in ['I.', 'II.', 'III.', 'IV.', 'V.', 'VI.', 'VII.', 'VIII.', 'IX.', 'X.', 'Đầu tư', 'Hiện đại']:
            if b_str.startswith(prefix) and not (b_str == 'STT' or c_str == 'STT' or c_str == 'Danh mục'):
                is_sec = True
                break

        if is_sec:
            current_table = {
                'id': f'mang_sec_{len(tables) + 1}',
                'title': b_str,
                'rows': []
            }
            tables.append(current_table)
            continue

        if b_str == 'STT' or c_str == 'STT' or c_str == 'Danh mục' or b_str.startswith('Phụ lục'):
            continue

        if current_table is not None:
            name = c_str if c_str else b_str
            if name or d_val is not None or e_val is not None or f_val is not None:
                stt = b_str if (b_str and b_str != name and b_str != c_str) else ''
                is_total = (name.lower() == 'tổng' or stt.lower() == 'tổng')
                is_sub = (stt == '-' or name.startswith('-'))

                # Giải số liệu 2027 và 2028
                y27 = resolve_cell_formula(e_form, e_val, matrix_27, matrix_28, dv_row_map)
                y28 = resolve_cell_formula(f_form, f_val, matrix_27, matrix_28, dv_row_map)
                tot = resolve_cell_formula(d_form, d_val, matrix_27, matrix_28, dv_row_map)
                if tot == 0.0 and (y27 > 0 or y28 > 0):
                    tot = y27 + y28

                strat_tot = safe_num(ws_val.cell(r, 12).value or ws_val.cell(r, 10).value)
                strat_26 = safe_num(ws_val.cell(r, 13).value or ws_val.cell(r, 11).value)
                strat_27 = safe_num(ws_val.cell(r, 14).value or ws_val.cell(r, 12).value)
                strat_28 = safe_num(ws_val.cell(r, 15).value or ws_val.cell(r, 13).value)
                strat_29 = safe_num(ws_val.cell(r, 16).value or ws_val.cell(r, 14).value)
                strat_30 = safe_num(ws_val.cell(r, 17).value or ws_val.cell(r, 15).value)

                note = str(ws_val.cell(r, 18).value or ws_val.cell(r, 16).value or ws_val.cell(r, 7).value or '').strip()

                if name and not (name.endswith(':') and tot == 0 and not is_total):
                    clean_name = name.lstrip('-').strip()
                    current_table['rows'].append({
                        'stt': stt,
                        'name': clean_name if not is_total else 'Tổng',
                        'raw_name': name,
                        'is_total': is_total,
                        'is_sub': is_sub,
                        'tong': tot,
                        'y2027': y27,
                        'y2028': y28,
                        'strat_tot': strat_tot,
                        'strat_26': strat_26,
                        'strat_27': strat_27,
                        'strat_28': strat_28,
                        'strat_29': strat_29,
                        'strat_30': strat_30,
                        'note': note
                    })

    # Cập nhật lại hàng tổng nếu các dòng con có dữ liệu
    for t in tables:
        tot_row = next((r for r in t['rows'] if r['is_total']), None)
        if tot_row and tot_row['tong'] == 0:
            sub_27 = sum(r['y2027'] for r in t['rows'] if not r['is_total'] and not r['is_sub'])
            sub_28 = sum(r['y2028'] for r in t['rows'] if not r['is_total'] and not r['is_sub'])
            if sub_27 > 0 or sub_28 > 0:
                tot_row['y2027'] = sub_27
                tot_row['y2028'] = sub_28
                tot_row['tong'] = sub_27 + sub_28

    return [t for t in tables if len(t['rows']) > 0]

def parse_sheet_dv(wb_val, wb_form, sheet_name, matrix_27, matrix_28, dv_row_map):
    if sheet_name not in wb_val.sheetnames:
        return []

    ws_val = wb_val[sheet_name]
    ws_form = wb_form[sheet_name]

    tables = []
    current_table = None

    for r in range(1, ws_val.max_row + 1):
        b_val = ws_val.cell(r, 2).value
        c_val = ws_val.cell(r, 3).value
        d_val = ws_val.cell(r, 4).value
        e_val = ws_val.cell(r, 5).value
        f_val = ws_val.cell(r, 6).value
        g_val = ws_val.cell(r, 7).value

        b_form = ws_form.cell(r, 2).value
        c_form = ws_form.cell(r, 3).value
        d_form = ws_form.cell(r, 4).value
        e_form = ws_form.cell(r, 5).value
        f_form = ws_form.cell(r, 6).value
        g_form = ws_form.cell(r, 7).value

        b_str = str(b_val or '').strip()
        c_str = str(c_val or '').strip()
        d_str = str(d_val or '').strip()

        is_sec = False
        for prefix in ['I.', 'II.', 'III.', 'IV.', 'V.', 'VI.', 'VII.', 'VIII.', 'IX.', 'X.', 'Đầu tư', 'Hiện đại']:
            if b_str.startswith(prefix) and not (b_str == 'STT' or c_str == 'STT' or c_str == 'Dịch vụ'):
                is_sec = True
                break

        if is_sec:
            current_table = {
                'id': f'dv_sec_{len(tables) + 1}',
                'title': b_str,
                'rows': []
            }
            tables.append(current_table)
            continue

        if b_str == 'STT' or c_str == 'STT' or c_str == 'Dịch vụ' or b_str.startswith('Phụ lục'):
            continue

        if current_table is not None:
            name = d_str if d_str else (c_str if c_str else b_str)
            service = c_str
            if name or e_val is not None or f_val is not None or g_val is not None:
                stt = b_str if (b_str and b_str != name and b_str != c_str and b_str != d_str) else ''
                is_total = (name.lower() == 'tổng' or d_str.lower() == 'tổng' or stt.lower() == 'tổng')
                is_sub = (stt == '-' or name.startswith('-') or d_str.startswith('-'))

                y27 = resolve_cell_formula(f_form, f_val, matrix_27, matrix_28, dv_row_map)
                y28 = resolve_cell_formula(g_form, g_val, matrix_27, matrix_28, dv_row_map)
                tot = resolve_cell_formula(e_form, e_val, matrix_27, matrix_28, dv_row_map)
                if tot == 0.0 and (y27 > 0 or y28 > 0):
                    tot = y27 + y28

                strat_tot = safe_num(ws_val.cell(r, 11).value)
                strat_26 = safe_num(ws_val.cell(r, 12).value)
                strat_27 = safe_num(ws_val.cell(r, 13).value)
                strat_28 = safe_num(ws_val.cell(r, 14).value)
                strat_29 = safe_num(ws_val.cell(r, 15).value)
                strat_30 = safe_num(ws_val.cell(r, 16).value)

                note = str(ws_val.cell(r, 17).value or ws_val.cell(r, 8).value or '').strip()

                if name and not (name.endswith(':') and tot == 0 and not is_total):
                    clean_name = name.lstrip('-').strip()
                    current_table['rows'].append({
                        'stt': stt,
                        'service': service,
                        'name': clean_name if not is_total else 'Tổng',
                        'raw_name': name,
                        'is_total': is_total,
                        'is_sub': is_sub,
                        'tong': tot,
                        'y2027': y27,
                        'y2028': y28,
                        'strat_tot': strat_tot,
                        'strat_26': strat_26,
                        'strat_27': strat_27,
                        'strat_28': strat_28,
                        'strat_29': strat_29,
                        'strat_30': strat_30,
                        'note': note
                    })

    # Cập nhật lại hàng tổng nếu các dòng con có dữ liệu
    for t in tables:
        tot_row = next((r for r in t['rows'] if r['is_total']), None)
        if tot_row and tot_row['tong'] == 0:
            sub_27 = sum(r['y2027'] for r in t['rows'] if not r['is_total'] and not r['is_sub'])
            sub_28 = sum(r['y2028'] for r in t['rows'] if not r['is_total'] and not r['is_sub'])
            if sub_27 > 0 or sub_28 > 0:
                tot_row['y2027'] = sub_27
                tot_row['y2028'] = sub_28
                tot_row['tong'] = sub_27 + sub_28

    return [t for t in tables if len(t['rows']) > 0]

def compute_overall_summary(wb_val, tables_mang, tables_dv):
    total_27 = 0.0
    total_28 = 0.0
    total_sum = 0.0
    total_strat = 0.0

    mang_summary = []
    for t in tables_mang:
        tot_row = next((r for r in t['rows'] if r['is_total']), None)
        if not tot_row and t['rows']:
            tot_val = sum(r['tong'] for r in t['rows'] if not r['is_sub'])
            y27_val = sum(r['y2027'] for r in t['rows'] if not r['is_sub'])
            y28_val = sum(r['y2028'] for r in t['rows'] if not r['is_sub'])
            strat_val = sum(r['strat_tot'] for r in t['rows'] if not r['is_sub'])
        else:
            tot_val = tot_row['tong'] if tot_row else 0.0
            y27_val = tot_row['y2027'] if tot_row else 0.0
            y28_val = tot_row['y2028'] if tot_row else 0.0
            strat_val = tot_row['strat_tot'] if tot_row else 0.0

        total_27 += y27_val
        total_28 += y28_val
        total_sum += tot_val
        total_strat += strat_val

        mang_summary.append({
            'name': t['title'],
            'y2027': y27_val,
            'y2028': y28_val,
            'tong': tot_val,
            'strat_tot': strat_val
        })

    for m in mang_summary:
        m['share'] = round((m['tong'] / total_sum * 100), 2) if total_sum > 0 else 0.0

    return {
        'total_investment': total_sum,
        'total_2027': total_27,
        'total_2028': total_28,
        'total_strategy_5y': total_strat,
        'share_2027': round((total_27 / total_sum * 100), 1) if total_sum > 0 else 0.0,
        'share_2028': round((total_28 / total_sum * 100), 1) if total_sum > 0 else 0.0,
        'by_mang': mang_summary
    }

def build_strategy_comparison(tables_mang, tables_dv):
    comparison = []
    for t in tables_mang:
        tot_row = next((r for r in t['rows'] if r['is_total']), None)
        if tot_row:
            qhdc_tot = tot_row['tong']
            qhdc_27 = tot_row['y2027']
            qhdc_28 = tot_row['y2028']
            strat_tot = tot_row['strat_tot']
            strat_26 = tot_row['strat_26']
            strat_27 = tot_row['strat_27']
            strat_28 = tot_row['strat_28']
            strat_29 = tot_row['strat_29']
            strat_30 = tot_row['strat_30']

            strat_27_28 = strat_27 + strat_28
            delta = qhdc_tot - strat_27_28
            rate = round((qhdc_tot / strat_27_28 * 100), 1) if strat_27_28 > 0 else 0.0

            status = 'Vừa khớp'
            status_color = 'green'
            if strat_27_28 > 0:
                if delta > 0.05:
                    status = f'Vượt CL +{delta:.2f} M$'
                    status_color = 'orange'
                elif delta < -0.05:
                    status = f'Dưới CL {delta:.2f} M$'
                    status_color = 'blue'

            comparison.append({
                'title': t['title'],
                'qhdc_2027': qhdc_27,
                'qhdc_2028': qhdc_28,
                'qhdc_total': qhdc_tot,
                'strat_2026': strat_26,
                'strat_2027': strat_27,
                'strat_2028': strat_28,
                'strat_2029': strat_29,
                'strat_2030': strat_30,
                'strat_27_28': strat_27_28,
                'strat_5y': strat_tot,
                'delta': delta,
                'rate': rate,
                'status': status,
                'status_color': status_color
            })

    return comparison
