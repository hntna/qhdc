import zipfile, io, re

def escape_xml(s):
    if s is None:
        return ''
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;')

def build_perfect_masterlist(template_zip_path, target_output_path, mang_data_dict):
    with zipfile.ZipFile(template_zip_path, 'r') as z_in:
        sheet1_xml = z_in.read('xl/worksheets/sheet1.xml').decode('utf-8')
        
        match_data = re.search(r'(<sheetData>)(.*?)(</sheetData>)', sheet1_xml, re.DOTALL)
        if not match_data:
            raise ValueError("Cannot find sheetData in sheet1.xml")
        
        orig_sheet_data = match_data.group(2)
        row_matches = re.findall(r'(<row\s+[^>]*?r="([0-9]+)"[^>]*?>.*?</row>)', orig_sheet_data, re.DOTALL)
        rows_1_to_9 = {}
        for full_row, r_num in row_matches:
            r_int = int(r_num)
            if r_int <= 9:
                rows_1_to_9[r_int] = full_row

        MANG_CONFIG = [
            {'code': 'VT', 'name': 'VÔ TUYẾN', 'tt': 'A'},
            {'code': 'ML', 'name': 'MẠNG LÕI', 'tt': 'B'},
            {'code': 'CDBR', 'name': 'CĐBR', 'tt': 'C'},
            {'code': 'CNTT', 'name': 'CNTT+VÍ', 'tt': 'D'},
            {'code': 'TD', 'name': 'TRUYỀN DẪN', 'tt': 'E'},
            {'code': 'CD', 'name': 'CƠ ĐIỆN', 'tt': 'F'},
            {'code': 'HT', 'name': 'TRIỂN KHAI HẠ TẦNG', 'tt': 'G'}
        ]

        current_row = 10
        section_mapping = {}
        new_rows_xml = []

        for mang in MANG_CONFIG:
            code = mang['code']
            items = mang_data_dict.get(code, [])
            sec_start = current_row

            sec_r = current_row
            current_row += 1

            COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
            item_row_map = {}
            child_rows_xml = []

            for idx, it in enumerate(items):
                r_num = current_row
                item_row_map[idx] = r_num
                current_row += 1

                cell_map = {}
                if it.get('tt'):
                    cell_map['A'] = f'<c r="A{r_num}" t="inlineStr"><is><t>{escape_xml(it["tt"])}</t></is></c>'
                if it.get('nd'):
                    cell_map['B'] = f'<c r="B{r_num}" t="inlineStr"><is><t>{escape_xml(it["nd"])}</t></is></c>'
                if it.get('dvt'):
                    cell_map['C'] = f'<c r="C{r_num}" t="inlineStr"><is><t>{escape_xml(it["dvt"])}</t></is></c>'
                if it.get('kl27') is not None:
                    cell_map['D'] = f'<c r="D{r_num}" s="323"><v>{it["kl27"]}</v></c>'
                if it.get('kl28') is not None:
                    cell_map['E'] = f'<c r="E{r_num}" s="323"><v>{it["kl28"]}</v></c>'
                if it.get('dg') is not None:
                    cell_map['F'] = f'<c r="F{r_num}" s="323"><v>{it["dg"]}</v></c>'
                
                if not it.get('isGroup'):
                    if it.get('kl27') is not None:
                        cell_map['G'] = f'<c r="G{r_num}" s="323"><f>F{r_num}*D{r_num}</f></c>'
                    if it.get('kl28') is not None:
                        cell_map['H'] = f'<c r="H{r_num}" s="323"><f>F{r_num}*E{r_num}</f></c>'
                
                if it.get('donViDT'):
                    cell_map['I'] = f'<c r="I{r_num}" t="inlineStr"><is><t>{escape_xml(it["donViDT"])}</t></is></c>'
                cell_map['J'] = f'<c r="J{r_num}" t="inlineStr"><is><t>{escape_xml(it.get("maMang") or code)}</t></is></c>'
                if it.get('maDV'):
                    cell_map['K'] = f'<c r="K{r_num}" t="inlineStr"><is><t>{escape_xml(it["maDV"])}</t></is></c>'
                if it.get('maLoai'):
                    cell_map['L'] = f'<c r="L{r_num}" t="inlineStr"><is><t>{escape_xml(it["maLoai"])}</t></is></c>'

                s_row = ' s="328"' if it.get('isGroup') else ''
                child_rows_xml.append((idx, r_num, s_row, cell_map, it))

            sec_end = current_row - 1
            section_mapping[code] = {'start': sec_start, 'end': sec_end}

            sec_cell_map = {
                'A': f'<c r="A{sec_r}" s="320" t="inlineStr"><is><t>{mang["tt"]}</t></is></c>',
                'B': f'<c r="B{sec_r}" s="321" t="inlineStr"><is><t>{mang["name"]}</t></is></c>',
            }
            if sec_end > sec_start:
                sec_cell_map['G'] = f'<c r="G{sec_r}" s="323"><f>SUBTOTAL(9,G{sec_start+1}:G{sec_end})</f></c>'
                sec_cell_map['H'] = f'<c r="H{sec_r}" s="323"><f>SUBTOTAL(9,H{sec_start+1}:H{sec_end})</f></c>'
            sec_cell_map['J'] = f'<c r="J{sec_r}" s="324" t="inlineStr"><is><t>{code}</t></is></c>'

            new_rows_xml.append(f'<row r="{sec_r}" spans="1:17" s="328" customFormat="1">{"".join(sec_cell_map[c] for c in COLS if c in sec_cell_map)}</row>')

            for idx, r_num, s_row, cell_map, it in child_rows_xml:
                if it.get('isGroup') and it.get('hasSubtotal') and it.get('subtotalStartRow') and it.get('subtotalEndRow'):
                    c_start = None
                    c_end = None
                    for k, k_it in enumerate(items):
                        o_r = k_it.get('origRow')
                        if o_r and o_r >= it['subtotalStartRow'] and o_r <= it['subtotalEndRow']:
                            if c_start is None:
                                c_start = item_row_map[k]
                            c_end = item_row_map[k]
                    if c_start and c_end:
                        cell_map['G'] = f'<c r="G{r_num}" s="323"><f>SUBTOTAL(9,G{c_start}:G{c_end})</f></c>'
                        cell_map['H'] = f'<c r="H{r_num}" s="323"><f>SUBTOTAL(9,H{c_start}:H{c_end})</f></c>'

                new_rows_xml.append(f'<row r="{r_num}" spans="1:17"{s_row}>{"".join(cell_map[c] for c in COLS if c in cell_map)}</row>')

        total_rows = current_row - 1

        cd_end = section_mapping.get('CD', {}).get('end')
        td_end = section_mapping.get('TD', {}).get('end')
        last_before_ht = cd_end or td_end or total_rows
        
        def update_row_subtotals(r_int, start_r, end_r):
            row_str = rows_1_to_9[r_int]
            row_str = re.sub(
                r'(<c\s+[^>]*?r="G' + str(r_int) + r'"[^>]*?>)(.*?)(</c>)',
                r'\g<1><f>SUBTOTAL(9,G' + str(start_r) + ':G' + str(end_r) + r')</f>\g<3>',
                row_str, flags=re.DOTALL
            )
            row_str = re.sub(
                r'(<c\s+[^>]*?r="H' + str(r_int) + r'"[^>]*?>)(.*?)(</c>)',
                r'\g<1><f>SUBTOTAL(9,H' + str(start_r) + ':H' + str(end_r) + r')</f>\g<3>',
                row_str, flags=re.DOTALL
            )
            rows_1_to_9[r_int] = row_str

        update_row_subtotals(9, 10, last_before_ht)
        update_row_subtotals(8, 9, total_rows)
        update_row_subtotals(7, 8, total_rows)

        all_rows = [rows_1_to_9[i] for i in range(1, 10) if i in rows_1_to_9]
        all_rows.extend(new_rows_xml)
        new_sheet_data = f'<sheetData>{"".join(all_rows)}</sheetData>'

        new_sheet1_xml = sheet1_xml[:match_data.start()] + new_sheet_data + sheet1_xml[match_data.end():]
        new_sheet1_xml = re.sub(r'<dimension\s+ref="[^"]*"', f'<dimension ref="A1:Q{total_rows}"', new_sheet1_xml)
        new_sheet1_xml = re.sub(r'<autoFilter\s+ref="[^"]*"', f'<autoFilter ref="A7:Q{total_rows}"', new_sheet1_xml)
        new_sheet1_xml = re.sub(r'<dataValidations\s+count="[^"]*">.*?</dataValidations>', '', new_sheet1_xml, flags=re.DOTALL)

        content_types_xml = z_in.read('[Content_Types].xml').decode('utf-8')
        content_types_xml = re.sub(r'<Override\s+[^>]*?PartName="/xl/calcChain\.xml"[^>]*?/>', '', content_types_xml)

        wb_rels_xml = z_in.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        wb_rels_xml = re.sub(r'<Relationship\s+[^>]*?Target="calcChain\.xml"[^>]*?/>', '', wb_rels_xml)

        wb_xml = z_in.read('xl/workbook.xml').decode('utf-8')
        if '<calcPr' in wb_xml:
            wb_xml = re.sub(r'<calcPr([^>]*?)/?>', r'<calcPr\1 fullCalcOnLoad="1" forceFullCalculation="1" calcMode="auto"/>', wb_xml)
        else:
            wb_xml = wb_xml.replace('</workbook>', '<calcPr fullCalcOnLoad="1" forceFullCalculation="1" calcMode="auto"/></workbook>')

        with zipfile.ZipFile(target_output_path, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename == 'xl/worksheets/sheet1.xml':
                    z_out.writestr(item, new_sheet1_xml.encode('utf-8'))
                elif item.filename == '[Content_Types].xml':
                    z_out.writestr(item, content_types_xml.encode('utf-8'))
                elif item.filename == 'xl/_rels/workbook.xml.rels':
                    z_out.writestr(item, wb_rels_xml.encode('utf-8'))
                elif item.filename == 'xl/workbook.xml':
                    z_out.writestr(item, wb_xml.encode('utf-8'))
                elif item.filename == 'xl/calcChain.xml':
                    continue
                else:
                    z_out.writestr(item, z_in.read(item.filename))

    print(f"Successfully generated {target_output_path} with {total_rows} rows!")
