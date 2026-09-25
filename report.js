/**
 * report.js - Báo cáo Tổng Hợp & So Sánh Chiến Lược QHĐC 2027-2028
 * Tự động tính toán số liệu thời gian thực từ dữ liệu các mảng theo công thức chuẩn của file mẫu.
 */

// Schema cấu trúc bảng & công thức trích xuất chuẩn từ Masterlist Mau
const REPORT_TEMPLATE_SCHEMA = {"tables_mang": [{"title": "I. Vô tuyến", "rows": [{"row_idx": 5, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUBTOTAL(9,E6:E12)", "form_f": "=SUBTOTAL(9,F6:F12)", "strat_tot": 38.68827488757404, "strat_26": 9.279389515226024, "strat_27": 11.107721492780065, "strat_28": 12.109146876129818, "strat_29": 3.916827626473486, "strat_30": 2.275189376964648}, {"row_idx": 6, "stt": "1", "name": "Mạng 5G", "is_total": false, "is_sub": false, "form_e": "=E7", "form_f": "=F7", "strat_tot": 20.853331875, "strat_26": 4.759689375, "strat_27": 7.04104875, "strat_28": 6.736123125, "strat_29": 1.35579375, "strat_30": 0.960676875}, {"row_idx": 7, "stt": "-", "name": "Vùng phủ", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!D7", "form_f": "='TH theo DV'!N7", "strat_tot": 20.853331875, "strat_26": 4.759689375, "strat_27": 7.04104875, "strat_28": 6.736123125, "strat_29": 1.35579375, "strat_30": 0.960676875}, {"row_idx": 8, "stt": "2", "name": "Mạng 4G", "is_total": false, "is_sub": false, "form_e": "=SUBTOTAL(9,E9:E10)", "form_f": "=SUBTOTAL(9,F9:F10)", "strat_tot": 15.73726851257404, "strat_26": 4.193305640226022, "strat_27": 3.4351027427800656, "strat_28": 5.118953751129819, "strat_29": 1.9294638764734862, "strat_30": 1.060442501964648}, {"row_idx": 9, "stt": "-", "name": "Vùng phủ", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!D13", "form_f": "='TH theo DV'!N13", "strat_tot": 4.225653752946972, "strat_26": 1.2676070024558102, "strat_27": 0.49012875, "strat_28": 0.460164250491162, "strat_29": 1.65847425, "strat_30": 0.3492795}, {"row_idx": 10, "stt": "-", "name": "Dung lượng", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!D14", "form_f": "='TH theo DV'!N14", "strat_tot": 11.51161475962707, "strat_26": 2.9256986377702123, "strat_27": 2.9449739927800653, "strat_28": 4.658789500638657, "strat_29": 0.27098962647348607, "strat_30": 0.711163001964648}, {"row_idx": 11, "stt": "3", "name": "Hiện đại hóa", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!D28", "form_f": "='TH theo DV'!N28", "strat_tot": 1.3273245, "strat_26": 0.1723245, "strat_27": 0.4775, "strat_28": 0.1, "strat_29": 0.4775, "strat_30": 0.1}, {"row_idx": 12, "stt": "4", "name": "UCTT", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!D27", "form_f": "='TH theo DV'!N27", "strat_tot": 0.7703500000000001, "strat_26": 0.15407, "strat_27": 0.15407, "strat_28": 0.15407, "strat_29": 0.15407, "strat_30": 0.15407}]}, {"title": "II. BRCĐ", "rows": [{"row_idx": 17, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUBTOTAL(9,E18:E22)", "form_f": "=SUBTOTAL(9,F18:F22)", "strat_tot": 3.079619, "strat_26": 2.091494, "strat_27": 1.638903, "strat_28": 1.776286, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 18, "stt": "1", "name": "Vùng phủ", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!F16", "form_f": "='TH theo DV'!P16", "strat_tot": 1.162536, "strat_26": 1.252383, "strat_27": 1.420917, "strat_28": 1.556171, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 19, "stt": "2", "name": "Dung lượng Internet", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!F17", "form_f": "='TH theo DV'!P17", "strat_tot": 0.86448, "strat_26": 0.64024, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 20, "stt": "3", "name": "Hiện đại hóa", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!F28", "form_f": "='TH theo DV'!P28", "strat_tot": 0.07, "strat_26": 0.07, "strat_27": 0.07, "strat_28": 0.07, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 21, "stt": "4", "name": "Kiên cố, củng cố", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!F29", "form_f": "='TH theo DV'!P29", "strat_tot": 0.86448, "strat_26": 0.0, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 22, "stt": "5", "name": "Vật tư ƯCTT", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!F27", "form_f": "='TH theo DV'!P27", "strat_tot": 0.118123, "strat_26": 0.128871, "strat_27": 0.147986, "strat_28": 0.150115, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "III. Mạng lõi", "rows": [{"row_idx": 27, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUBTOTAL(9,E28:E32)", "form_f": "=SUBTOTAL(9,F28:F32)", "strat_tot": 3.79, "strat_26": 2.61, "strat_27": 1.37, "strat_28": 1.6400000000000001, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 28, "stt": "1", "name": "Đầu tư dung lượng 5G", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!E7", "form_f": "='TH theo DV'!O7", "strat_tot": 0.23, "strat_26": 0.26, "strat_27": 0.34, "strat_28": 0.31, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 29, "stt": "2", "name": "Đầu tư dung lượng 4G", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!E14", "form_f": "='TH theo DV'!O14", "strat_tot": 2.14, "strat_26": 1.19, "strat_27": 1.03, "strat_28": 1.33, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 30, "stt": "3", "name": "Đầu tư hiện đại hóa mạng lưới", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!E28", "form_f": "='TH theo DV'!O28", "strat_tot": 1.42, "strat_26": 1.16, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 31, "stt": "4", "name": "Kiên cố, củng cố", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!E29", "form_f": "='TH theo DV'!O29", "strat_tot": 0.0, "strat_26": 0.0, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 32, "stt": "5", "name": "UCTT", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!E30", "form_f": "='TH theo DV'!O30", "strat_tot": 0.0, "strat_26": 0.0, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "IV. Công nghệ thông tin", "rows": [{"row_idx": 37, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUBTOTAL(9,E38:E43)", "form_f": "=SUBTOTAL(9,F38:F43)", "strat_tot": 3.5873586153121324, "strat_26": 6.459445924303102, "strat_27": 6.058848615312132, "strat_28": 3.737938324303102, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 38, "stt": "A", "name": "Chi phí cho phần CNTT", "is_total": false, "is_sub": false, "form_e": "=SUBTOTAL(9,E39:E42)", "form_f": "=SUBTOTAL(9,F39:F42)", "strat_tot": 2.9373586153121325, "strat_26": 5.809445924303102, "strat_27": 5.408848615312132, "strat_28": 3.087938324303102, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 39, "stt": "1", "name": "Tài nguyên cho đầu tư phát triển", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!G23+'TH theo DV'!G24", "form_f": "='TH theo DV'!Q23+'TH theo DV'!Q24", "strat_tot": 2.239422615312132, "strat_26": 2.1335019243031015, "strat_27": 2.255344615312132, "strat_28": 2.1844523243031015, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 40, "stt": "2", "name": "Tài nguyên cho kiên cố củng cố", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!G29", "form_f": "='TH theo DV'!Q29", "strat_tot": 0.083352, "strat_26": 0.66136, "strat_27": 2.53892, "strat_28": 0.13892, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 41, "stt": "3", "name": "Tài nguyên cho hiện đại hoá", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!G28", "form_f": "='TH theo DV'!Q28", "strat_tot": 0.514584, "strat_26": 2.914584, "strat_27": 0.514584, "strat_28": 0.514584, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 42, "stt": "4", "name": "Tài nguyên cho UCTT", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!G27", "form_f": "='TH theo DV'!Q27", "strat_tot": 0.1, "strat_26": 0.1, "strat_27": 0.1, "strat_28": 0.249982, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 43, "stt": "B", "name": "Chi phí cho hệ thống Ví điện tử", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!G31", "form_f": "='TH theo DV'!Q31", "strat_tot": 0.65, "strat_26": 0.65, "strat_27": 0.65, "strat_28": 0.65, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "V. Truyền dẫn quang", "rows": [{"row_idx": 48, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUBTOTAL(9,E49:E57)", "form_f": "=SUBTOTAL(9,F49:F57)", "strat_tot": 5.5319648599999995, "strat_26": 4.25533434, "strat_27": 2.27352152, "strat_28": 1.8653537999999998, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 49, "stt": "1", "name": "Mạng 5G", "is_total": false, "is_sub": false, "form_e": "=E50", "form_f": "=F50", "strat_tot": 1.421762518010634, "strat_26": 1.0507716009021422, "strat_27": 0.34650542317214467, "strat_28": 0.29930565872357306, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 50, "stt": "-", "name": "Vùng phủ", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!H7", "form_f": "='TH theo DV'!R7", "strat_tot": 0.3184852, "strat_26": 0.28023068, "strat_27": 0.09152207999999999, "strat_28": 0.0791085, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 51, "stt": "2", "name": "Mạng 4G/3G/2G", "is_total": false, "is_sub": false, "form_e": "", "form_f": "", "strat_tot": 0.2733749010437518, "strat_26": 0.3038983664115197, "strat_27": 0.18076838287681612, "strat_28": 0.17176819613290822, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 52, "stt": "-", "name": "Vùng phủ", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!H13", "form_f": "='TH theo DV'!R13", "strat_tot": 0.1342339, "strat_26": 0.1808579, "strat_27": 0.13275367999999999, "strat_28": 0.14199954, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 53, "stt": "-", "name": "Dung lượng", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!H14", "form_f": "='TH theo DV'!R14", "strat_tot": 0.1391410010437518, "strat_26": 0.12304046641151968, "strat_27": 0.04801470287681614, "strat_28": 0.02976865613290822, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 54, "stt": "3", "name": "Đáp ứng mạng CĐBR/TH", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!H15", "form_f": "='TH theo DV'!R15", "strat_tot": 1.1242816809456144, "strat_26": 0.9131186126863382, "strat_27": 0.3437019539510392, "strat_28": 0.31673418514351875, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 55, "stt": "4", "name": "Hiện đại hóa", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!H28", "form_f": "='TH theo DV'!R28", "strat_tot": 2.025, "strat_26": 1.3, "strat_27": 0.715, "strat_28": 0.39, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 56, "stt": "5", "name": "Củng cố bền vững", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!H29", "form_f": "='TH theo DV'!R29", "strat_tot": 0.35718738999999994, "strat_26": 0.35718738999999994, "strat_27": 0.35718738999999994, "strat_28": 0.35718738999999994, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 57, "stt": "6", "name": "UCTT", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!H27", "form_f": "='TH theo DV'!R27", "strat_tot": 0.33035836999999996, "strat_26": 0.33035836999999996, "strat_27": 0.33035836999999996, "strat_28": 0.33035836999999996, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "IV. Truyền dẫn IP", "rows": [{"row_idx": 62, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F62:G62)", "form_f": "=SUM(G62:H62)", "strat_tot": 3.613371, "strat_26": 2.164133, "strat_27": 1.8335649999999997, "strat_28": 1.677052, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 63, "stt": "1", "name": "Mạng 5G", "is_total": false, "is_sub": false, "form_e": "=SUM(F63:G63)", "form_f": "=SUM(G63:H63)", "strat_tot": 0.6106931434708015, "strat_26": 0.6807498625074259, "strat_27": 0.5538867027600978, "strat_28": 0.42180705621154235, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 64, "stt": "-", "name": "Vùng phủ 5G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!I7", "form_f": "='TH theo DV'!S7", "strat_tot": 0.4153059295154185, "strat_26": 0.3670992272727273, "strat_27": 0.07386586666666667, "strat_28": 0.06522025333333334, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 65, "stt": "2", "name": "Mạng 4G", "is_total": false, "is_sub": false, "form_e": "=SUM(F65:G65)", "form_f": "=SUM(G65:H65)", "strat_tot": 0.26647526195790056, "strat_26": 0.19843095591025497, "strat_27": 0.08441813333333333, "strat_28": 0.09236374666666668, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 66, "stt": "-", "name": "Vùng phủ 4G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!I9", "form_f": "='TH theo DV'!S9", "strat_tot": 0.09457807048458149, "strat_26": 0.12876677272727274, "strat_27": 0.08441813333333333, "strat_28": 0.09236374666666668, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 67, "stt": "-", "name": "Dung lượng 4G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!I14", "form_f": "='TH theo DV'!S14", "strat_tot": 0.17189719147331908, "strat_26": 0.06966418318298223, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 68, "stt": "3", "name": "Dung lượng BRCĐ", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!I15", "form_f": "='TH theo DV'!S15", "strat_tot": 0.9628335945712981, "strat_26": 0.8802871815823191, "strat_27": 0.9433411639065686, "strat_28": 0.894439197121791, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 69, "stt": "4", "name": "ƯCTT", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!I27", "form_f": "='TH theo DV'!S27", "strat_tot": 0.210969, "strat_26": 0.404665, "strat_27": 0.251919, "strat_28": 0.268442, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 70, "stt": "5", "name": "Tổng trạm mới", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!I21", "form_f": "='TH theo DV'!S21", "strat_tot": 1.5624, "strat_26": 0.0, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "V. Cơ điện tổng trạm", "rows": [{"row_idx": 75, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "", "form_f": "", "strat_tot": 3.2192156862745103, "strat_26": 1.6882352941176473, "strat_27": 0.7000000000000002, "strat_28": 0.7725490196078431, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 76, "stt": "1", "name": "Tổng trạm mới", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!J21", "form_f": "='TH theo DV'!T21", "strat_tot": 2.91921568627451, "strat_26": 1.3882352941176472, "strat_27": 0.4000000000000001, "strat_28": 0.5725490196078431, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 77, "stt": "2", "name": "Hiện đại hóa mạng lưới", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!J28", "form_f": "='TH theo DV'!T28", "strat_tot": 0.1, "strat_26": 0.1, "strat_27": 0.1, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 78, "stt": "3", "name": "Kiên cố, củng cố", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!J29", "form_f": "='TH theo DV'!T29", "strat_tot": 0.2, "strat_26": 0.2, "strat_27": 0.2, "strat_28": 0.2, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "V. Cơ điện BTS", "rows": [{"row_idx": 83, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUBTOTAL(9,E84:E91)", "form_f": "=SUBTOTAL(9,F84:F91)", "strat_tot": 2.2687779999999997, "strat_26": 2.105617, "strat_27": 1.859979, "strat_28": 1.830026, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 84, "stt": "1", "name": "Mạng 5G", "is_total": false, "is_sub": false, "form_e": "=E85", "form_f": "=F85", "strat_tot": 0.2718, "strat_26": 0.17044, "strat_27": 0.09051, "strat_28": 0.0425, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 85, "stt": "-", "name": "Vùng phủ 5G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!J7", "form_f": "='TH theo DV'!T7", "strat_tot": 0.2718, "strat_26": 0.17044, "strat_27": 0.09051, "strat_28": 0.0425, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 86, "stt": "2", "name": "Mạng 4G", "is_total": false, "is_sub": false, "form_e": "=SUBTOTAL(9,E87:E88)", "form_f": "=SUBTOTAL(9,F87:F88)", "strat_tot": 0.582286, "strat_26": 0.525872, "strat_27": 0.354464, "strat_28": 0.366821, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 87, "stt": "-", "name": "Vùng phủ 4G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!J13", "form_f": "='TH theo DV'!T13", "strat_tot": 0.32058, "strat_26": 0.442972, "strat_27": 0.3052, "strat_28": 0.349136, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 88, "stt": "-", "name": "Dung lượng 4G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!J14", "form_f": "='TH theo DV'!T14", "strat_tot": 0.261706, "strat_26": 0.0829, "strat_27": 0.049264, "strat_28": 0.017685, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 89, "stt": "3", "name": "Dung lượng BRCĐ", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!J15", "form_f": "='TH theo DV'!T15", "strat_tot": 0.024052, "strat_26": 0.018665, "strat_27": 0.024365, "strat_28": 0.030065, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 90, "stt": "4", "name": "Hiện đại hóa mạng lưới", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!J28", "form_f": "='TH theo DV'!T28", "strat_tot": 0.12, "strat_26": 0.12, "strat_27": 0.12, "strat_28": 0.12, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 91, "stt": "5", "name": "Kiên cố, củng cố", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!J29", "form_f": "='TH theo DV'!T29", "strat_tot": 1.27064, "strat_26": 1.27064, "strat_27": 1.27064, "strat_28": 1.27064, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "VI. Hạ tầng", "rows": [{"row_idx": 96, "stt": "", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUBTOTAL(9,E97:E104)", "form_f": "=SUBTOTAL(9,F97:F104)", "strat_tot": 2.637626829970733, "strat_26": 2.56156146589012, "strat_27": 2.5018883719725364, "strat_28": 2.491370246804815, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 97, "stt": "1", "name": "Mạng 5G", "is_total": false, "is_sub": false, "form_e": "=E98", "form_f": "=F98", "strat_tot": 0.123508, "strat_26": 0.108671, "strat_27": 0.023659, "strat_28": 0.020451, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 98, "stt": "-", "name": "Vùng phủ 5G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!K7", "form_f": "='TH theo DV'!U7", "strat_tot": 0.123508, "strat_26": 0.108671, "strat_27": 0.023659, "strat_28": 0.020451, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 99, "stt": "2", "name": "Mạng 4G", "is_total": false, "is_sub": false, "form_e": "=SUBTOTAL(9,E100:E101)", "form_f": "=SUBTOTAL(9,F100:F101)", "strat_tot": 0.5849941936070968, "strat_26": 0.5237658295264838, "strat_27": 0.5491047356089003, "strat_28": 0.5417946104411789, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 100, "stt": "-", "name": "Vùng phủ 4G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!K13", "form_f": "='TH theo DV'!U13", "strat_tot": 0.36317311207878045, "strat_26": 0.47943501366402536, "strat_27": 0.39243043487047646, "strat_28": 0.46136059823682396, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 101, "stt": "-", "name": "Dung lượng 4G", "is_total": false, "is_sub": true, "form_e": "='TH theo DV'!K14", "form_f": "='TH theo DV'!U14", "strat_tot": 0.2218210815283163, "strat_26": 0.04433081586245837, "strat_27": 0.15667430073842384, "strat_28": 0.08043401220435498, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 102, "stt": "4", "name": "Hiện đại hóa mạng lưới", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!K28", "form_f": "='TH theo DV'!U28", "strat_tot": 0.0, "strat_26": 0.0, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 103, "stt": "5", "name": "Kiên cố, củng cố", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!K29", "form_f": "='TH theo DV'!U29", "strat_tot": 1.9291246363636363, "strat_26": 1.9291246363636363, "strat_27": 1.9291246363636363, "strat_28": 1.9291246363636363, "strat_29": 0.0, "strat_30": 0.0}, {"row_idx": 104, "stt": "6", "name": "UCTT", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!K27", "form_f": "='TH theo DV'!U27", "strat_tot": 0.0, "strat_26": 0.0, "strat_27": 0.0, "strat_28": 0.0, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "VII. Phần mềm VHKT", "rows": [{"row_idx": 109, "stt": "1", "name": "Phần mềm VHKT", "is_total": false, "is_sub": false, "form_e": "", "form_f": "", "strat_tot": 0.188, "strat_26": 0.142, "strat_27": 0.234, "strat_28": 0.142, "strat_29": 0.0, "strat_30": 0.0}]}, {"title": "VIII. An toàn thông tin", "rows": [{"row_idx": 114, "stt": "1", "name": "Triển khai giải pháp ATTT", "is_total": false, "is_sub": false, "form_e": "='TH theo DV'!G25", "form_f": "='TH theo DV'!Q25", "strat_tot": 0.4441666666666667, "strat_26": 0.5090378787878789, "strat_27": 0.5390325757575758, "strat_28": 0.6366403787878789, "strat_29": 0.0, "strat_30": 0.0}]}], "tables_dv": [{"title": "I. Dịch vụ di động", "rows": [{"row_idx": 4, "stt": "Dịch vụ (4G+5G):", "service": "", "mang": "", "name": "Dịch vụ (4G+5G):", "is_total": false, "is_sub": false, "form_e": "", "form_f": "", "form_g": "", "strat_tot": 0.0}, {"row_idx": 7, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F7:G7)", "form_f": "", "form_g": "", "strat_tot": 16.42523740116671}, {"row_idx": 8, "stt": "1", "service": "Mạng di động", "mang": "Mạng 5G", "name": "Mạng 5G", "is_total": false, "is_sub": false, "form_e": "=SUM(F8:G8)", "form_f": "", "form_g": "", "strat_tot": 7.1214531570534545}, {"row_idx": 9, "stt": "2", "service": "", "mang": "Vùng phủ", "name": "Vùng phủ", "is_total": false, "is_sub": false, "form_e": "=SUM(F9:G9)", "form_f": "='TH theo DV'!L7", "form_g": "='TH theo DV'!V7", "strat_tot": 5.587957640882353}, {"row_idx": 10, "stt": "4", "service": "", "mang": "Mạng 4G", "name": "Mạng 4G", "is_total": false, "is_sub": false, "form_e": "=SUM(F10:G10)", "form_f": "", "form_g": "", "strat_tot": 9.303784244113256}, {"row_idx": 11, "stt": "5", "service": "", "mang": "Vùng phủ", "name": "Vùng phủ", "is_total": false, "is_sub": false, "form_e": "=SUM(F11:G11)", "form_f": "='TH theo DV'!L13", "form_g": "='TH theo DV'!V13", "strat_tot": 3.3427795287498023}, {"row_idx": 12, "stt": "6", "service": "", "mang": "Dung lượng", "name": "Dung lượng", "is_total": false, "is_sub": false, "form_e": "=SUM(F12:G12)", "form_f": "='TH theo DV'!L14", "form_g": "='TH theo DV'!V14", "strat_tot": 5.961004715363453}]}, {"title": "Đầu tư Vùng phủ 4G", "rows": [{"row_idx": 17, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F17:G17)", "form_f": "", "form_g": "", "strat_tot": 3.3427795287498023}, {"row_idx": 18, "stt": "1", "service": "Vùng phủ 4G", "mang": "Vô tuyến", "name": "Vô tuyến", "is_total": false, "is_sub": false, "form_e": "=SUM(F18:G18)", "form_f": "='TH theo DV'!D13", "form_g": "='TH theo DV'!N13", "strat_tot": 1.2676070024558102}, {"row_idx": 19, "stt": "2", "service": "", "mang": "Truyền dẫn", "name": "Truyền dẫn", "is_total": false, "is_sub": false, "form_e": "=SUM(F19:G19)", "form_f": "='TH theo DV'!H13", "form_g": "='TH theo DV'!R13", "strat_tot": 0.21275368000000003}, {"row_idx": 20, "stt": "3", "service": "", "mang": "IP", "name": "IP", "is_total": false, "is_sub": false, "form_e": "=SUM(F20:G20)", "form_f": "", "form_g": "", "strat_tot": 0.02881229411764706}, {"row_idx": 21, "stt": "4", "service": "", "mang": "Cơ điện", "name": "Cơ điện", "is_total": false, "is_sub": false, "form_e": "=SUM(F21:G21)", "form_f": "='TH theo DV'!J13", "form_g": "='TH theo DV'!T13", "strat_tot": 0.541232}, {"row_idx": 22, "stt": "5", "service": "", "mang": "Hạ tầng", "name": "Hạ tầng", "is_total": false, "is_sub": false, "form_e": "=SUM(F22:G22)", "form_f": "='TH theo DV'!K13", "form_g": "='TH theo DV'!U13", "strat_tot": 1.2923745521763452}]}, {"title": "Đầu tư Dung lượng 4G", "rows": [{"row_idx": 27, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F27:G27)", "form_f": "", "form_g": "", "strat_tot": 5.961004715363453}, {"row_idx": 28, "stt": "1", "service": "Dung lượng 4G", "mang": "Vô tuyến", "name": "Vô tuyến", "is_total": false, "is_sub": false, "form_e": "=SUM(F28:G28)", "form_f": "='TH theo DV'!D14", "form_g": "='TH theo DV'!N14", "strat_tot": 2.9256986377702123}, {"row_idx": 29, "stt": "2", "service": "", "mang": "Mạng lõi", "name": "Mạng lõi", "is_total": false, "is_sub": false, "form_e": "=SUM(F29:G29)", "form_f": "='TH theo DV'!E14", "form_g": "='TH theo DV'!O14", "strat_tot": 0.89}, {"row_idx": 30, "stt": "3", "service": "", "mang": "Truyền dẫn", "name": "Truyền dẫn", "is_total": false, "is_sub": false, "form_e": "=SUM(F30:G30)", "form_f": "='TH theo DV'!H14", "form_g": "='TH theo DV'!R14", "strat_tot": 0.17916343005965427}, {"row_idx": 31, "stt": "4", "service": "", "mang": "IP", "name": "IP", "is_total": false, "is_sub": false, "form_e": "=SUM(F31:G31)", "form_f": "", "form_g": "", "strat_tot": 1.5496218712314551}, {"row_idx": 32, "stt": "5", "service": "", "mang": "Cơ điện", "name": "Cơ điện", "is_total": false, "is_sub": false, "form_e": "=SUM(F32:G32)", "form_f": "='TH theo DV'!J14", "form_g": "='TH theo DV'!T14", "strat_tot": 0.165197}, {"row_idx": 33, "stt": "6", "service": "", "mang": "Hạ tầng", "name": "Hạ tầng", "is_total": false, "is_sub": false, "form_e": "=SUM(F33:G33)", "form_f": "='TH theo DV'!K14", "form_g": "='TH theo DV'!U14", "strat_tot": 0.2513237763021316}]}, {"title": "Đầu tư Vùng phủ 5G", "rows": [{"row_idx": 38, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F38:G38)", "form_f": "", "form_g": "", "strat_tot": 5.587957640882353}, {"row_idx": 39, "stt": "1", "service": "Vùng phủ 5G", "mang": "Vô tuyến", "name": "Vô tuyến", "is_total": false, "is_sub": false, "form_e": "=SUM(F39:G39)", "form_f": "='TH theo DV'!D7", "form_g": "='TH theo DV'!N7", "strat_tot": 4.759689375}, {"row_idx": 40, "stt": "3", "service": "", "mang": "Truyền dẫn", "name": "Truyền dẫn", "is_total": false, "is_sub": false, "form_e": "=SUM(F40:G40)", "form_f": "='TH theo DV'!H7", "form_g": "='TH theo DV'!R7", "strat_tot": 0.21404855999999997}, {"row_idx": 41, "stt": "2", "service": "", "mang": "IP", "name": "IP", "is_total": false, "is_sub": false, "form_e": "=SUM(F41:G41)", "form_f": "", "form_g": "", "strat_tot": 0.3021427058823529}, {"row_idx": 42, "stt": "4", "service": "", "mang": "Cơ điện", "name": "Cơ điện", "is_total": false, "is_sub": false, "form_e": "=SUM(F42:G42)", "form_f": "='TH theo DV'!J7", "form_g": "='TH theo DV'!T7", "strat_tot": 0.22907}, {"row_idx": 43, "stt": "5", "service": "", "mang": "Hạ tầng", "name": "Hạ tầng", "is_total": false, "is_sub": false, "form_e": "=SUM(F43:G43)", "form_f": "='TH theo DV'!K7", "form_g": "='TH theo DV'!U7", "strat_tot": 0.083007}]}, {"title": "Đầu tư Dung lượng 5G", "rows": [{"row_idx": 48, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F48:G48)", "form_f": "", "form_g": "", "strat_tot": 1.5334955161711015}, {"row_idx": 49, "stt": "1", "service": "Dung lượng 5G", "mang": "Vô tuyến", "name": "Vô tuyến", "is_total": false, "is_sub": false, "form_e": "=SUM(F49:G49)", "form_f": "", "form_g": "", "strat_tot": 0.0}, {"row_idx": 50, "stt": "2", "service": "", "mang": "Mạng lõi", "name": "Mạng lõi", "is_total": false, "is_sub": false, "form_e": "=SUM(F50:G50)", "form_f": "", "form_g": "", "strat_tot": 0.26}, {"row_idx": 51, "stt": "3", "service": "", "mang": "Truyền dẫn", "name": "Truyền dẫn", "is_total": false, "is_sub": false, "form_e": "=SUM(F51:G51)", "form_f": "", "form_g": "", "strat_tot": 0.5651843219547452}, {"row_idx": 52, "stt": "4", "service": "", "mang": "IP", "name": "IP", "is_total": false, "is_sub": false, "form_e": "=SUM(F52:G52)", "form_f": "", "form_g": "", "strat_tot": 0.7083111942163562}, {"row_idx": 53, "stt": "5", "service": "", "mang": "Cơ điện", "name": "Cơ điện", "is_total": false, "is_sub": false, "form_e": "=SUM(F53:G53)", "form_f": "", "form_g": "", "strat_tot": 0.0}, {"row_idx": 54, "stt": "6", "service": "", "mang": "Hạ tầng", "name": "Hạ tầng", "is_total": false, "is_sub": false, "form_e": "=SUM(F54:G54)", "form_f": "", "form_g": "", "strat_tot": 0.0}]}, {"title": "II. Dịch vụ CĐBR", "rows": []}, {"title": "Đầu tư Dịch vụ CĐBR", "rows": [{"row_idx": 61, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F61:G61)", "form_f": "", "form_g": "", "strat_tot": 3.4223591825377895}, {"row_idx": 62, "stt": "1", "service": "BRCĐ", "mang": "CORE", "name": "CORE", "is_total": false, "is_sub": false, "form_e": "=SUM(F62:G62)", "form_f": "='TH theo DV'!F17", "form_g": "='TH theo DV'!P17", "strat_tot": 0.072}, {"row_idx": 63, "stt": "2", "service": "", "mang": "Truy nhập", "name": "Truy nhập", "is_total": false, "is_sub": false, "form_e": "=SUM(F63:G63)", "form_f": "='TH theo DV'!F16", "form_g": "='TH theo DV'!P16", "strat_tot": 1.514476}, {"row_idx": 64, "stt": "", "service": "", "mang": "Công cụ, phần mềm", "name": "Công cụ, phần mềm", "is_total": false, "is_sub": false, "form_e": "=SUM(F64:G64)", "form_f": "='TH theo DV'!F28", "form_g": "='TH theo DV'!P28", "strat_tot": 0.07}, {"row_idx": 65, "stt": "4", "service": "", "mang": "Truyền dẫn", "name": "Truyền dẫn", "is_total": false, "is_sub": false, "form_e": "=SUM(F65:G65)", "form_f": "='TH theo DV'!H15", "form_g": "='TH theo DV'!R15", "strat_tot": 0.5823522479856005}, {"row_idx": 66, "stt": "5", "service": "", "mang": "IP", "name": "IP", "is_total": false, "is_sub": false, "form_e": "=SUM(F66:G66)", "form_f": "", "form_g": "", "strat_tot": 1.1560629345521887}, {"row_idx": 67, "stt": "6", "service": "", "mang": "Cơ điện", "name": "Cơ điện", "is_total": false, "is_sub": false, "form_e": "=SUM(F67:G67)", "form_f": "='TH theo DV'!J15", "form_g": "='TH theo DV'!T15", "strat_tot": 0.027468}]}, {"title": "III. Dịch vụ CNTT", "rows": []}, {"title": "Đầu tư Dịch vụ CNTT", "rows": [{"row_idx": 73, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F73:G73)", "form_f": "", "form_g": "", "strat_tot": 2.3003343243031016}, {"row_idx": 74, "stt": "2", "service": "CNTT", "mang": "Kinh doanh trực tiếp", "name": "Kinh doanh trực tiếp", "is_total": false, "is_sub": false, "form_e": "=SUM(F74:G74)", "form_f": "='TH theo DV'!L23", "form_g": "='TH theo DV'!V23", "strat_tot": 0.65}, {"row_idx": 75, "stt": "3", "service": "", "mang": "Hỗ trợ kinh doanh và VHKT", "name": "Hỗ trợ kinh doanh và VHKT", "is_total": false, "is_sub": false, "form_e": "=SUM(F75:G75)", "form_f": "='TH theo DV'!L24", "form_g": "='TH theo DV'!V24", "strat_tot": 1.6503343243031017}]}, {"title": "IV.Data Center (DC)", "rows": [{"row_idx": 80, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F80:G80)", "form_f": "", "form_g": "", "strat_tot": 7.018576470588235}, {"row_idx": 81, "stt": "1", "service": "Tổng trạm", "mang": "Phát triển mới", "name": "Phát triển mới", "is_total": false, "is_sub": false, "form_e": "=SUM(F81:G81)", "form_f": "", "form_g": "", "strat_tot": 6.684376470588235}, {"row_idx": 82, "stt": "2", "service": "", "mang": "Củng cố", "name": "Củng cố", "is_total": false, "is_sub": false, "form_e": "=SUM(F82:G82)", "form_f": "", "form_g": "", "strat_tot": 0.0}, {"row_idx": 83, "stt": "3", "service": "", "mang": "ƯCTT", "name": "ƯCTT", "is_total": false, "is_sub": false, "form_e": "=SUM(F83:G83)", "form_f": "", "form_g": "", "strat_tot": 0.3342}]}, {"title": "V.Các dịch vụ khác", "rows": []}, {"title": "Hiện đại hóa mạng lưới", "rows": [{"row_idx": 90, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F90:G90)", "form_f": "", "form_g": "", "strat_tot": 2.8421045000000005}, {"row_idx": 91, "stt": "1", "service": "Hiện đại hóa mạng lưới", "mang": "Vô tuyến", "name": "Vô tuyến", "is_total": false, "is_sub": false, "form_e": "=SUM(F91:G91)", "form_f": "='TH theo DV'!D28", "form_g": "='TH theo DV'!N28", "strat_tot": 0.1723245}, {"row_idx": 92, "stt": "2", "service": "", "mang": "Mạng lõi", "name": "Mạng lõi", "is_total": false, "is_sub": false, "form_e": "=SUM(F92:G92)", "form_f": "='TH theo DV'!E28", "form_g": "='TH theo DV'!O28", "strat_tot": 0.54}, {"row_idx": 93, "stt": "3", "service": "", "mang": "Truyền dẫn quang", "name": "Truyền dẫn quang", "is_total": false, "is_sub": false, "form_e": "=SUM(F93:G93)", "form_f": "='TH theo DV'!H28", "form_g": "='TH theo DV'!R28", "strat_tot": 1.407596}, {"row_idx": 94, "stt": "", "service": "", "mang": "Truyền dẫn IP", "name": "Truyền dẫn IP", "is_total": false, "is_sub": false, "form_e": "=SUM(F94:G94)", "form_f": "", "form_g": "", "strat_tot": 0.0}, {"row_idx": 95, "stt": "4", "service": "", "mang": "CNTT", "name": "CNTT", "is_total": false, "is_sub": false, "form_e": "=SUM(F95:G95)", "form_f": "='TH theo DV'!G28", "form_g": "='TH theo DV'!Q28", "strat_tot": 0.514584}, {"row_idx": 96, "stt": "", "service": "", "mang": "Cơ điện", "name": "Cơ điện", "is_total": false, "is_sub": false, "form_e": "=SUM(F96:G96)", "form_f": "='TH theo DV'!J28", "form_g": "='TH theo DV'!T28", "strat_tot": 0.2076}, {"row_idx": 97, "stt": "5", "service": "", "mang": "Hạ tầng", "name": "Hạ tầng", "is_total": false, "is_sub": false, "form_e": "=SUM(F97:G97)", "form_f": "='TH theo DV'!K28", "form_g": "='TH theo DV'!U28", "strat_tot": 0.0}]}, {"title": "Đầu tư kiên cố mạng lưới", "rows": [{"row_idx": 102, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F102:G102)", "form_f": "", "form_g": "", "strat_tot": 3.0201187899999997}, {"row_idx": 103, "stt": "1", "service": "Kiên cố, củng cố", "mang": "BRCĐ", "name": "BRCĐ", "is_total": false, "is_sub": false, "form_e": "=SUM(F103:G103)", "form_f": "='TH theo DV'!D29", "form_g": "='TH theo DV'!N29", "strat_tot": 0.1}, {"row_idx": 104, "stt": "2", "service": "", "mang": "Mạng lõi", "name": "Mạng lõi", "is_total": false, "is_sub": false, "form_e": "=SUM(F104:G104)", "form_f": "='TH theo DV'!E29", "form_g": "='TH theo DV'!O29", "strat_tot": 0.21}, {"row_idx": 105, "stt": "3", "service": "", "mang": "Truyền dẫn quang", "name": "Truyền dẫn quang", "is_total": false, "is_sub": false, "form_e": "=SUM(F105:G105)", "form_f": "='TH theo DV'!H29", "form_g": "='TH theo DV'!R29", "strat_tot": 1.11476879}, {"row_idx": 106, "stt": "4", "service": "", "mang": "Truyền dẫn IP", "name": "Truyền dẫn IP", "is_total": false, "is_sub": false, "form_e": "=SUM(F106:G106)", "form_f": "", "form_g": "", "strat_tot": 0.0}, {"row_idx": 107, "stt": "5", "service": "", "mang": "Cơ điện", "name": "Cơ điện", "is_total": false, "is_sub": false, "form_e": "=SUM(F107:G107)", "form_f": "='TH theo DV'!J29", "form_g": "='TH theo DV'!T29", "strat_tot": 1.41899}, {"row_idx": 108, "stt": "6", "service": "", "mang": "CNTT", "name": "CNTT", "is_total": false, "is_sub": false, "form_e": "=SUM(F108:G108)", "form_f": "='TH theo DV'!G29", "form_g": "='TH theo DV'!Q29", "strat_tot": 0.0}, {"row_idx": 109, "stt": "7", "service": "", "mang": "Hạ tầng", "name": "Hạ tầng", "is_total": false, "is_sub": false, "form_e": "=SUM(F109:G109)", "form_f": "='TH theo DV'!K29", "form_g": "='TH theo DV'!U29", "strat_tot": 0.17636}]}, {"title": "Đầu tư ƯCTT", "rows": [{"row_idx": 114, "stt": "", "service": "", "mang": "Tổng", "name": "Tổng", "is_total": true, "is_sub": false, "form_e": "=SUM(F114:G114)", "form_f": "", "form_g": "", "strat_tot": 1.0909237859999998}, {"row_idx": 115, "stt": "1", "service": "ƯCTT", "mang": "Vô tuyến", "name": "Vô tuyến", "is_total": false, "is_sub": false, "form_e": "=SUM(F115:G115)", "form_f": "='TH theo DV'!D27", "form_g": "='TH theo DV'!N27", "strat_tot": 0.15407}, {"row_idx": 116, "stt": "2", "service": "", "mang": "Mạng lõi", "name": "Mạng lõi", "is_total": false, "is_sub": false, "form_e": "=SUM(F116:G116)", "form_f": "='TH theo DV'!E27", "form_g": "='TH theo DV'!O27", "strat_tot": 0.0}, {"row_idx": 117, "stt": "3", "service": "", "mang": "Truyền dẫn quang", "name": "Truyền dẫn quang", "is_total": false, "is_sub": false, "form_e": "=SUM(F117:G117)", "form_f": "='TH theo DV'!H27", "form_g": "='TH theo DV'!R27", "strat_tot": 0.40655678599999984}, {"row_idx": 118, "stt": "4", "service": "", "mang": "Truyền dẫn IP", "name": "Truyền dẫn IP", "is_total": false, "is_sub": false, "form_e": "=SUM(F118:G118)", "form_f": "", "form_g": "", "strat_tot": 0.172652}, {"row_idx": 119, "stt": "5", "service": "", "mang": "Cơ điện", "name": "Cơ điện", "is_total": false, "is_sub": false, "form_e": "=SUM(F119:G119)", "form_f": "='TH theo DV'!J27", "form_g": "='TH theo DV'!T27", "strat_tot": 0.0}, {"row_idx": 120, "stt": "6", "service": "", "mang": "CNTT", "name": "CNTT", "is_total": false, "is_sub": false, "form_e": "=SUM(F120:G120)", "form_f": "='TH theo DV'!G27", "form_g": "='TH theo DV'!Q27", "strat_tot": 0.249982}, {"row_idx": 121, "stt": "7", "service": "", "mang": "BRCĐ", "name": "BRCĐ", "is_total": false, "is_sub": false, "form_e": "=SUM(F121:G121)", "form_f": "='TH theo DV'!F27", "form_g": "='TH theo DV'!U27", "strat_tot": 0.107663}]}], "dv_map": {"4": {"stt": "*", "nd": "Tổng", "madv": null}, "5": {"stt": null, "nd": "Tỷ trọng theo mảng", "madv": null}, "6": {"stt": "I", "nd": "Công nghệ mới triển khai diện rộng cho kinh doanh", "madv": null}, "7": {"stt": 1, "nd": "Mạng 5G", "madv": "5G"}, "8": {"stt": 2, "nd": "XGSPON", "madv": "XGSPON"}, "9": {"stt": 3, "nd": "AI/GPU", "madv": "AI/GPU"}, "10": {"stt": 4, "nd": "Private Mobile Network", "madv": "PMN"}, "11": {"stt": "II", "nd": "Mở rộng mạng lưới hiện tại cho kinh doanh", "madv": null}, "12": {"stt": 1, "nd": "Mạng 2/3/4G", "madv": null}, "13": {"stt": "-", "nd": "2/3/4G vùng phủ", "madv": "VPDĐ"}, "14": {"stt": "-", "nd": "2/3/4G dung lượng", "madv": "DLDĐ"}, "15": {"stt": 2, "nd": "Mạng BRCĐ&TH", "madv": null}, "16": {"stt": "-", "nd": "Vùng phủ GPON", "madv": "VPGPON"}, "17": {"stt": "-", "nd": "Dung lượng Internet", "madv": "DLGPON"}, "18": {"stt": "-", "nd": "Dung lượng Truyền hình", "madv": "TH"}, "19": {"stt": "-", "nd": "Kênh truyền", "madv": "KT"}, "20": {"stt": 3, "nd": "Phục vụ kinh doanh Cloud", "madv": "Cloud_KD"}, "21": {"stt": 4, "nd": "Data Center", "madv": "DC"}, "22": {"stt": 5, "nd": "Triển khai hạ tầng CNTT", "madv": null}, "23": {"stt": "-", "nd": "Trực tiếp kinh doanh", "madv": "TTKD"}, "24": {"stt": "-", "nd": "Hỗ trợ kinh doanh và quản trị", "madv": "HTKD"}, "25": {"stt": "-", "nd": "Phần mềm", "madv": "PM"}, "26": {"stt": "III", "nd": "Đảm bảo VHKT, kiên cố, bền vững trong thiên tai", "madv": null}, "27": {"stt": 1, "nd": "Đảm bảo dự phòng ƯCTT", "madv": "ƯCTT"}, "28": {"stt": 2, "nd": "Đảm bảo VHKT và nâng cao chất lượng mạng", "madv": "VHKT"}, "29": {"stt": 3, "nd": "Củng cố kiên cố thường trình", "madv": "KCTT"}, "30": {"stt": 4, "nd": "Bền vững mạng lưới trong thiên tai", "madv": "PCTT"}, "31": {"stt": "IV", "nd": "Ví điện tử", "madv": "VI"}}};

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
  const regex = /'?TH theo DV'?!([A-Z]+)(\d+)/gi;
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

  const blob = new Blob(['\uFEFF' + html], { type: 'application/vnd.ms-excel;charset=utf-8' });
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
      cells.push(row.cells[c].innerText.trim().replace(/\t/g, ' '));
    }
    text += cells.join('\t') + '\n';
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
        cells.push(row.cells[c].innerText.trim().replace(/\t/g, ' '));
      }
      fullText += cells.join('\t') + '\n';
    }
    fullText += '\n';
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

  // Lắng nghe click các nút chuyển Sub-tab báo cáo (Tổng hợp / theo Mảng / theo Dịch vụ / So sánh CL 5 năm)
  document.querySelectorAll('.report-subtab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const subtabId = btn.getAttribute('data-subtab');
      if (subtabId) switchReportSubtab(subtabId);
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
