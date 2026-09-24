import http.server
import socketserver
import os
import json
import threading
import subprocess
import base64
import glob
import urllib.parse
from datetime import datetime

PORT = 8080
DIRECTORY = os.path.dirname(os.path.abspath(__file__))
DRAFTS_DIR = os.path.join(DIRECTORY, 'saved_drafts')
os.makedirs(DRAFTS_DIR, exist_ok=True)

class MasterlistHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def handle_open_file(self):
        try:
            import glob
            kq_files = sorted(glob.glob(os.path.join(DIRECTORY, 'Masterlist 2027-2028_KQ_*.xlsx')), key=os.path.getmtime, reverse=True)
            target_path = kq_files[0] if kq_files else os.path.join(DIRECTORY, 'Masterlist 2027-2028_Mau.xlsx')
            if os.path.exists(target_path):
                def _do_open():
                    opened = False
                    if hasattr(os, 'startfile'):
                        try:
                            os.startfile(target_path)
                            opened = True
                        except Exception as ex:
                            print(f"os.startfile error: {ex}")
                    if not opened:
                        try:
                            subprocess.Popen(f'cmd /c start "" "{target_path}"', shell=True)
                            opened = True
                        except Exception as ex:
                            print(f"subprocess start error: {ex}")
                
                threading.Thread(target=_do_open, daemon=True).start()

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                response = {
                    'success': True,
                    'message': f'Đã mở file: {target_path}',
                    'path': target_path
                }
                self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))
            else:
                self.send_response(404)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': 'Chưa tìm thấy file Masterlist 2027-2028_Mau.xlsx'}, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))

    def handle_open_folder(self):
        try:
            def _do_open():
                opened = False
                if hasattr(os, 'startfile'):
                    try:
                        os.startfile(DIRECTORY)
                        opened = True
                    except Exception as ex:
                        print(f"os.startfile error: {ex}")
                if not opened:
                    try:
                        subprocess.Popen(f'explorer "{DIRECTORY}"', shell=True)
                        opened = True
                    except Exception as ex:
                        print(f"subprocess explorer error: {ex}")

            threading.Thread(target=_do_open, daemon=True).start()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            response = {
                'success': True,
                'message': f'Đã mở thư mục: {DIRECTORY}',
                'path': DIRECTORY
            }
            self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))

    def handle_get_strategy_profiles(self):
        try:
            target_path = os.path.join(DIRECTORY, 'strategy_profiles.json')
            if os.path.exists(target_path):
                with open(target_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(content.encode('utf-8'))
            else:
                self.send_response(404)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'error': 'Not found'}).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))

    def handle_save_strategy_profiles(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            parsed = json.loads(body.decode('utf-8'))
            target_path = os.path.join(DIRECTORY, 'strategy_profiles.json')
            with open(target_path, 'w', encoding='utf-8') as f:
                json.dump(parsed, f, ensure_ascii=False, indent=2)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True, 'message': 'Đã lưu cấu hình profile chiến lược thành công'}).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode('utf-8'))

    def handle_list_drafts(self):
        try:
            os.makedirs(DRAFTS_DIR, exist_ok=True)
            drafts = []
            for file_path in glob.glob(os.path.join(DRAFTS_DIR, '*.json')):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    draft_id = data.get('id') or os.path.splitext(os.path.basename(file_path))[0]
                    xlsx_path = os.path.join(DRAFTS_DIR, f"{draft_id}.xlsx")
                    has_xlsx = os.path.exists(xlsx_path)
                    file_size = os.path.getsize(xlsx_path) if has_xlsx else os.path.getsize(file_path)
                    drafts.append({
                        'id': draft_id,
                        'name': data.get('name', 'Bản lưu không tên'),
                        'note': data.get('note', ''),
                        'timestamp': data.get('timestamp') or os.path.getmtime(file_path),
                        'formattedTime': data.get('formattedTime', ''),
                        'rowCount': data.get('rowCount', 0),
                        'mangs': data.get('mangs', []),
                        'total2027': data.get('total2027', 0),
                        'total2028': data.get('total2028', 0),
                        'hasExcel': has_xlsx,
                        'fileSize': file_size
                    })
                except Exception as ex:
                    print(f"Error reading draft {file_path}: {ex}")

            drafts.sort(key=lambda x: x.get('timestamp', 0), reverse=True)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True, 'drafts': drafts}, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))

    def handle_load_draft(self, query):
        try:
            params = urllib.parse.parse_qs(query)
            draft_id = params.get('id', [None])[0]
            if not draft_id:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': 'Thiếu tham số id bản lưu'}).encode('utf-8'))
                return

            safe_id = os.path.basename(draft_id)
            target_path = os.path.join(DRAFTS_DIR, f"{safe_id}.json")
            if not os.path.exists(target_path):
                self.send_response(404)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': f'Không tìm thấy bản lưu với id: {draft_id}'}, ensure_ascii=False).encode('utf-8'))
                return

            with open(target_path, 'r', encoding='utf-8') as f:
                content = f.read()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(content.encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))

    def handle_download_draft_excel(self, query):
        try:
            params = urllib.parse.parse_qs(query)
            draft_id = params.get('id', [None])[0]
            if not draft_id:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': 'Thiếu tham số id bản lưu'}).encode('utf-8'))
                return

            safe_id = os.path.basename(draft_id)
            xlsx_path = os.path.join(DRAFTS_DIR, f"{safe_id}.xlsx")
            if not os.path.exists(xlsx_path):
                self.send_response(404)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': 'Chưa tìm thấy file Excel của bản lưu này'}, ensure_ascii=False).encode('utf-8'))
                return

            with open(xlsx_path, 'rb') as f:
                content = f.read()

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="Masterlist_Draft_{safe_id}.xlsx"')
            self.send_header('Content-Length', str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))

    def handle_save_draft(self):
        try:
            os.makedirs(DRAFTS_DIR, exist_ok=True)
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode('utf-8'))

            draft_id = payload.get('id')
            now = datetime.now()
            if not draft_id:
                draft_id = f"draft_{now.strftime('%Y%m%d_%H%M%S')}"
            else:
                draft_id = os.path.basename(draft_id)

            payload['id'] = draft_id
            if 'timestamp' not in payload:
                payload['timestamp'] = now.timestamp()
            if 'formattedTime' not in payload:
                payload['formattedTime'] = now.strftime('%H:%M %d/%m/%Y')

            excel_b64 = payload.pop('excelBase64', None)
            if excel_b64:
                try:
                    excel_bytes = base64.b64decode(excel_b64)
                    xlsx_path = os.path.join(DRAFTS_DIR, f"{draft_id}.xlsx")
                    with open(xlsx_path, 'wb') as f:
                        f.write(excel_bytes)
                    payload['excelFileName'] = f"{draft_id}.xlsx"
                except Exception as ex:
                    print(f"Lỗi ghi file Excel bản lưu: {ex}")

            json_path = os.path.join(DRAFTS_DIR, f"{draft_id}.json")
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            resp = {
                'success': True,
                'message': f'Đã lưu thành công bản lưu: {payload.get("name")}',
                'draft': {
                    'id': draft_id,
                    'name': payload.get('name'),
                    'formattedTime': payload.get('formattedTime'),
                    'rowCount': payload.get('rowCount', 0),
                    'mangs': payload.get('mangs', [])
                }
            }
            self.wfile.write(json.dumps(resp, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))

    def handle_delete_draft(self, query):
        try:
            params = urllib.parse.parse_qs(query)
            draft_id = params.get('id', [None])[0]
            if not draft_id:
                try:
                    content_length = int(self.headers.get('Content-Length', 0))
                    if content_length > 0:
                        body = self.rfile.read(content_length)
                        parsed = json.loads(body.decode('utf-8'))
                        draft_id = parsed.get('id')
                except Exception:
                    pass

            if not draft_id:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': 'Thiếu tham số id bản lưu'}).encode('utf-8'))
                return

            safe_id = os.path.basename(draft_id)
            json_path = os.path.join(DRAFTS_DIR, f"{safe_id}.json")
            xlsx_path = os.path.join(DRAFTS_DIR, f"{safe_id}.xlsx")

            deleted = False
            if os.path.exists(json_path):
                os.remove(json_path)
                deleted = True
            if os.path.exists(xlsx_path):
                os.remove(xlsx_path)
                deleted = True

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True, 'message': 'Đã xóa bản lưu'}).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))

    def handle_get_report_data(self, query):
        try:
            import parse_report_data
            params = urllib.parse.parse_qs(query)
            custom_file = params.get('file', [None])[0]
            best_file = parse_report_data.find_best_report_file(custom_file)
            if not best_file or not os.path.exists(best_file):
                self.send_response(404)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': 'Không tìm thấy file Masterlist hợp lệ'}, ensure_ascii=False).encode('utf-8'))
                return

            result = parse_report_data.parse_report_workbook(best_file)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps(result, ensure_ascii=False).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}, ensure_ascii=False).encode('utf-8'))

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        clean_path = parsed.path
        query = parsed.query

        if clean_path == '/api/ping':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'status': 'ok', 'directory': DIRECTORY}).encode('utf-8'))
            return
        elif clean_path == '/api/open-file':
            self.handle_open_file()
            return
        elif clean_path == '/api/open-folder':
            self.handle_open_folder()
            return
        elif clean_path == '/api/strategy-profiles':
            self.handle_get_strategy_profiles()
            return
        elif clean_path == '/api/drafts':
            self.handle_list_drafts()
            return
        elif clean_path == '/api/load-draft':
            self.handle_load_draft(query)
            return
        elif clean_path == '/api/draft-excel':
            self.handle_download_draft_excel(query)
            return
        elif clean_path == '/api/report-data':
            self.handle_get_report_data(query)
            return
        super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        clean_path = parsed.path
        query = parsed.query

        if clean_path == '/api/save-masterlist':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            from datetime import datetime
            now_str = datetime.now().strftime('%d%m%Y_%H%M')
            target_path = os.path.join(DIRECTORY, f'Masterlist 2027-2028_KQ_{now_str}.xlsx')

            if len(body) < 1000:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': 'Dữ liệu file không hợp lệ (kích thước quá nhỏ)'}, ensure_ascii=False).encode('utf-8'))
                return

            try:
                # Ghi file kết quả kèm dấu thời gian, bảo vệ file mẫu gốc
                with open(target_path, 'wb') as f:
                    f.write(body)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                response = {
                    'success': True,
                    'message': f'Đã lưu file kết quả: {os.path.basename(target_path)} thành công!',
                    'path': target_path,
                    'size': len(body)
                }
                self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))
            except PermissionError:
                # File đang mở trong Excel
                self.send_response(409)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                response = {
                    'success': False,
                    'locked': True,
                    'error': 'File Masterlist 2027-2028_Mau.xlsx đang được mở trong Microsoft Excel. Vui lòng đóng file Excel lại rồi bấm Cập nhật!'
                }
                self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                response = {
                    'success': False,
                    'error': str(e)
                }
                self.wfile.write(json.dumps(response, ensure_ascii=False).encode('utf-8'))
        elif clean_path == '/api/open-file':
            self.handle_open_file()
        elif clean_path == '/api/open-folder':
            self.handle_open_folder()
        elif clean_path == '/api/strategy-profiles':
            self.handle_save_strategy_profiles()
        elif clean_path == '/api/save-draft':
            self.handle_save_draft()
        elif clean_path == '/api/delete-draft':
            self.handle_delete_draft(query)
        else:
            self.send_response(404)
            self.end_headers()

class ThreadedTCPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    allow_reuse_address = True
    daemon_threads = True

if __name__ == '__main__':
    with ThreadedTCPServer(("", PORT), MasterlistHandler) as httpd:
        print(f"Masterlist server running at http://localhost:{PORT}")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("Server stopped.")
