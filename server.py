import http.server
import socketserver
import os
import json
import threading
import subprocess

PORT = 8080
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

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
            target_path = os.path.join(DIRECTORY, 'Masterlist 2027-2028_Mau.xlsx')
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

    def do_GET(self):
        if self.path == '/api/ping':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({'status': 'ok', 'directory': DIRECTORY}).encode('utf-8'))
            return
        elif self.path == '/api/open-file':
            self.handle_open_file()
            return
        elif self.path == '/api/open-folder':
            self.handle_open_folder()
            return
        elif self.path == '/api/strategy-profiles':
            self.handle_get_strategy_profiles()
            return
        super().do_GET()

    def do_POST(self):
        if self.path == '/api/save-masterlist':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            target_path = os.path.join(DIRECTORY, 'Masterlist 2027-2028_Mau.xlsx')

            if len(body) < 1000:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': 'Dữ liệu file không hợp lệ (kích thước quá nhỏ)'}, ensure_ascii=False).encode('utf-8'))
                return

            try:
                # Ghi trực tiếp vào file Masterlist trên đĩa
                with open(target_path, 'wb') as f:
                    f.write(body)

                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                response = {
                    'success': True,
                    'message': 'Đã cập nhật trực tiếp vào file Masterlist 2027-2028_Mau.xlsx thành công!',
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
        elif self.path == '/api/open-file':
            self.handle_open_file()
        elif self.path == '/api/open-folder':
            self.handle_open_folder()
        elif self.path == '/api/strategy-profiles':
            self.handle_save_strategy_profiles()
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
