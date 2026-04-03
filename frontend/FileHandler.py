import http.server
import socketserver
import json
import os
import subprocess
import sys
import logging

# Configure basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

PORT = 16225
# Set the base directory to the project root (one level up from `frontend/`)
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

class FolderOpenHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/open_folder':
            try:
                content_length = int(self.headers['Content-Length'])
                post_data = self.rfile.read(content_length)
                data = json.loads(post_data)
                relative_path = data.get('path')

                if not relative_path:
                    self._send_response(400, {'status': 'error', 'message': 'Path not provided'})
                    return

                # Prevent directory traversal attacks by ensuring the path is within the project
                full_path = os.path.abspath(os.path.join(BASE_DIR, relative_path))
                if not full_path.startswith(BASE_DIR):
                    self._send_response(403, {'status': 'error', 'message': 'Access denied: Path is outside the allowed directory.'})
                    return

                if not os.path.isdir(full_path):
                    self._send_response(404, {'status': 'error', 'message': f'Directory not found: {full_path}'})
                    return

                logging.info(f"Opening folder: {full_path}")
                self._open_directory(full_path)
                self._send_response(200, {'status': 'success', 'message': f'Opened {full_path}'})

            except Exception as e:
                logging.error(f"Server error: {e}")
                self._send_response(500, {'status': 'error', 'message': str(e)})
        else:
            self._send_response(404, {'status': 'error', 'message': 'Endpoint not found'})

    def _open_directory(self, path):
        """Opens a directory in the default file explorer."""
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin": # macOS
            subprocess.run(["open", path])
        else: # Linux
            subprocess.run(["xdg-open", path])

    def _send_response(self, status_code, content):
        self.send_response(status_code)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(content).encode('utf-8'))

if __name__ == "__main__":
    with socketserver.TCPServer(("0.0.0.0", PORT), FolderOpenHandler) as httpd:
        logging.info(f"FileHandler server started on port {PORT}")
        httpd.serve_forever()