from http.server import BaseHTTPRequestHandler
import json

class handler(BaseHTTPRequestHandler):
    
    def _set_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', 'https://kmoreland126.github.io')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
    
    def do_OPTIONS(self):
        self.send_response(200)
        self._set_cors_headers()
        self.end_headers()
        return
    
    def do_POST(self):
        try:
            # For now, return a message that this feature is coming soon
            # The actual download/remediation logic from python-server/server.py 
            # requires python-docx and lxml which don't work on Vercel
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self._set_cors_headers()
            self.end_headers()
            
            response = {
                'success': False,
                'message': 'Download/remediation feature is not available yet. The backend needs to be deployed to a platform that supports python-docx library (like Railway, Render, or Fly.io).',
                'note': 'Currently only upload and analysis is supported on Vercel.'
            }
            
            self.wfile.write(json.dumps(response).encode())
            
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self._set_cors_headers()
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())
