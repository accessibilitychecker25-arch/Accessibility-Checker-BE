from http.server import BaseHTTPRequestHandler
import json
import io
import zipfile

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
            content_type = self.headers.get('Content-Type')
            if not content_type or 'multipart/form-data' not in content_type:
                self.send_error(400, 'Expected multipart/form-data')
                return
            
            content_length = int(self.headers.get('Content-Length', 0))
            if content_length == 0:
                self.send_error(400, 'No file uploaded')
                return
            
            body = self.rfile.read(content_length)
            boundary = content_type.split('boundary=')[-1].encode()
            parts = body.split(b'--' + boundary)
            
            file_data = None
            filename = None
            
            for part in parts:
                if b'Content-Disposition' in part and b'filename=' in part:
                    lines = part.split(b'\r\n')
                    for line in lines:
                        if b'filename=' in line:
                            filename_start = line.find(b'filename="') + 10
                            filename_end = line.find(b'"', filename_start)
                            filename = line[filename_start:filename_end].decode('utf-8')
                    file_start = part.find(b'\r\n\r\n') + 4
                    file_data = part[file_start:-2]
                    break
            
            if not file_data or not filename:
                self.send_error(400, 'No file found in request')
                return
            
            if not filename.lower().endswith('.docx'):
                self.send_error(400, 'Please upload a .docx file')
                return
            
            report = self.analyze_docx(file_data, filename)
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self._set_cors_headers()
            self.end_headers()
            
            response = {'fileName': filename, 'suggestedFileName': filename, 'report': report}
            self.wfile.write(json.dumps(response).encode())
            
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self._set_cors_headers()
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode())
    
    def analyze_docx(self, file_data, filename):
        try:
            docx_zip = zipfile.ZipFile(io.BytesIO(file_data))
            report = {
                'fileName': filename,
                'suggestedFileName': filename,
                'summary': {'fixed': 0, 'flagged': 0},
                'details': {
                    'titleNeedsFixing': False,
                    'imagesMissingOrBadAlt': 0,
                    'gifsDetected': [],
                    'fileNameNeedsFixing': False,
                }
            }
            
            try:
                core_xml = docx_zip.read('docProps/core.xml').decode('utf-8')
                if '<dc:title></dc:title>' in core_xml or '<dc:title/>' in core_xml:
                    report['details']['titleNeedsFixing'] = True
                    report['summary']['flagged'] += 1
            except:
                pass
            
            try:
                rels_xml = docx_zip.read('word/_rels/document.xml.rels').decode('utf-8')
                image_count = rels_xml.count('relationships/image"')
                if image_count > 0:
                    report['details']['imagesMissingOrBadAlt'] = image_count
                    report['summary']['flagged'] += image_count
            except:
                pass
            
            gif_files = [n for n in docx_zip.namelist() if n.startswith('word/media/') and n.lower().endswith('.gif')]
            if gif_files:
                report['details']['gifsDetected'] = gif_files
                report['summary']['flagged'] += len(gif_files)
            
            if '_' in filename or filename.lower().startswith('document') or filename.lower().startswith('untitled'):
                report['details']['fileNameNeedsFixing'] = True
                report['summary']['flagged'] += 1
            
            return report
        except Exception as e:
            return {'fileName': filename, 'error': str(e), 'summary': {'fixed': 0, 'flagged': 0}, 'details': {}}
