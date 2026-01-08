"""
Tansu API Server
A lightweight HTTP server that provides variable data to the Word add-in.
Runs on localhost:5050 by default. Supports both HTTP and WebSocket connections.
"""

import json
import logging
import os
import hashlib
import struct
import base64
import ssl
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn
from urllib.parse import urlparse, parse_qs
import threading

from database import VariableDatabase
import platform

# Path to SSL certificates for local.tansu.co
CERTS_DIR = os.path.join(os.path.dirname(__file__), 'certs')

# Path to word-addin static files
ADDIN_DIR = os.path.join(os.path.dirname(__file__), 'word-addin')

# Content type mappings
CONTENT_TYPES = {
    '.html': 'text/html',
    '.js': 'application/javascript',
    '.css': 'text/css',
    '.png': 'image/png',
    '.ico': 'image/x-icon',
    '.json': 'application/json',
}

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

DEFAULT_PORT = 5050


class TansuAPIHandler(BaseHTTPRequestHandler):
    """HTTP request handler for Tansu API."""

    def log_message(self, format, *args):
        """Override to use our logger."""
        logger.debug(f"{self.address_string()} - {format % args}")

    def _send_json_response(self, data, status=200):
        """Send a JSON response."""
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode('utf-8'))

    def _send_error_response(self, message, status=400):
        """Send an error response."""
        self._send_json_response({'error': message}, status)

    def _serve_static_file(self, path):
        """Serve static files from the word-addin directory."""
        # Default to index.html for root
        if path == '/' or path == '':
            path = '/taskpane.html'

        # Security: prevent directory traversal
        safe_path = os.path.normpath(path).lstrip('/')
        if '..' in safe_path:
            self._send_error_response('Invalid path', 403)
            return

        file_path = os.path.join(ADDIN_DIR, safe_path)

        if os.path.isfile(file_path):
            # Determine content type
            ext = os.path.splitext(file_path)[1].lower()
            content_type = CONTENT_TYPES.get(ext, 'application/octet-stream')

            try:
                with open(file_path, 'rb') as f:
                    content = f.read()

                self.send_response(200)
                self.send_header('Content-Type', content_type)
                self.send_header('Content-Length', len(content))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(content)
            except Exception as e:
                logger.error(f"Error serving file {file_path}: {e}")
                self._send_error_response('Error reading file', 500)
        else:
            self._send_error_response(f'File not found: {path}', 404)

    def do_GET(self):
        """Handle GET requests including WebSocket upgrade."""
        # Check for WebSocket upgrade request
        if self.headers.get('Upgrade', '').lower() == 'websocket':
            self._handle_websocket()
            return

        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)

        try:
            db = VariableDatabase()

            if path == '/variables':
                # Get all variables
                variables = db.get_all_variables()
                # Convert to simpler format for VBA
                result = []
                for var in variables:
                    result.append({
                        'id': var['id'],
                        'name': var['name'],
                        'value': var['value'],
                        'unit': var.get('unit', '')
                    })
                self._send_json_response({'variables': result})

            elif path == '/variable':
                # Get single variable by name
                name = query.get('name', [None])[0]
                if not name:
                    self._send_error_response('Missing name parameter')
                    return
                var = db.get_variable_by_name(name)
                if var:
                    self._send_json_response({
                        'id': var['id'],
                        'name': var['name'],
                        'value': var['value'],
                        'unit': var.get('unit', '')
                    })
                else:
                    self._send_error_response(f'Variable not found: {name}', 404)

            elif path == '/ping':
                # Health check
                self._send_json_response({'status': 'ok', 'service': 'Tansu API'})

            else:
                # Try to serve static files from word-addin folder
                self._serve_static_file(path)

        except Exception as e:
            logger.error(f"API error: {e}")
            self._send_error_response(str(e), 500)

    def do_POST(self):
        """Handle POST requests."""
        parsed = urlparse(self.path)
        path = parsed.path

        try:
            if path == '/insert':
                # Read request body
                content_length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(content_length).decode('utf-8')
                data = json.loads(body) if body else {}

                var_name = data.get('name')
                if not var_name:
                    self._send_error_response('Missing name parameter')
                    return

                # Get variable from database
                db = VariableDatabase()
                var = db.get_variable_by_name(var_name)
                if not var:
                    self._send_error_response(f'Variable not found: {var_name}', 404)
                    return

                # Determine value to insert (with or without unit)
                with_unit = data.get('with_unit', False)
                value = var['value']
                if with_unit and var.get('unit'):
                    value = f"{value} {var['unit']}"

                # Insert into Word using platform-specific integration
                success = self._insert_into_word(var_name, value)

                if success:
                    self._send_json_response({'status': 'ok', 'inserted': var_name})
                else:
                    self._send_error_response('Failed to insert into Word - is Word open?', 500)

            else:
                self._send_error_response(f'Unknown endpoint: {path}', 404)

        except json.JSONDecodeError:
            self._send_error_response('Invalid JSON body')
        except Exception as e:
            logger.error(f"API error: {e}")
            self._send_error_response(str(e), 500)

    def _insert_into_word(self, var_name: str, var_value: str) -> bool:
        """Insert variable into Word using platform-specific integration."""
        try:
            if platform.system() == 'Darwin':
                from word_mac import WordIntegration
                word = WordIntegration()
                return word.insert_variable(var_name, var_value)
            elif platform.system() == 'Windows':
                from word_windows import WordIntegration
                word = WordIntegration()
                return word.insert_variable(var_name, var_value)
            else:
                logger.error(f"Unsupported platform: {platform.system()}")
                return False
        except Exception as e:
            logger.error(f"Error inserting into Word: {e}")
            return False

    def do_OPTIONS(self):
        """Handle CORS preflight."""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def _handle_websocket(self):
        """Handle WebSocket connection upgrade and communication."""
        try:
            # Perform WebSocket handshake
            key = self.headers.get('Sec-WebSocket-Key')
            if not key:
                self._send_error_response('Missing Sec-WebSocket-Key', 400)
                return

            # Calculate accept key
            GUID = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
            accept_key = base64.b64encode(
                hashlib.sha1((key + GUID).encode()).digest()
            ).decode()

            # Send handshake response
            self.send_response(101, 'Switching Protocols')
            self.send_header('Upgrade', 'websocket')
            self.send_header('Connection', 'Upgrade')
            self.send_header('Sec-WebSocket-Accept', accept_key)
            self.end_headers()

            logger.info("WebSocket connection established")

            # Handle WebSocket messages
            while True:
                try:
                    message = self._ws_receive()
                    if message is None:
                        break

                    data = json.loads(message)
                    response = self._handle_ws_message(data)
                    if response:
                        self._ws_send(json.dumps(response))
                except Exception as e:
                    logger.error(f"WebSocket message error: {e}")
                    break

            logger.info("WebSocket connection closed")

        except Exception as e:
            logger.error(f"WebSocket error: {e}")

    def _ws_receive(self):
        """Receive a WebSocket frame and return the payload."""
        try:
            # Read first 2 bytes
            header = self.rfile.read(2)
            if len(header) < 2:
                return None

            b1, b2 = header[0], header[1]
            opcode = b1 & 0x0F
            masked = (b2 & 0x80) != 0
            payload_len = b2 & 0x7F

            # Handle close frame
            if opcode == 0x8:
                return None

            # Extended payload length
            if payload_len == 126:
                ext = self.rfile.read(2)
                payload_len = struct.unpack('>H', ext)[0]
            elif payload_len == 127:
                ext = self.rfile.read(8)
                payload_len = struct.unpack('>Q', ext)[0]

            # Read mask key if masked
            mask_key = self.rfile.read(4) if masked else None

            # Read payload
            payload = self.rfile.read(payload_len)

            # Unmask if necessary
            if masked and mask_key:
                payload = bytes(b ^ mask_key[i % 4] for i, b in enumerate(payload))

            return payload.decode('utf-8')

        except Exception as e:
            logger.error(f"WebSocket receive error: {e}")
            return None

    def _ws_send(self, message):
        """Send a WebSocket text frame."""
        try:
            payload = message.encode('utf-8')
            length = len(payload)

            # Build frame header
            frame = bytearray()
            frame.append(0x81)  # FIN + text opcode

            if length < 126:
                frame.append(length)
            elif length < 65536:
                frame.append(126)
                frame.extend(struct.pack('>H', length))
            else:
                frame.append(127)
                frame.extend(struct.pack('>Q', length))

            frame.extend(payload)
            self.wfile.write(bytes(frame))
            self.wfile.flush()

        except Exception as e:
            logger.error(f"WebSocket send error: {e}")

    def _handle_ws_message(self, data):
        """Handle a WebSocket message and return response."""
        msg_type = data.get('type')

        if msg_type == 'get_variables':
            db = VariableDatabase()
            variables = db.get_all_variables()
            result = []
            for var in variables:
                result.append({
                    'id': var['id'],
                    'name': var['name'],
                    'value': var['value'],
                    'unit': var.get('unit', '')
                })
            return {'type': 'variables', 'variables': result}

        elif msg_type == 'insert':
            var_name = data.get('name')
            with_unit = data.get('with_unit', False)

            if not var_name:
                return {'type': 'insert_result', 'success': False, 'error': 'Missing name'}

            db = VariableDatabase()
            var = db.get_variable_by_name(var_name)
            if not var:
                return {'type': 'insert_result', 'success': False, 'error': f'Variable not found: {var_name}'}

            value = var['value']
            if with_unit and var.get('unit'):
                value = f"{value} {var['unit']}"

            success = self._insert_into_word(var_name, str(value))
            return {
                'type': 'insert_result',
                'success': success,
                'name': var_name,
                'error': None if success else 'Failed to insert into Word'
            }

        elif msg_type == 'ping':
            return {'type': 'pong'}

        return None


class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    """HTTP server that handles each request in a separate thread."""
    daemon_threads = True


class TansuAPIServer:
    """Manages the Tansu API HTTP server."""

    def __init__(self, port=DEFAULT_PORT):
        self.port = port
        self.server = None
        self.thread = None
        self._running = False
        self._ssl_enabled = False

    def start(self):
        """Start the API server in a background thread."""
        if self._running:
            return

        try:
            self.server = ThreadedHTTPServer(('127.0.0.1', self.port), TansuAPIHandler)

            # Enable SSL if certificates exist
            cert_file = os.path.join(CERTS_DIR, 'fullchain.pem')
            key_file = os.path.join(CERTS_DIR, 'privkey.pem')

            if os.path.exists(cert_file) and os.path.exists(key_file):
                context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
                context.load_cert_chain(cert_file, key_file)
                self.server.socket = context.wrap_socket(self.server.socket, server_side=True)
                self._ssl_enabled = True
                logger.info(f"SSL enabled with certificate: {cert_file}")

            self._running = True
            self.thread = threading.Thread(target=self._serve, daemon=True)
            self.thread.start()
            protocol = "https" if self._ssl_enabled else "http"
            logger.info(f"Tansu API server started on {protocol}://127.0.0.1:{self.port}")
        except OSError as e:
            if "Address already in use" in str(e):
                logger.warning(f"Port {self.port} already in use - API server may already be running")
            else:
                raise

    def _serve(self):
        """Serve requests until stopped."""
        while self._running:
            self.server.handle_request()

    def stop(self):
        """Stop the API server."""
        self._running = False
        if self.server:
            self.server.shutdown()
            logger.info("Tansu API server stopped")

    def is_running(self):
        """Check if server is running."""
        return self._running


# Global server instance
_server_instance = None


def start_api_server(port=DEFAULT_PORT):
    """Start the global API server instance."""
    global _server_instance
    if _server_instance is None:
        _server_instance = TansuAPIServer(port)
    _server_instance.start()
    return _server_instance


def stop_api_server():
    """Stop the global API server instance."""
    global _server_instance
    if _server_instance:
        _server_instance.stop()
        _server_instance = None


def main():
    """Run the API server standalone."""
    # Check for SSL certificates
    cert_file = os.path.join(CERTS_DIR, 'fullchain.pem')
    key_file = os.path.join(CERTS_DIR, 'privkey.pem')
    use_ssl = os.path.exists(cert_file) and os.path.exists(key_file)

    protocol = "https" if use_ssl else "http"
    domain = "local.tansu.co" if use_ssl else "127.0.0.1"

    print(f"Starting Tansu API server on {protocol}://{domain}:{DEFAULT_PORT}")
    print("Endpoints:")
    print("  GET  /variables - List all variables")
    print("  GET  /variable?name=VAR_NAME - Get single variable")
    print("  GET  /ping - Health check")
    print("  POST /insert - Insert variable into Word (body: {name, with_unit})")
    print("  WSS  /ws - Secure WebSocket for add-in")
    print("\nPress Ctrl+C to stop...")

    server = ThreadedHTTPServer(('127.0.0.1', DEFAULT_PORT), TansuAPIHandler)

    if use_ssl:
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        context.load_cert_chain(cert_file, key_file)
        server.socket = context.wrap_socket(server.socket, server_side=True)
        print(f"SSL enabled with: {cert_file}")
    else:
        print("WARNING: No SSL certificates found in certs/ directory")
        print("         Word add-in will not be able to connect")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        server.shutdown()


if __name__ == "__main__":
    main()
