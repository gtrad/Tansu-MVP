"""
Tansu API Server
A lightweight HTTP server that provides variable data to the Word VBA macro.
Runs on localhost:5050 by default.
"""

import json
import logging
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import threading

from database import VariableDatabase

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

    def do_GET(self):
        """Handle GET requests."""
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
                self._send_error_response(f'Unknown endpoint: {path}', 404)

        except Exception as e:
            logger.error(f"API error: {e}")
            self._send_error_response(str(e), 500)

    def do_OPTIONS(self):
        """Handle CORS preflight."""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()


class TansuAPIServer:
    """Manages the Tansu API HTTP server."""

    def __init__(self, port=DEFAULT_PORT):
        self.port = port
        self.server = None
        self.thread = None
        self._running = False

    def start(self):
        """Start the API server in a background thread."""
        if self._running:
            return

        try:
            self.server = HTTPServer(('127.0.0.1', self.port), TansuAPIHandler)
            self._running = True
            self.thread = threading.Thread(target=self._serve, daemon=True)
            self.thread.start()
            logger.info(f"Tansu API server started on http://127.0.0.1:{self.port}")
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
    print(f"Starting Tansu API server on http://127.0.0.1:{DEFAULT_PORT}")
    print("Endpoints:")
    print("  GET /variables - List all variables")
    print("  GET /variable?name=VAR_NAME - Get single variable")
    print("  GET /ping - Health check")
    print("\nPress Ctrl+C to stop...")

    server = HTTPServer(('127.0.0.1', DEFAULT_PORT), TansuAPIHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        server.shutdown()


if __name__ == "__main__":
    main()
