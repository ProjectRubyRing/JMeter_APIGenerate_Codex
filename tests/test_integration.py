"""Real wire-level tests. Enable with JMETER_TEST_HOME pointing at JMeter 5.6.3."""
import csv
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import time
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qsl, urlsplit

from restjmeter.inputs import COLUMNS
from restjmeter.workbook import write_xlsx
from test_support import TemporaryWorkspace

ROOT = Path(__file__).resolve().parents[1]
JMETER_HOME = os.environ.get('JMETER_TEST_HOME')


class RecordingHandler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def do_GET(self):
        self.respond()

    def do_POST(self):
        self.respond()

    def respond(self):
        body = self.rfile.read(int(self.headers.get('Content-Length', '0')))
        self.server.requests.put({'method': self.command, 'path': self.path,
                                  'headers': dict(self.headers), 'body': body})
        path = urlsplit(self.path).path
        if path == '/slow':
            time.sleep(1)
        status = 404 if path == '/missing' else 302 if path == '/redirect' else 200
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=UTF-8')
        self.send_header('Content-Length', '2')
        self.send_header('Connection', 'close')
        self.close_connection = True
        if status == 302:
            self.send_header('Location', '/unexpected-followup')
        try:
            self.end_headers()
            self.wfile.write(b'{}')
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            pass

    def log_message(self, format_, *args):
        pass


@unittest.skipUnless(JMETER_HOME, 'Set JMETER_TEST_HOME to run real JMeter 5.6.3 integration tests')
class JMeterIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingHTTPServer(('127.0.0.1', 0), RecordingHandler)
        cls.server.requests = queue.Queue()
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = 'http://127.0.0.1:%d' % cls.server.server_port

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=5)

    def setUp(self):
        self.temp = TemporaryWorkspace(prefix='jmeter test 日本語 ')
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.drain_requests()

    def drain_requests(self):
        requests = []
        while True:
            try:
                requests.append(self.server.requests.get_nowait())
            except queue.Empty:
                return requests

    def input_file(self, rows, suffix='.csv'):
        path = self.root / ('input file' + suffix)
        if suffix == '.xlsx':
            write_xlsx(path, [('APIs', [COLUMNS] + [[row.get(key, '') for key in COLUMNS] for row in rows], [20] * len(COLUMNS))])
        else:
            with path.open('w', encoding='utf-8-sig', newline='') as stream:
                writer = csv.DictWriter(stream, fieldnames=COLUMNS)
                writer.writeheader()
                writer.writerows(rows)
        return path

    def run_tool(self, rows, suffix='.csv', extra=None, wrapper=None, trailing_output=False):
        input_path = self.input_file(rows, suffix)
        output = self.root / 'results & spaces'
        if wrapper == 'powershell':
            command = [shutil.which('powershell.exe'), '-NoProfile', '-File', str(ROOT / 'scripts/Invoke-JMeterApi.ps1'),
                       '-InputPath', str(input_path), '-JMeterHome', JMETER_HOME,
                       '-OutputDirectory', str(output) + (os.sep if trailing_output else '')]
        else:
            launcher = (['bash', str(ROOT / 'scripts/jmeter-api.sh')] if wrapper == 'bash'
                        else [sys.executable, str(ROOT / 'jmeter_api.py')])
            command = launcher + ['--input', str(input_path), '--jmeter-home', JMETER_HOME, '--output', str(output)]
        result = subprocess.run(command + (extra or []), capture_output=True, timeout=120)
        folders = list(output.iterdir()) if output.exists() else []
        self.assertEqual(len(folders), 1, (result.stdout + result.stderr).decode(errors='replace'))
        folder = folders[0]
        summary_path = folder / 'summary.json'
        self.assertTrue(summary_path.exists(), (result.stdout + result.stderr).decode(errors='replace'))
        summary = json.loads(summary_path.read_text(encoding='utf-8'))
        diagnostics = (result.stdout + result.stderr).decode(errors='replace')
        for name in ('console.log', 'jmeter.log', 'report-console.log', 'report.log'):
            if (folder / name).exists():
                diagnostics += '\n' + name + '\n' + (folder / name).read_text(encoding='utf-8', errors='replace')[-16000:]
        return result.returncode, summary, folder, diagnostics

    def test_single_get_default_csv_and_html_report(self):
        row = {'name': 'single_GET', 'method': 'GET', 'url': self.base + '/echo?existing=a%2Bb',
               'query_params': '{"q":"日本語 & +","empty":""}', 'headers': '{"X-Test":"header-value"}'}
        code, summary, folder, diagnostics = self.run_tool([row], extra=['--html-report'])
        self.assertEqual(code, 0, diagnostics)
        self.assertEqual(summary['total_samples'], 1)
        self.assertEqual(summary['failed_samples'], 0)
        requests = self.drain_requests()
        self.assertEqual(len(requests), 1)
        self.assertEqual(requests[0]['body'], b'')
        self.assertEqual(requests[0]['headers']['X-Test'], 'header-value')
        self.assertEqual(parse_qsl(urlsplit(requests[0]['path']).query, keep_blank_values=True),
                         [('existing', 'a+b'), ('q', '日本語 & +'), ('empty', '')])
        self.assertIn('API_RESULT label=single_GET code=200', (folder / 'jmeter.log').read_text(encoding='utf-8'))
        self.assertIn('SAMPLE index=1 code=200', (folder / 'execution.log').read_text(encoding='utf-8'))
        self.assertTrue((folder / 'html-report/index.html').is_file(), diagnostics)

    def test_xlsx_post_json_form_raw_and_literal_jmeter_functions(self):
        raw_body = ' first, <&>\r\n日本語\\end\n'
        literal_body = '${__groovy(40+2)}\\path, 日本語\r\n'
        json_body = json.dumps({'text': '日本語', 'escape': '\\value', 'expression': '${__groovy(1+1)}'}, ensure_ascii=False)
        rows = [
            {'name': 'json', 'method': 'POST', 'url': self.base + '/json?original=1', 'query_params': '{"query":"value"}', 'body_type': 'json', 'body': json_body},
            {'name': 'form', 'method': 'POST', 'url': self.base + '/form?query=only', 'body_type': 'form',
             'body': '[{"name":"x","value":"a+b &"},{"name":"x","value":"日本語"},{"name":"empty","value":""}]'},
            {'name': 'raw', 'method': 'POST', 'url': self.base + '/raw', 'body_type': 'raw', 'body': raw_body},
            {'name': 'literal', 'method': 'POST', 'url': self.base + '/literal', 'body_type': 'raw', 'body': literal_body,
             'headers': '{"X-Literal":"${__groovy(40+2)}"}'},
            {'name': 'disabled', 'enabled': 'false', 'method': 'GET', 'url': self.base + '/disabled'},
        ]
        code, summary, _, diagnostics = self.run_tool(rows, '.xlsx')
        self.assertEqual(code, 0, diagnostics)
        self.assertEqual(summary['total_samples'], 4)
        requests = self.drain_requests()
        self.assertEqual(len(requests), 4)
        self.assertEqual(requests[0]['body'], json_body.encode('utf-8'))
        self.assertEqual(requests[0]['path'], '/json?original=1&query=value')
        self.assertTrue(requests[0]['headers']['Content-Type'].startswith('application/json'))
        self.assertEqual(requests[1]['path'], '/form?query=only')
        self.assertEqual(parse_qsl(requests[1]['body'].decode(), keep_blank_values=True),
                         [('x', 'a+b &'), ('x', '日本語'), ('empty', '')])
        self.assertEqual(requests[2]['body'], raw_body.encode('utf-8'))
        self.assertEqual(requests[3]['body'], literal_body.encode('utf-8'))
        self.assertEqual(requests[3]['headers']['X-Literal'], '${__groovy(40+2)}')

    def test_status_assertions_no_redirect_and_socket_timeout(self):
        rows = [
            {'name': 'expected404', 'method': 'GET', 'url': self.base + '/missing', 'expected_status': '404'},
            {'name': 'failed404', 'method': 'GET', 'url': self.base + '/missing'},
            {'name': 'redirect', 'method': 'GET', 'url': self.base + '/redirect'},
            {'name': 'slow', 'method': 'GET', 'url': self.base + '/slow', 'response_timeout_ms': '100'},
        ]
        code, summary, folder, diagnostics = self.run_tool(rows)
        self.assertEqual(code, 4, diagnostics)
        self.assertEqual(summary['total_samples'], 4)
        self.assertEqual(summary['failed_samples'], 3)
        self.assertEqual(summary['jmeter_exit_code'], 0)
        requests = self.drain_requests()
        self.assertEqual(len(requests), 4)
        self.assertNotIn('/unexpected-followup', [item['path'] for item in requests])
        with (folder / 'results.jtl').open(encoding='utf-8', newline='') as stream:
            samples = list(csv.DictReader(stream))
        self.assertEqual(samples[0]['success'], 'true')
        self.assertTrue(all(sample['success'] == 'false' for sample in samples[1:]))

    def test_thread_and_loop_count(self):
        rows = [{'name': 'counts', 'method': 'GET', 'url': self.base + '/echo'}]
        code, summary, _, diagnostics = self.run_tool(rows, extra=['--threads', '2', '--loops', '3'])
        self.assertEqual(code, 0, diagnostics)
        self.assertEqual(summary['expected_samples'], 6)
        self.assertEqual(summary['total_samples'], 6)
        self.assertEqual(len(self.drain_requests()), 6)

    def test_whole_run_timeout_is_failure(self):
        rows = [{'name': 'deadline', 'method': 'GET', 'url': self.base + '/slow'}]
        code, summary, _, diagnostics = self.run_tool(rows, extra=['--max-run-seconds', '1'])
        self.assertEqual(code, 5, diagnostics)
        self.assertEqual(summary['status'], 'timeout')

    @unittest.skipUnless(os.name != 'nt', 'Native POSIX Bash wrapper test')
    def test_posix_bash_default_one_request(self):
        rows = [{'name': 'Bash_default', 'method': 'GET', 'url': self.base + '/echo'}]
        code, summary, _, diagnostics = self.run_tool(rows, wrapper='bash')
        self.assertEqual(code, 0, diagnostics)
        self.assertEqual(summary['total_samples'], 1)
        self.assertEqual(len(self.drain_requests()), 1)

    @unittest.skipUnless(os.name == 'nt', 'Windows PowerShell wrapper test')
    def test_windows_powershell_success(self):
        rows = [{'name': 'PS_success', 'method': 'GET', 'url': self.base + '/echo'}]
        code, summary, _, diagnostics = self.run_tool(rows, wrapper='powershell', trailing_output=True)
        self.assertEqual(code, 0, diagnostics)
        self.assertEqual(summary['total_samples'], 1)
        self.assertEqual(len(self.drain_requests()), 1)

    @unittest.skipUnless(os.name == 'nt', 'Windows PowerShell wrapper test')
    def test_windows_powershell_default_one_request_and_failure_exit_code(self):
        rows = [{'name': 'PS_default', 'method': 'GET', 'url': self.base + '/missing'}]
        code, summary, _, diagnostics = self.run_tool(rows, '.xlsx', wrapper='powershell')
        self.assertEqual(code, 4, diagnostics)
        self.assertEqual(summary['total_samples'], 1)
        self.assertEqual(len(self.drain_requests()), 1)


if __name__ == '__main__':
    unittest.main()
