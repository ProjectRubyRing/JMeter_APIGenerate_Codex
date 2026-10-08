import csv
import subprocess
import sys
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET

from restjmeter.inputs import validate_request
from restjmeter.plan import write_plan
from restjmeter.runner import read_results, run_process
from restjmeter.workbook import InputError
from test_support import TemporaryWorkspace

ROOT = Path(__file__).resolve().parents[1]


class RunnerTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryWorkspace()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.requests = [validate_request({'name': name, 'method': 'GET', 'url': 'http://example.invalid/'})
                         for name in ('one', 'two')]

    def jtl(self, labels, successful=True):
        path = self.root / 'results.jtl'
        with path.open('w', encoding='utf-8', newline='') as stream:
            writer = csv.writer(stream)
            writer.writerow(['timeStamp', 'elapsed', 'label', 'responseCode', 'success'])
            for label in labels:
                writer.writerow(['1700000000000', '10', label, '200' if successful else '500', str(successful).lower()])
        return path

    def test_jtl_checks_each_api_count_not_only_total(self):
        result = read_results(self.jtl(['one', 'one']), self.requests, 1, 1, self.root)
        self.assertEqual(result['total_samples'], result['expected_samples'])
        self.assertFalse(result['sample_counts_match'])

    def test_jtl_aggregates_failure_and_elapsed_time(self):
        result = read_results(self.jtl(['one', 'two'], False), self.requests, 1, 1, self.root)
        self.assertTrue(result['sample_counts_match'])
        self.assertEqual(result['failed_samples'], 2)
        self.assertEqual(result['mean_elapsed_ms'], 10)

    def test_empty_truncated_or_malformed_jtl_is_error(self):
        path = self.root / 'invalid.jtl'
        for data in ['', 'label,success\na,true\n',
                     'timeStamp,elapsed,label,responseCode,success\n1,2,a,200\n',
                     'timeStamp,elapsed,label,responseCode,success\n1,x,a,200,true\n']:
            with self.subTest(data=data):
                path.write_text(data, encoding='utf-8')
                with self.assertRaises(InputError):
                    read_results(path, self.requests, 1, 1, self.root)

    def test_jmx_xml_preserves_multiline_raw_body_and_has_no_extra_sampler(self):
        body = ' first\r\nsecond <&> "日本語"\n'
        request = validate_request({'name': 'post', 'method': 'POST', 'url': 'http://example.invalid/?x=1',
                                    'body_type': 'raw', 'body': body})
        path = self.root / 'test.jmx'
        write_plan(path, [request])
        root = ET.parse(path)
        self.assertEqual(len(root.findall('.//HTTPSamplerProxy')), 1)
        self.assertEqual(root.find(".//stringProp[@name='Argument.value']").text, body)
        self.assertEqual(root.find(".//stringProp[@name='ThreadGroup.num_threads']").text, '1')
        self.assertEqual(root.find(".//stringProp[@name='LoopController.loops']").text, '1')
        self.assertEqual(root.find(".//boolProp[@name='HTTPSampler.follow_redirects']").text, 'false')

    def test_timeout_reaps_child_process(self):
        with self.assertRaises(subprocess.TimeoutExpired):
            run_process([sys.executable, '-c', 'import time; time.sleep(30)'], self.root, 'console.log', 1)

    def test_cli_generation_handles_spaces_and_japanese_paths(self):
        folder = self.root / 'space & 日本語'
        folder.mkdir()
        input_path = folder / 'input.csv'
        input_path.write_text('name,url,method\na,http://example.invalid,GET\n', encoding='utf-8')
        result = subprocess.run([sys.executable, str(ROOT / 'jmeter_api.py'), '--input', str(input_path),
                                 '--output', str(folder / 'out'), '--generate-only'], capture_output=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr.decode(errors='replace'))
        outputs = list((folder / 'out').iterdir())
        self.assertEqual(len(outputs), 1)
        self.assertTrue((outputs[0] / 'JMeter-User-Guide.xlsx').is_file())
        self.assertFalse((outputs[0] / 'results.jtl').exists())


if __name__ == '__main__':
    unittest.main()
