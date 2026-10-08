import csv
import json
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from restjmeter.documentation import COLUMNS, EXAMPLES, create_templates
from restjmeter.inputs import read_requests
from restjmeter.workbook import InputError, MAIN, NS, REL, read_xlsx, write_xlsx
from test_support import TemporaryWorkspace


class InputTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryWorkspace()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def csv_file(self, overrides=None, rows=None, encoding='utf-8-sig'):
        row = dict(zip(COLUMNS, EXAMPLES[0]))
        row.update(overrides or {})
        path = self.root / 'input.csv'
        with path.open('w', encoding=encoding, newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=COLUMNS)
            writer.writeheader()
            writer.writerows(rows or [row])
        return path

    def modify_xlsx(self, path, transform):
        with zipfile.ZipFile(path) as archive:
            files = {name: archive.read(name) for name in archive.namelist()}
        transform(files)
        with zipfile.ZipFile(path, 'w') as archive:
            for name, data in files.items():
                archive.writestr(name, data)

    def test_templates_are_equivalent_and_only_one_api_is_enabled(self):
        create_templates(self.root)
        csv_requests = read_requests(self.root / 'api-template.csv')
        self.assertEqual(csv_requests, read_requests(self.root / 'api-template.xlsx'))
        self.assertEqual(len(csv_requests), 1)
        self.assertTrue((self.root / 'api-template.csv').read_bytes().startswith(b'\xef\xbb\xbf'))
        with self.assertRaises(InputError):
            create_templates(self.root)

    def test_unicode_multiline_body_and_duplicate_post_query(self):
        body = '  first, "quoted"\r\n日本語\\tail\n'
        path = self.csv_file({'method': 'POST', 'url': 'https://example.invalid/a%2Fb?existing=a%2Bb',
             'query_params': '[{"name":"tag","value":"a &"},{"name":"tag","value":"日本語"}]',
             'body_type': 'raw', 'body': body})
        request = read_requests(path)[0]
        self.assertEqual(request.body, body)
        self.assertIn('/a%2Fb?existing=a%2Bb&tag=a+%26&tag=', request.url)
        self.assertEqual(request.method, 'POST')

    def test_form_body_keeps_duplicates_blank_and_special_values(self):
        body = '[{"name":"x","value":"a+b &"},{"name":"x","value":""}]'
        request = read_requests(self.csv_file({'method': 'POST', 'body_type': 'form', 'body': body}))[0]
        self.assertEqual(request.body, 'x=a%2Bb+%26&x=')
        self.assertIn(('Content-Type', 'application/x-www-form-urlencoded; charset=UTF-8'), request.headers)

    def test_url_unicode_ipv6_and_port(self):
        request = read_requests(self.csv_file({'url': 'http://[::1]:8081/日本語 a?x=%2B'}))[0]
        self.assertTrue(request.url.startswith('http://[::1]:8081/%E6%97%A5'))
        self.assertIn('%20a?x=%2B', request.url)

    def test_csv_without_bom(self):
        self.assertEqual(len(read_requests(self.csv_file(encoding='utf-8'))), 1)

    def test_invalid_fields_fail_with_row_context(self):
        cases = [
            {'method': 'DELETE'}, {'url': 'file:///tmp/test'}, {'url': 'https://user:pass@example.invalid/'},
            {'url': 'https://example.invalid/#fragment'}, {'url': 'https://example.invalid:0/'},
            {'url': 'https://example.invalid:abc/'}, {'url': 'https://example.invalid/%xx'},
            {'body': 'not allowed'}, {'method': 'POST', 'body_type': 'json', 'body': '{bad}'},
            {'method': 'POST', 'body_type': 'json', 'body': '{"x":NaN}'},
            {'query_params': '{"x":"a","x":"b"}'}, {'query_params': '{"x":12}'},
            {'headers': '{"X-A":"one","x-a":"two"}'}, {'headers': '{"X-A":"a\\r\\nb"}'},
            {'headers': '{"Content-Length":"1"}'}, {'headers': '{"X-Test":"日本語"}'},
            {'headers': '{"Content-Type":""}'},
            {'headers': '{"Content-Type":"application/json; charset=Shift_JIS"}'},
            {'connect_timeout_ms': '0'}, {'response_timeout_ms': '-1'}, {'response_timeout_ms': '1.5'},
            {'expected_status': '299-200'}, {'expected_status': '600'}, {'enabled': 'maybe'},
            {'name': '${__groovy(1)}'}, {'name': 'line\nbreak'},
        ]
        for case in cases:
            with self.subTest(case=list(case)):
                with self.assertRaisesRegex(InputError, 'Row 2:'):
                    read_requests(self.csv_file(case))

    def test_errors_do_not_echo_credentials_or_body(self):
        marker = 'private-placeholder-value'
        with self.assertRaises(InputError) as caught:
            read_requests(self.csv_file({'method': 'POST', 'body_type': 'json', 'body': marker}))
        self.assertNotIn(marker, str(caught.exception))

    def test_disabled_rows_are_skipped_and_zero_enabled_is_error(self):
        base = dict(zip(COLUMNS, EXAMPLES[0]))
        rows = [base, {'enabled': 'false', 'method': 'INVALID', 'url': 'bad'}]
        self.assertEqual(len(read_requests(self.csv_file(rows=rows))), 1)
        with self.assertRaisesRegex(InputError, 'No enabled'):
            read_requests(self.csv_file({'enabled': 'false'}))

    def test_duplicate_names_and_extra_columns_rejected(self):
        base = dict(zip(COLUMNS, EXAMPLES[0]))
        with self.assertRaisesRegex(InputError, 'unique'):
            read_requests(self.csv_file(rows=[base, base]))
        path = self.root / 'extra.csv'
        path.write_text('name,url,method\na,http://example.invalid,GET,extra\n', encoding='utf-8')
        with self.assertRaisesRegex(InputError, 'beyond'):
            read_requests(path)
        path.write_text('name,url,method,unknown\na,http://example.invalid,GET,\n', encoding='utf-8')
        with self.assertRaisesRegex(InputError, 'Unknown'):
            read_requests(path)

    def test_required_columns_only_and_default_status(self):
        path = self.root / 'minimal.csv'
        path.write_text('method,name,url\nGET,minimal,https://example.invalid\n', encoding='utf-8')
        request = read_requests(path)[0]
        self.assertEqual(request.url, 'https://example.invalid/')
        self.assertEqual(request.connect_timeout_ms, 10000)
        self.assertIn('299', request.status_pattern)

    def test_xlsx_preserves_excel_escape_sequences_and_crlf(self):
        path = self.root / 'escapes.xlsx'
        values = ['literal _x0041_ and _x005F_', 'first\r\nsecond', '=not-a-formula', '日本語']
        write_xlsx(path, [('APIs', [values], [30] * 4)])
        self.assertEqual(read_xlsx(path)[0][1], values)

    def test_xlsx_shared_strings_sparse_cells_and_sheet_relationship(self):
        path = self.root / 'shared.xlsx'
        write_xlsx(path, [('Notes', [['notes']], [20]), ('APIs', [COLUMNS, EXAMPLES[0]], [20] * len(COLUMNS))])

        def transform(files):
            sheet = ET.fromstring(files['xl/worksheets/sheet2.xml'])
            cell = sheet.find(".//s:c[@r='B2']", NS)
            cell.clear()
            cell.attrib.update({'r': 'B2', 't': 's'})
            ET.SubElement(cell, '{' + MAIN + '}v').text = '0'
            row = sheet.find(".//s:row[@r='2']", NS)
            empty = row.find("s:c[@r='H2']", NS)
            row.remove(empty)
            files['xl/worksheets/sheet2.xml'] = ET.tostring(sheet)
            files['xl/sharedStrings.xml'] = ('<sst xmlns="' + MAIN + '"><si><r><t>shared</t></r><r><t>_name</t></r></si></sst>').encode()
            relations = ET.fromstring(files['xl/_rels/workbook.xml.rels'])
            ET.SubElement(relations, 'Relationship', Id='shared', Type=REL + '/sharedStrings', Target='/xl/sharedStrings.xml')
            files['xl/_rels/workbook.xml.rels'] = ET.tostring(relations)

        self.modify_xlsx(path, transform)
        request = read_requests(path)[0]
        self.assertEqual(request.name, 'shared_name')
        self.assertEqual(request.body, '')

    def test_xlsx_formula_and_external_relationship_are_rejected(self):
        path = self.root / 'formula.xlsx'
        write_xlsx(path, [('APIs', [COLUMNS, EXAMPLES[0]], [20] * len(COLUMNS))])

        def formula(files):
            sheet = ET.fromstring(files['xl/worksheets/sheet1.xml'])
            cell = sheet.find(".//s:c[@r='B2']", NS)
            ET.SubElement(cell, '{' + MAIN + '}f').text = '1+1'
            files['xl/worksheets/sheet1.xml'] = ET.tostring(sheet)

        self.modify_xlsx(path, formula)
        with self.assertRaisesRegex(InputError, 'formulas'):
            read_requests(path)

        def external(files):
            relations = ET.fromstring(files['xl/_rels/workbook.xml.rels'])
            relations[0].set('TargetMode', 'External')
            files['xl/_rels/workbook.xml.rels'] = ET.tostring(relations)

        self.modify_xlsx(path, external)
        with self.assertRaisesRegex(InputError, 'External'):
            read_requests(path)


if __name__ == '__main__':
    unittest.main()
