"""Small OOXML reader/writer for text tables; Excel installation is not required."""
import posixpath
import re
import math
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'s': MAIN, 'r': REL}
MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_XML_BYTES = 50 * 1024 * 1024
MAX_ROWS = 10001
MAX_COLS = 64


class InputError(ValueError):
    """An actionable error that is safe to display without input values."""


def xml_bytes(root):
    # XML processors normalize literal CR. Character references preserve request bytes.
    return ET.tostring(root, encoding='utf-8', xml_declaration=True).replace(b'\r', b'&#13;')


def check_xml_text(value):
    if any(ord(ch) < 32 and ch not in '\t\r\n' or 0xD800 <= ord(ch) <= 0xDFFF
           or ord(ch) in (0xFFFE, 0xFFFF) for ch in value):
        raise InputError('XML cannot represent a control character in the input.')


def column_name(number):
    result = ''
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def column_number(reference):
    match = re.fullmatch(r'([A-Z]+)([1-9][0-9]*)', reference)
    if not match:
        raise InputError('Invalid XLSX cell reference.')
    number = 0
    for char in match[1]:
        number = number * 26 + ord(char) - 64
    if number > MAX_COLS or int(match[2]) > MAX_ROWS:
        raise InputError('XLSX exceeds the supported table size (10001 rows / 64 columns).')
    return number - 1, int(match[2])


def unescape_excel(value):
    # A single pass preserves literal escapes represented as _x005F_xNNNN_.
    return re.sub(r'_x([0-9A-Fa-f]{4})_', lambda m: chr(int(m[1], 16)), value)


def _rich_text(element):
    if element is None:
        return ''
    return unescape_excel(''.join(node.text or '' for node in element.findall('s:t', NS)
                                 + element.findall('s:r/s:t', NS)))


def read_xlsx(path, sheet_name='APIs'):
    """Return (Excel row number, cell strings), including shared/sparse Excel cells."""
    try:
        if Path(path).stat().st_size > MAX_FILE_BYTES:
            raise InputError('Input file exceeds 10 MiB.')
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)) or len(names) > 2000:
                raise InputError('Invalid or excessively complex XLSX archive.')
            if sum(info.file_size for info in archive.infolist()) > MAX_XML_BYTES:
                raise InputError('Expanded XLSX exceeds 50 MiB.')

            def read_xml(name):
                data = archive.read(name)
                if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
                    raise InputError('XLSX XML declarations with entities are not supported.')
                return ET.fromstring(data)

            workbook = read_xml('xl/workbook.xml')
            if workbook.tag != '{' + MAIN + '}workbook':
                raise InputError('Save as a standard Excel Workbook (.xlsx).')
            relations = {el.get('Id'): el for el in read_xml('xl/_rels/workbook.xml.rels')}
            sheet = next((el for el in workbook.findall('s:sheets/s:sheet', NS)
                          if el.get('name') == sheet_name), None)
            if sheet is None:
                raise InputError("XLSX must contain a worksheet named 'APIs'.")

            def part_path(relation):
                if relation is None or relation.get('TargetMode') == 'External':
                    raise InputError('External or missing XLSX relationship.')
                target = relation.get('Target', '')
                if '\\' in target:
                    raise InputError('Invalid XLSX relationship path.')
                resolved = posixpath.normpath(posixpath.join('xl', target))
                if target.startswith('/'):
                    resolved = posixpath.normpath(target.lstrip('/'))
                if not resolved.startswith('xl/'):
                    raise InputError('XLSX relationship is outside the workbook.')
                return resolved

            shared = []
            for relation in relations.values():
                if relation.get('Type') == REL + '/sharedStrings':
                    shared = [_rich_text(el) for el in read_xml(part_path(relation))]
            worksheet = read_xml(part_path(relations.get(sheet.get('{' + REL + '}id'))))
            if worksheet.find('s:mergeCells', NS) is not None:
                raise InputError('Merged cells are not supported on the APIs worksheet.')
            rows = []
            seen_rows = set()
            for row in worksheet.findall('s:sheetData/s:row', NS):
                row_number = int(row.get('r', '0'))
                if not 1 <= row_number <= MAX_ROWS or row_number in seen_rows:
                    raise InputError('Invalid or duplicate XLSX row number.')
                seen_rows.add(row_number)
                values = {}
                for cell in row.findall('s:c', NS):
                    col, cell_row = column_number(cell.get('r', ''))
                    if cell_row != row_number or col in values:
                        raise InputError('Duplicate or misplaced XLSX cell.')
                    location = 'APIs!' + cell.get('r', '')
                    if cell.find('s:f', NS) is not None:
                        raise InputError(location + ': formulas are not allowed; paste values instead.')
                    kind = cell.get('t', 'n')
                    value = cell.findtext('s:v', default='', namespaces=NS)
                    if kind == 's':
                        index = int(value)
                        if not 0 <= index < len(shared):
                            raise InputError(location + ': invalid shared string index.')
                        value = shared[index]
                    elif kind == 'inlineStr':
                        value = _rich_text(cell.find('s:is', NS))
                    elif kind == 'b':
                        if value not in ('0', '1'):
                            raise InputError(location + ': invalid boolean.')
                        value = 'true' if value == '1' else 'false'
                    elif kind == 'str':
                        value = unescape_excel(value)
                    elif kind != 'n':
                        raise InputError(location + ': unsupported cell type; use text or numbers.')
                    values[col] = value
                if any(values.values()):
                    rows.append((row_number, [values.get(i, '') for i in range(max(values) + 1)]))
            return sorted(rows)
    except InputError:
        raise
    except (zipfile.BadZipFile, ET.ParseError, KeyError, ValueError, RuntimeError) as exc:
        raise InputError('Invalid XLSX workbook structure; save again as .xlsx.') from exc


def write_xlsx(path, sheets):
    """Write (sheet name, rows, widths) tables as text, with filters and frozen headers."""
    workbook = ET.Element('workbook', xmlns=MAIN, attrib={'xmlns:r': REL})
    sheet_list = ET.SubElement(workbook, 'sheets')
    relations = ET.Element('Relationships', xmlns=PKG)
    types = ET.Element('Types', xmlns='http://schemas.openxmlformats.org/package/2006/content-types')
    ET.SubElement(types, 'Default', Extension='rels', ContentType='application/vnd.openxmlformats-package.relationships+xml')
    ET.SubElement(types, 'Default', Extension='xml', ContentType='application/xml')
    for part, content in [('workbook', 'sheet.main'), ('styles', 'styles')]:
        ET.SubElement(types, 'Override', PartName='/xl/' + part + '.xml',
                      ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.' + content + '+xml')
    root_rels = ET.Element('Relationships', xmlns=PKG)
    ET.SubElement(root_rels, 'Relationship', Id='rId1', Type=REL + '/officeDocument', Target='xl/workbook.xml')
    styles = '''<styleSheet xmlns="%s"><fonts count="2"><font><sz val="11"/><name val="Yu Gothic"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Yu Gothic"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF24476B"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyNumberFormat="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="49" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>''' % MAIN
    with zipfile.ZipFile(path, 'x', compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('_rels/.rels', xml_bytes(root_rels))
        archive.writestr('xl/styles.xml', styles.encode('utf-8'))
        for number, (name, rows, widths) in enumerate(sheets, 1):
            ET.SubElement(sheet_list, 'sheet', name=name, sheetId=str(number), attrib={'r:id': 'rId' + str(number)})
            ET.SubElement(relations, 'Relationship', Id='rId' + str(number), Type=REL + '/worksheet', Target='worksheets/sheet%d.xml' % number)
            ET.SubElement(types, 'Override', PartName='/xl/worksheets/sheet%d.xml' % number,
                          ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
            sheet = ET.Element('worksheet', xmlns=MAIN)
            views = ET.SubElement(sheet, 'sheetViews')
            view = ET.SubElement(views, 'sheetView', workbookViewId='0')
            ET.SubElement(view, 'pane', ySplit='1', topLeftCell='A2', activePane='bottomLeft', state='frozen')
            ET.SubElement(sheet, 'sheetFormatPr', defaultRowHeight='30')
            cols = ET.SubElement(sheet, 'cols')
            for i, width in enumerate(widths, 1):
                ET.SubElement(cols, 'col', min=str(i), max=str(i), width=str(width), customWidth='1')
            data = ET.SubElement(sheet, 'sheetData')
            for row_number, values in enumerate(rows, 1):
                lines = 1
                for index, item in enumerate(values):
                    width = max(8, widths[index] - 3) if index < len(widths) else 20
                    needed = sum(max(1, math.ceil(sum(2 if unicodedata.east_asian_width(char) in ('W', 'F') else 1
                                                      for char in line) / width))
                                 for line in str(item).split('\n'))
                    lines = max(lines, needed)
                row = ET.SubElement(data, 'row', r=str(row_number), ht=str(min(409, max(30, lines * 17))), customHeight='1')
                for col_number, value in enumerate(values, 1):
                    value = str(value)
                    check_xml_text(value)
                    if len(value.encode('utf-16-le')) // 2 > 32767:
                        raise InputError('Excel cells support at most 32767 UTF-16 characters.')
                    value = re.sub(r'_x[0-9A-Fa-f]{4}_', lambda m: '_x005F_' + m[0][1:], value)
                    cell = ET.SubElement(row, 'c', r=column_name(col_number) + str(row_number),
                                         t='inlineStr', s='1' if row_number == 1 else '0')
                    text = ET.SubElement(ET.SubElement(cell, 'is'), 't', attrib={'xml:space': 'preserve'})
                    text.text = value
            if rows:
                ET.SubElement(sheet, 'autoFilter', ref='A1:%s%d' % (column_name(len(rows[0])), len(rows)))
            if name == 'APIs':
                validations = ET.SubElement(sheet, 'dataValidations', count='3')
                for column, choices in [('A', 'true,false'), ('D', 'GET,POST'), ('G', 'none,json,form,raw')]:
                    validation = ET.SubElement(validations, 'dataValidation', type='list', allowBlank='1',
                                               showErrorMessage='1', errorTitle='Invalid choice',
                                               error='Select a value from the list.', sqref=column + '2:' + column + '10001')
                    ET.SubElement(validation, 'formula1').text = '"' + choices + '"'
            archive.writestr('xl/worksheets/sheet%d.xml' % number, xml_bytes(sheet))
        ET.SubElement(relations, 'Relationship', Id='rIdStyles', Type=REL + '/styles', Target='styles.xml')
        archive.writestr('xl/workbook.xml', xml_bytes(workbook))
        archive.writestr('xl/_rels/workbook.xml.rels', xml_bytes(relations))
        archive.writestr('[Content_Types].xml', xml_bytes(types))
