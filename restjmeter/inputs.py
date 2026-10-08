"""CSV/XLSX schema and request validation shared by both operating systems."""
import csv
import ipaddress
import json
import re
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote, urlencode, urlsplit, urlunsplit

from .workbook import InputError, MAX_FILE_BYTES, MAX_ROWS, check_xml_text, read_xlsx

COLUMNS = ['enabled', 'name', 'url', 'method', 'query_params', 'headers', 'body_type',
           'body', 'connect_timeout_ms', 'response_timeout_ms', 'expected_status']
REQUIRED = {'name', 'url', 'method'}


@dataclass(frozen=True)
class Request:
    name: str
    url: str
    method: str
    headers: list
    body_type: str
    body: str
    connect_timeout_ms: int
    response_timeout_ms: int
    status_pattern: str


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise InputError('Duplicate JSON keys; use a name/value array for repeated parameters.')
        result[key] = value
    return result


def _invalid_constant(value):
    raise InputError('NaN and Infinity are not valid JSON.')


def json_value(value, field):
    try:
        return json.loads(value, object_pairs_hook=_unique_object, parse_constant=_invalid_constant)
    except json.JSONDecodeError as exc:
        raise InputError('%s: invalid JSON at line %d, column %d.' % (field, exc.lineno, exc.colno)) from exc
    except RecursionError as exc:
        raise InputError(field + ': JSON is nested too deeply.') from exc
    except ValueError as exc:
        if isinstance(exc, InputError):
            raise
        raise InputError(field + ': JSON contains an unsupported numeric value.') from exc


def parameter_pairs(value, field):
    if not value.strip():
        return []
    parsed = json_value(value, field)
    if isinstance(parsed, dict):
        pairs = list(parsed.items())
    elif isinstance(parsed, list):
        pairs = []
        for pair in parsed:
            if not isinstance(pair, dict) or set(pair) != {'name', 'value'}:
                raise InputError(field + ': array elements must contain exactly name and value.')
            pairs.append((pair['name'], pair['value']))
    else:
        raise InputError(field + ': use a JSON object or an array of name/value objects.')
    for key, item in pairs:
        if not isinstance(key, str) or not isinstance(item, str):
            raise InputError(field + ': parameter names and values must be JSON strings.')
        check_xml_text(key)
        check_xml_text(item)
    return pairs


def bounded_integer(value, field, maximum=2147483647):
    if (len(str(value)) > len(str(maximum)) or not re.fullmatch(r'[0-9]+', str(value))
            or not 1 <= int(value) <= maximum):
        raise InputError('%s: specify an integer from 1 to %d.' % (field, maximum))
    return int(value)


def normalize_url(value, query_pairs):
    if not value or any(ord(c) < 32 or ord(c) == 127 for c in value):
        raise InputError('url: an absolute HTTP/HTTPS URL without control characters is required.')
    try:
        parts = urlsplit(value)
        port = parts.port
        hostname = parts.hostname
        if (parts.scheme.lower() not in ('http', 'https') or not hostname or
                parts.username is not None or parts.password is not None or '#' in value):
            raise InputError('url: use HTTP/HTTPS without user information or fragments.')
        if parts.netloc.endswith(':') or port == 0:
            raise InputError('url: port must be between 1 and 65535.')
        if ':' in hostname:
            hostname = '[' + str(ipaddress.IPv6Address(hostname)) + ']'
        else:
            hostname = hostname.encode('idna').decode('ascii')
            if not re.fullmatch(r'[A-Za-z0-9.-]+', hostname) or len(hostname) > 253:
                raise InputError('url: invalid hostname.')
        netloc = hostname + (':' + str(port) if port is not None else '')
        if re.search(r'%(?![0-9A-Fa-f]{2})', parts.path + parts.query):
            raise InputError('url: invalid percent escape; use %25 for a literal percent sign.')
        path = quote(parts.path or '/', safe="/%:@!$&'()*+,;=-._~")
        query = quote(parts.query, safe="%/:?@!$&'()*+,;=-._~")
        extra = urlencode(query_pairs, encoding='utf-8')
        if extra:
            query += ('&' if query else '') + extra
        return urlunsplit((parts.scheme.lower(), netloc, path, query, ''))
    except (ValueError, UnicodeError) as exc:
        if isinstance(exc, InputError):
            raise
        raise InputError('url: invalid host, port, or Unicode encoding.') from exc


def status_pattern(value):
    if not value:
        value = '200-299'
    statuses = set()
    for segment in value.split(','):
        match = re.fullmatch(r'([1-5][0-9]{2})(?:-([1-5][0-9]{2}))?', segment.strip())
        if not match:
            raise InputError('expected_status: use codes/ranges such as 200,201 or 200-299.')
        start, end = int(match[1]), int(match[2] or match[1])
        if start > end:
            raise InputError('expected_status: range start must not exceed its end.')
        statuses.update(range(start, end + 1))
    return '^(?:' + '|'.join(str(code) for code in sorted(statuses)) + ')$'


def validate_request(row):
    name = row.get('name', '').strip()
    if not name or len(name) > 120 or any(ord(c) < 32 or ord(c) == 127 for c in name):
        raise InputError('name: use 1-120 characters without control characters.')
    if '${' in name:
        raise InputError('name: JMeter expressions are not allowed in labels.')
    method = row.get('method', '').strip().upper()
    if method not in ('GET', 'POST'):
        raise InputError('method: only GET and POST are supported.')
    url = normalize_url(row.get('url', '').strip(), parameter_pairs(row.get('query_params', ''), 'query_params'))
    headers = parameter_pairs(row.get('headers', ''), 'headers')
    seen = set()
    for name_, value in headers:
        key = name_.lower()
        if not re.fullmatch(r"[!#$%&'*+.^_`|~0-9A-Za-z-]+", name_):
            raise InputError('headers: invalid HTTP header name.')
        if any(ord(ch) < 32 or ord(ch) > 126 for ch in value):
            raise InputError('headers: values must be printable ASCII (encode non-ASCII values).')
        if key in seen:
            raise InputError('headers: duplicate header names are not allowed (case insensitive).')
        if key in ('content-length', 'transfer-encoding'):
            raise InputError('headers: Content-Length and Transfer-Encoding are managed by JMeter.')
        seen.add(key)
    body_type = row.get('body_type', '').strip().lower() or 'none'
    body = row.get('body', '')
    if body_type not in ('none', 'json', 'form', 'raw'):
        raise InputError('body_type: choose none, json, form, or raw.')
    if method == 'GET' and (body_type != 'none' or body):
        raise InputError('GET cannot have a request body; use query_params.')
    if body_type == 'none' and body:
        raise InputError('body: specify body_type when a body is present.')
    if body_type == 'json':
        json_value(body, 'body')
    elif body_type == 'form':
        body = urlencode(parameter_pairs(body, 'body'), encoding='utf-8')
    defaults = {'json': 'application/json; charset=UTF-8',
                'form': 'application/x-www-form-urlencoded; charset=UTF-8',
                'raw': 'text/plain; charset=UTF-8'}
    content_type = next((v for k, v in headers if k.lower() == 'content-type'), '')
    if 'content-type' in seen and not content_type.strip():
        raise InputError('headers: Content-Type must not be empty.')
    if content_type:
        charset = re.search(r'charset\s*=\s*"?([^;"\s]+)', content_type, re.IGNORECASE)
        if charset and charset[1].lower() not in ('utf-8', 'utf8'):
            raise InputError('headers: request body encoding is UTF-8; use a UTF-8 charset.')
        media = content_type.split(';')[0].strip().lower()
        if body_type == 'json' and media != 'application/json' and not media.endswith('+json'):
            raise InputError('headers: json body requires a JSON Content-Type.')
        if body_type == 'form' and media != 'application/x-www-form-urlencoded':
            raise InputError('headers: form body requires application/x-www-form-urlencoded.')
    elif body_type in defaults:
        headers.append(('Content-Type', defaults[body_type]))
    return Request(name, url, method, headers, body_type, body,
                   bounded_integer(row.get('connect_timeout_ms', '').strip() or '10000', 'connect_timeout_ms'),
                   bounded_integer(row.get('response_timeout_ms', '').strip() or '30000', 'response_timeout_ms'),
                   status_pattern(row.get('expected_status', '').strip()))


def read_requests(path):
    path = Path(path)
    if path.stat().st_size > MAX_FILE_BYTES:
        raise InputError('Input file exceeds 10 MiB.')
    if path.suffix.lower() == '.xlsx':
        rows = read_xlsx(path)
    elif path.suffix.lower() == '.csv':
        csv.field_size_limit(MAX_FILE_BYTES)
        try:
            with path.open(encoding='utf-8-sig', newline='') as source:
                reader = csv.reader(source, strict=True)
                rows = [(i, row) for i, row in enumerate(reader, 1) if any(row)]
        except (UnicodeError, csv.Error) as exc:
            raise InputError('CSV must be valid UTF-8 (BOM allowed), comma separated, with quoted multiline fields.') from exc
    else:
        raise InputError('Input extension must be .xlsx or .csv.')
    if not rows or rows[0][0] != 1:
        raise InputError('Place column names in the first row.')
    headers = [value.strip() for value in rows[0][1]]
    while headers and not headers[-1]:
        headers.pop()
    if len(headers) != len(set(headers)) or not REQUIRED.issubset(headers):
        raise InputError('Column names must be unique and include name, url, method.')
    if set(headers) - set(COLUMNS):
        raise InputError('Unknown or blank column name; use the supplied template column names.')
    if len(rows) > MAX_ROWS:
        raise InputError('At most 10000 API rows are supported.')
    requests = []
    seen_names = set()
    for number, values in rows[1:]:
        try:
            if any(values[len(headers):]):
                raise InputError('Data exists beyond the defined columns.')
            row = dict(zip(headers, values + [''] * (len(headers) - len(values))))
            enabled = row.get('enabled', '').strip().lower() or 'true'
            if enabled not in ('true', 'false', '1', '0'):
                raise InputError('enabled: use true/false or 1/0.')
            if enabled in ('false', '0'):
                continue
            for value in row.values():
                check_xml_text(value)
            request = validate_request(row)
            if request.name in seen_names:
                raise InputError('name: enabled API names must be unique.')
            seen_names.add(request.name)
            requests.append(request)
        except InputError as exc:
            raise InputError('Row %d: %s' % (number, exc)) from exc
    if not requests:
        raise InputError('No enabled API rows were found.')
    return requests
