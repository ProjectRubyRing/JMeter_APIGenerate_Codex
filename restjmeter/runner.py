"""Start Java directly, avoiding .bat/shell quoting and checking actual JTL results."""
import csv
import json
import os
import re
import shutil
import subprocess
import zipfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from .workbook import InputError


def resolve_runtime(jmeter_home, java_path=None):
    if not jmeter_home:
        raise InputError('Set JMETER_HOME or specify --jmeter-home (JMeter 5.6.3 directory).')
    home = Path(jmeter_home).expanduser().resolve()
    needed = ['bin/ApacheJMeter.jar', 'bin/jmeter.properties', 'bin/log4j2.xml', 'lib/ext/ApacheJMeter_core.jar']
    if any(not (home / file).is_file() for file in needed):
        raise InputError('JMeter home must contain the complete Apache JMeter 5.6.3 binary distribution.')
    try:
        with zipfile.ZipFile(home / 'lib/ext/ApacheJMeter_core.jar') as jar:
            manifest = jar.read('META-INF/MANIFEST.MF').decode('utf-8')
        if not re.search(r'^Implementation-Version: 5\.6\.3\s*$', manifest, re.MULTILINE):
            raise InputError('Only Apache JMeter 5.6.3 is supported; check --jmeter-home.')
    except (KeyError, zipfile.BadZipFile, UnicodeError) as exc:
        raise InputError('Cannot verify the JMeter core JAR version.') from exc
    executable = 'java.exe' if os.name == 'nt' else 'java'
    if java_path:
        java = shutil.which(java_path)
    elif os.environ.get('JAVA_HOME'):
        java = str(Path(os.environ['JAVA_HOME']) / 'bin' / executable)
        if not Path(java).is_file():
            java = None
    else:
        java = shutil.which(executable)
    if not java:
        raise InputError('Java was not found; set JAVA_HOME, PATH, or --java. Java 17 is recommended.')
    try:
        probe = subprocess.run([java, '-version'], capture_output=True, timeout=20, check=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise InputError('Could not start Java or Java version detection timed out.') from exc
    version_text = (probe.stdout + probe.stderr).decode('utf-8', errors='replace')
    match = re.search(r'version\s+"(\d+)(?:\.(\d+))?', version_text)
    major = int(match[2] if match and match[1] == '1' else match[1]) if match else 0
    if probe.returncode != 0 or major < 8:
        raise InputError('Java 8 or later is required; Java 17 is recommended.')
    return home, str(Path(java).resolve()), major


def java_command(home, java, output):
    # Do not launch batch files, invoke a shell, or inherit launcher-only JVM_ARGS.
    return [java, '-Xms128m', '-Xmx512m', '-Djava.awt.headless=true', '-Dfile.encoding=UTF-8',
            '-Duser.language=en', '-Duser.country=US', '-jar', str(home / 'bin/ApacheJMeter.jar'),
            '-d', str(home), '-p', str(home / 'bin/jmeter.properties'),
            '-q', str(output / 'runtime.properties'), '-i', str(home / 'bin/log4j2.xml')]


def append_log(output, message):
    timestamp = datetime.now(timezone.utc).isoformat(timespec='seconds')
    with (output / 'execution.log').open('a', encoding='utf-8', newline='\n') as log:
        log.write(timestamp + ' ' + message + '\n')


def run_process(command, output, log_name, timeout):
    """A single JVM is the child; timeout/Ctrl+C always terminates and reaps it."""
    with (output / log_name).open('wb') as console:
        process = subprocess.Popen(command, cwd=output, stdout=console, stderr=subprocess.STDOUT,
                                   stdin=subprocess.DEVNULL, shell=False)
        try:
            return process.wait(timeout=timeout)
        except (subprocess.TimeoutExpired, KeyboardInterrupt):
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=10)
            raise


def read_results(path, requests, threads, loops, output):
    expected = Counter({request.name: threads * loops for request in requests})
    actual = Counter()
    total = failed = 0
    elapsed_total = 0
    try:
        with path.open(encoding='utf-8-sig', newline='') as source:
            reader = csv.DictReader(source, strict=True)
            required = {'label', 'success', 'responseCode', 'elapsed', 'timeStamp'}
            if not reader.fieldnames or not required.issubset(reader.fieldnames):
                raise InputError('results.jtl is missing required CSV columns.')
            for row in reader:
                if None in row or any(row.get(field) is None for field in required):
                    raise InputError('results.jtl contains an incomplete sample.')
                if row['success'] not in ('true', 'false'):
                    raise InputError('results.jtl contains an invalid success flag.')
                elapsed = int(row['elapsed'])
                timestamp = int(row['timeStamp'])
                if elapsed < 0 or timestamp < 0:
                    raise InputError('results.jtl contains an invalid time value.')
                total += 1
                failed += row['success'] != 'true'
                elapsed_total += elapsed
                actual[row['label']] += 1
                # Only fixed-format metrics are logged; request/response content is not copied.
                code = row['responseCode'] if re.fullmatch(r'[1-5][0-9]{2}', row['responseCode']) else 'transport_error'
                append_log(output, 'SAMPLE index=%d code=%s elapsed_ms=%d success=%s' %
                           (total, code, elapsed, row['success']))
    except (OSError, UnicodeError, csv.Error, ValueError) as exc:
        if isinstance(exc, InputError):
            raise
        raise InputError('Cannot read complete JTL results; inspect jmeter.log and console.log.') from exc
    return {'expected_samples': sum(expected.values()), 'total_samples': total,
            'failed_samples': failed, 'passed_samples': total - failed,
            'sample_counts_match': actual == expected,
            'mean_elapsed_ms': round(elapsed_total / total, 3) if total else None}


def save_summary(output, summary):
    (output / 'summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def execute(home, java, java_major, output, requests, threads, loops, timeout, html_report=False):
    summary = {'jmeter_version': '5.6.3', 'java_major': java_major,
               'expected_samples': len(requests) * threads * loops,
               'status': 'running', 'html_report': 'not_requested'}
    save_summary(output, summary)
    append_log(output, 'START expected_samples=%d' % summary['expected_samples'])
    command = java_command(home, java, output) + ['-n', '-t', str(output / 'test-plan.jmx'),
              '-j', 'jmeter.log', '-Jrest.results.file=results.jtl']
    exit_code = 3
    try:
        code = run_process(command, output, 'console.log', timeout)
        summary['jmeter_exit_code'] = code
        summary.update(read_results(output / 'results.jtl', requests, threads, loops, output))
        if code != 0 or not summary['sample_counts_match']:
            summary['status'] = 'execution_error'
            exit_code = 3
        elif summary['failed_samples']:
            summary['status'] = 'api_failed'
            exit_code = 4
        else:
            summary['status'] = 'passed'
            exit_code = 0
        if html_report:
            report_command = java_command(home, java, output) + [
                '-g', str(output / 'results.jtl'), '-o', str(output / 'html-report'), '-j', 'report.log']
            report_code = run_process(report_command, output, 'report-console.log', timeout)
            report_ok = report_code == 0 and (output / 'html-report/index.html').is_file()
            summary['html_report'] = 'generated' if report_ok else 'failed'
            if not report_ok:
                summary['status'] = 'report_error'
                exit_code = 3
    except subprocess.TimeoutExpired:
        summary['status'] = 'timeout'
        exit_code = 5
    except KeyboardInterrupt:
        summary['status'] = 'interrupted'
        exit_code = 130
    except (OSError, InputError):
        summary['status'] = 'execution_error'
        append_log(output, 'ERROR unable to complete execution or read JTL; inspect JMeter logs')
        exit_code = 3
    summary['exit_code'] = exit_code
    save_summary(output, summary)
    append_log(output, 'END status=%s exit_code=%d' % (summary['status'], exit_code))
    return exit_code, summary
