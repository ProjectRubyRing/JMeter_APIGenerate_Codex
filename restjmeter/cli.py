"""User-facing command line interface and artifact creation."""
import argparse
import os
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

from .documentation import create_templates, write_guide
from .inputs import bounded_integer, read_requests
from .plan import RUNTIME_PROPERTIES, write_plan
from .runner import execute, resolve_runtime
from .workbook import InputError


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description='Generate and run Apache JMeter 5.6.3 REST API plans from XLSX or UTF-8 CSV.',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('--input', type=Path, help='Input .xlsx (APIs sheet) or UTF-8 .csv')
    parser.add_argument('--jmeter-home', default=os.environ.get('JMETER_HOME'), help='JMeter 5.6.3 installation directory')
    parser.add_argument('--java', help='Java executable; otherwise JAVA_HOME/bin/java or PATH')
    parser.add_argument('--output', type=Path, default=Path('output'), help='Parent directory for a new, unique run directory')
    parser.add_argument('--threads', default='1', help='Concurrent users, 1-10000')
    parser.add_argument('--loops', default='1', help='Executions per API per user, 1-1000000')
    parser.add_argument('--max-run-seconds', default='300', help='Timeout in seconds for each JMeter process')
    parser.add_argument('--generate-only', action='store_true', help='Generate JMX, settings and Excel guide without sending requests')
    parser.add_argument('--html-report', action='store_true', help='Also generate an HTML dashboard after execution')
    parser.add_argument('--create-templates', action='store_true', help='Create CSV/XLSX templates and the Excel guide without running JMeter')
    parser.add_argument('--template-dir', type=Path, default=Path('templates'), help='Destination for --create-templates; existing files are not overwritten')
    args = parser.parse_args(argv)
    if args.create_templates and args.input:
        parser.error('--create-templates and --input cannot be combined')
    if not args.create_templates and not args.input:
        parser.error('--input is required unless --create-templates is specified')
    if args.generate_only and args.html_report:
        parser.error('--html-report requires execution; remove --generate-only')
    return args


def main(argv=None):
    # RHEL output directories/files containing request definitions are private by default.
    # Windows permissions follow the user's directory ACL.
    if os.name != 'nt':
        os.umask(0o077)
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(errors='backslashreplace')
        sys.stderr.reconfigure(errors='backslashreplace')
    try:
        args = parse_args(argv)
        if args.create_templates:
            create_templates(args.template_dir.expanduser().resolve())
            print('Templates created: ' + str(args.template_dir.resolve()))
            return 0
        threads = bounded_integer(args.threads, 'threads', 10000)
        loops = bounded_integer(args.loops, 'loops', 1000000)
        timeout = bounded_integer(args.max_run_seconds, 'max-run-seconds')
        requests = read_requests(args.input.expanduser().resolve())
        runtime = None if args.generate_only else resolve_runtime(args.jmeter_home, args.java)
        parent = args.output.expanduser().resolve()
        parent.mkdir(parents=True, exist_ok=True)
        name = datetime.now(timezone.utc).strftime('run-%Y%m%dT%H%M%S-') + uuid.uuid4().hex[:12]
        output = parent / name
        # Recent Python versions apply 0700 as a restrictive Windows ACL too.
        # On Windows inherit the selected parent directory's ACL instead.
        output.mkdir(mode=0o777 if os.name == 'nt' else 0o700)
        write_plan(output / 'test-plan.jmx', requests, threads, loops)
        with (output / 'runtime.properties').open('w', encoding='utf-8', newline='\n') as properties:
            properties.write(RUNTIME_PROPERTIES)
        write_guide(output / 'JMeter-User-Guide.xlsx')
        print('Output: ' + str(output), flush=True)
        print('Planned requests: %d (%d APIs x %d users x %d loops)' %
              (len(requests) * threads * loops, len(requests), threads, loops), flush=True)
        if args.generate_only:
            print('Generated JMX, runtime.properties and Excel guide.')
            return 0
        home, java, java_major = runtime
        print('Running JMeter 5.6.3; progress is written to console.log and jmeter.log.', flush=True)
        exit_code, summary = execute(home, java, java_major, output, requests, threads, loops, timeout, args.html_report)
        print('Result: %s; samples=%s; failed=%s; exit_code=%d' %
              (summary['status'], summary.get('total_samples', 'unknown'),
               summary.get('failed_samples', 'unknown'), exit_code))
        return exit_code
    except InputError as exc:
        print('ERROR: ' + str(exc), file=sys.stderr)
        return 2
    except OSError as exc:
        # Do not print file contents or an entire command line on error.
        print('ERROR: filesystem/process operation failed (%s). Check paths and permissions.' %
              type(exc).__name__, file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print('Interrupted.', file=sys.stderr)
        return 130


if __name__ == '__main__':
    sys.exit(main())
