#!/usr/bin/env python3
"""Command entry point kept compatible with Python 3.9 (RHEL 9)."""
import sys

if sys.version_info < (3, 9):
    print('ERROR: Python 3.9 or later is required.', file=sys.stderr)
    sys.exit(2)

from restjmeter.cli import main

if __name__ == '__main__':
    sys.exit(main())
