#!/usr/bin/env python3
"""
Generate a bcrypt hash for NASH_PASSWORD_HASH.

Usage:
  python make_password.py                 # prompts for password (hidden)
  python make_password.py 'mypassword'    # one-shot (avoid this in shared shells)
  python make_password.py --help

Output is a base64-encoded bcrypt hash with a "b64:" prefix.
This avoids docker-compose interpolating the $ signs inside bcrypt hashes.
app.py decodes the b64: prefix automatically.

Set the result in .env:
  NASH_PASSWORD_HASH=b64:<output from this script>
"""

import base64
import getpass
import sys

if any(a in ("-h", "--help") for a in sys.argv[1:]):
    print(__doc__)
    sys.exit(0)

try:
    import bcrypt
except ImportError:
    print("Missing dependency. Run: pip install bcrypt", file=sys.stderr)
    sys.exit(1)

if len(sys.argv) > 1:
    pw = sys.argv[1]
else:
    pw = getpass.getpass("Password: ")
    confirm = getpass.getpass("Confirm:  ")
    if pw != confirm:
        print("Passwords do not match.", file=sys.stderr)
        sys.exit(1)

if len(pw) < 6:
    print("Refusing to hash a password shorter than 6 chars.", file=sys.stderr)
    sys.exit(1)

raw = bcrypt.hashpw(pw.encode(), bcrypt.gensalt(rounds=12))
print("b64:" + base64.b64encode(raw).decode())
