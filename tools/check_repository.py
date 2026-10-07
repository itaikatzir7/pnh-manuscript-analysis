"""Check source integrity and Python syntax. Does not import analysis programs."""
import ast
import hashlib
import json
from pathlib import Path
import sys
import tokenize

ROOT = Path(__file__).resolve().parents[1]

def main():
    errors = []
    inventory = json.loads((ROOT / 'source_inventory.json').read_text(encoding='utf-8'))
    for item in inventory['sources']:
        path = (ROOT / item['path']).resolve()
        if not path.is_relative_to(ROOT) or not path.is_file():
            errors.append('Missing or unsafe source: ' + item['path'])
            continue
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if digest != item['packaged_sha256']:
            errors.append('Source hash differs from packaging record: ' + item['path'])
        else:
            print('HASH OK:', item['path'])
    for path in sorted(ROOT.rglob('*.py')):
        if any(part in {'.venv','venv','outputs'} for part in path.relative_to(ROOT).parts):
            continue
        try:
            with tokenize.open(path) as stream:
                ast.parse(stream.read(), filename=str(path))
        except (SyntaxError, UnicodeError, LookupError) as exc:
            errors.append(str(exc))
    for path in (ROOT / 'config').glob('*.json'):
        json.loads(path.read_text(encoding='utf-8'))
    for error in errors:
        print('FAIL:', error)
    print('Integrity/syntax check only; no numerical or GUI validation is implied.')
    return 1 if errors else 0

if __name__ == '__main__':
    sys.exit(main())
