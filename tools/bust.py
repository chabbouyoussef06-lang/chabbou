#!/usr/bin/env python3
"""Cache-busting: appends ?v=<content hash> to every local CSS/JS reference in the HTML pages.
Run it after changing any file in css/ or js/ (python3 tools/bust.py) so visitors never keep an old copy."""
import hashlib, pathlib, re

ROOT = pathlib.Path(__file__).resolve().parent.parent
REF = re.compile(r'((?:href|src)="(?:\.\./)?)((?:css|js)/[\w.-]+\.(?:css|js))(?:\?v=\w+)?(")')

def digest(rel):
    return hashlib.md5((ROOT / rel).read_bytes()).hexdigest()[:8]

changed = 0
for page in list(ROOT.glob('*.html')) + list(ROOT.glob('*/index.html')):
    text = page.read_text(encoding='utf-8')
    new = REF.sub(lambda m: f'{m.group(1)}{m.group(2)}?v={digest(m.group(2))}{m.group(3)}', text)
    if new != text:
        page.write_text(new, encoding='utf-8'); changed += 1
print(f'{changed} page(s) updated')
