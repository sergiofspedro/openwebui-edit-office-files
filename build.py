#!/usr/bin/env python3
"""Build tool.py from src/ modules."""
import os

SRC = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(SRC), 'tool.py')

files = sorted(f for f in os.listdir(os.path.join(SRC, 'src')) if f.endswith('.py') and f != '__init__.py')

with open(OUT, 'w', encoding='utf-8') as out:
    for fname in files:
        fpath = os.path.join(SRC, 'src', fname)
        with open(fpath, 'r', encoding='utf-8') as f:
            content = f.read()
        out.write(content)
        if not content.endswith('\n'):
            out.write('\n')

print(f"Built {OUT} from {len(files)} source files ({os.path.getsize(OUT)} bytes)")
