import re
lines = open('src/tool.py', encoding='utf-8').read().split('\n')

starts = {}
for i, l in enumerate(lines, 1):
    m = re.match(r'^    (?:async )?def ([a-zA-Z_]\w*)', l)
    if m:
        starts[m.group(1)] = i
names = list(starts)

def get_end(name):
    idx = names.index(name)
    return (starts[names[idx+1]] - 1) if idx+1 < len(names) else len(lines)

# scan every method for triple-quoted strings spanning lines
print("=== multi-line triple-quoted strings per method ===")
total = 0
for n in names:
    s, e = starts[n], get_end(n)
    body = lines[s-1:e]
    in_str = None
    for i, l in enumerate(body):
        # crude scanner: find unescaped ''' or """
        for q in ('"""', "'''"):
            pos = 0
            while True:
                j = l.find(q, pos)
                if j == -1:
                    break
                # count preceding backslashes
                bs = 0
                k = j - 1
                while k >= 0 and l[k] == '\\':
                    bs += 1; k -= 1
                if bs % 2 == 0:
                    if in_str == q:
                        in_str = None
                    elif in_str is None:
                        in_str = q
                    pos = j + 3
                else:
                    pos = j + 3
        if in_str is not None and in_str not in l:
            pass
    # re-scan to report spans
    in_str = None
    spans = []
    for i, l in enumerate(body):
        for q in ('"""', "'''"):
            pos = 0
            while True:
                j = l.find(q, pos)
                if j == -1: break
                bs = 0; k = j-1
                while k >= 0 and l[k] == '\\':
                    bs += 1; k -= 1
                if bs % 2 == 0:
                    if in_str == q: in_str = None
                    elif in_str is None: in_str = q
                pos = j + 3
        if in_str is not None:
            spans.append(s + i)
    if spans:
        total += len(spans)
        print(f"{n} (lines {s}-{e}): string-continuation lines: {spans}")

print("total continuation lines:", total)
