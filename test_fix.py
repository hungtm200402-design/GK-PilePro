import sys, io, tokenize, token as token_mod
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

reverse_map = {}
for b in range(0x00, 0x100):
    c = bytes([b]).decode('latin-1')
    reverse_map[ord(c)] = b
for b in range(0x80, 0x100):
    try:
        c = bytes([b]).decode('iso-8859-2')
        co = ord(c)
        if co >= 0x100:
            reverse_map[co] = b
    except:
        pass
for b in range(0x80, 0xA0):
    try:
        c = bytes([b]).decode('cp1252')
        co = ord(c)
        if co >= 0x100:
            reverse_map.setdefault(co, b)
    except:
        pass

def fix_mojibake_new(s):
    orig_bytes = []
    for c in s:
        co = ord(c)
        if co in reverse_map:
            orig_bytes.append(reverse_map[co])
        else:
            for b in c.encode('utf-8'):
                orig_bytes.append(b)
    try:
        return bytes(orig_bytes).decode('utf-8')
    except UnicodeDecodeError:
        return s

raw = open('app.py.mojibake-bak', 'rb').read()
if raw.startswith(b'\xef\xbb\xbf'):
    raw = raw[3:]
text = raw.decode('utf-8')
readline = io.StringIO(text).readline

tokens_to_check = {437, 495, 629, 989, 1223}
for tok in tokenize.generate_tokens(readline):
    if tok.type == token_mod.STRING and tok.start[0] in tokens_to_check:
        inner = tok.string[1:-1]
        result = fix_mojibake_new(inner)
        changed = result != inner
        print("L%d: changed=%s result=%r" % (tok.start[0], changed, result))
