import sys, io, tokenize, token as token_mod, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# Build reverse map
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

def fix_mojibake_str(s):
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

def likely_has_mojibake(s):
    return bool(re.search(r'[\x80-\xff\u0100-\u024f\u2000-\u21ff]', s))

escape_pattern = re.compile(
    r'(\\(?:u[0-9a-fA-F]{4}|U[0-9a-fA-F]{8}|x[0-9a-fA-F]{2}'
    r'|N\{[^}]+\}|[nrtbf\\\'"0-9]))'
)

raw = open('app.py.mojibake-bak', 'rb').read()
if raw.startswith(b'\xef\xbb\xbf'):
    raw = raw[3:]
text = raw.decode('utf-8')

readline = io.StringIO(text).readline
for tok in tokenize.generate_tokens(readline):
    if tok.type == token_mod.STRING and tok.start[0] == 989:
        ts = tok.string
        i = 0
        while i < len(ts) and ts[i] in ('r', 'R', 'b', 'B', 'f', 'F', 'u', 'U'):
            i += 1
        prefix = ts[:i]
        rest = ts[i:]
        
        print("Prefix:", repr(prefix))
        
        # For f-strings, the inner may contain {expressions}
        # We need to handle these carefully  
        if rest.startswith('"""') or rest.startswith("'''"):
            quote = rest[:3]
            inner = rest[3:-3]
        else:
            quote = rest[0]
            inner = rest[1:-1]
        
        print("inner first 200:", repr(inner[:200]))
        print()
        
        # Test escape_pattern.split
        parts = escape_pattern.split(inner[:200])
        print("Parts count:", len(parts))
        for p in parts[:5]:
            print("  part:", repr(p[:60]))
        
        # Does fix work char by char?
        fixed = fix_mojibake_str(inner)
        print()
        print("Changed:", fixed != inner)
        if fixed != inner:
            print("Fixed first 100:", fixed[:100])
        else:
            # Find why it fails
            orig_bytes = []
            for c in inner:
                co = ord(c)
                if co in reverse_map:
                    orig_bytes.append(reverse_map[co])
                else:
                    for b in c.encode('utf-8'):
                        orig_bytes.append(b)
            try:
                result = bytes(orig_bytes).decode('utf-8')
                print("Direct fix works:", result[:80])
            except UnicodeDecodeError as e:
                print("Decode error:", e)
                # Find problematic byte sequence
                rb = bytes(orig_bytes)
                print("Problematic bytes around error:", rb[max(0,e.start-5):e.start+10].hex())
        break
