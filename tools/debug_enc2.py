import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

raw = open('app.py.mojibake-bak', 'rb').read()
if raw.startswith(b'\xef\xbb\xbf'):
    raw = raw[3:]

count = 0
for i in range(len(raw)-6):
    if raw[i:i+2] == b'Th' and raw[i+2] in (0xc4, 0xc3, 0xc6):
        ctx = raw[i:i+15]
        decoded = ctx.decode('utf-8', errors='replace')
        print(f'Found at {i}: {ctx.hex()} = {decoded!r}')
        count += 1
        if count > 5:
            break

# Also: what does the actual 'ThĂ nh' line look like in bytes?
# Search for line 873 content
lines = raw.split(b'\n')
for j, line in enumerate(lines, 1):
    if b'Th' in line and b'\xc4\x82' in line:
        print(f'Line {j}: {line[:80].hex()}')
        print(f'         {line[:80].decode("utf-8", errors="replace")!r}')
        break
    if b'Th' in line and b'\xc3\x83' in line:
        print(f'Line {j} (Ã): {line[:80].hex()}')
        break
