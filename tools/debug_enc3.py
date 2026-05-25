import sys, io, tokenize, token as token_mod
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

raw = open('app.py.mojibake-bak', 'rb').read()
if raw.startswith(b'\xef\xbb\xbf'):
    raw = raw[3:]
text = raw.decode('utf-8')
lines = text.split('\n')

for i in range(984, 992):
    print("L%d: %s" % (i+1, lines[i].strip()[:120]))

print()
print("---")
# Find f-string in backup
readline = io.StringIO(text).readline
for tok in tokenize.generate_tokens(readline):
    if tok.type == token_mod.STRING and tok.start[0] >= 984 and tok.start[0] <= 1000:
        ts = tok.string
        i2 = 0
        while i2 < len(ts) and ts[i2] in ('r', 'R', 'b', 'B', 'f', 'F', 'u', 'U'):
            i2 += 1
        prefix = ts[:i2]
        print("Token L%d prefix=%r len=%d" % (tok.start[0], prefix, len(ts)))
