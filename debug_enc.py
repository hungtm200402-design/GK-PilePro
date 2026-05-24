import sys, io, tokenize, token as token_mod, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def fix_mojibake_str(s):
    try:
        return s.encode('latin-1').decode('utf-8')
    except:
        return s

src = open('app.py.mojibake-bak', 'rb').read()
if src.startswith(b'\xef\xbb\xbf'):
    src = src[3:]
text = src.decode('utf-8')

count = 0
changed_count = 0
skipped_r = 0
skipped_b = 0
skipped_nochange = 0

readline = io.StringIO(text).readline
for tok in tokenize.generate_tokens(readline):
    if tok.type not in (token_mod.STRING, token_mod.COMMENT):
        continue
    tok_str = tok.string
    if not re.search(r'[\x80-\xff]', tok_str):
        continue

    count += 1

    if tok.type == token_mod.COMMENT:
        fixed = fix_mojibake_str(tok_str)
        if fixed != tok_str:
            changed_count += 1
        else:
            skipped_nochange += 1
        continue

    # STRING token
    i = 0
    while i < len(tok_str) and tok_str[i] in ('r', 'R', 'b', 'B', 'f', 'F', 'u', 'U'):
        i += 1
    prefix = tok_str[:i]
    rest = tok_str[i:]

    if 'r' in prefix.lower():
        skipped_r += 1
        continue
    if 'b' in prefix.lower():
        skipped_b += 1
        continue

    if rest.startswith('"""') or rest.startswith("'''"):
        quote = rest[:3]
        inner = rest[3:-3]
    elif rest and rest[0] in ('"', "'"):
        quote = rest[0]
        inner = rest[1:-1]
    else:
        skipped_nochange += 1
        continue

    # Try fix
    fixed_inner = fix_mojibake_str(inner)
    if fixed_inner != inner:
        changed_count += 1
    else:
        if skipped_nochange < 5:
            print(f'NOT CHANGED at {tok.start}: inner={repr(inner[:60])}')
        skipped_nochange += 1

print(f'Total with mojibake: {count}')
print(f'Would change: {changed_count}')
print(f'Skipped r-strings: {skipped_r}')
print(f'Skipped b-strings: {skipped_b}')
print(f'Skipped (no change possible): {skipped_nochange}')
