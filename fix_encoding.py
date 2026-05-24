"""
Fix mojibake (double-encoded UTF-8) in app.py.

Root cause: The file's Vietnamese UTF-8 text was read using iso-8859-2/latin-1/cp1252,
each byte becoming a unicode codepoint, then re-saved as UTF-8.

Fix: For each char in a mojibake string, map it back to its original byte using a
combined reverse lookup table (latin-1 + iso-8859-2 + cp1252), then decode as UTF-8.
"""
import sys
import io
import re
import tokenize
import token as token_mod
import shutil
from pathlib import Path


def build_reverse_map():
    """
    Build a reverse map: unicode codepoint -> original byte value.

    The file was corrupted by decoding UTF-8 bytes through a mix of encodings:
    - Most bytes (0x00-0xFF) were decoded as latin-1 (identity mapping)
    - Byte 0xC3 was decoded as iso-8859-2 giving Ă (U+0102) instead of Ã (U+00C3)
    - Byte 0xE3 was decoded as iso-8859-2 giving ă (U+0103) instead of ã (U+00E3)
    - cp1252 special chars for bytes 0x80-0x9F (e.g. TM=0x99->U+2122, OE=0x8C->U+0152)

    IMPORTANT: We do NOT include iso-8859-2 overrides for bytes 0xC0-0xFF range
    (except 0xC3, 0xE3) because those bytes are UTF-8 sequence leaders and
    their iso-8859-2 mappings (like 0xD0->Đ U+0110) would incorrectly convert
    valid Vietnamese Unicode chars that appear directly in strings.
    """
    reverse_map = {}
    # Latin-1: identity mapping 0-255
    for b in range(0x00, 0x100):
        c = bytes([b]).decode('latin-1')
        reverse_map[ord(c)] = b

    # Specific iso-8859-2 overrides for bytes that are commonly confused:
    # 0xC3 -> Ă (U+0102): verified from file analysis
    # 0xE3 -> ă (U+0103): lowercase counterpart
    for b in [0xC3, 0xE3]:
        c = bytes([b]).decode('iso-8859-2')
        reverse_map[ord(c)] = b

    # CP1252 special chars for the 0x80-0x9F control char range:
    # These bytes (like 0x8C=Œ/U+0152, 0x99=™/U+2122) appear in the mojibake text
    # as their cp1252 unicode equivalents since they're "undefined" in strict latin-1.
    for b in range(0x80, 0xA0):
        try:
            c = bytes([b]).decode('cp1252')
            co = ord(c)
            if co >= 0x100:  # Special cp1252 chars not in latin-1 range
                reverse_map.setdefault(co, b)
        except Exception:
            pass
    return reverse_map


REVERSE_MAP = build_reverse_map()


def fix_mojibake_str(s):
    """
    Fix a mojibake string by mapping each codepoint back to its original byte,
    then decoding the resulting bytes as UTF-8.
    Returns original string if fix fails or produces same result.
    """
    orig_bytes = []
    for c in s:
        co = ord(c)
        if co in REVERSE_MAP:
            orig_bytes.append(REVERSE_MAP[co])
        else:
            # Unknown codepoint - keep as utf-8 bytes (already correct encoding)
            for b in c.encode('utf-8'):
                orig_bytes.append(b)
    try:
        return bytes(orig_bytes).decode('utf-8')
    except UnicodeDecodeError:
        return s


def likely_has_mojibake(s):
    """
    Detect likely mojibake: non-ASCII chars that may be re-encoded original bytes.
    Covers Latin-1 (0x80-0xFF), Latin Extended-A/B (U+0100-U+024F),
    and cp1252 special chars (U+2000-U+21FF).
    """
    return bool(re.search(r'[\x80-\xff\u0100-\u024f\u2000-\u21ff]', s))


def fix_file(src_path, dst_path):
    src_path = Path(src_path)
    dst_path = Path(dst_path)

    raw = src_path.read_bytes()

    # Detect and strip BOM
    bom = b''
    if raw.startswith(b'\xef\xbb\xbf'):
        bom = b'\xef\xbb\xbf'
        raw = raw[3:]

    source_text = raw.decode('utf-8')

    # Build offset map: (row, col) -> offset in flattened text
    lines_list = source_text.split('\n')
    line_offsets = []
    offset = 0
    for line in lines_list:
        line_offsets.append(offset)
        offset += len(line) + 1  # +1 for '\n'

    def line_col_to_offset(row, col):
        return line_offsets[row - 1] + col

    # Tokenize
    tokens = []
    try:
        readline = io.StringIO(source_text).readline
        for tok in tokenize.generate_tokens(readline):
            tokens.append(tok)
    except tokenize.TokenError as e:
        print("Tokenize error: %s" % e, file=sys.stderr)
        return

    replacements = []
    errors = 0

    # Escape sequence pattern in Python strings (don't fix escape sequences)
    escape_pattern = re.compile(
        r'(\\(?:u[0-9a-fA-F]{4}|U[0-9a-fA-F]{8}|x[0-9a-fA-F]{2}'
        r'|N\{[^}]+\}|[nrtbf\\\'"0-9]))'
    )

    for tok in tokens:
        try:
            if tok.type == token_mod.STRING:
                tok_str = tok.string

                if not likely_has_mojibake(tok_str):
                    continue

                # Find string prefix
                i = 0
                while i < len(tok_str) and tok_str[i] in ('r', 'R', 'b', 'B', 'f', 'F', 'u', 'U'):
                    i += 1
                prefix = tok_str[:i]
                rest = tok_str[i:]

                # Skip raw and byte strings
                if 'r' in prefix.lower() or 'b' in prefix.lower():
                    continue

                # Determine quote style and extract inner content
                if rest.startswith('"""') or rest.startswith("'''"):
                    quote = rest[:3]
                    inner = rest[3:-3]
                elif rest and rest[0] in ('"', "'"):
                    quote = rest[0]
                    inner = rest[1:-1]
                else:
                    continue

                if not likely_has_mojibake(inner):
                    continue

                # Split on Python escape sequences to preserve them
                parts = escape_pattern.split(inner)
                fixed_parts = []
                changed = False
                for part in parts:
                    if escape_pattern.match(part):
                        fixed_parts.append(part)  # Keep escape sequences unchanged
                    else:
                        fixed = fix_mojibake_str(part)
                        if fixed != part:
                            changed = True
                        fixed_parts.append(fixed)

                if not changed:
                    continue

                fixed_inner = ''.join(fixed_parts)
                new_tok = prefix + quote + fixed_inner + quote

                start_offset = line_col_to_offset(tok.start[0], tok.start[1])
                end_offset = line_col_to_offset(tok.end[0], tok.end[1])
                replacements.append((start_offset, end_offset, new_tok))

            elif tok.type == token_mod.COMMENT:
                comment = tok.string
                if not likely_has_mojibake(comment):
                    continue
                fixed_comment = fix_mojibake_str(comment)
                if fixed_comment == comment:
                    continue
                start_offset = line_col_to_offset(tok.start[0], tok.start[1])
                end_offset = line_col_to_offset(tok.end[0], tok.end[1])
                replacements.append((start_offset, end_offset, fixed_comment))

        except Exception as e:
            errors += 1
            if errors <= 5:
                print("  Warning at %s: %s" % (tok.start, e), file=sys.stderr)

    # Sort descending to apply from end (preserves earlier offsets)
    replacements.sort(key=lambda x: x[0], reverse=True)

    # Apply replacements
    text_list = list(source_text)
    for start, end, new_text in replacements:
        text_list[start:end] = list(new_text)

    result = ''.join(text_list)

    print("Applied %d replacements." % len(replacements))
    if errors:
        print("Warnings: %d tokens had errors (skipped)." % errors)
    dst_path.write_bytes(bom + result.encode('utf-8'))
    print("Written to %s" % dst_path)


if __name__ == '__main__':
    src = Path('app.py')
    bak = Path('app.py.mojibake-bak')
    dst = Path('app.py')

    if not bak.exists():
        shutil.copy2(src, bak)
        print("Backup created: %s" % bak)
    else:
        # Restore from backup before re-running to ensure idempotent result
        shutil.copy2(bak, src)
        print("Restored from backup: %s" % bak)

    tmp = Path('app.py.fixed-tmp')
    fix_file(src, tmp)

    # Verify syntax
    try:
        import ast
        code = tmp.read_bytes()
        if code.startswith(b'\xef\xbb\xbf'):
            code = code[3:]
        ast.parse(code.decode('utf-8'))
        print("Syntax check: OK")
        tmp.replace(dst)
        print("Done! app.py has been fixed.")
    except SyntaxError as e:
        print("Syntax error in fixed file: %s" % e)
        print("Original file preserved. Fixed file at app.py.fixed-tmp for inspection.")
        sys.exit(1)
