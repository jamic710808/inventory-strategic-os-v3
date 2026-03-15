#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reassemble_v3.py
Reassembles Inventory_Strategic_OS_V3.html from clean source parts,
then fixes literal newlines inside JS string literals.
"""

import re

BSLASH = chr(92)

def fix_js_strings(text):
    """
    Replace literal CR/LF inside JS string literals with escape sequences.
    Handles: '//' comments, '/* */' comments, regex literals, template literals,
    single-quoted strings, double-quoted strings.
    """
    result = []
    i = 0
    n = len(text)

    def is_regex_start(pos):
        """Heuristic: / at pos is a regex start if preceded by operator or keyword."""
        # Walk back over whitespace
        j = pos - 1
        while j >= 0 and text[j] in ' \t\r\n':
            j -= 1
        if j < 0:
            return True
        c = text[j]
        # After these chars, / starts a regex
        return c in '=({[,;!&|?:~^%+-*<>'

    while i < n:
        ch = text[i]

        # // single-line comment — pass through unchanged
        if ch == '/' and i + 1 < n and text[i + 1] == '/':
            end = text.find('\n', i)
            if end == -1:
                end = n
            result.append(text[i:end])
            i = end
            continue

        # /* */ multi-line comment — pass through unchanged
        if ch == '/' and i + 1 < n and text[i + 1] == '*':
            end = text.find('*/', i + 2)
            if end == -1:
                end = n - 2
            result.append(text[i:end + 2])
            i = end + 2
            continue

        # Regex literal: /.../ preceded by operator
        if ch == '/' and i + 1 < n and text[i + 1] not in ('/', '*') and is_regex_start(i):
            j = i + 1
            s = [ch]
            while j < n:
                c2 = text[j]
                if c2 == BSLASH and j + 1 < n:
                    s.append(c2)
                    j += 1
                    s.append(text[j])
                    j += 1
                elif c2 == '[':
                    # character class — consume until ]
                    s.append(c2)
                    j += 1
                    while j < n and text[j] != ']':
                        if text[j] == BSLASH and j + 1 < n:
                            s.append(text[j])
                            j += 1
                        s.append(text[j])
                        j += 1
                    if j < n:
                        s.append(text[j])
                        j += 1
                elif c2 == '/':
                    s.append(c2)
                    j += 1
                    # consume flags
                    while j < n and text[j].isalpha():
                        s.append(text[j])
                        j += 1
                    break
                elif c2 in '\r\n':
                    # Unterminated regex — stop, let JS engine handle it
                    break
                else:
                    s.append(c2)
                    j += 1
            result.append(''.join(s))
            i = j
            continue

        # Single-quoted string
        if ch == "'":
            j = i + 1
            s = [ch]
            while j < n:
                c2 = text[j]
                if c2 == BSLASH and j + 1 < n:
                    s.append(c2)
                    j += 1
                    s.append(text[j])
                    j += 1
                elif c2 == '\r':
                    s.append(BSLASH + 'r')
                    j += 1
                elif c2 == '\n':
                    s.append(BSLASH + 'n')
                    j += 1
                elif c2 == "'":
                    s.append(c2)
                    j += 1
                    break
                else:
                    s.append(c2)
                    j += 1
            result.append(''.join(s))
            i = j
            continue

        # Double-quoted string
        if ch == '"':
            j = i + 1
            s = [ch]
            while j < n:
                c2 = text[j]
                if c2 == BSLASH and j + 1 < n:
                    s.append(c2)
                    j += 1
                    s.append(text[j])
                    j += 1
                elif c2 == '\r':
                    s.append(BSLASH + 'r')
                    j += 1
                elif c2 == '\n':
                    s.append(BSLASH + 'n')
                    j += 1
                elif c2 == '"':
                    s.append(c2)
                    j += 1
                    break
                else:
                    s.append(c2)
                    j += 1
            result.append(''.join(s))
            i = j
            continue

        # Template literal — newlines are valid, pass through unchanged
        if ch == '`':
            j = i + 1
            s = [ch]
            while j < n:
                c2 = text[j]
                if c2 == BSLASH and j + 1 < n:
                    s.append(c2)
                    j += 1
                    s.append(text[j])
                    j += 1
                elif c2 == '`':
                    s.append(c2)
                    j += 1
                    break
                else:
                    s.append(c2)
                    j += 1
            result.append(''.join(s))
            i = j
            continue

        result.append(ch)
        i += 1

    return ''.join(result)


# ── 1. Read original part files ──────────────────────────────────
with open('v3_parts/part_html.txt', 'r', encoding='utf-8') as f:
    html_part = f.read()

with open('v3_parts/part_js.txt', 'r', encoding='utf-8') as f:
    js_part = f.read()

print(f"HTML part: {len(html_part)} chars, {len(html_part.splitlines())} lines")
print(f"JS   part: {len(js_part)} chars,  {len(js_part.splitlines())} lines")

# ── 2. Fix string literal newlines in JS ─────────────────────────
orig_nl = js_part.count('\n')
fixed_js = fix_js_strings(js_part)
fixed_nl = fixed_js.count('\n')
print(f"JS string fix: {orig_nl} lines before → {fixed_nl} after  "
      f"({orig_nl - fixed_nl} newlines escaped inside strings)")

# ── 3. Assemble: html_part ends with "<script>\n", append JS then close ─
# Determine the structure of html_part
if html_part.rstrip().endswith('<script>'):
    # html_part already opens the script tag — just append JS + closing tags
    new_html = html_part + fixed_js + '\n</script>\n</body>\n</html>\n'
elif '<!-- JS_PLACEHOLDER -->' in html_part:
    new_html = html_part.replace('<!-- JS_PLACEHOLDER -->', f'<script>\n{fixed_js}\n</script>')
else:
    assert '</body>' in html_part, "Cannot find insertion point in html_part"
    new_html = html_part.replace('</body>', f'<script>\n{fixed_js}\n</script>\n</body>')

# ── 4. Write output ───────────────────────────────────────────────
OUT = 'Inventory_Strategic_OS_V3.html'
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(new_html)
print(f"Written {OUT}: {len(new_html)} chars")

# ── 5. Write debug JS ─────────────────────────────────────────────
import os
os.makedirs('v3_debug', exist_ok=True)
with open('v3_debug/full_js.js', 'w', encoding='utf-8') as f:
    f.write(fixed_js)
print(f"Written v3_debug/full_js.js: {len(fixed_js)} chars, {len(fixed_js.splitlines())} lines")
