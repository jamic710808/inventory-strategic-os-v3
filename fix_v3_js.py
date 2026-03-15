#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fix literal newlines inside JavaScript string literals in V3 HTML"""

with open('Inventory_Strategic_OS_V3.html', 'r', encoding='utf-8') as f:
    content = f.read()

start = content.find('<script>') + len('<script>')
end = content.rfind('</script>')
js = content[start:end]

BSLASH = chr(92)

def fix_js(text):
    result = []
    i = 0
    n = len(text)

    while i < n:
        ch = text[i]

        # // single-line comment
        if ch == '/' and i+1 < n and text[i+1] == '/':
            end_nl = text.find('\n', i)
            if end_nl == -1:
                end_nl = n
            result.append(text[i:end_nl])
            i = end_nl
            continue

        # /* */ multi-line comment
        if ch == '/' and i+1 < n and text[i+1] == '*':
            end_cm = text.find('*/', i+2)
            if end_cm == -1:
                end_cm = n - 2
            result.append(text[i:end_cm+2])
            i = end_cm + 2
            continue

        # Single-quoted string
        if ch == "'":
            j = i + 1
            s = [ch]
            while j < n:
                c2 = text[j]
                if c2 == BSLASH and j+1 < n:
                    s.append(c2)
                    j += 1
                    s.append(text[j])
                    j += 1
                elif c2 == '\n':
                    s.append(BSLASH + 'n')
                    j += 1
                elif c2 == '\r':
                    s.append(BSLASH + 'r')
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
                if c2 == BSLASH and j+1 < n:
                    s.append(c2)
                    j += 1
                    s.append(text[j])
                    j += 1
                elif c2 == '\n':
                    s.append(BSLASH + 'n')
                    j += 1
                elif c2 == '\r':
                    s.append(BSLASH + 'r')
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

        # Template literal - keep as is (newlines are valid)
        if ch == '`':
            j = i + 1
            s = [ch]
            while j < n:
                c2 = text[j]
                if c2 == BSLASH and j+1 < n:
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


fixed_js = fix_js(js)
orig_nl = js.count('\n')
fixed_nl = fixed_js.count('\n')
print(f'Original lines: {orig_nl}  Fixed lines: {fixed_nl}  Strings escaped: {orig_nl - fixed_nl}')

new_content = content[:start] + fixed_js + content[end:]
with open('Inventory_Strategic_OS_V3.html', 'w', encoding='utf-8') as f:
    f.write(new_content)

# Also write fixed JS for syntax check
with open('v3_debug/full_js.js', 'w', encoding='utf-8') as f:
    f.write(fixed_js)

print('Saved Inventory_Strategic_OS_V3.html and v3_debug/full_js.js')
