#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Targeted fix: replace literal CR/LF inside .join() string with escape sequence"""

with open('Inventory_Strategic_OS_V3.html', 'rb') as f:
    data = f.read()

# Pattern: .join(' followed by CR (0x0d) then newline
# The broken string is .join('\r\n') where \r is literal 0x0d

# Find all occurrences of .join(' + CR byte
import re

broken1 = b".join('\\r\\n"  # .join(' + CR + LF (CRLF)
broken2 = b".join('\\r'"    # .join(' + CR + ' (just CR)

print(f"Looking for .join(CR+LF): {data.count(broken1)} occurrences")
print(f"Looking for .join(CR+'): {data.count(broken2)} occurrences")

# Also search for any 0x0d inside single-quoted strings near .join
idx = 0
found = []
while True:
    pos = data.find(b".join('", idx)
    if pos == -1:
        break
    # Check next 10 bytes for CR
    snippet = data[pos:pos+15]
    if b'\r' in snippet:
        found.append((pos, repr(snippet)))
    idx = pos + 1

print(f"\nFound {len(found)} .join() calls with CR inside:")
for pos, snip in found:
    print(f"  offset {pos}: {snip}")

# Fix: replace .join('\r\n') with .join('\n') (keep newline escape)
# and .join('\r') with .join('\r\n') (proper CRLF escape)

# The intended code is: .join('\r\n') meaning CRLF line ending for CSV
# But the literal bytes are .join(' + 0x0d 0x0a + ')
# We want: .join('\\r\\n')  (the escape sequences as text)

fixed = data
# Replace literal .join('\r\n') -> .join('\\r\\n')
old = b".join('\r\n"
new = b".join('\\r\\n"
count = fixed.count(old)
print(f"\nReplacing .join(literal CRLF): {count} occurrences")
fixed = fixed.replace(old, new)

# Also fix the closing: after the newline escape we need the closing quote
# Pattern was: .join('\r\n') -> broken as .join('\r  \n')
# After above fix it becomes: .join('\r\n  \n') - need to check
# Actually let's look at full context around each fix

with open('Inventory_Strategic_OS_V3.html', 'wb') as f:
    f.write(fixed)

print("Written fixed file.")

# Also write debug JS
import re as re2
content = fixed.decode('utf-8', errors='replace')
start = content.find('<script>') + len('<script>')
end = content.rfind('</script>')
js = content[start:end]
with open('v3_debug/full_js.js', 'w', encoding='utf-8') as f:
    f.write(js)
print("Written v3_debug/full_js.js")
