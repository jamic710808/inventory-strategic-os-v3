parts = [
    'C:/Users/jamic/庫存分析/v3_parts/part_js_p1.txt',
    'C:/Users/jamic/庫存分析/v3_parts/part_js_p2.txt',
    'C:/Users/jamic/庫存分析/v3_parts/part_js_p3.txt',
    'C:/Users/jamic/庫存分析/v3_parts/part_js_p4.txt',
]

out_path = 'C:/Users/jamic/庫存分析/v3_parts/part_js.txt'

total = 0
with open(out_path, 'w', encoding='utf-8') as out:
    for i, path in enumerate(parts):
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        out.write(content)
        if i < len(parts) - 1:
            out.write('\n\n')
        total += len(content)
        print(f"  Part {i+1}: {len(content)} chars from {path}")

print(f"\nTotal assembled: {total} chars -> {out_path}")

# Verify
with open(out_path, 'r', encoding='utf-8') as f:
    final = f.read()
print(f"Final file size: {len(final)} chars, {len(final.splitlines())} lines")
