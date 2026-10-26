#!/usr/bin/env python3
"""统计书稿字数：汉字数、中文标点数、总字数（不含空白）。"""
import re
import sys
from pathlib import Path

book_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "book"

han_re = re.compile(r"[一-鿿㐀-䶿]")
cjk_punct_re = re.compile(r"[　-〿＀-￯‘-‟…—]")

total_han = total_punct = total_nonspace = 0
print(f"{'文件':<40}{'汉字':>8}{'中文标点':>10}{'非空白字符':>12}")
for f in sorted(book_dir.glob("*.md")):
    if f.name.startswith("全本"):
        continue
    text = f.read_text(encoding="utf-8")
    han = len(han_re.findall(text))
    punct = len(cjk_punct_re.findall(text))
    nonspace = len(re.sub(r"\s", "", text))
    total_han += han
    total_punct += punct
    total_nonspace += nonspace
    print(f"{f.name:<40}{han:>8}{punct:>10}{nonspace:>12}")
print("-" * 70)
print(f"{'合计':<40}{total_han:>8}{total_punct:>10}{total_nonspace:>12}")
print(f"汉字+中文标点（出版口径近似）：{total_han + total_punct}")
