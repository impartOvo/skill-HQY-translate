# -*- coding: utf-8 -*-
"""probe.py — 探查上传表结构：站点/语言/属性数/数据行/已填属性/取值唯一数。

用法：
    python probe.py "D:/path/a.xlsx" "D:/path/b.xlsm" [--values] [--diff]

--values  打印每个已填属性的所有不同取值（默认最多 8 个）
--diff    与第一个文件做属性名差异对比（谁多谁少）
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import flatfile as ff  # noqa: E402

ff.utf8_stdio()


def describe(path, show_values=False, max_values=8):
    ws = ff.open_source(path)
    cols = ff.attr_map(ws)
    mp, lang = ff.meta(ws)
    rows = ff.data_rows(ws, cols)
    print("=" * 96)
    print(os.path.basename(path))
    print("  工作表=%-12s max_row=%-4s max_col=%-4s marketplace=%-15s lang=%s"
          % (ws.title, ws.max_row, ws.max_column, mp or "-", lang or "-"))
    print("  属性数=%d  数据行=%d（第 %s 行）"
          % (len(cols), len(rows), ("%d–%d" % (rows[0], rows[-1])) if rows else "无"))
    filled = []
    for a, c in cols.items():
        vals = [ws.cell(r, c).value for r in rows]
        nz = [v for v in vals if not ff.is_blank(v)]
        if nz:
            uniq = sorted({str(v).strip() for v in nz})
            filled.append((a, len(nz), len(uniq), uniq, c))
    print("  已填属性=%d" % len(filled))
    for a, n, u, uniq, c in sorted(filled):
        if show_values:
            shown = uniq if u <= max_values else uniq[:max_values] + ["…(%d)" % u]
            print("    col%-4d %-56s 填 %2d 行 值: %r" % (c, a, n, shown))
        else:
            print("    col%-4d %-56s 填 %2d 行 唯一值 %d" % (c, a, n, u))
    return cols


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--values", action="store_true")
    ap.add_argument("--diff", action="store_true")
    args = ap.parse_args()

    maps = [describe(f, args.values) for f in args.files]
    if args.diff and len(maps) >= 2:
        base = maps[0]
        for path, m in zip(args.files[1:], maps[1:]):
            print("=" * 96)
            only_b = sorted(set(base) - set(m))
            only_t = sorted(set(m) - set(base))
            print("属性差异 %s（基准） vs %s" % (os.path.basename(args.files[0]), os.path.basename(path)))
            print("  基准多出 %d 项:" % len(only_b))
            for a in only_b:
                print("    -", a)
            print("  对比文件多出 %d 项:" % len(only_t))
            for a in only_t:
                print("    +", a)


if __name__ == "__main__":
    main()
