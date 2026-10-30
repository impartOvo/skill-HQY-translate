# -*- coding: utf-8 -*-
"""enums.py — 导出某站点模板里「官方允许值」（definedNames → Dropdown Lists），不靠猜。

用法：
    # 单个属性
    python enums.py TEMPLATE.xlsm --pt SWIMWEAR --attr "color[marketplace_id=A1805IZSGTT6HS][language_tag=nl_NL]#1.value"
    # 模板第 5 行里所有属性都导一遍
    python enums.py TEMPLATE.xlsm --pt SWIMWEAR --all --out reference/allowed_nl.tsv

输出 TSV：attribute <TAB> allowed_values（多值用 | 分隔）
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import flatfile as ff  # noqa: E402

ff.utf8_stdio()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("--pt", default="SWIMWEAR", help="product type，如 SWIMWEAR / SHIRT / APPAREL")
    ap.add_argument("--attr", action="append", default=[], help="可重复；给属性原名或 alias")
    ap.add_argument("--all", action="store_true", help="导出模板第 5 行的全部属性")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    av = ff.AllowedValues(args.template)
    ws = ff.open_source(args.template)
    cols = ff.attr_map(ws)
    raw = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(ff.ATTR_ROW, c).value
        if v not in (None, ""):
            raw.setdefault(ff.alias(v), str(v))

    targets = []
    if args.all:
        targets = list(raw)
    for a in args.attr:
        if a in raw:
            targets.append(a)
            continue
        # 允许只给前缀（如 color / special_feature / shapewear_size）
        pre = [k for k in raw if k == a or k.startswith(a)]
        if pre:
            targets.extend(pre)
            print("[match] %-40s -> %s" % (a, pre))
        else:
            targets.append(ff.alias(a))

    lines = []
    for a in targets:
        orig = raw.get(a, a)
        vals = av.get(args.pt, orig)
        if vals is None:
            lines.append("%s\t<无下拉列表/自由文本>" % a)
            print("[free] %-58s 无受控取值" % a)
            continue
        lines.append("%s\t%s" % (a, " | ".join(vals)))
        print("[ok]   %-58s %d 个取值" % (a, len(vals)))

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write("attribute\tallowed_values\n")
            f.write("\n".join(lines) + "\n")
        print("\n写出 %s（%d 行）" % (args.out, len(lines)))
    av.close()


if __name__ == "__main__":
    main()
