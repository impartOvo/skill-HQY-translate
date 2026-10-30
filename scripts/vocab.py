# -*- coding: utf-8 -*-
"""vocab.py — 从「多站点已填写/已上传」的上传表里反推跨语言取值对照表。

原理：同一店铺不同站点是同一个 SKU、同一套属性名，所以按 SKU + 属性名对齐，
     就能得到「德语取值 → 各语言官方取值」的实测对照（这些值是真被亚马逊接受过的）。

用法：
    python vocab.py --pair "德国/DE.xlsx|nl_NL|荷兰/NL.xlsm" \
                    --pair "德国/DE.xlsx|fr_FR|法国/FR.xlsm" \
                    --pair "德国/DE.xlsx|en_GB|英国/UK.xlsm" \
                    --out reference/vocab_de_multi.tsv [--drop-identity]

输出 TSV 列：attribute  de_value  <locale1>  <locale2> ...
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import flatfile as ff  # noqa: E402

ff.utf8_stdio()


def collect(pairs):
    data = defaultdict(lambda: defaultdict(lambda: defaultdict(set)))   # attr -> locale -> de -> {tgt}
    stats = []
    for src_path, locale, tgt_path in pairs:
        if not (os.path.exists(src_path) and os.path.exists(tgt_path)):
            print("[skip] 缺文件: %s / %s" % (src_path, tgt_path))
            continue
        ws_s = ff.open_source(src_path)
        cs = ff.attr_map(ws_s)
        ws_t = ff.open_source(tgt_path)
        ct = ff.attr_map(ws_t)
        sku_s, sku_t = ff.sku_col(cs), ff.sku_col(ct)
        rs = {str(ws_s.cell(r, sku_s).value).strip(): r for r in ff.data_rows(ws_s, cs)}
        rt = {str(ws_t.cell(r, sku_t).value).strip(): r for r in ff.data_rows(ws_t, ct)}
        common = sorted(set(rs) & set(rt))
        shared = set(cs) & set(ct)
        for sku in common:
            for a in shared:
                dv = ws_s.cell(rs[sku], cs[a]).value
                tv = ws_t.cell(rt[sku], ct[a]).value
                if ff.is_blank(dv) or ff.is_blank(tv):
                    continue
                data[a][locale][str(dv).strip()].add(str(tv).strip())
        stats.append((os.path.basename(tgt_path), locale, len(rt), len(common), len(shared)))
    return data, stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pair", action="append", required=True,
                    help="源表|语言标签|目标表，可重复；源表是同一个德语/英语上传表")
    ap.add_argument("--out", default="reference/vocab.tsv")
    ap.add_argument("--drop-identity", action="store_true", help="丢掉各语言取值与德文完全相同的行")
    args = ap.parse_args()

    pairs = []
    for p in args.pair:
        parts = p.split("|")
        if len(parts) != 3:
            raise SystemExit("--pair 格式错误（用 | 分隔三段）: " + p)
        pairs.append(tuple(x.strip() for x in parts))

    data, stats = collect(pairs)
    for name, loc, nrows, ncommon, nshared in stats:
        print("[pair] %-44s %-6s 行=%-3d 对齐=%-3d 共有属性=%d" % (name, loc, nrows, ncommon, nshared))

    locales = sorted({loc for _, loc, *_ in stats})
    rows = []
    for a in sorted(data):
        de_values = set()
        for loc in data[a]:
            de_values |= set(data[a][loc])
        for dv in sorted(de_values):
            cells = []
            for loc in locales:
                vs = data[a][loc].get(dv, set())
                cells.append(" | ".join(sorted(vs)))
            if args.drop_identity and all(c in ("", dv) for c in cells):
                continue
            rows.append((a, dv, cells))

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write("attribute\tde_value\t" + "\t".join(locales) + "\n")
        for a, dv, cells in rows:
            f.write("%s\t%s\t%s\n" % (a, dv, "\t".join(cells)))
    print("\n[out] %s  属性 %d 个，行 %d 条，语言 %s" % (args.out, len(data), len(rows), locales))


if __name__ == "__main__":
    main()
