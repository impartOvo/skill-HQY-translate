# -*- coding: utf-8 -*-
"""pick_template.py — 按类目自动挑目标站空白模板。

表里基本都带类目词（bikini / tankini / monokini / badpak / one piece…）：
  1) 从源表读 item_type_name / swimwear_form_type / item_name，判定这款属于哪一类；
  2) 扫目标国家文件夹里的空白模板，看各自 recommended_browse_nodes 覆盖哪些类目
     （节点数 ≤5 说明是按单个 ASIN 下载的"预筛模板"，只能干那一个类目）；
  3) 输出推荐模板 + 备选 + 不匹配原因。

用法：
    python pick_template.py --source "德国/惠全优 德国 2399 上传表 .xlsx" \
                            --dir "荷兰" [--pt SWIMWEAR]
    python pick_template.py --scan "荷兰"          # 只列该目录里所有模板的类目
"""
from __future__ import annotations

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import flatfile as ff  # noqa: E402

# 类目关键词（多语言）：出现在 browse node / item_type_name / 商品标题里
CAT_PATTERNS = {
    "monokini": r"monokini",
    "tankini": r"tankini",
    "bikini": r"bikini",
    "onepiece": r"badpak|one[- ]?piece|une pi[eè]ce|una pieza|eendelig|zwempak uit|one pieces",
}
CAT_LABEL = {"monokini": "Monokini 连体比基尼", "tankini": "Tankini 套装",
             "bikini": "比基尼套装", "onepiece": "连体泳衣", "unknown": "无法判定"}


def guess_category_from_text(text: str):
    t = (text or "").lower()
    # monokini/tankini 优先于 bikini（"monokini" 里也含 "bikini"）
    for cat in ("monokini", "tankini", "bikini", "onepiece"):
        if re.search(CAT_PATTERNS[cat], t):
            return cat
    return None


def source_category(paths, pt="SWIMWEAR"):
    """从源表判定类目：先看 swimwear_form_type / item_type_name，再看标题"""
    ws = ff.open_source(paths)
    cols = ff.attr_map(ws)
    rows = ff.data_rows(ws, cols)

    def vals(alias):
        c = cols.get(alias)
        if not c:
            return []
        return sorted({str(ws.cell(r, c).value).strip() for r in rows if not ff.is_blank(ws.cell(r, c).value)})

    for alias in ("swimwear_form_type#1.value", "item_type_name#1.value"):
        for v in vals(alias):
            cat = guess_category_from_text(v)
            if cat:
                return cat, "%s = %s" % (alias.split("#")[0], v)
    for v in vals("item_name#1.value"):
        cat = guess_category_from_text(v)
        if cat:
            return cat, "标题含类目词：%s" % v[:48]
    return "unknown", "源表没有可识别的类目词"


def template_categories(path, pt="SWIMWEAR"):
    """读模板 browse node 允许值 → 覆盖的类目集合 + 是否被预筛"""
    ws = ff.open_source(path)
    cols = ff.attr_map(ws)
    c = cols.get("recommended_browse_nodes#1.value")
    if not c:
        return set(), 0
    av = ff.AllowedValues(path)
    vals = av.get(pt, str(ws.cell(5, c).value)) or []
    av.close()
    cats = set()
    for v in vals:
        cat = guess_category_from_text(v)
        if cat:
            cats.add(cat)
    return cats, len(vals)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=None, help="德国源表，用于判定类目")
    ap.add_argument("--dir", default=None, help="目标国家文件夹，或用国家名（自动拼 ZINIAO_EU）")
    ap.add_argument("--scan", default=None, help="只扫描某目录里所有模板的类目")
    ap.add_argument("--pt", default="SWIMWEAR")
    args = ap.parse_args()

    target_dir = args.scan or args.dir
    if target_dir and not os.path.isdir(target_dir):
        root = os.environ.get("ZINIAO_EU", "")
        cand = os.path.join(root, target_dir) if root else target_dir
        if os.path.isdir(cand):
            target_dir = cand
    if not target_dir or not os.path.isdir(target_dir):
        raise SystemExit("目录不存在：%r（可用环境变量 ZINIAO_EU 指到站点根目录）" % target_dir)

    blank = []
    for f in sorted(os.listdir(target_dir)):
        if not f.lower().endswith((".xlsm", ".xlsx")) or f.startswith("~$") or "dsh生成" in f:
            continue
        p = os.path.join(target_dir, f)
        try:
            ws = ff.open_source(p)
            cols = ff.attr_map(ws)
            if ff.data_rows(ws, cols):        # 有数据行 = 已填过的表，不当模板
                continue
            cats, n = template_categories(p, args.pt)
        except BaseException:
            continue
        blank.append({"path": p, "name": f, "cats": cats, "nodes": n})

    print("目标目录：%s" % target_dir)
    print("%-42s %-7s %-6s %s" % ("模板", "节点数", "性质", "覆盖类目"))
    for b in blank:
        kind = "预筛(单类目)" if b["nodes"] <= 5 else "全量"
        print("%-42s %-7d %-6s %s" % (b["name"][:42], b["nodes"], kind,
                                      " / ".join(sorted(CAT_LABEL.get(c, c) for c in b["cats"])) or "-"))

    if not args.source:
        print("\n（没给 --source，只看模板清单）")
        return

    cat, why = source_category(args.source, args.pt)
    print("\n源表类目判定：%s  —— %s" % (CAT_LABEL.get(cat, cat), why))

    fit = [b for b in blank if cat in b["cats"]]
    # 排序：① 预筛模板且只覆盖本类目（=后台按该 ASIN 生成的，最贴）② 其余预筛 ③ 全量模板按节点数
    fit.sort(key=lambda b: (0 if (b["nodes"] <= 5 and b["cats"] == {cat}) else 1,
                            0 if b["nodes"] <= 5 else 1,
                            -b["nodes"]))
    if not fit:
        print("\n[!] 没有模板覆盖「%s」类目，需要到后台按该 ASIN 重新下载模板。" % CAT_LABEL.get(cat, cat))
        return
    best = fit[0]
    print("\n推荐模板：%s（节点 %d，%s）" % (best["name"], best["nodes"],
                                       "预筛" if best["nodes"] <= 5 else "全量"))
    if len(fit) > 1:
        print("备选：")
        for b in fit[1:]:
            print("   - %s（节点 %d）" % (b["name"], b["nodes"]))
    print("\n下一步：把 config 的 target.path 指向推荐模板，再跑 fill.py")


if __name__ == "__main__":
    ff.utf8_stdio()
    main()
