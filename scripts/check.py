# -*- coding: utf-8 -*-
"""check.py — 上传前体检：结构完整性 / 必填 / 枚举合法性 / 文案规则 / 变体关系。

用法：
    python check.py --template 空白模板.xlsm --file 已填.xlsm [--source 源站点上传表.xlsx] [--pt SWIMWEAR]

--source 给了就再做一次「源站点 vs 目标站点」逐格对照（缺失/多余）。
退出码：0 全部通过；1 有 FAIL。
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import flatfile as ff  # noqa: E402
import openpyxl  # noqa: E402
from lxml import etree  # noqa: E402

ff.utf8_stdio()

# 文案上限（可按类目/站点调）：标题 75、副标题 125、五点 500（第 2 条 700）、描述 2000、后台关键词 250
LIMITS = {"item_name": 75, "title_differentiation": 125, "product_description": 2000,
          "bullet_point": 500, "bullet_point#2": 700, "generic_keyword": 250}
FORBID = set('!?$_{}^¬¦~#<>*|;"\\/')
PROMO = ["gratis", "free shipping", "bestseller", "topkwaliteit", "beste kwaliteit", "100%",
         "aanbieding", "korting", "goedkoop", "koop nu", "sale", "günstig", "beste preis", "gratuit"]
CONTACT = ["@", "www.", "http", "tel.", "whatsapp", "klantenservice", "kundenservice", "service client"]
REQ_WORDS = ("Vereist", "Erforderlich", "Obligatoire", "Obligatorio", "Required")
DEF_SHEET_RE = re.compile(r"definit|finici|definici|definition", re.I)


def fail(msgs, m):
    msgs["FAIL"].append(m)


def warn(msgs, m):
    msgs["WARN"].append(m)


def ok(msgs, m):
    msgs["OK"].append(m)


def find_def_sheet(wb):
    for n in wb.sheetnames:
        if DEF_SHEET_RE.search(n):
            return wb[n]
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--file", required=True)
    ap.add_argument("--source", default=None)
    ap.add_argument("--pt", default="SWIMWEAR")
    args = ap.parse_args()
    msgs = {"OK": [], "WARN": [], "FAIL": []}

    # ---------- A. 包结构 ----------
    zt, zf = zipfile.ZipFile(args.template), zipfile.ZipFile(args.file)
    same = zt.namelist() == zf.namelist()
    (ok if same else warn)(msgs, "zip 条目 %d vs %d，集合一致=%s（注释：Excel/WPS 另存会重排图片部件，不影响上传）"
                          % (len(zt.namelist()), len(zf.namelist()), same))
    bad = []
    for n in zf.namelist():
        if n.endswith(".xml") or n.endswith(".rels"):
            try:
                etree.fromstring(zf.read(n))
            except Exception as e:
                bad.append((n, str(e)[:50]))
    (ok if not bad else fail)(msgs, "XML 解析异常部件: %s" % (bad or "无"))

    wb0 = openpyxl.load_workbook(args.template, data_only=True,
                                 keep_vba=args.template.lower().endswith(".xlsm"))
    s0 = wb0[ff.detect_sheet(wb0)]
    wb = openpyxl.load_workbook(args.file, keep_vba=args.file.lower().endswith(".xlsm"))
    s = wb[ff.detect_sheet(wb)]
    cols = ff.attr_map(s)
    rows = ff.data_rows(s, cols)

    # ---------- B. 表头未破坏 ----------
    diff = [(r, c) for r in ff.HEADER_ROWS for c in range(1, s0.max_column + 1)
            if s0.cell(r, c).value != s.cell(r, c).value]
    (ok if not diff else fail)(msgs, "第1–6行（表头+示例）与空白模板差异 %d 处 %s" % (len(diff), diff[:4]))

    # ---------- C. 必填（模板自带的 Vereist 列）----------
    defs = find_def_sheet(wb0)
    if defs is not None:
        req_col = None
        for r in range(1, min(defs.max_row, 6) + 1):
            for c in range(1, defs.max_column + 1):
                v = str(defs.cell(r, c).value or "")
                if v.strip() in REQ_WORDS:
                    req_col = c
                    break
            if req_col:
                break
        hard, cond = [], []
        for r in range(2, defs.max_row + 1):
            field = defs.cell(r, 2).value
            status = str(defs.cell(r, req_col).value or "").strip() if req_col else ""
            if field in (None, ""):
                continue
            a = ff.alias(field)
            c = cols.get(a)
            if c is None:
                continue
            filled = [x for x in rows if not ff.is_blank(s.cell(x, c).value)]
            label = str(defs.cell(r, 3).value or a)[:40]
            if status.lower().startswith(("vereist", "erforderlich", "obligatoire", "obligatorio", "required")) \
                    and len(filled) < len(rows):
                hard.append((label, a, "%d/%d" % (len(filled), len(rows))))
            elif status.lower().startswith(("voorwaardelijk", "bedingt", "sous condition",
                                            "condicional", "conditional")) and not filled:
                cond.append((label, a))
        (ok if not hard else fail)(msgs, "硬性必填缺填: %s" % (hard or "0 项"))
        if cond:
            warn(msgs, "条件必填且整列为空 %d 项（多为电池/危险品/尺寸区间等不适用项，需人工确认）: %s"
                 % (len(cond), [x[1] for x in cond[:12]]))

    # ---------- D. 枚举合法性 ----------
    av = ff.AllowedValues(args.template)
    bad_enum, checked = [], 0
    for a, c in cols.items():
        lst = av.get(args.pt, str(s.cell(ff.ATTR_ROW, c).value))
        vals = {a: [s.cell(r, c).value for r in rows]}
        for r in rows:
            v = s.cell(r, c).value
            if ff.is_blank(v) or not isinstance(v, str):
                continue          # 数字列（价格/尺寸/重量）不参与枚举校验
            if lst is None:
                continue
            checked += 1
            if v.strip() not in lst:
                bad_enum.append((r, a, v, lst[:5]))
    (ok if not bad_enum else fail)(msgs, "枚举校验：受控文本单元格 %d 个，越界 %d %s"
                                   % (checked, len(bad_enum), bad_enum[:5]))
    av.close()

    # ---------- E. 文案规则 ----------
    def text_rule(label, t, limit, forbid=False, promo=True, contact=True):
        t = "" if t is None else str(t)
        if not t:
            fail(msgs, "%s 为空" % label)
            return
        if len(t) > limit:
            fail(msgs, "%s 超长 %d>%d" % (label, len(t), limit))
        if t != t.strip():
            warn(msgs, "%s 首尾空格" % label)
        if "  " in t:
            warn(msgs, "%s 连续空格" % label)
        if forbid and {ch for ch in t if ch in FORBID}:
            fail(msgs, "%s 含禁用符号 %s" % (label, sorted({ch for ch in t if ch in FORBID})))
        low = t.lower()
        for p in (PROMO if promo else []):
            if p in low:
                fail(msgs, "%s 含促销词「%s」" % (label, p))
        for p in (CONTACT if contact else []):
            if p in low:
                fail(msgs, "%s 含联系方式「%s」" % (label, p))
        if any(unicodedata.east_asian_width(ch) in ("F", "W") for ch in t):
            warn(msgs, "%s 含全角字符" % label)

    probe = rows[:1] + rows[1:2]
    for r in probe:
        text_rule("r%d 标题" % r, s.cell(r, cols.get("item_name#1.value", 0)).value, LIMITS["item_name"], True) \
            if "item_name#1.value" in cols else None
        if "title_differentiation#1.value" in cols:
            text_rule("r%d 副标题" % r, s.cell(r, cols["title_differentiation#1.value"]).value,
                      LIMITS["title_differentiation"], True)
        if "product_description#1.value" in cols:
            text_rule("r%d 描述" % r, s.cell(r, cols["product_description#1.value"]).value,
                      LIMITS["product_description"])
        for i in range(1, 6):
            a = "bullet_point#%d.value" % i
            if a in cols:
                text_rule("r%d 五点%d" % (r, i), s.cell(r, cols[a]).value,
                          LIMITS.get("bullet_point#%d" % i, LIMITS["bullet_point"]))
        if "generic_keyword#1.value" in cols:
            text_rule("r%d 后台关键词" % r, s.cell(r, cols["generic_keyword#1.value"]).value,
                      LIMITS["generic_keyword"], True, promo=False, contact=False)
    if "item_name#1.value" in cols and "generic_keyword#1.value" in cols:
        title = str(s.cell(rows[0], cols["item_name#1.value"]).value or "")
        kw = str(s.cell(rows[0], cols["generic_keyword#1.value"]).value or "").lower()
        tw = {w.strip(",.").lower() for w in title.split()}
        dup = sorted({w for w in kw.split() if w in tw})
        if dup:
            warn(msgs, "后台关键词与标题重复 %s（亚马逊建议换同义词）" % dup)

    # ---------- F. 变体结构 ----------
    if "parentage_level#1.value" in cols:
        PL = cols["parentage_level#1.value"]
        parents = [r for r in rows if str(s.cell(r, PL).value).strip().lower().startswith(
            ("eltern", "ouder", "parent", "principal", "padre", "vater"))]
        for r in rows:
            if r in parents:
                for a in ("shapewear_size#1.size", "color#1.value"):
                    if a in cols and not ff.is_blank(s.cell(r, cols[a]).value):
                        fail(msgs, "r%d 父体填了 %s" % (r, a))
            else:
                for a in ("child_parent_sku_relationship#1.parent_sku", "shapewear_size#1.size", "color#1.value"):
                    if a in cols and ff.is_blank(s.cell(r, cols[a]).value):
                        fail(msgs, "r%d 子体缺 %s" % (r, a))
        if "variation_theme#1.name" in cols:
            vt = {s.cell(r, cols["variation_theme#1.name"]).value for r in rows}
            if len(vt) > 1:
                fail(msgs, "变体主题不一致: %s" % vt)
        combos = [(s.cell(r, cols["shapewear_size#1.size"]).value, s.cell(r, cols["color#1.value"]).value)
                  for r in rows if "shapewear_size#1.size" in cols and "color#1.value" in cols]
        if combos and len(set(combos)) != len(combos):
            fail(msgs, "尺码×颜色组合重复")
    for a in ("list_price#1.value_with_tax",):
        if a in cols and "purchasable_offer#1.our_price#1.schedule#1.value_with_tax" in cols:
            for r in rows:
                lp = s.cell(r, cols[a]).value
                op = s.cell(r, cols["purchasable_offer#1.our_price#1.schedule#1.value_with_tax"]).value
                if isinstance(lp, (int, float)) and isinstance(op, (int, float)) and lp < op:
                    fail(msgs, "r%d 目录价 < 售价" % r)

    # ---------- G. 与源站点逐格对照 ----------
    if args.source:
        src = ff.open_source(args.source)
        cs = ff.attr_map(src)
        srows = ff.data_rows(src, cs)
        skua, skut = "contribution_sku#1.value", "contribution_sku#1.value"
        map_t = {str(s.cell(r, cols[skut]).value).strip(): r for r in rows if skut in cols}
        miss, extra = [], []
        for rs in srows:
            sku = str(src.cell(rs, cs[skua]).value).strip() if skua in cs else None
            if sku is None or sku not in map_t:
                continue
            rt = map_t[sku]
            for a, csrc in cs.items():
                sv = src.cell(rs, csrc).value
                ct = cols.get(a)
                if ct is None:
                    continue
                tv = s.cell(rt, ct).value
                if ff.is_blank(sv):
                    if not ff.is_blank(tv):
                        extra.append((sku, a, tv))
                elif ff.is_blank(tv):
                    miss.append((sku, a, sv))
        (ok if not miss and not extra else warn)(
            msgs, "与源表逐格对照：缺失 %d 格 / 多余 %d 格（父体行、源站点没有的列属正常）" % (len(miss), len(extra)))

    print("== 通过 ==")
    for m in msgs["OK"]:
        print("  ✓", m)
    print("== 待注意 ==")
    for m in msgs["WARN"]:
        print("  !", m)
    print("== 不合规 ==")
    for m in msgs["FAIL"]:
        print("  ✗", m)
    print("\nSUMMARY  FAIL=%d WARN=%d  (数据行 %d，属性 %d)"
          % (len(msgs["FAIL"]), len(msgs["WARN"]), len(rows), len(cols)))
    sys.exit(1 if msgs["FAIL"] else 0)


if __name__ == "__main__":
    main()
