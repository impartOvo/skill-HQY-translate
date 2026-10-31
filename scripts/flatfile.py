# -*- coding: utf-8 -*-
"""flatfile.py — 亚马逊多站点上传表（Flat File / 库存模板）操作公共库。

设计要点
1) 一切按「属性名」而不是列号定位：列会随模板版本移动，属性名不会。
   alias() 把 `color[marketplace_id=XXX][language_tag=nl_NL]#1.value` 归一成 `color#1.value`，
   于是同一份配置可以在荷兰/英国/法国/西班牙/德国模板之间通用。
2) 枚举值只认目标模板自己的 definedNames → 'Dropdown Lists' 区间，不靠猜。
   注意：definedName 里 `[ ] # =` 全被删掉了，构造键时也必须删。
3) 写入走 XML 手术（只替换目标 sheet part + sharedStrings），其余部件字节级保留，
   openpyxl/Excel 保存会丢图片、改 sharedStrings、动 drawing 引用。
"""
from __future__ import annotations

import io
import os
import re
import sys
import json
import zipfile
import shutil
from xml.etree import ElementTree

import openpyxl

TEMPLATE_SHEETS = ("Vorlage", "Sjabloon", "Modèle", "Plantilla", "Template", "Modelo", "Plantilla")
ATTR_ROW = 5          # 第 5 行 = 系统属性名
DATA_START = 7        # 第 7 行起 = 数据区（第 6 行是灰色示例行）
HEADER_ROWS = (1, 2, 3, 4, 5, 6)


def utf8_stdio():
    """把 stdout 切成 UTF-8。幂等：重复调用不会再包一层
    （否则旧 wrapper 被回收时会关掉底层 buffer，后面 print 就报 I/O operation on closed file）"""
    try:
        if getattr(sys.stdout, "_dsh_utf8", False):
            return
        wrapper = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        wrapper._dsh_utf8 = True
        sys.stdout = wrapper
    except Exception:
        pass


# ---------------------------------------------------------------- 属性归一化
def alias(attr) -> str:
    """`color[marketplace_id=X][language_tag=nl_NL]#1.value` -> `color#1.value`"""
    if attr is None:
        return ""
    s = str(attr).strip()
    s = re.sub(r"\[[^\]]*\]", "", s)
    return s


def name_key(attr) -> str:
    """definedName 用的键：再删掉 `#`（Amazon 生成名字时会去掉 [ ] # =）。"""
    return re.sub(r"[\[\]#=]", "", str(attr or ""))


# ---------------------------------------------------------------- 模板读取
def detect_sheet(wb) -> str:
    for s in TEMPLATE_SHEETS:
        if s in wb.sheetnames:
            return s
    raise SystemExit("找不到模板工作表（Vorlage/Sjabloon/Modèle/Plantilla/Template）: %s" % wb.sheetnames)


def open_source(path):
    wb = openpyxl.load_workbook(path, data_only=True, keep_vba=path.lower().endswith(".xlsm"))
    ws = wb[detect_sheet(wb)]
    return ws


def attr_map(ws, row: int = ATTR_ROW):
    """alias -> 列号（1-based），同 alias 取第一列"""
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row, c).value
        if v not in (None, ""):
            m.setdefault(alias(v), c)
    return m


def meta(ws):
    """返回 (marketplace_id, language_tag, product_type 示例值)"""
    mp = lang = ""
    for c in range(1, min(ws.max_column, 60) + 1):
        v = str(ws.cell(ATTR_ROW, c).value or "")
        m = re.search(r"marketplace_id=([A-Z0-9]+)", v)
        if m and not mp:
            mp = m.group(1)
        m = re.search(r"language_tag=([A-Za-z_]+)", v)
        if m and not lang:
            lang = m.group(1)
    return mp, lang


def sku_col(cols):
    return cols.get("contribution_sku#1.value") or cols.get("seller_sku#1.value")


def data_rows(ws, cols, start: int = DATA_START):
    """SKU 非空的行号列表"""
    c = sku_col(cols)
    if not c:
        return [r for r in range(start, ws.max_row + 1)
                if any(ws.cell(r, x).value not in (None, "") for x in range(1, min(ws.max_column, 12) + 1))]
    return [r for r in range(start, ws.max_row + 1) if ws.cell(r, c).value not in (None, "")]


# ---------------------------------------------------------------- 官方允许值
class AllowedValues:
    """从 workbook.xml 的 definedNames 解析出「某 product type + 某属性」的官方允许值。"""

    def __init__(self, template_path: str):
        self.path = template_path
        self.z = zipfile.ZipFile(template_path)
        wbx = self.z.read("xl/workbook.xml").decode("utf-8")
        self.names = re.findall(r'<definedName name="([^"]+)"[^>]*>([^<]*)</definedName>', wbx)
        self.wb = openpyxl.load_workbook(template_path, data_only=True,
                                         keep_vba=template_path.lower().endswith(".xlsm"))
        self.dd = self.wb["Dropdown Lists"] if "Dropdown Lists" in self.wb.sheetnames else None
        self._cache = {}

    def _range(self, ref):
        m = re.match(r"'Dropdown Lists'!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$", ref)
        if not m or self.dd is None:
            return None
        c1, r1 = m.group(1), int(m.group(2))
        c2, r2 = (m.group(3) or c1), int(m.group(4) or m.group(2))
        out = []
        for c in range(openpyxl.utils.column_index_from_string(c1),
                       openpyxl.utils.column_index_from_string(c2) + 1):
            for r in range(r1, r2 + 1):
                v = self.dd.cell(r, c).value
                if v not in (None, ""):
                    out.append(str(v).strip())
        return out

    def get(self, product_type: str, attr: str):
        key = (product_type + name_key(attr)).replace(" ", "")
        if key in self._cache:
            return self._cache[key]
        res = None
        for n, ref in self.names:
            if n.replace(" ", "") == key:
                res = self._range(ref)
                if res is not None:
                    break
        self._cache[key] = res
        return res

    def close(self):
        try:
            self.z.close()
        except Exception:
            pass


# ---------------------------------------------------------------- XML 手术写入
def colname(n: int) -> str:
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def sheet_part(z: zipfile.ZipFile, sheet_title: str) -> str:
    wb = ElementTree.fromstring(z.read("xl/workbook.xml"))
    rels = ElementTree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2t = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    rel_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    for sheet in wb.findall(".//{*}sheet"):
        if sheet.attrib.get("name") == sheet_title:
            target = rid2t[sheet.attrib[rel_attr]]
            return "xl/" + target.lstrip("/").removeprefix("xl/")
    raise SystemExit("sheet 未找到: " + sheet_title)


def write_rows_xml(target_path: str, out_path: str, sheet_title: str,
                   records, max_col: int, last_row: int):
    """records: [(row, [(col, value, is_number), ...]), ...]"""
    from xml.sax.saxutils import escape

    z = zipfile.ZipFile(target_path)
    part = sheet_part(z, sheet_title)

    sst = z.read("xl/sharedStrings.xml").decode("utf-8")
    count0 = int(re.search(r'<sst[^>]*count="(\d+)"', sst).group(1))
    uniq0 = int(re.search(r'<sst[^>]*uniqueCount="(\d+)"', sst).group(1))
    existing = re.findall(r"<si>(.*?)</si>", sst, re.S)
    index = {}
    for i, si in enumerate(existing):
        index.setdefault(re.sub(r"<[^>]+>", "", si), i)

    new_si = []

    def sid(text):
        if text in index:
            return index[text]
        idx = len(existing) + len(new_si)
        index[text] = idx
        new_si.append('<si><t xml:space="preserve">%s</t></si>' % escape(text))
        return idx

    rows_xml, total = [], 0
    for r, cells in records:
        parts = ['<row r="%d" spans="1:%d">' % (r, max_col)]
        for c, v, isnum in cells:
            ref = "%s%d" % (colname(c), r)
            if isnum:
                parts.append('<c r="%s"><v>%s</v></c>' % (ref, repr(float(v))))
            else:
                parts.append('<c r="%s" t="s"><v>%d</v></c>' % (ref, sid(str(v))))
            total += 1
        parts.append("</row>")
        rows_xml.append("".join(parts))

    sheet = z.read(part).decode("utf-8")
    assert "</sheetData>" in sheet, "目标 sheet 无 sheetData"
    sheet_new = sheet.replace("</sheetData>", "".join(rows_xml) + "</sheetData>", 1)
    sheet_new = re.sub(r'(<dimension ref="[A-Z]+\d+:[A-Z]+)\d+"',
                       lambda m: '%s%d"' % (m.group(1), last_row), sheet_new, count=1)

    sst_new = sst.replace("</sst>", "".join(new_si) + "</sst>")
    sst_new = re.sub(r'(<sst[^>]*\bcount=")(\d+)(")', lambda m: m.group(1) + str(count0 + total) + m.group(3),
                     sst_new, count=1)
    sst_new = re.sub(r'(uniqueCount=")(\d+)(")', lambda m: m.group(1) + str(uniq0 + len(new_si)) + m.group(3),
                     sst_new, count=1)

    repl = {part: sheet_new.encode("utf-8"), "xl/sharedStrings.xml": sst_new.encode("utf-8")}
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in z.infolist():
            data = repl.get(item.filename, z.read(item.filename))
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = item.external_attr
            zo.writestr(zi, data)
    z.close()
    os.replace(tmp, out_path)
    return {"cells": total, "new_strings": len(new_si), "part": part}


# ---------------------------------------------------------------- 小工具
def load_config(path: str):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def norm_value(v):
    return str(v).strip() if v is not None else ""


def is_blank(v):
    return v is None or (isinstance(v, str) and not v.strip())

