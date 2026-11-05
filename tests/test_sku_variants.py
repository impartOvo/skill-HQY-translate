import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from sku_variants import compare


class SkuVariantTests(unittest.TestCase):
    def test_known_colors_and_size(self):
        self.assertEqual([x[1] for x in compare("X2C1YY1167-zise-L", "Lila", "L")], ["ok", "ok"])
        self.assertEqual([x[1] for x in compare("abc-junlv-L", "Grün", "L")], ["ok", "ok"])
        self.assertEqual([x[1] for x in compare("abc-zise-L", "Grün", "M")], ["mismatch", "mismatch"])

    def test_unknown_values_are_not_passes(self):
        self.assertEqual([x[1] for x in compare("abc-unknown-L", "Lila", "40")], ["unknown", "unknown"])
        self.assertEqual(compare("parent", "", "")[0][1], "unknown")

    def test_check_uses_sku_and_template_dropdown(self):
        with tempfile.TemporaryDirectory() as directory:
            template = Path(directory) / "template.xlsx"
            output = Path(directory) / "output.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Template"
            attrs = ["contribution_sku#1.value", "parentage_level#1.value", "color#1.value", "shapewear_size#1.size"]
            for i, attr in enumerate(attrs, 1):
                ws.cell(5, i, attr)
            dd = wb.create_sheet("Dropdown Lists")
            dd["A1"], dd["A2"] = "Lila", "Grün"
            dd["B1"], dd["B2"] = "L", "M"
            for name, ref in [("SWIMWEARcolor1.value", "'Dropdown Lists'!$A$1:$A$2"),
                              ("SWIMWEARshapewear_size1.size", "'Dropdown Lists'!$B$1:$B$2")]:
                wb.defined_names.add(DefinedName(name, attr_text=ref))
            wb.save(template)
            ws.append([])
            for col, value in enumerate(["abc-zise-L", "child", "Grün", "M"], 1):
                ws.cell(7, col, value)
            wb.save(output)
            result = subprocess.run([sys.executable, str(ROOT / "scripts/check.py"), "--template", str(template),
                                     "--file", str(output)], text=True, encoding="utf-8", capture_output=True,
                                    env={**os.environ, "PYTHONIOENCODING": "utf-8"})
            self.assertEqual(result.returncode, 1, result.stderr + result.stdout)
            self.assertIn("SKU 颜色 zise 与表格颜色 Grün 不一致", result.stdout)
            self.assertIn("SKU 尺码 L 与表格尺码 M 不一致", result.stdout)
            self.assertIn("越界 0", result.stdout)  # Both wrong values are valid dropdown options.


if __name__ == "__main__":
    unittest.main()
