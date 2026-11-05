"""Compare a child SKU's final color/size tokens with its filled values."""
import re
import unicodedata

# These are recognized meanings, not values to write into a template.
SKU_COLORS = {
    "zise": "purple", "junlv": "green", "lvse": "green",
    "heise": "black", "baise": "white", "hongse": "red",
    "lanse": "blue", "huise": "gray", "fense": "pink",
    "huangse": "yellow", "zongse": "brown", "mise": "beige",
    "pink": "pink", "haijunlan": "navy", "heimeihong": "black_pink",
}
COLOR_VALUES = {
    "purple": ("purple", "violet", "violett", "lila", "paars", "morado", "viola", "紫色"),
    "green": ("green", "grün", "groen", "vert", "verde", "军绿色", "绿色"),
    "black": ("black", "schwarz", "zwart", "noir", "negro", "nero", "黑色"),
    "white": ("white", "weiß", "weiss", "wit", "blanc", "blanco", "bianco", "白色"),
    "red": ("red", "rot", "rood", "rouge", "rojo", "rosso", "红色"),
    "blue": ("blue", "blau", "blauw", "bleu", "azul", "blu", "蓝色"),
    "gray": ("grey", "gray", "grau", "grijs", "gris", "grigio", "灰色"),
    "pink": ("pink", "rosa", "rose", "roze", "粉色"),
    "yellow": ("yellow", "gelb", "geel", "jaune", "amarillo", "giallo", "黄色"),
    "brown": ("brown", "braun", "bruin", "marron", "marrone", "棕色"),
    "beige": ("beige", "米色"),
    "navy": ("navy", "marine blau", "marineblau", "dunkelblau", "bleu marine"),
    "black_pink": ("schwarz&rosa", "schwarz und rosa", "black&pink", "black and pink"),
}
SIZES = {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "2XL", "3XL", "4XL", "5XL"}


def normalized(value):
    return unicodedata.normalize("NFKC", str(value or "")).strip().casefold()


def compare(sku, color, size):
    """Return (field, status, explanation); uncertain comparisons never pass."""
    parts = str(sku or "").strip().split("-")
    if len(parts) < 3 or not parts[-2] or not parts[-1]:
        return [("SKU", "unknown", "无法按款号-颜色-尺码识别末两段")]
    color_code, size_code = parts[-2].casefold(), parts[-1].upper()
    results = []
    family = SKU_COLORS.get(color_code)
    actual_color = normalized(color)
    if not family:
        results.append(("颜色", "unknown", "SKU 颜色编码 %r 未收录" % parts[-2]))
    elif not actual_color:
        results.append(("颜色", "unknown", "目标颜色为空或未找到颜色列"))
    elif actual_color in {normalized(v) for v in COLOR_VALUES[family]}:
        results.append(("颜色", "ok", "%s 对应 %s" % (parts[-2], color)))
    elif any(actual_color in {normalized(v) for v in vals} for vals in COLOR_VALUES.values()):
        results.append(("颜色", "mismatch", "SKU 颜色 %s 与表格颜色 %s 不一致" % (parts[-2], color)))
    else:
        results.append(("颜色", "unknown", "无法判断 SKU 颜色 %s 与表格颜色 %s 的对应关系" % (parts[-2], color)))
    actual_size = normalized(size).upper()
    # Amazon may decorate the size token, e.g. "XXL (xx_l)".
    size_match = re.match(r"^(XXXL|XXL|XXS|2XL|3XL|4XL|5XL|XL|XS|[SML])(?:\b|\s|\()", actual_size)
    if size_match:
        actual_size = size_match.group(1)
    if size_code not in SIZES:
        results.append(("尺码", "unknown", "SKU 尺码编码 %r 未收录" % parts[-1]))
    elif not actual_size:
        results.append(("尺码", "unknown", "目标尺码为空或未找到尺码列"))
    elif actual_size == size_code:
        results.append(("尺码", "ok", "%s 对应 %s" % (parts[-1], size)))
    elif actual_size in SIZES:
        results.append(("尺码", "mismatch", "SKU 尺码 %s 与表格尺码 %s 不一致" % (parts[-1], size)))
    else:
        results.append(("尺码", "unknown", "SKU 尺码 %s 与表格尺码 %s 是否对应需确认" % (parts[-1], size)))
    return results
