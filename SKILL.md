---
name: amazon-flatfile-l10n
description: 亚马逊同店铺跨站点同链接商品的上传表（Flat File / 库存模板）多语言填写。按属性名自动对齐可翻译列、按目标站点官方下拉列表校验枚举值、XML 级写入保格式、上传前规则体检。适用于德国↔荷兰/英国/法国/比利时/西班牙/意大利等站点之间搬运 Listing（标题、五点、描述、后台关键词、材质尺码颜色等），也适用于新建站点的批量上架。
version: 1.1.0
license: MIT
---

# 亚马逊多站点上传表翻译填写（amazon-flatfile-l10n）

## 何时用

- 同一店铺、不同站点、**同一链接（同一套 SKU/EAN/ASIN）** 的商品要补上架：德国站填好了，要荷兰/英国/法国/比利时/西班牙站。
- 站点之间品类模板一致（都是 SWIMWEAR / SHIRT / APPAREL 之类），需要把源站点内容翻成目标站语言并填进目标站**空白模板**。
- 只做「填数据区」，**不重建文件、不改表头、不动格式**。

## 硬规则（违反就算失败）

1. **枚举值只能填目标站点 official dropdown 里的值**。先从目标模板 `workbook.xml → definedNames → 'Dropdown Lists'` 读出允许集合，再挑值；查不到就停下报错，绝不猜近似值。
2. **源站点没写的字段，目标站点也不写**。特别注意 EU 新增合规声明（`compliance_has_lining`、`compliance_has_inner_brief`、`compliance_waist_tightness_type`、`compliance_lining_material`）：这是实质申报，客户没确认就不要替他们声明。
3. **父体行不填 size / 颜色 / 价格**；子体行必须有 `parent_sku`、size、颜色，且 `variation_theme` 全表一致。
4. **文案上限**：主标题 75、副标题 125、五点各 500（内部标准：第 2 条 ≤700）、描述 2000、后台关键词 250。禁用 `! ? $ _ { } ^ ¬ ¦ ~ # < > * | ; " \ /`、全角标点、促销词、联系方式、HTML。
5. **写入只能走 XML 手术**（zip 内替换目标 sheet part + sharedStrings）。用 openpyxl 保存 xlsm 会丢图片、改 sharedStrings、动 drawing，等于毁掉模板。
6. **交付文件不要用 WPS 二次保存**（会把 44 个 zip 条目变成 55 个、重编码图片）。产出命名 `<目标表名>_<语言>填写_德国<编号>.xlsm`，与目标表同目录，原表不动。

## 模板结构（EU 各站通用）

| 行 | 内容 |
|---|---|
| 1 | 隐藏签名（`&TemplateSignature=` base64，别动） |
| 3 | 分组标题 |
| 4 | 本地化字段名（人看的） |
| 5 | **系统属性名**（脚本对齐用，含 `marketplace_id=` / `language_tag=`） |
| 6 | 灰色示例行（别照抄样式） |
| 7+ | **数据区** |

工作表名按站点变：德国 `Vorlage`、荷兰 `Sjabloon`、法国/比利时 `Modèle`、西班牙 `Plantilla`、英国 `Template`。
必填口径看模板自带的数据定义页（`Gegevensdefinities` / `Datendefinitionen` / `Définitions des données` / `Definiciones de datos`），最后一列是 `Vereist / Erforderlich / Obligatoire / Obligatorio / Required`。

## 标准流程

```bash
# 1) 摸结构：站点、语言、属性数、数据行、已填属性、两表属性差异
python scripts/probe.py "德国.xlsx" "目标模板.xlsm" --values --diff

# 2) 导出目标站官方允许值（枚举白名单）
python scripts/enums.py "目标模板.xlsm" --pt SWIMWEAR --all --out reference/allowed_target.tsv

# 3) （可选）从历史成交文件反推跨语言取值对照表
python scripts/vocab.py --pair "德国.xlsx|nl_NL|荷兰已填.xlsm" \
                        --pair "德国.xlsx|fr_FR|法国已填.xlsm" \
                        --out reference/vocab.tsv

# 4) 写配置（text 文案 / copy 原样搬 / numeric 数字 / value_map 兜底 / vocab 引用对照表）
#    样例见 config/de_to_nl_2399.json

# 5) 写入（XML 手术，原表不动）
python scripts/fill.py config/de_to_nl_2399.json

# 6) 体检（FAIL 必须为 0 才能上传）
python scripts/check.py --template "目标模板.xlsm" --file "产出.xlsm" --source "德国.xlsx" --pt SWIMWEAR
```

## 列对齐（关键设计）

按**属性名**不按列号。归一化规则：
`color[marketplace_id=A1805IZSGTT6HS][language_tag=nl_NL]#1.value` → `color#1.value`

于是同一份配置可以在德国/荷兰/法国/西班牙/英国模板之间通用，模板升级挪列也不受影响。
源表有值但目标模板没这属性（例：`item_shape` 荷兰模板没有）时，脚本会列出来，人工判断丢弃即可。

## 枚举取值怎么查（最坑的一步）

```python
from flatfile import AllowedValues
av = AllowedValues("目标模板.xlsm")
av.get("SWIMWEAR", "color[marketplace_id=A1805IZSGTT6HS][language_tag=nl_NL]#1.value")
# → ['Beige','Blauw','Brons',...,'Zwart']
```
注意：definedName 生成时 `[ ] # =` 全被删掉，拼键时也要删，否则 2600+ 个名字一个都匹配不上。
尺码是复合键（size_system + age_range + size_class 三段拼起来），必须走 `AllowedValues`，别手拼。

## 交付前必过

`check.py` 的 11 项：zip 条目集合 / XML 可解析 / 1–6 行零差异 / 硬性必填全填 / 枚举零越界 / 源→目标逐格零缺失零多余 / 文案不超长不违规 / 变体关系正确 / 价格关系正确 / 条件必填逐条确认 / 合规声明不自造。

## 参考数据（随仓库）

- `reference/vocab_de_eu.tsv`：78 个属性、222 条实测对照（德→荷/英/法/法比/西），来自真实上传过的文件。
- `reference/attributes_swimwear_nl.tsv`：SWIMWEAR 223 个属性 + 必填级别 + 荷兰列号。
- `reference/marketplaces.tsv`：6 个站点的 marketplace_id / language_tag / 工作表名。
- `reference/marketplace_specific_values.tsv`：**每站都不同**的枚举取值（merchant_shipping_group、browse node、尺码写法、GTIN 豁免词、parentage 译法），照抄源站必错。
- `reference/copy_samples.md`：同一款商品在 5 种语言里的真实标题、五点、描述（翻译风格基准）。
- `docs/教程.xlsx`：十一页签教程表格，含流程、属性处理、跨语言取值、踩坑、验收清单。
- 现成配置：`config/de_to_nl_2399.json`、`de_to_fr_2399.json`、`de_to_es_2399.json`、`de_to_en_1391.json`、`de_to_be_2375.json`（文案取自各自历史上架文件）。

## 已验证的战绩（可当回归基线）

| 源 | 目标 | 行数 | 写入格数 | check.py |
|---|---|---|---|---|
| 德国 2399（Tankini 套装） | 荷兰 | 21（1 父 + 4 色 × 5 码） | 1568 | FAIL 0 |
| 德国 2399 | 英国 / 法国 / 西班牙 / 比利时 | 各 21 | 各 1568 | FAIL 0 |
| 德国 1391（Monokini 连体） | 英国 / 法国 / 西班牙 / 比利时 | 各 21 | 各 1238 | FAIL 0 |
| 德国 2375（比基尼套装） | 比利时（法语文案） | 21 | 1611 | FAIL 0 |

9 份真实跑批的配置留在 `config/runs/`（路径已换成 `{{ZINIAO_EU}}`、`{{SKILL}}` 占位符，设同名环境变量或用 `vard` 段即可复跑）。回归命令：

```bash
python scripts/fill.py config/runs/cfg_2399_fr_FR.json --dry-run   # 只看映射，不写文件
```

## 跑批实测踩坑（2026-10-30，10 个组合一次跑通后补的）

1. **`copy` 必须优先于 vocab/value_map**：vocab 是从历史表反推的，里面也有 `contribution_sku`、`model_number`、`product_id_value` 这些"值是 SKU/ASIN"的属性，先走 vmap 就会报「值映射缺项」而不写文件。
2. **父体行不算必填**：`size/颜色/价格/配送/parent_sku` 这些只有子体行该有；check.py 用 `VARIANT_ONLY` + `parent_rows` 按子体行计数，否则每次都是假 FAIL。
3. **条件必填要先判再判硬必填**：ES 模板的 `Obligatorio condicionalmente` 会被 `startswith("obligatorio")` 误判成硬必填，一次刷 70+ 条假 FAIL。各站状态词：`Vereist / Erforderlich / Obligatoire / Obligatorio / Obligatorio condicionalmente / Required`。
4. **模板可能被"预筛"过**：在后台按某个 ASIN 下载的模板，browse node 下拉会被筛成该商品**唯一一个**节点。荷兰那份「空白上传模版」只剩 `Bikini's > Sets` ——拿它上 Tankini/连体款就是错类目。判断方法：`enums.py --attr recommended_browse_nodes`，节点数 ≤5 就要重新下载通用模板。
5. **后台关键词不要撞标题**：`check.py` 会 WARN 列出重复词；新写的文案一律换成标题里没有的同义词/场景词。

## 换品类时要做的事

1. `probe.py` 确认 product type（`product_type#1.value`，如 SWIMWEAR / SHIRT）。
2. `enums.py --pt <新PT>` 重新导允许值。
3. `vocab.py` 用新品类历史文件重建对照表（没有历史文件就手工挑值写 `value_map`）。
4. 文案上限按类目微调 `check.py` 的 `LIMITS`。
