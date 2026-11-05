---
name: amazon-flatfile-l10n
description: 亚马逊上传表跨站翻译与已有表格补充新 SKU。按属性名对齐字段、按目标模板下拉值填写和校验，并可参照同父体同尺码子体补齐新增 SKU 的属性；保留 Excel 模板结构。
version: 1.2.0
license: MIT
---

# 亚马逊多站点上传表翻译填写（amazon-flatfile-l10n）

## 何时用

- 同一店铺、不同站点、**同一链接（同一套 SKU/EAN/ASIN）** 的商品要补上架：德国站填好了，要荷兰/英国/法国/比利时/西班牙站。
- 站点之间品类模板一致（都是 SWIMWEAR / SHIRT / APPAREL 之类），需要把源站点内容翻成目标站语言并填进目标站**空白模板**。
- 在已有上传表中增加了新的子 SKU，需参照同一父体已有变体补齐属性；这类任务可以不做翻译，只按 SKU 中的颜色、尺码更新对应属性并检查一致性。
- 只做「填数据区」，**不重建文件、不改表头、不动格式**。

## 硬规则（违反就算失败）

1. **枚举值只能填目标站点 official dropdown 里的值**。先从目标模板 `workbook.xml → definedNames → 'Dropdown Lists'` 读出允许集合，再挑值。
   **优先级：① 模板/历史表里已经出现过的值 → ② 允许列表里语义最近的那个 → ③ 都没有就留空**，留空项写进 `<产出>.pending.tsv` 清单逐条问客户「要不要填、填哪个」，绝不猜。`fill.py --strict` 才会改成硬中断。
2. **不从空缺处编造商品事实**。跨站翻译时源站没写的字段不自行补；在已有表补新子 SKU 时，可从同父体、同尺码的已填子体复制共用属性，但颜色、尺码需按新 SKU 和目标模板更新，合规声明须有依据，不能只因兄弟变体已填写就推断。
   特别注意 EU 新增合规声明（`compliance_has_lining`、`compliance_has_inner_brief`、`compliance_waist_tightness_type`、`compliance_lining_material`）：这是实质申报，客户没确认就不要替他们声明。
3. **父体行不填 size / 颜色 / 价格**；子体行必须有 `parent_sku`、size、颜色，且 `variation_theme` 全表一致。
   逐个子 SKU 核对颜色和尺码：例如 `DEMOSPU1167-zise-L` 的 `zise` 是紫色、`L` 是 L 码；`junlv` 是军绿色，可以对应目标下拉里的绿色（如 `Grün`）；`Pink` 在德语目标模板中对应 `Rosa`。填值仍须符合本次目标模板的允许值。`check.py` 会报告识别出的冲突；未识别的编码、目标颜色或未确认的尺码换算须人工处理，不能当作通过。像 `XXL (xx_l)` 这样的模板写法按尺码主体 `XXL` 核对；跨尺码系统换算仍须以模板允许值和已有同款变体为准。用户确认后再补充编码映射；不要凭 SKU 猜测未收录的编码，也不要自动改商品事实。
4. **文案上限**：主标题 75、副标题 125、五点各 500（内部标准：第 2 条 ≤700）、描述 2000、后台关键词 250。禁用 `! ? $ _ { } ^ ¬ ¦ ~ # < > * | ; " \ /`、全角标点、促销词、联系方式、HTML。
5. **写入只能走 XML 手术**（zip 内替换目标 sheet part + sharedStrings）。用 openpyxl 保存 xlsm 会丢图片、改 sharedStrings、动 drawing，等于毁掉模板。
6. **交付文件不要用 WPS 二次保存**（会把 44 个 zip 条目变成 55 个、重编码图片）。产出**必须**命名 `<国家>-<sku 简短编号>-上传表-dsh生成.xlsm`，落到目标表同目录、原表不动；`dsh生成` 这段一定要带，用来区分手工填的表。config 里写 `"naming": {"country": "荷兰", "sku": "2399"}` 即自动生成该文件名。

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
# 0) 按类目自动挑模板：表里基本都有 bikini / tankini / monokini 之类的类目词
python scripts/pick_template.py --source "德国.xlsx" --dir 荷兰      # 或 --scan 荷兰 只看清单

# 1) 摸结构：站点、语言、属性数、数据行、已填属性、两表属性差异
python scripts/probe.py "德国.xlsx" "目标模板.xlsm" --values --diff

# 2) 导出目标站官方允许值（枚举白名单）
python scripts/enums.py "目标模板.xlsm" --pt SWIMWEAR --all --out reference/allowed_target.tsv

# 3) （可选）从历史成交文件反推跨语言取值对照表
python scripts/vocab.py --pair "德国.xlsx|nl_NL|荷兰已填.xlsm" \
                        --pair "德国.xlsx|fr_FR|法国已填.xlsm" \
                        --out reference/vocab.tsv

# 4) 写配置（text 文案 / copy 原样搬 / numeric 数字 / value_map 兜底 / vocab 引用对照表 / naming 命名）
#    样例见 config/runs/cfg_2399_fr_FR.json

# 5) 写入（XML 手术，原表不动）；下拉里没有的取值会留空并列进 <产出>.pending.tsv
python scripts/fill.py config/runs/cfg_2399_fr_FR.json

# 6) 体检（FAIL 必须为 0 才能上传）
python scripts/check.py --template "目标模板.xlsm" --file "产出.xlsm" --source "德国.xlsx" --pt SWIMWEAR
```

## 在已有上传表中增加子 SKU（不翻译）

当用户已经在一份已填写的上传表中添加了 SKU、但新行属性尚空时，可参照同一父体的现有子体填写。先核对数据行中的 SKU、父 SKU 和变体主题，再将 SKU 末段尺码与现有子体的尺码匹配，选择**同父体、同尺码**的子体行作为参考。复制可确认属于同一商品的共用属性，保留新 SKU 本身、已有推荐操作和目标模板格式；父体行不能当子体参考。

颜色编码从 SKU 中解析后，按目标站语言换成目标模板实际允许的颜色值，再写入颜色标准值及颜色展示值。例如 `...-Pink-S` 在德国模板中填 `Rosa`，`...-junlv-L` 可映射为目标下拉中的绿色值 `Grün`。尺码使用同尺码参考子体中已经符合模板上下文的写法，例如 SKU 后缀 `XXL` 可对应表格中的 `XXL (xx_l)`。颜色和尺码字段填写后，要再次逐行检查与 SKU 是否相符，并确认受控字段来自该模板的下拉选项。

按本次用户要求，5 个 Pink 新 SKU 分别从同尺码子体行复制共用属性，颜色填 `Rosa`，价格及促销价格保持空白。不要从其他颜色的子体复制 ASIN/GTIN 等唯一商品标识，也不要复制颜色专属图片；这些信息缺失时保留空白并向用户说明。价格只有在用户明确要求时才留空或填写，不要把该实验处理方式推成所有任务的默认值。新行原有 SKU 和操作提示必须保留。

对 `.xlsm` 只改目标数据表及所需的 shared strings，保留工作簿的其他 zip 部件、数据验证和媒体文件。不要用 openpyxl 或 WPS 保存目标工作簿。完成后复核新增 SKU 数量、父体/尺码/颜色对应、目标下拉值、指定保持为空的价格字段，并确认输出文件可重新打开。

## 选模板：按类目自动挑（`pick_template.py`）

同一国家可能有好几份空白模板，**别乱用**：

- 从源表读 `swimwear_form_type` / `item_type_name` / 标题 → 判定类目（bikini / tankini / monokini / onepiece）；
- 扫目标国家文件夹里**没有数据行**的模板，看各自 `recommended_browse_nodes` 覆盖哪些类目；
- 排序：① 预筛模板且只覆盖本类目（后台按该 ASIN 生成的，最贴）→ ② 其它预筛 → ③ 全量模板（节点多者优先）。

实测（荷兰）：2399 Tankini → `takini set 模板.xlsm`；1391 Monokini → `荷兰 monokini空白模版.xlsm`（142 节点全量）；2375 Bikini → `荷兰空白上传模版.xlsm`（1 节点，专为该类目下载）。
**节点数 ≤5 就是预筛模板**，只能干那一个类目，拿去上别的类目 = 挂错类目。

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

`check.py` 会检查非数据部件内容、XML、表头、必填字段、目标模板下拉值、全部数据行的文案、变体和价格。与源表逐格对照及条件必填项会给出待注意提示；这些提示仍需人工核对。

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

