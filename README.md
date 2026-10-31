# Amazon 多站点上传表翻译填写

把一个站点已填好的亚马逊上传表，转换到另一个站点的空白模板中。德国站既可以作为源站，也可以作为目标站。能否直接复用现有配置，取决于两个站点的模板字段和语言映射；没有现成映射时，需要先补充配置。脚本按系统属性名对齐列，读取目标模板的下拉选项，最后生成一份新的上传表供检查。

这个仓库目前主要针对 EU 站点的泳衣模板。换品类或模板版本时，先检查字段和下拉值，再填写。

目前仓库已有以德语上传表为源的荷兰语、英语（英国）、法语（法国和比利时）及西班牙语映射示例。其他方向（包括翻译回德语）可以按同一流程处理，但需要准备对应方向的词表、文案和目标模板允许值；不能直接把现有德语源表配置反过来用。

## 开始前准备

- Python 3.9 或更高版本，并安装 `openpyxl`、`lxml`。
- 一份已填好的源站上传表，以及目标站从 Amazon 下载的空白模板。
- 确认目标模板对应的商品类目。不要把某个 ASIN 的预筛模板用于其他类目。

```bash
pip install openpyxl lxml
```

仓库里的 `config/runs/` 是历史任务配置，用来参考字段映射和文案写法。它们的源文件、目标模板路径需要按自己的电脑修改；仓库没有附带那些上传表。

## 一次填写怎么做

以下命令从仓库根目录运行，把文件名和 product type 换成自己的实际值。

```bash
# 1. 看源表和目标模板有哪些属性
python scripts/probe.py "源站上传表.xlsx" "目标站空白模板.xlsm" --values --diff

# 2. 查看目标站模板的允许值
python scripts/enums.py "目标站空白模板.xlsm" --pt SWIMWEAR --all --out allowed_target.tsv

# 3. 复制一份相近的配置，修改 source.path、target.path、文案和取值映射
#    字段说明见 config/schema.md

# 4. 先预览将要填写的内容
python scripts/fill.py "我的配置.json" --dry-run

# 5. 生成新文件
python scripts/fill.py "我的配置.json" --strict

# 6. 检查产物
python scripts/check.py --template "目标站空白模板.xlsm" --file "生成的上传表.xlsm" --source "源站上传表.xlsx" --pt SWIMWEAR
```

`--dry-run` 不写文件。`--strict` 会在取值映射缺项时中止；不加这个参数时，脚本会跳过未确定的值，并生成同名的 `.pending.tsv` 清单。清单中的字段需要人工确认后再上传。

检查结果里的 `FAIL=0` 是继续人工审核的前提。还要核对商品类目、文案、价格、变体关系以及需要卖家确认的合规信息。

## 配置里的几类字段

| 字段 | 用途 |
|---|---|
| `source.path`、`target.path` | 源表和目标空白模板的路径 |
| `text` | 目标语言的标题、五点、描述等文案 |
| `copy` | SKU、品牌、型号等可以原样搬运的字段 |
| `value_map` | 源值到目标值的明确映射 |
| `vocab` | 从 `reference/vocab_de_eu.tsv` 读取历史对照 |
| `numeric` | 需要按数字写入的字段 |
| `naming` | 生成文件名中的国家和 SKU |

同一字段的处理顺序是 `text` → `copy` → `value_map` / `vocab`。`value_map` 可以覆盖词表中同一属性的映射。详细格式和 alias 写法见 [配置说明](config/schema.md)。

脚本只处理源表中已有值、目标模板中也存在的属性。没有覆盖到的属性会在运行输出中列出来；这些提示要逐项看，不能只看是否生成了文件。

## 文件在哪里

| 路径 | 内容 |
|---|---|
| `SKILL.md` | 给 AI 助手使用的操作说明 |
| `scripts/probe.py`、`scripts/enums.py` | 查看模板结构和允许值 |
| `scripts/fill.py`、`scripts/check.py` | 填表和检查 |
| `scripts/pick_template.py` | 根据泳衣类目推荐目标模板 |
| `config/schema.md` | 配置字段说明 |
| `config/runs/` | 历史任务配置示例 |
| `reference/` | 站点信息、历史取值对照和文案样例 |
| `docs/教程.xlsx` | 表格版操作教程 |

## 使用范围

脚本按目前接触过的 EU 上传表结构处理：第 5 行是系统属性名，第 7 行开始是数据区。其他模板结构需要先用 `probe.py` 核对。目标站下拉项和必填规则可能随模板变化，始终以本次下载的模板为准。

项目采用 MIT License。
