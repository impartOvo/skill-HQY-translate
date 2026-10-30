# amazon-flatfile-l10n

> 仓库：**impartOvo/skill-amazon-translate** · 技能名：`amazon-flatfile-l10n`

亚马逊**同店铺 · 不同站点 · 同一链接**商品的**上传表（Flat File / 库存模板）多语言填写**工具包 + AI 技能。

> 德国站填好的表 → 荷兰 / 英国 / 法国 / 比利时 / 西班牙 站，只写数据区、不动表头格式、枚举值一律取目标站官方下拉列表。
> 已用真实数据验证：3 款泳衣（2375 / 2399 / 1391）× 5 个站点，21 行/款，逐格 0 差异、枚举 0 越界。

---

## 它能解决什么

| 痛点 | 本项目的做法 |
|---|---|
| 各站模板列顺序不同、模板一升级列号全变 | **按属性名对齐**（`color#1.value`），不按列号 |
| 枚举字段（颜色/尺码/领型/材质…）翻错就被拒 | 从模板 `definedNames → Dropdown Lists` 读**官方允许值**再挑，越界直接报错停下 |
| 用 openpyxl 保存 xlsm 会丢图片、毁模板 | **XML 手术**：只替换目标 sheet part + sharedStrings，其余 zip 条目字节级保留 |
| 上次翻的词和上次不一样 | `reference/vocab_de_eu.tsv` 是从**真实上传过的文件**反推的 222 条对照，跨批次一致 |
| 产出和手工填的表混在一起分不清 | 统一命名 **<国家>-<sku>-上传表-dsh生成.xlsm**（dsh生成 必带） |
| 上传后才发现超字符/踩禁用词 | `check.py` 上传前体检，11 项，FAIL=0 才交付 |

---

## 快速开始

```bash
# 依赖：Python 3.9+，openpyxl、lxml
pip install openpyxl lxml

# 1. 摸清两表结构（站点/语言/属性/数据行/属性差异）
python scripts/probe.py "德国 2399 上传表.xlsx" "takini set 模板.xlsm" --values --diff

# 2. 导出目标站官方允许值
python scripts/enums.py "takini set 模板.xlsm" --pt SWIMWEAR --all --out allowed_nl.tsv

# 3. 改 config 里的路径与文案，然后写入
python scripts/fill.py config/de_to_nl_2399.json

# 4. 体检
python scripts/check.py --template "takini set 模板.xlsm" \
                        --file "takini set 模板_荷兰填写_德国2399.xlsm" \
                        --source "德国 2399 上传表.xlsx" --pt SWIMWEAR
```

`fill.py` 输出示例：

```
[rows] 源数据行 21，写入 21 行，1568 格
[warn] 源表有值但未映射的属性（1）:
        - item_shape#1.value                目标模板无此属性
[write] {'cells': 1568, 'new_strings': 54, 'part': 'xl/worksheets/sheet5.xml'}
[ok] 原表未改动；仅替换 sheet part 与 sharedStrings，其余部件字节级保留
```

---

## 仓库结构

```
SKILL.md                      # AI 技能入口（可直接放进 .dsh/skills/ 或作为技能上传）
README.md
LICENSE
scripts/
  flatfile.py                 # 公共库：属性归一化 / 官方允许值解析 / XML 手术写入
  probe.py                    # 探查结构与属性差异
  enums.py                    # 导出某站点某 product type 的官方允许值
  vocab.py                    # 按 SKU+属性名反推跨语言取值对照表
  fill.py                     # 配置驱动的翻译填写（XML 手术）
  check.py                    # 上传前体检（11 项）
  make_tutorial.py            # 生成 docs/教程.xlsx
  legacy/                     # 第一代单任务脚本（仅供追溯，勿用；见 legacy/README.md）
config/
  de_to_nl_2399.json          # 德国→荷兰 实测样例（泳衣 Tankini 套装，21 行）
  de_to_fr_2399.json          # 德国→法国（法语文案取自真实上架文件）
  de_to_es_2399.json          # 德国→西班牙
  de_to_en_1391.json          # 德国→英国（连体泳衣款）
  de_to_be_2375.json          # 德国→比利时（法语文案）
  runs/                       # ★ 9 份真实跑批配置（2399/1391 × 荷英法西比），路径用 {{ZINIAO_EU}} {{SKILL}} 占位
  vars.example.json           # 占位符变量示例（也可用同名环境变量）
  schema.md                   # 配置字段说明
reference/
  vocab_de_eu.tsv             # 78 属性 × 222 条 实测跨语言对照（德/荷/英/法/法比/西）
  attributes_swimwear_nl.tsv  # SWIMWEAR 223 属性 + 必填级别 + 荷兰列号
  marketplaces.tsv            # 6 站点 marketplace_id / language_tag / 工作表名
  marketplace_specific_values.tsv  # 每站不同的枚举（shipping group / browse node / 尺码写法 / GTIN 豁免词）
  char_limits.tsv             # 字段字符上限
  copy_samples.md             # 同款商品 5 语言真实文案样例
docs/
  教程.xlsx                    # 十一页签教程表格
```

---

## 作为 AI 技能使用

三种方式任选：

1. **DeepSeek Harness / 兼容 skills 的客户端**：把整个仓库目录放到 `~/.dsh/skills/amazon-flatfile-l10n/`（`SKILL.md` 在根目录即可被识别）。
2. **直接上传仓库给 AI**：把 `SKILL.md` + `reference/` + `config/schema.md` 一起给模型，配合 `scripts/` 就能跑完整流程。
3. **只当命令行工具**：`python scripts/fill.py ...`，人不介入。

---

## 配置怎么写

见 [`config/schema.md`](config/schema.md)，核心四段：

```json
{
  "source": {"path": "德国站上传表.xlsx"},
  "target": {"path": "目标站空白模板.xlsm"},
  "vocab":  {"path": "reference/vocab_de_eu.tsv", "locale": "nl_NL"},
  "text":   {"item_name#1.value": "…", "bullet_point#1.value": "…"},
  "copy":   ["contribution_sku#1.value", "brand#1.value", "…"],
  "numeric":["list_price#1.value_with_tax", "…"],
  "value_map": {"color#1.value": {"Schwarz": "Zwart"}}
}
```

优先级：`text` > `value_map`（含 vocab 载入）> `copy` > 报「配置未覆盖」。

---

## 数据来源与边界

- `reference/` 里的对照表与文案样例来自本店铺**已上传成功**的历史文件（德国 2375/2399/1391 与其荷兰、英国、法国、比利时、西班牙成品表），已脱敏为通用映射，未包含买家数据。
- 本工具**不做**：排名/收益预测、合规声明代填、图片处理、广告投放。合规字段（GPSR/DSA/包装法）只搬源站点已有值，缺就留空等客户确认。

## 已知限制

- 只覆盖 EU 扁平文件结构（第 5 行属性名 / 第 7 行起数据）；美国站模板结构不同需微调 `flatfile.ATTR_ROW`。
- 尺码复合键（size_system × age_range × size_class）依赖目标模板自身的 definedNames，跨模板版本可能变化，务必用 `enums.py` 现场导。
- 意大利/波兰/瑞典等站点暂无历史文件，`vocab_de_eu.tsv` 里没有对应列；补法：拿该站一份已上传的表跑一次 `vocab.py` 即可扩列。

## License

MIT
