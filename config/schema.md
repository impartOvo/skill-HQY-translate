# 配置文件说明（fill.py 的输入）

一份配置 = 一次「源站点表 → 目标站点表」的填写任务。样例：`config/de_to_nl_2399.json`。

```json
{
  "source":  {"path": "…德国 2399 上传表.xlsx", "attr_row": 5, "data_start": 7},
  "target":  {"path": "…takini set 模板.xlsm",  "attr_row": 5, "data_start": 7},
  "output_suffix": "荷兰填写_德国2399",
  "vocab":   {"path": "reference/vocab_de_eu.tsv", "locale": "nl_NL"},
  "text":    { "item_name#1.value": "…" },
  "copy":    ["contribution_sku#1.value"],
  "numeric": ["list_price#1.value_with_tax"],
  "value_map": { "color#1.value": {"Schwarz": "Zwart"} }
}
```

## 字段

| 字段 | 必填 | 说明 |
|---|---|---|
| `source.path` | ✅ | 源站点上传表（通常是已填好的德国站表，.xlsx/.xlsm 都行） |
| `target.path` | ✅ | 目标站点**空白模板**（Amazon 下载的原始模板；不要用已填过的表当模板） |
| `source.attr_row` / `target.attr_row` | | 系统属性名所在行，默认 5 |
| `source.data_start` / `target.data_start` | | 数据区起始行，默认 7 |
| `output` | | 输出路径；不给就用 `<target 目录>/<target 名>_<output_suffix>.<ext>` |
| `output_suffix` | | 输出文件名后缀，建议 `<语言>填写_德国<编号>` |
| `vocab` | | `{path, locale}`：引用 `vocab.py` 产出的跨语言对照表，按 locale 列自动生成 value_map |
| `text` | | 常量文案，键=属性 alias，值=目标语文案；**只在源表该字段有值的行上写**（父体行自然跳过） |
| `copy` | | 原样搬运的属主 alias 列表（SKU、parent_sku、品牌、EAN/ASIN、型号、邮箱、FBA 通道…） |
| `numeric` | | 必须以数字写入的 alias（源表常把价格存成文本，不转会被当成字符串） |
| `value_map` | | 手工取值映射，键=属性 alias、值=源值→目标值；与 vocab 合并，**value_map 优先** |

## 三条约定

1. **alias 写法**：把属性名里的 `[marketplace_id=…]`、`[language_tag=…]` 去掉，保留 `#序号` 与后缀。
   `color[marketplace_id=A1805IZSGTT6HS][language_tag=nl_NL]#1.value` → `color#1.value`
   好处：同一份配置在德国/荷兰/法国/西班牙/英国模板之间通用。
2. **优先级**：`text` > `value_map`/`vocab` > `copy` > 「配置未覆盖」告警。
3. **安全阀**：`value_map` 查不到对应值、`numeric` 转不成数字时，脚本**不写这一格并报错退出（exit 2）**，绝不留半成品。

## 怎么拿到准确的 alias 清单

```bash
python scripts/probe.py "目标模板.xlsm" --values
```

会列出该模板全部属性 alias、所在列、填充状态；把要处理的复制进配置即可。
