# 客户主数据合并管道

把三个来源的客户记录归一到统一口径，合并成一份主数据 `master.csv`；必填字段缺失的
记录分流到 `quarantine.csv`，并输出机器可读的运行报告 `run_report.json`。

## 仓库内容

| 路径 | 说明 |
| --- | --- |
| `merge.py` | 合并管道（读配置 → 抽取归一 → 合并 → 分流 → 写产物） |
| `mapping.json` | 映射配置：字段对应、来源优先级、必填字段、类型规则、输出路径 |
| `src_api/*.json` | 业务系统接口按日拉取的 JSON，每日一份（含升级前后两种客户号字段名） |
| `src_export.csv` | 定期导出的 CSV |
| `src_manual.xlsx` | 运营手工维护的 Excel（工作表 `客户`） |
| `tests/test_merge.py` | pytest 测试 |
| `master.csv` / `quarantine.csv` / `run_report.json` | 运行产物（由脚本生成） |

环境：Python 3 + pandas + openpyxl，直接本机运行，不依赖 Docker。

## 运行

```bash
python3 merge.py --config mapping.json            # 数据与产物都在仓库根目录
python3 merge.py --config mapping.json -v         # 带 DEBUG 日志
python3 merge.py --config mapping.json --output-dir out   # 产物另放目录
```

`--data-dir` 可指定数据根目录（默认取配置文件所在目录），配置里的相对路径都相对它解析。
正常结束返回 0；配置或输入不满足约定时抛出 `MergeError`，打印原因并返回 2。

测试：

```bash
python3 -m pytest -q
```

## 映射配置 `mapping.json`

**代码里没有任何业务字段名**：字段对应关系、改名归一、优先级、必填字段、类型规则全部
写在配置里。接口升级把客户号从 `cust_no` 改成 `customer_code` 这种情况，只要在
`sources[].field_map` 里把两个来源列名都映射到同一个目标字段即可，代码不用改。

```json
{
  "target_fields": ["customer_id", "name", "level", "credit_limit", "phone", "updated_at"],
  "key_field": "customer_id",
  "required_fields": ["customer_id", "name", "level"],
  "contact_fields": ["phone"],
  "conflict_fields": ["name", "level", "credit_limit", "phone"],
  "date_field": "updated_at",
  "field_labels": { "customer_id": "客户号", "...": "..." },
  "field_rules": { "customer_id": { "kind": "identifier" }, "...": {} },
  "sources": [
    {
      "name": "api", "priority": 1, "type": "json_dir", "path": "src_api",
      "date_formats": ["%Y-%m-%d %H:%M:%S"],
      "field_map": { "customer_code": "customer_id", "cust_no": "customer_id", "...": "..." }
    },
    { "name": "csv",   "priority": 2, "type": "csv",   "path": "src_export.csv",  "field_map": { "...": "..." } },
    { "name": "excel", "priority": 3, "type": "excel", "path": "src_manual.xlsx",
      "sheet": "客户", "header_row": 1, "field_map": { "...": "..." } }
  ],
  "output": { "master": "master.csv", "quarantine": "quarantine.csv", "report": "run_report.json" }
}
```

关键配置项：

- `target_fields`：目标字段集合，也是 `master.csv` 的列顺序。
- `key_field`：合并键，必须同时出现在 `required_fields` 里。
- `required_fields`：必填字段，缺失即进 quarantine。
- `contact_fields`：联系方式字段，跨来源取值改用“日期最新”规则（见下）。
- `conflict_fields`：参与跨来源冲突统计的字段（不应包含键与日期字段）。
- `date_field`：记录日期字段，用于最新取值与行级更新时间。
- `field_labels`：报告与隔离说明里的人读字段名。
- `field_rules`：每个目标字段的归一规则，`kind` 支持
  - `text`：去首尾空白；
  - `identifier`：去前缀（如客户号的 `C`）、去分隔符、可选仅保留数字与去前导零
    （客户号 `C000046` / `000046` / `046` 都归一为 `46`）；
  - `decimal`：去千分位等噪声字符，按 `decimal_places` 四舍五入（`111,373.71` → `111373.71`，
    `1000` → `1000.00`）；
  - `datetime`：按来源 `date_formats` 解析，输出统一为 `output_format`
    （Excel 的 `2026-10-07` 与接口的 `2026-10-07 20:00:00` 都变成 `YYYY-MM-DD HH:MM:SS`）。
- `sources`：每个来源的读取方式与列映射；`type` 支持 `json_dir`（目录下所有 `*.json`，
  `records_key` 可指定记录数组所在键，兼容嵌套结构）、`csv`、`excel`；
  同一目标字段可声明多个来源列名（别名），解析时按声明顺序取第一个非空值，
  多个别名同时有值且不一致会记 warning 并计入 `alias_conflicts`。
- 路径、工作表名、表头行、编码、日期格式均可配置，没有写死在代码里。

配置在启动时校验：未知目标字段、重复来源名/优先级、键字段不在 `required_fields`、
来源路径不存在、字段规则缺项等都会立即报错退出，不会带着坏配置跑。

## 合并规则

1. **归一**：按 `field_rules` 把各源原始值转成统一类型与格式；无法归一的值按空值处理，
   同时写 warning 并计入报告的 `invalid_values`。
2. **必填闸门（源记录级）**：缺 `customer_id`、`name` 或 `level` 的源记录整体进
   `quarantine.csv`，不参与合并。客户号为空记 `missing_customer_id`，客户号无法归一记
   `invalid_customer_id`；其余记 `missing_<字段>`。该客户若还有其他来源的完整记录，
   仍可正常成行进入 master。
3. **逐字段取值**：通过闸门的记录按客户号分组，同一目标字段收集所有来源的候选值：
   - 普通字段：来源优先级优先（`api` 1 > `csv` 2 > `excel` 3）；同一优先级内取日期最新，
     同日期取信息更完整的记录，再取文件内靠后的记录；
   - 联系方式字段（`contact_fields`）：**不适用来源优先级**，只要任何来源有非空值就补全，
     取日期最新的那个；日期相同再按来源优先级、完整度、记录顺序兜底。
4. **行级更新时间**：`updated_at` 取该客户所有参与合并记录里最新的日期，而不是某个
   单一来源的日期，便于下游判断这条主数据的“新鲜度”。
5. **跨来源冲突**：对每个客户、每个冲突字段，先让各来源各自归出一个值
   （联系方式取本来源最新），再比较：≥2 个来源都有值且互不相同，记 1 处冲突。
   单一来源内部随日期的历史变化不会被误记成跨来源冲突。

## 产物

### `master.csv`

列 = `target_fields`，按客户号升序：

| 列 | 说明 |
| --- | --- |
| `customer_id` | 归一后的客户号（纯数字、无前导零，如 `46`） |
| `name` / `level` | 姓名 / 等级 |
| `credit_limit` | 信用额度，固定两位小数（如 `128553.49`） |
| `phone` | 联系方式，仅保留数字 |
| `updated_at` | 行级最新更新时间，`YYYY-MM-DD HH:MM:SS` |

### `quarantine.csv`

前 6 列同 master，另有诊断列：

| 列 | 说明 |
| --- | --- |
| `reason_codes` | 原因代码，多个用 `|` 分隔（`missing_customer_id` / `invalid_customer_id` / `missing_name` / `missing_level`） |
| `reason_detail` | 人读说明 |
| `sources` | 来源名 |
| `record_locators` | 源记录定位，如 `csv:src_export.csv#122`、`api:cust_2026-10-01.json#3`（`#` 后为数据行/记录序号） |
| `raw_key` | 归一前的原始客户号（客户号无法归一或为空时便于回查） |

### `run_report.json`

- `counts`：`source_rows_total`、`source_rows_rejected`、`customers_distinct`、
  `master_rows`、`quarantine_rows`、`quarantine_keyless_rows`。
- `sources.<来源>`：`rows_read` / `rows_used` / `rows_rejected` / `rows_without_key`、
  `customers` / `customers_used`、`master_rows_covered`（该来源贡献数据覆盖了多少 master 行，
  即“各来源贡献行数”）、`winning_values`（贡献的字段值个数）、`invalid_values`、
  `alias_conflicts`、`unmapped_columns`。
- `quarantine`：`total`、`reason_distribution`（按原因代码计数，一行多因重复计）、
  `reason_combinations`（原因组合计数）、`by_source`。
- `conflicts`：`total_field_conflicts`、`customers_with_conflicts`、`by_field`、
  `by_source_pair`（如 `api|csv`）以及最多 20 条 `examples`。

本次样例数据的运行结果：

```
合并完成: 源记录 899 行，客户 289 个 -> master 289 行 / quarantine 18 行（其中无客户号 4 行）
来源贡献: api 读入 671 行采用 671 行 / 273 个客户，覆盖 master 273 行
          csv 读入 142 行采用 130 行（隔离 12 行）/ 127 个客户，覆盖 master 45 行
          excel 读入 86 行采用 80 行（隔离 6 行）/ 85 个客户，覆盖 master 22 行
跨来源字段冲突: 105 处（phone 95, level 9, credit_limit 1），涉及 100 个客户
quarantine 原因: missing_name 6, missing_level 8, missing_customer_id 4
```

样例数据里缺姓名/等级的源记录都能被其他来源的完整记录补上，所以 master 覆盖了全部
289 个客户；这些缺字段的**源记录**仍按规则进 quarantine 留痕，不静默丢弃。

## 测试

`tests/test_merge.py` 共 19 个用例，覆盖：

- 归一规则：客户号去前缀/前导零、千分位与两位小数、日期多格式解析；
- 配置驱动：接口 `cust_no` → `customer_code` 改名、表头改名只需改配置；
- 合并规则：来源优先级、同源多天取最新、联系方式按日期最新、缺字段跨来源补全；
- 闸门与分流：缺必填字段/空客户号/非法客户号进 quarantine、原因分布、源内重复行
  （晚行缺字段被隔离但早行完整记录仍可成行）；
- 冲突计数：有冲突/无冲突两种情形；
- 输出与校验：列顺序、金额/日期格式、未知目标字段与缺客户号列的快速失败、空数据日容忍；
- 端到端：仓库样例数据跑通，校验 289 行 master、18 行 quarantine、原因分布与冲突计数。

## 边界与已知限制

- 时间字段不带时区，全程按 naive datetime 处理（源数据均为本地时间），避免跨源时区换算失真。
- 客户号必须能归一为数字，否则整条记录进 quarantine；不做模糊匹配/人工别名映射。
- 必填闸门在**源记录级**执行：缺字段的记录整体不使用（不做部分字段利用），
  其客户可由其他来源的完整记录成行。
- 来源类型固定为 `json_dir` / `csv` / `excel`；新增来源类型需在 `merge.py` 里加一个读取分支。
- 报告中的冲突样例最多保留 20 条，完整冲突只给计数，避免报告随数据规模膨胀。
