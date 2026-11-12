# 客户主数据三源合并管道

把业务系统接口 JSON、每周导出 CSV、运营手工 Excel 三个来源的客户数据，
按声明式映射配置归一后合并为一份主数据。

## 目录结构

```
.
├── merge.py          # 合并管道（读取、映射、归一、去重、合并、隔离、报告）
├── mapping.yaml      # 映射配置：字段对应、类型归一、主键、合并优先级
├── test_merge.py     # pytest 测试（合成数据，不依赖样例数据）
├── src_api/          # 来源 1：接口按日拉取的 JSON，每天一个文件
├── src_export.csv    # 来源 2：每周导出的 CSV
├── src_manual.xlsx   # 来源 3：运营手工维护的 Excel
├── master.csv        # 产物：合并后的客户主数据
├── quarantine.csv    # 产物：必填字段缺失的隔离行
└── report.json       # 产物：运行报告
```

## 环境与运行

依赖：Python 3.12、pandas、openpyxl、PyYAML、pytest（无需 Docker）。

```bash
# 在仓库根目录运行，数据使用相对路径
python3 merge.py

# 可选参数
python3 merge.py --base-dir . --config mapping.yaml --out-dir . --log-level INFO
```

运行测试：

```bash
python3 -m pytest test_merge.py -v
```

## 映射配置 mapping.yaml

所有字段映射与归一规则都在配置中声明，`merge.py` 不含来源字段名的硬编码。
接口升级导致的字段改名（`cust_no` → `customer_code`）通过把两个来源字段
映射到同一目标字段解决，代码无需改动：

```yaml
sources:
  api:
    field_map:
      cust_no: customer_id        # 升级前旧字段
      customer_code: customer_id  # 升级后新字段，归一到同一目标
```

配置分四部分：

| 配置块 | 作用 |
|---|---|
| `key_field` | 客户主键的目标字段名；主键无法识别的行直接剔除并计入报告 |
| `target_schema` | 目标字段、类型（int/str/float/date）、是否必填、master 列顺序 |
| `normalization` | 各类型的归一规则（C 前缀、千分位、小数位、日期输出格式） |
| `sources.<源>.field_map` | 来源字段 → 目标字段的对应关系 |
| `merge.priority` | 普通字段的取数优先级，本项目为 `api > csv > excel` |
| `merge.contact_fields` | 联系方式字段，走补全规则而非优先级覆盖 |

## 类型与格式归一

- **客户号（int）**：API 的 `C000046`、CSV 的 `0207`、Excel 文本会员号 `046`
  统一去掉 `C` 前缀和前导零后转整数，三源同一客户才能对上。
- **信用额度（float）**：CSV 的千分位逗号（`"111,373.71"`）去除后转浮点，
  统一保留两位小数。
- **联系方式（str）**：按文本处理并保持文本写出，避免被表格软件读成科学计数法；
  空白字符串与 `None` 同等视为缺失。
- **日期（date）**：`"2026-10-07 20:00:00"` 与 `"2026-10-07"` 统一输出
  `YYYY-MM-DD`。

## 合并规则

1. **来源内去重**：同一客户在同一来源出现多条（API 跨天快照、CSV/Excel 重复行）时，
   保留该来源更新时间最新的一条，重复次数计入报告。
2. **普通字段取数优先级**：`name`、`level`、`credit_limit`、`updated_at` 等字段
   按 API → CSV → Excel 取第一个非空值，与记录时间先后无关。
3. **联系方式例外**：`phone` 不做优先级覆盖——任一来源有非空值即补全；
   多个来源都有非空且值不同时，取**来源更新时间最新**的那个。
4. **跨源冲突计数**：同一客户的同一业务字段在多个来源存在互不相同的非空值时，
   按字段累计冲突数。`updated_at` 是记录时间戳而非业务字段，各源口径不同属常态，
   不计入冲突；联系方式的冲突照常统计。
5. **必填校验与隔离**：`customer_id`、`name`、`level` 为必填（由
   `target_schema` 的 `required` 声明）。合并后仍缺必填字段的客户不进
   `master.csv`，写入 `quarantine.csv` 并记录缺失原因；主键本身缺失的行
   无法归属任何客户，在读取阶段直接剔除并在报告中单列计数。

## 产物说明

### master.csv

UTF-8 BOM 编码（Excel 直接打开不乱码），列顺序同 `target_schema`：

| 列 | 说明 |
|---|---|
| customer_id | 归一后的整数客户号，唯一 |
| name / level | 按优先级取数的姓名、等级 |
| credit_limit | 浮点，两位小数 |
| phone | 文本联系方式，可能为空 |
| updated_at | `YYYY-MM-DD` |

### quarantine.csv

隔离行，含 `missing_fields`（缺失的必填字段，分号分隔）、`present_sources`
（该客户出现在哪些来源）以及各目标字段的当前取值，便于人工排查补录。

### report.json

| 指标 | 含义 |
|---|---|
| `raw_row_count` | 各来源读取到的原始行数 |
| `dropped_unidentified_rows` | 主键无法识别被剔除的行数 |
| `duplicates_within_source` | 各来源来源内去重合并的重复次数 |
| `master_customer_count` / `quarantine_count` | 主数据行数 / 隔离行数 |
| `quarantine_reason_distribution` | 隔离原因（缺失字段组合）分布 |
| `customers_present_per_source` | 各来源覆盖的客户数（同一客户可在多源重复计数） |
| `source_combo_distribution` | 客户按来源组合的分布（如 `api+csv`） |
| `source_contributing_rows` | 各来源实际贡献了取值的客户数 |
| `cross_source_field_conflicts` | 按字段统计的跨源冲突数及合计 |

## 样例数据运行结果（约 300 客户）

```
原始行数: api=671, csv=142, excel=86
无法识别主键剔除行: csv=4（空 cust_id）
来源内重复合并: api=398, csv=11, excel=1
master 客户数: 288，隔离 1 行（客户 77 缺 level，仅存在于 CSV）
来源贡献: api=273, csv=46, excel=26
跨源字段冲突: phone=74, level=8, credit_limit=6
```

勾稽关系：`master_customer_count + quarantine_count` = 三源客户号并集 289。

## 异常处理约定

- 数据文件缺失、JSON 解析失败、单值归一失败均通过 logging 记录
  （`ERROR`/`WARNING`），不静默吞掉；归一失败的非主键字段按缺失处理，
  后续由必填校验在隔离环节兜底。
- 配置结构错误（优先级与来源不一致、field_map 缺必填目标字段等）在启动时
  fail fast 抛出 `ValueError`。
