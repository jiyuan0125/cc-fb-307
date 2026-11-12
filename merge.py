"""客户主数据三源合并管道。

数据源：
  - src_api/        业务系统接口按日拉取的 JSON（字段升级前后两种命名并存）
  - src_export.csv  每周导出的 CSV（数字带千分位、客户号可能带前导零）
  - src_manual.xlsx 运营手工维护的 Excel（会员号为带前导零的文本）

字段映射、类型归一规则、合并优先级全部声明在 mapping.yaml 中，本文件不含
任何来源字段名的硬编码。

产物（默认输出到仓库根目录）：
  - master.csv      合并后的客户主数据
  - quarantine.csv  必填字段缺失的隔离行（含隔离原因）
  - report.json     运行报告（计数、原因分布、来源贡献、冲突统计）
"""

from __future__ import annotations

import argparse
import glob
import json
import logging
import math
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd
import yaml

logger = logging.getLogger("merge")


# --------------------------------------------------------------------------- #
# 配置加载与校验
# --------------------------------------------------------------------------- #


def load_config(path: str | Path) -> dict[str, Any]:
    """读取映射配置并做结构校验，配置缺项直接 fail fast。"""
    with open(path, encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh)

    schema = {col["name"]: col for col in cfg["target_schema"]}
    key_field = cfg.get("key_field", "customer_id")
    if key_field not in schema:
        raise ValueError(f"key_field {key_field!r} 未在 target_schema 中声明")

    required_fields = {
        col["name"] for col in cfg["target_schema"] if col.get("required")
    }
    if key_field not in required_fields:
        raise ValueError(f"key_field {key_field!r} 必须同时声明为 required")

    priority = cfg["merge"]["priority"]
    source_names = set(cfg["sources"])
    if set(priority) != source_names:
        raise ValueError(
            f"merge.priority {priority} 与 sources {sorted(source_names)} 不一致"
        )

    contact_fields = set(cfg["merge"]["contact_fields"])
    unknown_contacts = contact_fields - set(schema)
    if unknown_contacts:
        raise ValueError(f"contact_fields 含未知目标字段: {sorted(unknown_contacts)}")

    for src_name, src_cfg in cfg["sources"].items():
        mapped_targets = set(src_cfg["field_map"].values())
        missing = required_fields - mapped_targets
        if missing:
            raise ValueError(f"来源 {src_name} 的 field_map 缺少目标字段: {missing}")
        if src_cfg["date_field"] not in mapped_targets:
            raise ValueError(f"来源 {src_name} 的 date_field 未在 field_map 中声明")

    return cfg


# --------------------------------------------------------------------------- #
# 类型归一
# --------------------------------------------------------------------------- #


class NormalizeError(ValueError):
    """单值归一失败，调用方记录日志后按缺失处理或隔离。"""


def is_blank(value: Any) -> bool:
    """None、空串、纯空白、NaN 统一视为缺失。"""
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return isinstance(value, str) and not value.strip()


def normalize_value(value: Any, dtype: str, rules: dict[str, Any]) -> Any:
    """按 mapping.yaml 的 type 与 normalization 规则归一单个值。

    空白值统一返回 None；无法解析的值抛 NormalizeError，由调用方决定
    隔离还是记日志跳过。
    """
    if is_blank(value):
        return None

    if dtype == "str":
        return str(value).strip()

    if dtype == "int":
        text = str(value).strip()
        prefix = rules.get("int", {}).get("strip_prefix")
        if prefix:
            text = text.removeprefix(prefix)
        if not re.fullmatch(r"\d+", text):
            raise NormalizeError(f"无法归一为整数: {value!r}")
        return int(text)

    if dtype == "float":
        frules = rules.get("float", {})
        text = str(value).strip().replace(frules.get("thousands_sep", ""), "")
        try:
            number = float(text)
        except ValueError as exc:
            raise NormalizeError(f"无法归一为浮点数: {value!r}") from exc
        return round(number, frules.get("decimals", 2))

    if dtype == "date":
        text = str(value).strip()
        # 兼容 "YYYY-MM-DD" 与 "YYYY-MM-DD HH:MM:SS"，取日期部分
        match = re.fullmatch(r"(\d{4}-\d{2}-\d{2})(?:[ T]\d{2}:\d{2}:\d{2})?", text)
        if not match:
            raise NormalizeError(f"无法归一为日期: {value!r}")
        out_fmt = rules.get("date", {}).get("output_format", "%Y-%m-%d")
        return pd.Timestamp(match.group(1)).strftime(out_fmt)

    raise NormalizeError(f"配置中声明了未知类型 {dtype!r}")


# --------------------------------------------------------------------------- #
# 数据结构
# --------------------------------------------------------------------------- #


@dataclass
class Record:
    """一条已映射、已归一的来源记录。"""

    source: str
    fields: dict[str, Any]
    raw: dict[str, Any]
    source_file: str
    key_field: str
    date_field: str

    @property
    def key(self) -> Any:
        return self.fields.get(self.key_field)

    @property
    def updated_at(self) -> Any:
        return self.fields.get(self.date_field)


@dataclass
class SourceLoadResult:
    name: str
    records: list[Record]
    raw_row_count: int
    dropped_rows: list[dict[str, Any]] = field(default_factory=list)


# --------------------------------------------------------------------------- #
# 来源读取与映射
# --------------------------------------------------------------------------- #


def map_raw_row(
    raw: dict[str, Any],
    source: str,
    source_cfg: dict[str, Any],
    schema: dict[str, dict[str, Any]],
    norm_rules: dict[str, Any],
    source_file: str,
    key_field: str,
) -> Record | None:
    """把一行来源数据按 field_map 映射并归一为目标字段。

    主键归一失败的行无法参与合并，返回 None 并由调用方计入 dropped_rows。
    其余字段归一失败时记日志并按缺失处理（必填缺失后续在隔离环节拦截）。
    """
    mapped: dict[str, Any] = {}
    for src_field, target in source_cfg["field_map"].items():
        if target in mapped and mapped[target] is not None:
            # 同一目标字段的候选字段（如 cust_no/customer_code）取第一个非空
            continue
        raw_value = raw.get(src_field)
        dtype = schema[target]["type"]
        try:
            mapped[target] = normalize_value(raw_value, dtype, norm_rules)
        except NormalizeError as exc:
            logger.warning(
                "来源 %s 文件 %s 字段 %s 归一失败，按缺失处理: %s",
                source,
                source_file,
                src_field,
                exc,
            )
            mapped.setdefault(target, None)

    if mapped.get(key_field) is None:
        logger.warning(
            "来源 %s 文件 %s 存在无法识别主键 %s 的行，已剔除: %s",
            source,
            source_file,
            key_field,
            raw,
        )
        return None

    return Record(
        source=source,
        fields=mapped,
        raw=raw,
        source_file=source_file,
        key_field=key_field,
        date_field=source_cfg["date_field"],
    )


def load_api(
    src_cfg: dict[str, Any],
    schema: dict[str, dict[str, Any]],
    norm_rules: dict[str, Any],
    base_dir: Path,
    key_field: str,
) -> SourceLoadResult:
    """读取 JSON 目录，每个文件是一组记录。"""
    records: list[Record] = []
    dropped: list[dict[str, Any]] = []
    raw_count = 0
    for path in sorted(glob.glob(str(base_dir / src_cfg["path_glob"]))):
        source_file = Path(path).name
        try:
            with open(path, encoding="utf-8") as fh:
                rows = json.load(fh)
        except (json.JSONDecodeError, OSError) as exc:
            logger.error("跳过无法读取的 JSON 文件 %s: %s", path, exc)
            continue
        if not isinstance(rows, list):
            logger.error("JSON 文件 %s 顶层不是数组，跳过", path)
            continue
        raw_count += len(rows)
        for raw in rows:
            record = map_raw_row(
                raw, "api", src_cfg, schema, norm_rules, source_file, key_field
            )
            if record is None:
                dropped.append(
                    {"source": "api", "source_file": source_file, "row": raw}
                )
            else:
                records.append(record)
    return SourceLoadResult("api", records, raw_count, dropped)


def load_csv(
    src_cfg: dict[str, Any],
    schema: dict[str, dict[str, Any]],
    norm_rules: dict[str, Any],
    base_dir: Path,
    key_field: str,
) -> SourceLoadResult:
    """读取导出 CSV。全部按字符串读入，归一交给 normalize_value。"""
    path = base_dir / src_cfg["path"]
    records: list[Record] = []
    dropped: list[dict[str, Any]] = []
    try:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8")
    except OSError as exc:
        logger.error("无法读取 CSV %s: %s", path, exc)
        return SourceLoadResult("csv", [], 0, dropped)

    for idx, row in df.iterrows():
        raw = row.to_dict()
        record = map_raw_row(
            raw, "csv", src_cfg, schema, norm_rules, path.name, key_field
        )
        if record is None:
            dropped.append(
                {
                    "source": "csv",
                    "source_file": path.name,
                    "row": raw,
                    "row_no": idx + 2,
                }
            )
        else:
            records.append(record)
    return SourceLoadResult("csv", records, len(df), dropped)


def load_excel(
    src_cfg: dict[str, Any],
    schema: dict[str, dict[str, Any]],
    norm_rules: dict[str, Any],
    base_dir: Path,
    key_field: str,
) -> SourceLoadResult:
    """读取手工 Excel，表头为中文，会员号按文本读取以保留前导零。"""
    path = base_dir / src_cfg["path"]
    records: list[Record] = []
    dropped: list[dict[str, Any]] = []
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except OSError as exc:
        logger.error("无法读取 Excel %s: %s", path, exc)
        return SourceLoadResult("excel", [], 0, dropped)

    ws = wb[src_cfg.get("sheet")] if src_cfg.get("sheet") else wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return SourceLoadResult("excel", [], 0, dropped)

    headers = [str(h).strip() for h in rows[0]]
    raw_count = len(rows) - 1
    for row in rows[1:]:
        raw = dict(zip(headers, row))
        record = map_raw_row(
            raw, "excel", src_cfg, schema, norm_rules, path.name, key_field
        )
        if record is None:
            dropped.append({"source": "excel", "source_file": path.name, "row": raw})
        else:
            records.append(record)
    return SourceLoadResult("excel", records, raw_count, dropped)


LOADERS = {"json_dir": load_api, "csv": load_csv, "excel": load_excel}


def load_all_sources(
    cfg: dict[str, Any], base_dir: Path
) -> dict[str, SourceLoadResult]:
    schema = {col["name"]: col for col in cfg["target_schema"]}
    norm_rules = cfg["normalization"]
    key_field = cfg["key_field"]
    results: dict[str, SourceLoadResult] = {}
    for name in cfg["merge"]["priority"]:
        src_cfg = cfg["sources"][name]
        loader = LOADERS[src_cfg["type"]]
        results[name] = loader(src_cfg, schema, norm_rules, base_dir, key_field)
        logger.info(
            "来源 %s 读取完成: 原始 %s 行，有效 %s 行，剔除 %s 行",
            name,
            results[name].raw_row_count,
            len(results[name].records),
            len(results[name].dropped_rows),
        )
    return results


# --------------------------------------------------------------------------- #
# 来源内去重：同一主键保留更新时间最新的一条
# --------------------------------------------------------------------------- #


def dedupe_within_source(
    result: SourceLoadResult,
) -> tuple[dict[Any, Record], int]:
    latest: dict[Any, Record] = {}
    dup_count = 0
    for record in result.records:
        key = record.key
        old = latest.get(key)
        if old is None:
            latest[key] = record
            continue
        dup_count += 1
        old_ts = old.updated_at or ""
        new_ts = record.updated_at or ""
        if new_ts > old_ts:
            latest[key] = record
    logger.info(
        "来源 %s 去重: 合并重复出现 %s 次，保留客户 %s 个",
        result.name,
        dup_count,
        len(latest),
    )
    return latest, dup_count


# --------------------------------------------------------------------------- #
# 跨来源合并
# --------------------------------------------------------------------------- #


def _missing_required(
    fields: dict[str, Any], schema: dict[str, dict[str, Any]]
) -> list[str]:
    return [
        col["name"]
        for col in schema.values()
        if col.get("required") and is_blank(fields.get(col["name"]))
    ]


def merge_sources(
    cfg: dict[str, Any],
    source_data: dict[str, SourceLoadResult],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    """按优先级合并三源，返回 master 行、quarantine 行和统计信息。"""
    schema = {col["name"]: col for col in cfg["target_schema"]}
    target_names = list(schema)
    key_field: str = cfg["key_field"]
    priority: list[str] = cfg["merge"]["priority"]
    contact_fields: list[str] = cfg["merge"]["contact_fields"]
    date_fields = {name: cfg["sources"][name]["date_field"] for name in priority}
    # 记录时间戳各源口径不同，跨源不一致是常态，不计入字段冲突
    conflict_fields = [
        f
        for f in target_names
        if f not in contact_fields and f != key_field and f not in date_fields.values()
    ]
    ordinary_fields = [f for f in target_names if f not in contact_fields]

    # 来源内去重
    latest_by_source: dict[str, dict[Any, Record]] = {}
    dup_stats: dict[str, int] = {}
    for name in priority:
        latest_by_source[name], dup_stats[name] = dedupe_within_source(
            source_data[name]
        )

    all_keys = set()
    for latest in latest_by_source.values():
        all_keys.update(latest.keys())

    master_rows: list[dict[str, Any]] = []
    quarantine_rows: list[dict[str, Any]] = []
    contributing: Counter[str] = Counter()
    present_counter: Counter[str] = Counter()
    conflict_counter: Counter[str] = Counter()
    source_combo_counter: Counter[str] = Counter()

    for key in sorted(all_keys):
        present = [name for name in priority if key in latest_by_source[name]]
        present_counter.update(present)
        source_combo_counter["+".join(present)] += 1

        merged: dict[str, Any] = {f: None for f in target_names}
        chosen_sources: dict[str, str] = {}

        # 普通字段：按配置的优先级取第一个非空值，同时统计跨源冲突
        for fname in ordinary_fields:
            non_null = [
                (name, latest_by_source[name][key])
                for name in present
                if not is_blank(latest_by_source[name][key].fields.get(fname))
            ]
            distinct = {rec.fields.get(fname) for _, rec in non_null}
            if fname in conflict_fields and len(distinct) > 1:
                conflict_counter[fname] += 1
            if non_null:
                chosen_name, chosen_rec = non_null[0]
                merged[fname] = chosen_rec.fields.get(fname)
                chosen_sources[fname] = chosen_name

        # 联系方式字段：任一来源有非空值即补全；多个非空值取该来源更新时间最新的
        for fname in contact_fields:
            candidates = [
                (name, latest_by_source[name][key])
                for name in present
                if not is_blank(latest_by_source[name][key].fields.get(fname))
            ]
            distinct = {rec.fields.get(fname) for _, rec in candidates}
            if len(distinct) > 1:
                conflict_counter[fname] += 1
            if candidates:
                winner_name, winner = max(
                    candidates,
                    key=lambda item: item[1].fields.get(date_fields[item[0]]) or "",
                )
                merged[fname] = winner.fields.get(fname)
                chosen_sources[fname] = winner_name

        # 贡献口径：该客户每个非空取值实际来自哪些来源
        contributing.update(set(chosen_sources.values()))

        missing = _missing_required(merged, schema)
        if missing:
            quarantine_rows.append(
                {
                    key_field: key,
                    "missing_fields": ";".join(missing),
                    "present_sources": "+".join(present),
                    **{f: merged.get(f) for f in target_names if f != key_field},
                }
            )
            continue

        master_rows.append({f: merged.get(f) for f in target_names})

    stats = {
        "duplicates_within_source": dup_stats,
        "customers_present_per_source": dict(present_counter),
        "source_combo_distribution": dict(source_combo_counter),
        "conflict_field_counts": dict(conflict_counter),
        "conflict_total": sum(conflict_counter.values()),
        "contributing_rows": dict(contributing),
    }
    return master_rows, quarantine_rows, stats


# --------------------------------------------------------------------------- #
# 输出与报告
# --------------------------------------------------------------------------- #


def write_outputs(
    cfg: dict[str, Any],
    master_rows: list[dict[str, Any]],
    quarantine_rows: list[dict[str, Any]],
    stats: dict[str, Any],
    source_data: dict[str, SourceLoadResult],
    out_dir: Path,
) -> dict[str, Any]:
    target_names = [col["name"] for col in cfg["target_schema"]]
    key_field = cfg["key_field"]
    str_columns = [col["name"] for col in cfg["target_schema"] if col["type"] == "str"]

    master_df = pd.DataFrame(
        [{f: row.get(f) for f in target_names} for row in master_rows],
        columns=target_names,
    )
    # 联系方式等字符串列必须保持文本，防止 pandas 写 CSV 时把数字串转成科学计数法
    for col in str_columns:
        master_df[col] = master_df[col].astype("string")
    master_df.to_csv(out_dir / "master.csv", index=False, encoding="utf-8-sig")

    q_columns = [key_field, "missing_fields", "present_sources"] + [
        f for f in target_names if f != key_field
    ]
    quarantine_df = pd.DataFrame(quarantine_rows, columns=q_columns)
    for col in str_columns:
        if col in quarantine_df.columns:
            quarantine_df[col] = quarantine_df[col].astype("string")
    quarantine_df.to_csv(out_dir / "quarantine.csv", index=False, encoding="utf-8-sig")

    reason_distribution = Counter(row["missing_fields"] for row in quarantine_rows)
    report = {
        "raw_row_count": {
            name: result.raw_row_count for name, result in source_data.items()
        },
        "dropped_unidentified_rows": {
            name: len(result.dropped_rows) for name, result in source_data.items()
        },
        "duplicates_within_source": stats["duplicates_within_source"],
        "master_customer_count": len(master_rows),
        "quarantine_count": len(quarantine_rows),
        "quarantine_reason_distribution": dict(reason_distribution),
        "customers_present_per_source": stats["customers_present_per_source"],
        "source_combo_distribution": stats["source_combo_distribution"],
        "source_contributing_rows": stats["contributing_rows"],
        "cross_source_field_conflicts": stats["conflict_field_counts"],
        "cross_source_conflict_total": stats["conflict_total"],
    }
    with open(out_dir / "report.json", "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)
    return report


def print_report(report: dict[str, Any]) -> None:
    logger.info("-------------------- 运行报告 --------------------")
    logger.info("原始行数: %s", report["raw_row_count"])
    logger.info("无法识别客户号剔除行: %s", report["dropped_unidentified_rows"])
    logger.info("来源内重复合并次数: %s", report["duplicates_within_source"])
    logger.info("master 客户数: %s", report["master_customer_count"])
    logger.info("隔离行数: %s", report["quarantine_count"])
    logger.info("隔离原因分布: %s", report["quarantine_reason_distribution"])
    logger.info("各来源覆盖客户数: %s", report["customers_present_per_source"])
    logger.info("来源组合分布: %s", report["source_combo_distribution"])
    logger.info("各来源贡献行数: %s", report["source_contributing_rows"])
    logger.info("跨来源字段冲突计数: %s", report["cross_source_field_conflicts"])
    logger.info("冲突合计: %s", report["cross_source_conflict_total"])


def run(base_dir: Path, config_path: Path, out_dir: Path) -> dict[str, Any]:
    cfg = load_config(config_path)
    source_data = load_all_sources(cfg, base_dir)
    master_rows, quarantine_rows, stats = merge_sources(cfg, source_data)
    out_dir.mkdir(parents=True, exist_ok=True)
    report = write_outputs(
        cfg, master_rows, quarantine_rows, stats, source_data, out_dir
    )
    print_report(report)
    return report


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="客户主数据三源合并管道")
    parser.add_argument(
        "--base-dir",
        default=".",
        help="数据根目录（含 src_api/、src_export.csv、src_manual.xlsx），默认当前目录",
    )
    parser.add_argument("--config", default="mapping.yaml", help="映射配置路径")
    parser.add_argument("--out-dir", default=".", help="产物输出目录，默认当前目录")
    parser.add_argument(
        "--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"]
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    run(Path(args.base_dir), Path(args.config), Path(args.out_dir))


if __name__ == "__main__":
    main()
