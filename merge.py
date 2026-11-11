#!/usr/bin/env python3
"""merge.py — 客户主数据三源合并管道。

三源输入：业务系统接口按日拉取的 JSON（src_api/）、定期导出的 CSV（src_export.csv）、
运营手工维护的 Excel（src_manual.xlsx）。管道把它们归一到统一口径后合并成 master.csv，
必填字段缺失的行分流到 quarantine.csv，并输出 run_report.json。

所有业务约定都来自映射配置文件（默认 mapping.json）：
字段对应关系、同一目标字段的多个来源别名（例如接口升级把 cust_no 改名成
customer_code）、来源优先级、必填字段、联系方式字段、类型规则、输出路径。
本文件只负责通用逻辑，不硬编码任何业务字段名。

用法：
    python3 merge.py --config mapping.json

依赖：pandas、openpyxl。
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import sys
from collections import Counter, defaultdict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import date as date_cls
from datetime import datetime, timezone
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import pandas as pd

LOGGER = logging.getLogger("merge")

SUPPORTED_KINDS = ("text", "identifier", "decimal", "datetime")
SUPPORTED_SOURCE_TYPES = ("json_dir", "csv", "excel")

# 时间缺失时统一按“最早”参与比较，保证排序键可比。
_DATE_EMPTY = (-1, -1, -1, -1, -1, -1)


class MergeError(RuntimeError):
    """配置或输入数据不满足管道约定时抛出，用于快速失败。"""


# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class FieldRule:
    """单个目标字段的归一规则。"""

    kind: str
    strip_chars: tuple[str, ...] = ()
    prefixes: tuple[str, ...] = ()
    digits_only: bool = False
    strip_leading_zeros: bool = False
    decimal_places: int = 2
    output_format: str = "%Y-%m-%d %H:%M:%S"
    formats: tuple[str, ...] = ()


@dataclass(frozen=True)
class SourceSpec:
    """一个数据来源的读取与映射约定。"""

    name: str
    priority: int
    source_type: str
    path: Path
    field_map: Mapping[str, str]
    encoding: str = "utf-8-sig"
    sheet: str | int | None = None
    header_row: int = 1
    records_key: str | None = None
    date_formats: tuple[str, ...] = ()


@dataclass(frozen=True)
class Config:
    """整份映射配置。"""

    path: Path
    base_dir: Path
    target_fields: tuple[str, ...]
    key_field: str
    required_fields: tuple[str, ...]
    contact_fields: tuple[str, ...]
    conflict_fields: tuple[str, ...]
    date_field: str | None
    field_rules: Mapping[str, FieldRule]
    field_labels: Mapping[str, str]
    sources: tuple[SourceSpec, ...]
    output: Mapping[str, str]

    def label(self, field: str) -> str:
        return self.field_labels.get(field, field)


def _as_list(raw: Any, what: str) -> tuple[str, ...]:
    if not isinstance(raw, list) or not all(isinstance(x, str) for x in raw):
        raise MergeError(f"配置项 {what} 必须是字符串列表")
    return tuple(raw)


def _build_rule(field: str, raw: Any) -> FieldRule:
    if not isinstance(raw, Mapping):
        raise MergeError(f"field_rules[{field}] 必须是对象")
    unknown = set(raw) - {
        "kind",
        "strip_chars",
        "prefixes",
        "digits_only",
        "strip_leading_zeros",
        "decimal_places",
        "output_format",
        "formats",
    }
    if unknown:
        raise MergeError(f"field_rules[{field}] 含未知配置键: {sorted(unknown)}")
    kind = raw.get("kind")
    if kind not in SUPPORTED_KINDS:
        raise MergeError(
            f"field_rules[{field}].kind={kind!r} 不受支持，可用: {SUPPORTED_KINDS}"
        )
    return FieldRule(
        kind=kind,
        strip_chars=tuple(raw.get("strip_chars", ())),
        prefixes=tuple(raw.get("prefixes", ())),
        digits_only=bool(raw.get("digits_only", False)),
        strip_leading_zeros=bool(raw.get("strip_leading_zeros", False)),
        decimal_places=int(raw.get("decimal_places", 2)),
        output_format=str(raw.get("output_format", "%Y-%m-%d %H:%M:%S")),
        formats=tuple(raw.get("formats", ())),
    )


def _build_source(
    raw: Any, index: int, base_dir: Path, cfg: Mapping[str, Any]
) -> SourceSpec:
    if not isinstance(raw, Mapping):
        raise MergeError(f"sources[{index}] 必须是对象")
    required = {"name", "priority", "type", "path", "field_map"}
    missing = required - set(raw)
    if missing:
        raise MergeError(f"sources[{index}] 缺少配置项: {sorted(missing)}")
    source_type = raw["type"]
    if source_type not in SUPPORTED_SOURCE_TYPES:
        raise MergeError(
            f"sources[{index}].type={source_type!r} 不受支持，可用: {SUPPORTED_SOURCE_TYPES}"
        )
    field_map = raw["field_map"]
    if not isinstance(field_map, Mapping) or not field_map:
        raise MergeError(f"sources[{index}].field_map 必须是非空对象")
    for column, target in field_map.items():
        if not isinstance(column, str) or not isinstance(target, str):
            raise MergeError(f"sources[{index}].field_map 的键值必须是字符串")
        if target not in cfg["target_fields"]:
            raise MergeError(
                f"sources[{index}].field_map 把列 {column!r} 映射到未知目标字段 {target!r}"
            )
    path = Path(raw["path"])
    if not path.is_absolute():
        path = base_dir / path
    if not path.exists():
        raise MergeError(f"来源 {raw['name']} 的路径不存在: {path}")
    sheet = raw.get("sheet", None)
    return SourceSpec(
        name=str(raw["name"]),
        priority=int(raw["priority"]),
        source_type=str(source_type),
        path=path,
        field_map={str(k): str(v) for k, v in field_map.items()},
        encoding=str(raw.get("encoding", "utf-8-sig")),
        sheet=sheet if isinstance(sheet, (str, int)) else None,
        header_row=int(raw.get("header_row", 1)),
        records_key=raw.get("records_key") or None,
        date_formats=tuple(raw.get("date_formats") or ()),
    )


def load_config(config_path: str | Path, base_dir: str | Path | None = None) -> Config:
    """读取并校验映射配置。"""
    path = Path(config_path)
    if not path.is_absolute():
        path = Path.cwd() / path
    if not path.is_file():
        raise MergeError(f"配置文件不存在: {path}")
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        LOGGER.error("配置文件 %s 不是合法 JSON: %s", path, exc)
        raise MergeError(f"配置文件 {path} 不是合法 JSON: {exc}") from exc
    if not isinstance(raw, Mapping):
        raise MergeError(f"配置文件 {path} 顶层必须是对象")

    root = Path(base_dir) if base_dir is not None else path.parent

    target_fields = _as_list(raw.get("target_fields"), "target_fields")
    key_field = str(raw.get("key_field") or "")
    if key_field not in target_fields:
        raise MergeError(f"key_field={key_field!r} 不在 target_fields 中")
    required_fields = _as_list(raw.get("required_fields", []), "required_fields")
    if key_field not in required_fields:
        raise MergeError(f"required_fields 必须包含 key_field {key_field!r}")
    contact_fields = _as_list(raw.get("contact_fields", []), "contact_fields")
    conflict_fields = _as_list(raw.get("conflict_fields", []), "conflict_fields")
    date_field = raw.get("date_field")
    if date_field is not None and date_field not in target_fields:
        raise MergeError(f"date_field={date_field!r} 不在 target_fields 中")
    for field in (*required_fields, *contact_fields, *conflict_fields):
        if field not in target_fields:
            raise MergeError(f"字段规则引用了未知目标字段 {field!r}")
    if date_field in conflict_fields:
        raise MergeError(
            f"date_field {date_field!r} 不能列入 conflict_fields（每行取最新日期，不参与冲突判定）"
        )

    rules_raw = raw.get("field_rules")
    if not isinstance(rules_raw, Mapping):
        raise MergeError("field_rules 必须是对象")
    field_rules = {
        field: _build_rule(field, rules_raw.get(field)) for field in target_fields
    }

    labels_raw = raw.get("field_labels", {})
    if not isinstance(labels_raw, Mapping):
        raise MergeError("field_labels 必须是对象")
    field_labels = {str(k): str(v) for k, v in labels_raw.items()}

    sources_raw = raw.get("sources")
    if not isinstance(sources_raw, list) or not sources_raw:
        raise MergeError("sources 必须是非空数组")
    sources = tuple(
        _build_source(item, idx, root, raw) for idx, item in enumerate(sources_raw)
    )
    names = [item.name for item in sources]
    if len(set(names)) != len(names):
        raise MergeError(f"sources 名称重复: {names}")
    priorities = [item.priority for item in sources]
    if len(set(priorities)) != len(priorities):
        raise MergeError(f"sources 优先级重复: {priorities}")

    output_raw = raw.get("output", {})
    default_output = {
        "master": "master.csv",
        "quarantine": "quarantine.csv",
        "report": "run_report.json",
    }
    if not isinstance(output_raw, Mapping):
        raise MergeError("output 必须是对象")
    output = {**default_output, **{str(k): str(v) for k, v in output_raw.items()}}

    return Config(
        path=path,
        base_dir=root,
        target_fields=target_fields,
        key_field=key_field,
        required_fields=required_fields,
        contact_fields=contact_fields,
        conflict_fields=conflict_fields,
        date_field=str(date_field) if date_field is not None else None,
        field_rules=field_rules,
        field_labels=field_labels,
        sources=sources,
        output=output,
    )


# ---------------------------------------------------------------------------
# 取值归一
# ---------------------------------------------------------------------------


def is_missing(value: Any) -> bool:
    """统一的“空值”判定：None / NaN / 全空白字符串。"""
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):  # 非标量值不按缺失处理
        return False


def normalize_value(
    raw: Any, rule: FieldRule, date_formats: Sequence[str]
) -> Any | None:
    """按字段规则把原始值归一；无法归一返回 None（由调用方记录告警）。"""
    if rule.kind == "text":
        text = str(raw).strip()
        return text or None

    if rule.kind == "identifier":
        text = str(raw).strip()
        upper = text.upper()
        for prefix in rule.prefixes:
            if upper.startswith(prefix.upper()):
                text = text[len(prefix) :].strip()
                upper = text.upper()
        for token in rule.strip_chars:
            text = text.replace(token, "")
        text = text.strip()
        if rule.digits_only and not text.isdigit():
            return None
        if rule.strip_leading_zeros:
            text = text.lstrip("0") or "0"
        return text or None

    if rule.kind == "decimal":
        text = str(raw).strip()
        for token in rule.strip_chars:
            text = text.replace(token, "")
        text = text.strip()
        if not text:
            return None
        try:
            number = Decimal(text)
        except InvalidOperation:
            # audible: 非数字取值由调用方统一记 warning 并计入 invalid_values
            return None
        quantum = Decimal(1).scaleb(-rule.decimal_places)
        return format(
            number.quantize(quantum, rounding=ROUND_HALF_UP), f".{rule.decimal_places}f"
        )

    if rule.kind == "datetime":
        # 源数据里的时间都是不带时区的本地时间，全程保持 naive，避免跨源时区换算失真。
        if isinstance(raw, datetime):
            return raw
        if isinstance(raw, date_cls):
            return datetime(raw.year, raw.month, raw.day)  # noqa: DTZ001
        text = str(raw).strip()
        if not text:
            return None
        for fmt in (*date_formats, *rule.formats, rule.output_format):
            try:
                return datetime.strptime(text, fmt)  # noqa: DTZ007
            except ValueError:
                # audible: 该格式不匹配，换下一个；全都不匹配时由调用方记 warning
                continue
        return None

    raise MergeError(f"未知的字段归一类型: {rule.kind}")


def display_value(value: Any, rule: FieldRule) -> str:
    """输出用的字符串形式。"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime(rule.output_format)
    return str(value)


def _date_key(value: datetime | None) -> tuple[int, ...]:
    if value is None:
        return _DATE_EMPTY
    return (value.year, value.month, value.day, value.hour, value.minute, value.second)


def _reject_reasons(
    config: Config, key: str | None, raw_key: str, values: Mapping[str, Any]
) -> tuple[str, ...]:
    """源记录必填校验：返回原因代码元组，空元组表示通过。"""
    reasons: list[str] = []
    for field in config.required_fields:
        value = values.get(field)
        blank = value is None or display_value(value, config.field_rules[field]) == ""
        if not blank:
            continue
        if field == config.key_field and raw_key:
            reasons.append(f"invalid_{config.key_field}")
        else:
            reasons.append(f"missing_{field}")
    return tuple(reasons)


def reason_text(config: Config, code: str, raw_key: str = "") -> str:
    """把原因代码转成人读说明。"""
    if code.startswith("invalid_"):
        field = code[len("invalid_") :]
        return f"{config.label(field)}无法归一: {raw_key!r}"
    if code.startswith("missing_"):
        return f"必填字段缺失: {config.label(code[len('missing_') :])}"
    return code


# ---------------------------------------------------------------------------
# 来源读取
# ---------------------------------------------------------------------------


@dataclass
class Candidate:
    """一个目标字段在某条源记录里出现的候选值。"""

    value: Any
    source: str
    priority: int
    record_date: datetime | None
    completeness: int
    order: int
    locator: str


@dataclass
class SourceRecord:
    """归一后的单条源记录。"""

    source: str
    priority: int
    order: int
    locator: str
    raw_key: str
    key: str | None
    values: dict[str, Any]
    record_date: datetime | None
    completeness: int
    reject_reasons: tuple[str, ...]


def _resolve_dotted(payload: Any, dotted: str, path: Path) -> Any:
    current = payload
    for part in dotted.split("."):
        if isinstance(current, Mapping) and part in current:
            current = current[part]
        else:
            raise MergeError(f"{path}: 找不到 records_key={dotted!r} 指向的记录数组")
    return current


def _iter_source_rows(
    spec: SourceSpec, stats: dict[str, Any]
) -> tuple[list[tuple[str, Mapping[str, Any]]], list[str]]:
    """读取一个来源，返回 (行列表, 实际列名列表)。行 = (定位串, 原始行字典)。"""
    rows: list[tuple[str, Mapping[str, Any]]] = []

    if spec.source_type == "json_dir":
        if not spec.path.is_dir():
            raise MergeError(f"来源 {spec.name}: {spec.path} 不是目录")
        files = sorted(p for p in spec.path.glob("*.json") if p.is_file())
        if not files:
            raise MergeError(f"来源 {spec.name}: 目录 {spec.path} 下没有 JSON 文件")
        stats["files"] = len(files)
        for file_path in files:
            try:
                payload = json.loads(file_path.read_text(encoding=spec.encoding))
            except json.JSONDecodeError as exc:
                LOGGER.error("来源 %s: JSON 解析失败 %s: %s", spec.name, file_path, exc)
                raise MergeError(
                    f"来源 {spec.name}: JSON 解析失败 {file_path}: {exc}"
                ) from exc
            if spec.records_key:
                records = _resolve_dotted(payload, spec.records_key, file_path)
            elif isinstance(payload, list):
                records = payload
            else:
                raise MergeError(
                    f"来源 {spec.name}: {file_path.name} 顶层不是数组，"
                    "请在配置中声明 records_key"
                )
            if not isinstance(records, list):
                raise MergeError(f"来源 {spec.name}: {file_path.name} 的记录不是数组")
            for row_no, row in enumerate(records, start=1):
                if not isinstance(row, Mapping):
                    raise MergeError(
                        f"来源 {spec.name}: {file_path.name} 第 {row_no} 条记录不是对象"
                    )
                rows.append((f"{file_path.name}#{row_no}", row))
        columns: list[str] = []
        for _, row in rows:
            for column in row:
                if column not in columns:
                    columns.append(column)
        return rows, columns

    if spec.source_type == "csv":
        try:
            frame = pd.read_csv(
                spec.path,
                dtype=str,
                keep_default_na=False,
                encoding=spec.encoding,
            )
        except (pd.errors.ParserError, UnicodeDecodeError, OSError) as exc:
            LOGGER.error("来源 %s: CSV 读取失败 %s: %s", spec.name, spec.path, exc)
            raise MergeError(
                f"来源 {spec.name}: CSV 读取失败 {spec.path}: {exc}"
            ) from exc
        stats["files"] = 1
        columns = [str(c) for c in frame.columns]
        for offset, row in enumerate(frame.to_dict(orient="records"), start=2):
            rows.append(
                (f"{spec.path.name}#{offset}", {str(k): v for k, v in row.items()})
            )
        return rows, columns

    if spec.source_type == "excel":
        try:
            frame = pd.read_excel(
                spec.path,
                sheet_name=spec.sheet if spec.sheet is not None else 0,
                header=spec.header_row - 1,
                dtype=object,
            )
        except (ValueError, ImportError, OSError) as exc:
            LOGGER.error("来源 %s: Excel 读取失败 %s: %s", spec.name, spec.path, exc)
            raise MergeError(
                f"来源 {spec.name}: Excel 读取失败 {spec.path}: {exc}"
            ) from exc
        frame = frame.dropna(how="all")
        stats["files"] = 1
        columns = [str(c) for c in frame.columns]
        first_data_row = spec.header_row + 1
        for offset, row in enumerate(
            frame.to_dict(orient="records"), start=first_data_row
        ):
            rows.append(
                (f"{spec.path.name}#{offset}", {str(k): v for k, v in row.items()})
            )
        return rows, columns

    raise MergeError(f"来源 {spec.name}: 未实现的来源类型 {spec.source_type!r}")


def _build_records(
    spec: SourceSpec, config: Config, stats: dict[str, Any]
) -> list[SourceRecord]:
    """把来源原始行归一成 SourceRecord，并更新该来源的统计。"""
    rows, columns = _iter_source_rows(spec, stats)
    stats["rows_read"] = len(rows)
    if not rows:
        LOGGER.warning("来源 %s: 没有读到任何记录，本来源跳过", spec.name)
        return []

    target_aliases = {
        field: [col for col, tgt in spec.field_map.items() if tgt == field]
        for field in config.target_fields
    }
    known_columns = set(spec.field_map)
    unmapped = [c for c in columns if c not in known_columns]
    if unmapped:
        stats["unmapped_columns"] = unmapped
        LOGGER.warning("来源 %s: 存在未映射的列，已忽略: %s", spec.name, unmapped)

    key_aliases = target_aliases[config.key_field]
    present_key_aliases = [c for c in key_aliases if c in columns]
    if not present_key_aliases:
        raise MergeError(
            f"来源 {spec.name}: 缺少客户号列（配置声明的别名: {key_aliases}，"
            f"实际列: {columns}）"
        )
    for field, aliases in target_aliases.items():
        if field == config.key_field:
            continue
        if not any(alias in columns for alias in aliases):
            LOGGER.warning(
                "来源 %s: 缺少目标字段 %s 的列（声明别名: %s），该字段在本来源按空值处理",
                spec.name,
                field,
                aliases,
            )

    records: list[SourceRecord] = []
    for order, (locator, raw_row) in enumerate(rows):
        values: dict[str, Any] = {}
        invalid: list[tuple[str, str, Any]] = []
        alias_conflict = 0
        raw_key = ""
        for alias in key_aliases:
            if alias in raw_row and not is_missing(raw_row.get(alias)):
                raw_key = str(raw_row[alias]).strip()
                break

        for field in config.target_fields:
            rule = config.field_rules[field]
            selected: Any = None
            selected_alias: str | None = None
            for alias in target_aliases[field]:
                if alias not in raw_row:
                    continue
                raw_value = raw_row[alias]
                if is_missing(raw_value):
                    continue
                normalized = normalize_value(raw_value, rule, spec.date_formats)
                if normalized is None:
                    invalid.append((field, alias, raw_value))
                    continue
                if selected is None:
                    selected = normalized
                    selected_alias = alias
                elif display_value(selected, rule) != display_value(normalized, rule):
                    alias_conflict += 1
                    LOGGER.warning(
                        "来源 %s %s: 同一目标字段 %s 的多个别名取值不一致"
                        "（%s=%r，%s=%r），采用先声明者",
                        spec.name,
                        locator,
                        field,
                        selected_alias,
                        raw_row.get(selected_alias),
                        alias,
                        raw_value,
                    )
            if selected is not None:
                values[field] = selected

        for field, alias, raw_value in invalid:
            stats["invalid_values"][field] = stats["invalid_values"].get(field, 0) + 1
            LOGGER.warning(
                "来源 %s %s: 字段 %s（列 %s）取值 %r 无法归一，按空值处理",
                spec.name,
                locator,
                field,
                alias,
                raw_value,
            )
        stats["alias_conflicts"] += alias_conflict

        record_date = values.get(config.date_field) if config.date_field else None
        if (
            config.date_field
            and record_date is not None
            and not isinstance(record_date, datetime)
        ):
            record_date = None
        key = values.get(config.key_field)
        key = str(key) if key is not None else None
        completeness = sum(
            1
            for field, value in values.items()
            if field != config.date_field and value is not None
        )
        reject_reasons = _reject_reasons(config, key, raw_key, values)
        records.append(
            SourceRecord(
                source=spec.name,
                priority=spec.priority,
                order=order,
                locator=locator,
                raw_key=raw_key,
                key=key,
                values=values,
                record_date=record_date,
                completeness=completeness,
                reject_reasons=reject_reasons,
            )
        )
    return records


# ---------------------------------------------------------------------------
# 合并
# ---------------------------------------------------------------------------


def choose_candidate(
    field: str, candidates: Sequence[Candidate], config: Config
) -> Candidate | None:
    """按配置规则挑出一个字段的最终取值。

    普通字段：来源优先级优先（数字小者优先），同优先级内取日期最新，
    日期相同再取信息更完整的记录，最后取文件内靠后的记录。
    联系方式：跨来源取日期最新（不适用来源优先级），日期相同再取信息更完整的记录。
    """
    if not candidates:
        return None
    if field in config.contact_fields:
        return max(
            candidates,
            key=lambda c: (
                _date_key(c.record_date),
                c.completeness,
                -c.priority,
                c.order,
            ),
        )
    return min(
        candidates,
        key=lambda c: (
            c.priority,
            tuple(-x for x in _date_key(c.record_date)),
            -c.completeness,
            -c.order,
        ),
    )


def _detect_conflicts(
    field: str,
    records: Sequence[SourceRecord],
    config: Config,
) -> dict[str, Any] | None:
    """判断某字段是否构成跨来源冲突。

    先让每个来源各自按本来源记录归出一个值（联系方式取本来源最新），
    再比较不同来源的值：只有 ≥2 个来源都有值且值互不相同时才算跨来源冲突。
    这样单一来源内部随日期的历史变化不会被误记成跨来源冲突。
    """
    rule = config.field_rules[field]
    by_source: dict[str, list[SourceRecord]] = defaultdict(list)
    for record in records:
        if record.values.get(field) is not None:
            by_source[record.source].append(record)
    if len(by_source) < 2:
        return None

    per_source_value: dict[str, str] = {}
    for source, group in by_source.items():
        winner = choose_candidate(field, _candidates(field, group), config)
        if winner is not None:
            per_source_value[source] = display_value(winner.value, rule)

    distinct = set(per_source_value.values())
    if len(per_source_value) < 2 or len(distinct) < 2:
        return None
    return {
        "field": field,
        "sources": sorted(per_source_value),
        "values": sorted(distinct),
        "pair": "|".join(sorted(per_source_value)),
    }


def _candidates(field: str, group: Sequence[SourceRecord]) -> list[Candidate]:
    return [
        Candidate(
            value=record.values[field],
            source=record.source,
            priority=record.priority,
            record_date=record.record_date,
            completeness=record.completeness,
            order=record.order,
            locator=record.locator,
        )
        for record in group
        if field in record.values
    ]


def _quarantine_row_for_record(config: Config, record: SourceRecord) -> dict[str, Any]:
    """被拒的源记录：保留归一后的可用字段+定位信息，便于人工回查。"""
    row = {field: record.values.get(field) for field in config.target_fields}
    codes = list(record.reject_reasons)
    row.update(
        {
            "reason_codes": "|".join(codes),
            "reason_detail": "；".join(
                reason_text(config, code, record.raw_key) for code in codes
            ),
            "sources": record.source,
            "record_locators": f"{record.source}:{record.locator}",
            "raw_key": record.raw_key,
        }
    )
    return row


def merge_records(
    records: Sequence[SourceRecord], config: Config
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    """源记录必填闸门 → 按客户号字段级合并。

    缺必填字段的源记录整体进 quarantine，不参与合并；其客户仍可由其他来源的
    完整记录成行。通过闸门的记录按客户号分组后逐字段取值。
    """
    grouped: dict[str, list[SourceRecord]] = defaultdict(list)
    rejected = [record for record in records if record.reject_reasons]
    for record in records:
        if record.reject_reasons:
            continue
        grouped[record.key].append(record)

    stats: dict[str, Any] = {
        "conflicts_by_field": Counter(),
        "conflicts_by_pair": Counter(),
        "customers_with_conflicts": 0,
        "winning_values_by_source": Counter(),
        "master_customers_by_source": defaultdict(set),
        "conflict_details": [],
        "rejected_source_rows": len(rejected),
        "customers_distinct": len({record.key for record in records if record.key}),
    }
    master_rows: list[dict[str, Any]] = []
    quarantine_rows: list[dict[str, Any]] = []

    for key, group in grouped.items():
        resolved: dict[str, Any] = {config.key_field: key}
        for field in config.target_fields:
            if field in (config.key_field, config.date_field):
                continue
            winner = choose_candidate(field, _candidates(field, group), config)
            if winner is None:
                continue
            resolved[field] = winner.value
            stats["winning_values_by_source"][winner.source] += 1
            stats["master_customers_by_source"][winner.source].add(key)

        if config.date_field:
            dates = [r.record_date for r in group if r.record_date is not None]
            resolved[config.date_field] = max(dates, key=_date_key) if dates else None

        customer_conflicts: list[dict[str, Any]] = []
        for field in config.conflict_fields:
            detail = _detect_conflicts(field, group, config)
            if detail is None:
                continue
            customer_conflicts.append(detail)
            stats["conflicts_by_field"][field] += 1
            stats["conflicts_by_pair"][detail["pair"]] += 1
        if customer_conflicts:
            stats["customers_with_conflicts"] += 1
            if len(stats["conflict_details"]) < 20:  # 报告里只留少量样例
                stats["conflict_details"].append(
                    {config.key_field: key, "conflicts": customer_conflicts}
                )

        master_rows.append(
            {field: resolved.get(field) for field in config.target_fields}
        )

    quarantine_rows.extend(
        _quarantine_row_for_record(config, record) for record in rejected
    )

    return master_rows, quarantine_rows, stats


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------


def _sort_key_text(value: str) -> tuple[int, Any]:
    return (0, int(value)) if value.isdigit() else (1, value)


def write_master(path: Path, rows: Sequence[Mapping[str, Any]], config: Config) -> None:
    columns = list(config.target_fields)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle, fieldnames=columns, extrasaction="ignore", lineterminator="\n"
        )
        writer.writeheader()
        for row in sorted(
            rows, key=lambda r: _sort_key_text(str(r.get(config.key_field) or ""))
        ):
            writer.writerow(
                {
                    field: display_value(row.get(field), config.field_rules[field])
                    for field in columns
                }
            )


def write_quarantine(
    path: Path, rows: Sequence[Mapping[str, Any]], config: Config
) -> None:
    columns = [
        *config.target_fields,
        "reason_codes",
        "reason_detail",
        "sources",
        "record_locators",
        "raw_key",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle, fieldnames=columns, extrasaction="ignore", lineterminator="\n"
        )
        writer.writeheader()
        for row in sorted(
            rows,
            key=lambda r: (
                str(r.get("reason_codes", "")),
                _sort_key_text(str(r.get(config.key_field) or "")),
            ),
        ):
            payload = {
                field: display_value(row.get(field), config.field_rules[field])
                for field in config.target_fields
            }
            payload.update(
                {
                    "reason_codes": row.get("reason_codes", ""),
                    "reason_detail": row.get("reason_detail", ""),
                    "sources": row.get("sources", ""),
                    "record_locators": row.get("record_locators", ""),
                    "raw_key": row.get("raw_key", ""),
                }
            )
            writer.writerow(payload)


def _source_report(source_stats: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "files": source_stats.get("files", 0),
        "rows_read": source_stats.get("rows_read", 0),
        "rows_used": source_stats.get("rows_used", 0),
        "rows_rejected": source_stats.get("rows_rejected", 0),
        "rows_without_key": source_stats.get("rows_without_key", 0),
        "customers": source_stats.get("customers", 0),
        "customers_used": source_stats.get("customers_used", 0),
        "master_rows_covered": source_stats.get("master_rows_covered", 0),
        "winning_values": source_stats.get("winning_values", 0),
        "invalid_values": source_stats.get("invalid_values", {}),
        "alias_conflicts": source_stats.get("alias_conflicts", 0),
        "unmapped_columns": source_stats.get("unmapped_columns", []),
    }


def build_report(
    config: Config,
    source_stats: Mapping[str, dict[str, Any]],
    master_rows: Sequence[Mapping[str, Any]],
    quarantine_rows: Sequence[Mapping[str, Any]],
    merge_stats: Mapping[str, Any],
) -> dict[str, Any]:
    reason_counter = Counter()
    combination_counter = Counter()
    quarantine_sources = Counter()
    keyless_rows = 0
    for row in quarantine_rows:
        codes = str(row.get("reason_codes", ""))
        code_list = [c for c in codes.split("|") if c]
        combination_counter[codes] += 1
        for code in code_list:
            reason_counter[code] += 1
        for source in str(row.get("sources", "")).split(","):
            if source:
                quarantine_sources[source] += 1
        if (
            f"missing_{config.key_field}" in code_list
            or f"invalid_{config.key_field}" in code_list
        ):
            keyless_rows += 1

    total_field_conflicts = sum(merge_stats["conflicts_by_field"].values())
    return {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "config": str(config.path),
        "counts": {
            "source_rows_total": sum(
                s.get("rows_read", 0) for s in source_stats.values()
            ),
            "source_rows_rejected": merge_stats["rejected_source_rows"],
            "customers_distinct": merge_stats["customers_distinct"],
            "master_rows": len(master_rows),
            "quarantine_rows": len(quarantine_rows),
            "quarantine_keyless_rows": keyless_rows,
        },
        "sources": {
            name: _source_report(stats) for name, stats in source_stats.items()
        },
        "quarantine": {
            "total": len(quarantine_rows),
            "reason_distribution": dict(sorted(reason_counter.items())),
            "reason_combinations": dict(sorted(combination_counter.items())),
            "by_source": dict(sorted(quarantine_sources.items())),
        },
        "conflicts": {
            "total_field_conflicts": total_field_conflicts,
            "customers_with_conflicts": merge_stats["customers_with_conflicts"],
            "by_field": dict(sorted(merge_stats["conflicts_by_field"].items())),
            "by_source_pair": dict(sorted(merge_stats["conflicts_by_pair"].items())),
            "examples": list(merge_stats["conflict_details"]),
        },
    }


def print_summary(report: Mapping[str, Any], config: Config) -> None:
    counts = report["counts"]
    print(
        f"合并完成: 源记录 {counts['source_rows_total']} 行，客户 {counts['customers_distinct']} 个 -> "
        f"master {counts['master_rows']} 行 / quarantine {counts['quarantine_rows']} 行"
        f"（其中无客户号 {counts['quarantine_keyless_rows']} 行）"
    )
    print("各来源贡献:")
    for name, stats in report["sources"].items():
        print(
            f"  - {name}: 读入 {stats['rows_read']} 行，采用 {stats['rows_used']} 行，"
            f"隔离 {stats['rows_rejected']} 行 / {stats['customers']} 个客户，"
            f"覆盖 master {stats['master_rows_covered']} 行，"
            f"贡献字段值 {stats['winning_values']} 个"
        )
    conflict = report["conflicts"]
    print(
        f"跨来源字段冲突: {conflict['total_field_conflicts']} 处，"
        f"涉及 {conflict['customers_with_conflicts']} 个客户"
    )
    for field, count in conflict["by_field"].items():
        print(f"  - {config.label(field)}（{field}）: {count} 处")
    reasons = report["quarantine"]["reason_distribution"]
    if reasons:
        print("quarantine 原因分布:")
        for code, count in reasons.items():
            print(f"  - {code}: {count}")


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------


def run(
    config_path: str | Path = "mapping.json",
    base_dir: str | Path | None = None,
    output_dir: str | Path | None = None,
    write: bool = True,
) -> tuple[dict[str, Any], Config]:
    """执行完整管道，返回 (运行报告, 配置)。"""
    config = load_config(config_path, base_dir=base_dir)
    source_stats: dict[str, dict[str, Any]] = {}
    all_records: list[SourceRecord] = []

    for spec in config.sources:
        stats: dict[str, Any] = {
            "invalid_values": {},
            "alias_conflicts": 0,
            "unmapped_columns": [],
            "rows_read": 0,
            "rows_without_key": 0,
            "customers": 0,
            "master_rows_covered": 0,
            "winning_values": 0,
        }
        records = _build_records(spec, config, stats)
        source_stats[spec.name] = stats
        all_records.extend(records)
        LOGGER.info(
            "来源 %s: 读入 %d 行，%d 个客户",
            spec.name,
            stats["rows_read"],
            len({record.key for record in records if record.key}),
        )

    master_rows, quarantine_rows, merge_stats = merge_records(all_records, config)

    for name, stats in source_stats.items():
        source_records = [record for record in all_records if record.source == name]
        stats["customers"] = len(
            {record.key for record in source_records if record.key}
        )
        stats["customers_used"] = len(
            {
                record.key
                for record in source_records
                if record.key and not record.reject_reasons
            }
        )
        stats["rows_used"] = sum(
            1 for record in source_records if not record.reject_reasons
        )
        stats["rows_rejected"] = sum(
            1 for record in source_records if record.reject_reasons
        )
        stats["rows_without_key"] = sum(
            1 for record in source_records if record.key is None
        )
        stats["master_rows_covered"] = len(
            merge_stats["master_customers_by_source"].get(name, set())
        )
        stats["winning_values"] = merge_stats["winning_values_by_source"].get(name, 0)
        LOGGER.info(
            "来源 %s: 采用 %d 行，隔离 %d 行",
            name,
            stats["rows_used"],
            stats["rows_rejected"],
        )

    report = build_report(
        config, source_stats, master_rows, quarantine_rows, merge_stats
    )

    if write:
        out_root = Path(output_dir) if output_dir is not None else config.base_dir
        master_path = out_root / config.output["master"]
        quarantine_path = out_root / config.output["quarantine"]
        report_path = out_root / config.output["report"]
        write_master(master_path, master_rows, config)
        write_quarantine(quarantine_path, quarantine_rows, config)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        LOGGER.info(
            "已写出 %s / %s / %s",
            master_path,
            quarantine_path,
            report_path,
        )

    return report, config


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="客户主数据三源合并管道")
    parser.add_argument("--config", default="mapping.json", help="映射配置文件路径")
    parser.add_argument(
        "--data-dir", default=None, help="数据根目录（默认取配置文件所在目录）"
    )
    parser.add_argument(
        "--output-dir", default=None, help="输出目录（默认取数据根目录）"
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="输出调试日志")
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
        stream=sys.stderr,
    )
    try:
        report, config = run(
            config_path=args.config,
            base_dir=args.data_dir,
            output_dir=args.output_dir,
        )
    except MergeError as exc:
        LOGGER.error("合并失败: %s", exc)
        return 2
    print_summary(report, config)
    return 0


if __name__ == "__main__":
    sys.exit(main())
