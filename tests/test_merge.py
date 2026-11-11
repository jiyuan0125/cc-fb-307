"""merge.py 的单元测试与端到端测试。

覆盖点：类型/格式归一、配置驱动的字段改名、来源优先级、联系方式按最新日期、
必填字段闸门与 quarantine 原因分布、跨来源冲突计数、输出格式、样例数据端到端。
"""

from __future__ import annotations

import copy
import csv
import json
import re
from pathlib import Path

import pandas as pd
import pytest

import merge

REPO_ROOT = Path(__file__).resolve().parents[1]

BASE_CONFIG = {
    "target_fields": [
        "customer_id",
        "name",
        "level",
        "credit_limit",
        "phone",
        "updated_at",
    ],
    "key_field": "customer_id",
    "required_fields": ["customer_id", "name", "level"],
    "contact_fields": ["phone"],
    "conflict_fields": ["name", "level", "credit_limit", "phone"],
    "date_field": "updated_at",
    "field_labels": {
        "customer_id": "客户号",
        "name": "姓名",
        "level": "等级",
        "credit_limit": "信用额度",
        "phone": "联系方式",
        "updated_at": "更新日期",
    },
    "field_rules": {
        "customer_id": {
            "kind": "identifier",
            "prefixes": ["C"],
            "digits_only": True,
            "strip_leading_zeros": True,
        },
        "name": {"kind": "text"},
        "level": {"kind": "text"},
        "credit_limit": {"kind": "decimal", "strip_chars": [","], "decimal_places": 2},
        "phone": {
            "kind": "identifier",
            "strip_chars": [" ", "-"],
            "digits_only": True,
            "strip_leading_zeros": False,
        },
        "updated_at": {"kind": "datetime", "output_format": "%Y-%m-%d %H:%M:%S"},
    },
    "sources": [
        {
            "name": "api",
            "priority": 1,
            "type": "json_dir",
            "path": "src_api",
            "encoding": "utf-8",
            "date_formats": ["%Y-%m-%d %H:%M:%S"],
            "field_map": {
                "customer_code": "customer_id",
                "cust_no": "customer_id",
                "name": "name",
                "level": "level",
                "credit_limit": "credit_limit",
                "phone": "phone",
                "updated_at": "updated_at",
            },
        },
        {
            "name": "csv",
            "priority": 2,
            "type": "csv",
            "path": "src_export.csv",
            "encoding": "utf-8",
            "date_formats": ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d"],
            "field_map": {
                "cust_id": "customer_id",
                "name": "name",
                "level": "level",
                "credit_limit": "credit_limit",
                "phone": "phone",
                "last_updated": "updated_at",
            },
        },
        {
            "name": "excel",
            "priority": 3,
            "type": "excel",
            "path": "src_manual.xlsx",
            "sheet": "客户",
            "header_row": 1,
            "date_formats": ["%Y-%m-%d"],
            "field_map": {
                "会员号": "customer_id",
                "姓名": "name",
                "等级": "level",
                "信用额度": "credit_limit",
                "联系方式": "phone",
                "更新日期": "updated_at",
            },
        },
    ],
    "output": {
        "master": "out/master.csv",
        "quarantine": "out/quarantine.csv",
        "report": "out/run_report.json",
    },
}

CSV_COLUMNS = ["cust_id", "name", "level", "credit_limit", "phone", "last_updated"]
EXCEL_COLUMNS = ["会员号", "姓名", "等级", "信用额度", "联系方式", "更新日期"]


# ---------------------------------------------------------------------------
# 测试脚手架
# ---------------------------------------------------------------------------


def make_config(tmp_path: Path, enabled: list[str] | None = None) -> Path:
    """写一份配置；enabled 用来只保留本测试实际提供的来源。"""
    config = copy.deepcopy(BASE_CONFIG)
    if enabled is not None:
        config["sources"] = [s for s in config["sources"] if s["name"] in enabled]
    config_path = tmp_path / "mapping.json"
    config_path.write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return config_path


def write_api(tmp_path: Path, files: dict[str, list[dict]]) -> None:
    api_dir = tmp_path / "src_api"
    api_dir.mkdir(exist_ok=True)
    for name, rows in files.items():
        (api_dir / name).write_text(
            json.dumps(rows, ensure_ascii=False), encoding="utf-8"
        )


def write_csv(tmp_path: Path, rows: list[dict]) -> None:
    path = tmp_path / "src_export.csv"
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_COLUMNS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def write_excel(tmp_path: Path, rows: list[dict]) -> None:
    frame = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    frame.to_excel(tmp_path / "src_manual.xlsx", sheet_name="客户", index=False)


def run_pipeline(
    tmp_path: Path,
    config_path: Path | None = None,
    enabled: list[str] | None = None,
):
    path = config_path or make_config(tmp_path, enabled)
    return merge.run(path, base_dir=tmp_path, output_dir=tmp_path)


def read_csv_path(path: Path) -> list[dict]:
    with path.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def read_output(tmp_path: Path, filename: str) -> list[dict]:
    return read_csv_path(tmp_path / "out" / filename)


def row_by_id(rows: list[dict], customer_id: str) -> dict:
    matches = [row for row in rows if row["customer_id"] == customer_id]
    assert matches, f"master 中找不到客户 {customer_id}"
    return matches[0]


def api_row(**overrides) -> dict:
    row = {
        "customer_code": "C000001",
        "name": "张三",
        "level": "金卡",
        "credit_limit": 1000.5,
        "phone": "13800000001",
        "updated_at": "2026-10-01 06:00:00",
    }
    row.update(overrides)
    return row


def csv_row(**overrides) -> dict:
    row = {
        "cust_id": "0001",
        "name": "张三",
        "level": "金卡",
        "credit_limit": "1,000.50",
        "phone": "13800000001",
        "last_updated": "2026-10-01 20:00:00",
    }
    row.update(overrides)
    return row


def excel_row(**overrides) -> dict:
    row = {
        "会员号": "0001",
        "姓名": "张三",
        "等级": "金卡",
        "信用额度": 1000.5,
        "联系方式": "13800000001",
        "更新日期": "2026-10-01",
    }
    row.update(overrides)
    return row


@pytest.fixture
def rule_config(tmp_path: Path):
    write_csv(tmp_path, [])
    config_path = make_config(tmp_path, enabled=["csv"])
    return merge.load_config(config_path, base_dir=tmp_path)


# ---------------------------------------------------------------------------
# 归一规则
# ---------------------------------------------------------------------------


def test_identifier_normalization(rule_config) -> None:
    rule = rule_config.field_rules["customer_id"]
    cases = {
        "C000046": "46",
        "046": "46",
        "0183": "183",
        " 0007 ": "7",
        "0": "0",
        "122": "122",
    }
    for raw, expected in cases.items():
        assert merge.normalize_value(raw, rule, ()) == expected
    assert merge.normalize_value("CABC", rule, ()) is None
    assert merge.normalize_value("", rule, ()) is None


def test_decimal_normalization(rule_config) -> None:
    rule = rule_config.field_rules["credit_limit"]
    assert merge.normalize_value("111,373.71", rule, ()) == "111373.71"
    assert merge.normalize_value("48,283.10", rule, ()) == "48283.10"
    assert merge.normalize_value(1000, rule, ()) == "1000.00"
    assert merge.normalize_value(132397.1, rule, ()) == "132397.10"
    assert merge.normalize_value("abc", rule, ()) is None


def test_datetime_normalization(rule_config) -> None:
    rule = rule_config.field_rules["updated_at"]
    formats = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d")
    value = merge.normalize_value("2026-10-07 20:00:00", rule, formats)
    assert merge.display_value(value, rule) == "2026-10-07 20:00:00"
    value = merge.normalize_value("2026-10-07", rule, formats)
    assert merge.display_value(value, rule) == "2026-10-07 00:00:00"
    assert merge.normalize_value("07/10/2026", rule, formats) is None


# ---------------------------------------------------------------------------
# 配置驱动：接口改名与列名映射
# ---------------------------------------------------------------------------


def test_api_field_rename_is_config_driven(tmp_path: Path) -> None:
    """升级前 cust_no 与升级后 customer_code 都归一到同一客户号。"""
    write_api(
        tmp_path,
        {
            "cust_2026-10-01.json": [
                api_row(
                    cust_no="C000001",
                    customer_code=None,
                    updated_at="2026-10-01 06:00:00",
                )
            ],
            "cust_2026-10-02.json": [
                api_row(
                    cust_no=None,
                    customer_code="C000001",
                    updated_at="2026-10-02 06:00:00",
                )
            ],
        },
    )
    report, _ = run_pipeline(tmp_path, enabled=["api"])
    master = read_output(tmp_path, "master.csv")
    assert len(master) == 1
    assert report["counts"]["quarantine_rows"] == 0
    # 旧字段名行 cust_no=None 会被跳过；新文件提供最新记录
    assert row_by_id(master, "1")["updated_at"] == "2026-10-02 06:00:00"


def test_empty_api_file_is_tolerated(tmp_path: Path) -> None:
    """某天接口没有数据（空数组）时不应中断整条管道。"""
    write_api(
        tmp_path,
        {"cust_2026-10-01.json": [api_row()], "cust_2026-10-02.json": []},
    )
    report, _ = run_pipeline(tmp_path, enabled=["api"])
    master = read_output(tmp_path, "master.csv")
    assert len(master) == 1
    assert report["sources"]["api"]["rows_read"] == 1
    assert report["sources"]["api"]["files"] == 2


def test_mapping_change_needs_no_code_change(tmp_path: Path) -> None:
    """改配置里的列名映射即可适配不同的表头，逻辑代码不用动。"""
    path = tmp_path / "src_export.csv"
    renamed_columns = ["客户编号", "姓名", "等级", "信用额度", "手机", "更新时间"]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=renamed_columns, lineterminator="\n")
        writer.writeheader()
        writer.writerow(
            {
                "客户编号": "0001",
                "姓名": "张三",
                "等级": "金卡",
                "信用额度": "1,000.50",
                "手机": "13800000001",
                "更新时间": "2026-10-01 20:00:00",
            }
        )
    config = copy.deepcopy(BASE_CONFIG)
    config["sources"] = [config["sources"][1]]
    config["sources"][0]["field_map"] = dict(
        zip(renamed_columns, BASE_CONFIG["target_fields"])
    )
    config_path = tmp_path / "mapping.json"
    config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
    report, _ = run_pipeline(tmp_path, config_path)
    master = read_output(tmp_path, "master.csv")
    assert len(master) == 1
    assert row_by_id(master, "1")["credit_limit"] == "1000.50"
    assert report["counts"]["master_rows"] == 1


# ---------------------------------------------------------------------------
# 合并规则
# ---------------------------------------------------------------------------


def test_source_priority_and_contact_date_rule(tmp_path: Path) -> None:
    """普通字段按 API > CSV > Excel；联系方式取日期最新。"""
    write_api(tmp_path, {"cust_2026-10-01.json": [api_row(phone="13800000001")]})
    write_csv(
        tmp_path,
        [
            csv_row(
                cust_id="0001",
                name="张三改",
                level="银卡",
                credit_limit="999.00",
                phone="13800000002",
                last_updated="2026-10-05 20:00:00",
            )
        ],
    )
    write_excel(
        tmp_path,
        [
            excel_row(
                姓名="张三再改",
                等级="白金",
                信用额度=777,
                联系方式="13800000003",
                更新日期="2026-10-10",
            )
        ],
    )
    _, _ = run_pipeline(tmp_path, enabled=["api", "csv", "excel"])
    master = read_output(tmp_path, "master.csv")
    row = row_by_id(master, "1")
    assert row["name"] == "张三"
    assert row["level"] == "金卡"
    assert row["credit_limit"] == "1000.50"
    # 三个来源的联系方式都不同，取日期最新的 Excel（2026-10-10）
    assert row["phone"] == "13800000003"
    # 行级更新时间取所有来源里最新的一个
    assert row["updated_at"] == "2026-10-10 00:00:00"


def test_contact_rule_ignores_source_priority(tmp_path: Path) -> None:
    write_api(
        tmp_path,
        {
            "cust_2026-10-01.json": [
                api_row(phone="13800000001", updated_at="2026-10-01 06:00:00")
            ]
        },
    )
    write_csv(
        tmp_path,
        [
            csv_row(
                cust_id="0001",
                phone="13800000009",
                last_updated="2026-10-09 20:00:00",
            )
        ],
    )
    _, _ = run_pipeline(tmp_path, enabled=["api", "csv"])
    master = read_output(tmp_path, "master.csv")
    row = row_by_id(master, "1")
    assert row["phone"] == "13800000009"
    assert row["name"] == "张三"  # 非联系方式字段仍按优先级取 API


def test_latest_record_within_same_source(tmp_path: Path) -> None:
    """同一来源多天记录：普通字段取本来源最新，联系方式同样取最新非空。"""
    write_api(
        tmp_path,
        {
            "cust_2026-10-01.json": [
                api_row(
                    name="张三",
                    level="银卡",
                    phone=None,
                    updated_at="2026-10-01 06:00:00",
                )
            ],
            "cust_2026-10-05.json": [
                api_row(
                    name="张三最新",
                    level="金卡",
                    phone="13800000005",
                    updated_at="2026-10-05 06:00:00",
                )
            ],
        },
    )
    _, _ = run_pipeline(tmp_path, enabled=["api"])
    master = read_output(tmp_path, "master.csv")
    row = row_by_id(master, "1")
    assert row["name"] == "张三最新"
    assert row["level"] == "金卡"
    assert row["phone"] == "13800000005"
    assert row["updated_at"] == "2026-10-05 06:00:00"


def test_fill_contact_from_any_nonempty_source(tmp_path: Path) -> None:
    write_api(tmp_path, {"cust_2026-10-01.json": [api_row(phone=None)]})
    write_csv(
        tmp_path,
        [
            csv_row(
                cust_id="0001", phone="13800000002", last_updated="2026-10-05 20:00:00"
            )
        ],
    )
    write_excel(tmp_path, [excel_row(联系方式=None, 更新日期="2026-10-10")])
    _, _ = run_pipeline(tmp_path, enabled=["api", "csv", "excel"])
    master = read_output(tmp_path, "master.csv")
    assert row_by_id(master, "1")["phone"] == "13800000002"


# ---------------------------------------------------------------------------
# 必填闸门与 quarantine
# ---------------------------------------------------------------------------


def test_missing_required_field_row_goes_to_quarantine(tmp_path: Path) -> None:
    write_api(tmp_path, {"cust_2026-10-01.json": [api_row()]})
    write_csv(
        tmp_path,
        [csv_row(cust_id="0001", level="", last_updated="2026-10-05 20:00:00")],
    )
    report, _ = run_pipeline(tmp_path, enabled=["api", "csv"])
    quarantine = read_output(tmp_path, "quarantine.csv")
    assert len(quarantine) == 1
    row = quarantine[0]
    assert row["customer_id"] == "1"
    assert row["reason_codes"] == "missing_level"
    assert row["sources"] == "csv"
    assert row["record_locators"] == "csv:src_export.csv#2"
    assert report["quarantine"]["reason_distribution"] == {"missing_level": 1}
    master = read_output(tmp_path, "master.csv")
    assert len(master) == 1  # 客户仍由 API 记录成行


def test_missing_and_invalid_customer_id(tmp_path: Path) -> None:
    write_csv(
        tmp_path,
        [
            csv_row(cust_id="", name="无号客户"),
            csv_row(cust_id="ABC", name="号码异常"),
        ],
    )
    report, _ = run_pipeline(tmp_path, enabled=["csv"])
    quarantine = read_output(tmp_path, "quarantine.csv")
    assert len(quarantine) == 2
    reasons = {row["reason_codes"]: row for row in quarantine}
    assert set(reasons) == {"missing_customer_id", "invalid_customer_id"}
    assert reasons["invalid_customer_id"]["raw_key"] == "ABC"
    assert reasons["invalid_customer_id"]["reason_detail"] == "客户号无法归一: 'ABC'"
    assert report["counts"]["quarantine_keyless_rows"] == 2
    assert read_output(tmp_path, "master.csv") == []


def test_duplicate_rows_within_source_use_complete_row(tmp_path: Path) -> None:
    """同源重复行：晚行缺等级会被隔离，早的完整行仍可成行。"""
    write_csv(
        tmp_path,
        [
            csv_row(cust_id="0077", level="银卡", last_updated="2026-10-07 20:00:00"),
            csv_row(cust_id="0077", level="", last_updated="2026-10-13 21:00:00"),
        ],
    )
    report, _ = run_pipeline(tmp_path, enabled=["csv"])
    master = read_output(tmp_path, "master.csv")
    assert len(master) == 1
    assert row_by_id(master, "77")["level"] == "银卡"
    assert report["quarantine"]["reason_distribution"] == {"missing_level": 1}


# ---------------------------------------------------------------------------
# 冲突计数
# ---------------------------------------------------------------------------


def test_cross_source_conflict_counting(tmp_path: Path) -> None:
    write_api(tmp_path, {"cust_2026-10-01.json": [api_row(level="金卡")]})
    write_csv(tmp_path, [csv_row(cust_id="0001", level="银卡")])
    report, _ = run_pipeline(tmp_path, enabled=["api", "csv"])
    conflicts = report["conflicts"]
    assert conflicts["total_field_conflicts"] == 1
    assert conflicts["by_field"] == {"level": 1}
    assert conflicts["by_source_pair"] == {"api|csv": 1}
    assert conflicts["customers_with_conflicts"] == 1


def test_no_conflict_when_sources_agree(tmp_path: Path) -> None:
    write_api(tmp_path, {"cust_2026-10-01.json": [api_row()]})
    write_csv(tmp_path, [csv_row(cust_id="0001")])
    report, _ = run_pipeline(tmp_path, enabled=["api", "csv"])
    assert report["conflicts"]["total_field_conflicts"] == 0
    assert report["conflicts"]["customers_with_conflicts"] == 0


# ---------------------------------------------------------------------------
# 输出与配置校验
# ---------------------------------------------------------------------------


def test_output_formats_and_headers(tmp_path: Path) -> None:
    write_api(tmp_path, {"cust_2026-10-01.json": [api_row(credit_limit=132397.1)]})
    write_csv(tmp_path, [csv_row(cust_id="0001", level="")])
    _, config = run_pipeline(tmp_path, enabled=["api", "csv"])
    master = read_output(tmp_path, "master.csv")
    row = row_by_id(master, "1")
    assert row["credit_limit"] == "132397.10"
    assert row["customer_id"] == "1"
    assert list(master[0]) == list(config.target_fields)
    quarantine = read_output(tmp_path, "quarantine.csv")
    assert list(quarantine[0]) == [
        *config.target_fields,
        "reason_codes",
        "reason_detail",
        "sources",
        "record_locators",
        "raw_key",
    ]
    report_path = tmp_path / "out" / "run_report.json"
    assert report_path.is_file()
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert set(report) >= {"counts", "sources", "quarantine", "conflicts"}


def test_config_rejects_unknown_target_field(tmp_path: Path) -> None:
    config = copy.deepcopy(BASE_CONFIG)
    config["sources"][0]["field_map"]["cust_no"] = "bogus_field"
    config_path = tmp_path / "mapping.json"
    config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
    write_api(tmp_path, {"cust_2026-10-01.json": [api_row()]})
    with pytest.raises(merge.MergeError, match="未知目标字段"):
        merge.load_config(config_path, base_dir=tmp_path)


def test_source_without_key_column_fails_loudly(tmp_path: Path) -> None:
    write_api(
        tmp_path,
        {"cust_2026-10-01.json": [{"name": "无号", "level": "金卡"}]},
    )
    with pytest.raises(merge.MergeError, match="缺少客户号列"):
        run_pipeline(tmp_path, enabled=["api"])


# ---------------------------------------------------------------------------
# 仓库样例数据端到端
# ---------------------------------------------------------------------------


def test_sample_data_end_to_end(tmp_path: Path) -> None:
    config_path = REPO_ROOT / "mapping.json"
    report, _ = merge.run(config_path, base_dir=REPO_ROOT, output_dir=tmp_path)
    # 仓库配置把产物写在数据根目录，这里 output_dir=tmp_path
    master = read_csv_path(tmp_path / "master.csv")
    quarantine = read_csv_path(tmp_path / "quarantine.csv")

    counts = report["counts"]
    assert counts["source_rows_total"] == 899
    assert counts["customers_distinct"] == 289
    assert counts["master_rows"] == 289
    assert counts["quarantine_rows"] == 18
    assert counts["source_rows_rejected"] == 18

    # master：客户号唯一、必填字段齐全、类型与格式统一
    assert len(master) == 289
    assert len({row["customer_id"] for row in master}) == 289
    for row in master:
        assert row["customer_id"] and row["name"] and row["level"]
        assert row["customer_id"].isdigit()
        assert row["customer_id"] == str(int(row["customer_id"]))
        assert row["credit_limit"] and len(row["credit_limit"].split(".")[-1]) == 2
        assert re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", row["updated_at"])

    # 跨格式关联：手工表的 "066"/"0169" 等带零会员号与 CSV 的 "122" 等价归一
    assert row_by_id(master, "240")["name"] == "张伟6"
    assert row_by_id(master, "240")["updated_at"] == "2026-10-09 00:00:00"
    assert row_by_id(master, "17")["name"] == "马超8"
    assert row_by_id(master, "17")["credit_limit"] == "122958.30"
    assert row_by_id(master, "146")["credit_limit"] == "132397.10"

    # quarantine：每条都有原因代码，且来源与定位可回查
    assert len(quarantine) == 18
    for row in quarantine:
        assert row["reason_codes"]
        assert row["record_locators"]
        assert row["sources"] in {"api", "csv", "excel"}
    assert report["quarantine"]["reason_distribution"] == {
        "missing_customer_id": 4,
        "missing_level": 8,
        "missing_name": 6,
    }
    assert report["quarantine"]["by_source"] == {"csv": 12, "excel": 6}

    # 各来源贡献行数与采用/隔离行数自洽
    source_report = report["sources"]
    assert {name: stats["rows_read"] for name, stats in source_report.items()} == {
        "api": 671,
        "csv": 142,
        "excel": 86,
    }
    for stats in source_report.values():
        assert stats["rows_used"] + stats["rows_rejected"] == stats["rows_read"]
    assert sum(stats["rows_rejected"] for stats in source_report.values()) == 18

    # 冲突计数：总数 = 分字段之和，样例受上限约束
    conflicts = report["conflicts"]
    assert conflicts["total_field_conflicts"] == sum(conflicts["by_field"].values())
    assert conflicts["total_field_conflicts"] == 105
    assert conflicts["by_field"] == {"credit_limit": 1, "level": 9, "phone": 95}
    assert len(conflicts["examples"]) <= 20
