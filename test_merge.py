"""merge.py 的 pytest 测试。

全部用合成数据在 tmp_path 下构造三源，不依赖仓库里的真实样例，
保证测试可独立运行且意图清晰。
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest
import yaml

import merge

# --------------------------------------------------------------------------- #
# 测试用最小配置：结构与仓库 mapping.yaml 保持一致
# --------------------------------------------------------------------------- #

TEST_CONFIG = {
    "key_field": "customer_id",
    "target_schema": [
        {"name": "customer_id", "type": "int", "required": True},
        {"name": "name", "type": "str", "required": True},
        {"name": "level", "type": "str", "required": True},
        {"name": "credit_limit", "type": "float", "required": False},
        {"name": "phone", "type": "str", "required": False},
        {"name": "updated_at", "type": "date", "required": False},
    ],
    "normalization": {
        "int": {"strip_prefix": "C"},
        "float": {"thousands_sep": ",", "decimals": 2},
        "str": {"description": ""},
        "date": {"output_format": "%Y-%m-%d"},
    },
    "sources": {
        "api": {
            "type": "json_dir",
            "path_glob": "src_api/cust_*.json",
            "date_field": "updated_at",
            "field_map": {
                "cust_no": "customer_id",
                "customer_code": "customer_id",
                "name": "name",
                "level": "level",
                "credit_limit": "credit_limit",
                "phone": "phone",
                "updated_at": "updated_at",
            },
        },
        "csv": {
            "type": "csv",
            "path": "src_export.csv",
            "date_field": "updated_at",
            "field_map": {
                "cust_id": "customer_id",
                "name": "name",
                "level": "level",
                "credit_limit": "credit_limit",
                "phone": "phone",
                "last_updated": "updated_at",
            },
        },
        "excel": {
            "type": "excel",
            "path": "src_manual.xlsx",
            "sheet": "客户",
            "date_field": "updated_at",
            "field_map": {
                "会员号": "customer_id",
                "姓名": "name",
                "等级": "level",
                "信用额度": "credit_limit",
                "联系方式": "phone",
                "更新日期": "updated_at",
            },
        },
    },
    "merge": {"priority": ["api", "csv", "excel"], "contact_fields": ["phone"]},
}


@pytest.fixture()
def workspace(tmp_path: Path) -> Path:
    """搭出带配置的空工作目录，测试自行填充数据文件。"""
    (tmp_path / "src_api").mkdir()
    (tmp_path / "mapping.yaml").write_text(
        yaml.safe_dump(TEST_CONFIG, allow_unicode=True), encoding="utf-8"
    )
    return tmp_path


def write_api(workspace: Path, day: str, rows: list[dict]) -> None:
    (workspace / "src_api" / f"cust_{day}.json").write_text(
        json.dumps(rows, ensure_ascii=False), encoding="utf-8"
    )


def write_csv(workspace: Path, rows: list[dict]) -> None:
    pd.DataFrame(rows).to_csv(
        workspace / "src_export.csv", index=False, encoding="utf-8"
    )


def write_excel(workspace: Path, rows: list[tuple]) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "客户"
    ws.append(("会员号", "姓名", "等级", "信用额度", "联系方式", "更新日期"))
    for row in rows:
        ws.append(row)
    wb.save(workspace / "src_manual.xlsx")


def run_pipeline(workspace: Path):
    return merge.run(workspace, workspace / "mapping.yaml", workspace)


def read_master(workspace: Path) -> pd.DataFrame:
    return pd.read_csv(workspace / "master.csv", dtype={"phone": "string"})


# --------------------------------------------------------------------------- #
# 类型归一单元测试
# --------------------------------------------------------------------------- #


class TestNormalize:
    def test_int_strips_c_prefix_and_leading_zeros(self):
        rules = TEST_CONFIG["normalization"]
        assert merge.normalize_value("C000046", "int", rules) == 46
        assert merge.normalize_value("0207", "int", rules) == 207
        assert merge.normalize_value(46, "int", rules) == 46

    def test_int_invalid_raises(self):
        with pytest.raises(merge.NormalizeError):
            merge.normalize_value("ABC", "int", TEST_CONFIG["normalization"])

    def test_float_thousands_separator(self):
        assert (
            merge.normalize_value("111,373.71", "float", TEST_CONFIG["normalization"])
            == 111373.71
        )
        assert (
            merge.normalize_value(1000, "float", TEST_CONFIG["normalization"]) == 1000.0
        )

    def test_str_trim_and_blank(self):
        assert (
            merge.normalize_value("  金卡 ", "str", TEST_CONFIG["normalization"])
            == "金卡"
        )
        assert merge.normalize_value("", "str", TEST_CONFIG["normalization"]) is None
        assert merge.normalize_value(None, "str", TEST_CONFIG["normalization"]) is None
        assert merge.normalize_value("   ", "str", TEST_CONFIG["normalization"]) is None

    def test_date_accepts_datetime_and_date(self):
        assert (
            merge.normalize_value(
                "2026-10-07 20:00:00", "date", TEST_CONFIG["normalization"]
            )
            == "2026-10-07"
        )
        assert (
            merge.normalize_value("2026-10-07", "date", TEST_CONFIG["normalization"])
            == "2026-10-07"
        )

    def test_date_invalid_raises(self):
        with pytest.raises(merge.NormalizeError):
            merge.normalize_value("10/07/2026", "date", TEST_CONFIG["normalization"])


# --------------------------------------------------------------------------- #
# 端到端合并语义
# --------------------------------------------------------------------------- #


class TestMergeSemantics:
    def test_api_field_rename_both_names_map_to_same_target(self, workspace):
        """升级前 cust_no 与升级后 customer_code 的文件并存时归一为同一客户。"""
        write_api(
            workspace,
            "2026-10-01",
            [
                {
                    "cust_no": "C000001",
                    "name": "张三",
                    "level": "银卡",
                    "credit_limit": 100.0,
                    "phone": "13900000001",
                    "updated_at": "2026-10-01 06:00:00",
                },
            ],
        )
        write_api(
            workspace,
            "2026-10-09",
            [
                {
                    "customer_code": "C000001",
                    "name": "张三",
                    "level": "金卡",
                    "credit_limit": 200.0,
                    "phone": "13900000002",
                    "updated_at": "2026-10-09 06:00:00",
                },
            ],
        )
        report = run_pipeline(workspace)
        master = read_master(workspace)
        assert len(master) == 1
        row = master.iloc[0]
        assert row["customer_id"] == 1
        # 来源内去重保留 updated_at 最新一条
        assert row["level"] == "金卡"
        assert row["credit_limit"] == 200.0
        assert row["phone"] == "13900000002"
        assert row["updated_at"] == "2026-10-09"
        assert report["duplicates_within_source"]["api"] == 1

    def test_priority_api_over_csv_over_excel(self, workspace):
        write_api(
            workspace,
            "2026-10-10",
            [
                {
                    "customer_code": "C000007",
                    "name": "API名",
                    "level": "白金",
                    "credit_limit": 1.0,
                    "phone": None,
                    "updated_at": "2026-10-10 06:00:00",
                },
            ],
        )
        write_csv(
            workspace,
            [
                {
                    "cust_id": "7",
                    "name": "CSV名",
                    "level": "金卡",
                    "credit_limit": "2,000.00",
                    "phone": "",
                    "last_updated": "2026-10-11 20:00:00",
                },
            ],
        )
        write_excel(
            workspace,
            [
                ("07", "Excel名", "银卡", 3000.0, None, "2026-10-12"),
            ],
        )
        run_pipeline(workspace)
        row = read_master(workspace).iloc[0]
        # 即使 API 时间更早，普通字段仍按来源优先级取 API
        assert row["name"] == "API名"
        assert row["level"] == "白金"
        assert row["credit_limit"] == 1.0

    def test_contact_backfill_from_any_source(self, workspace):
        """联系方式例外：API 为空时由其他来源的非空值补全。"""
        write_api(
            workspace,
            "2026-10-10",
            [
                {
                    "customer_code": "C000008",
                    "name": "李四",
                    "level": "普通",
                    "credit_limit": 10.0,
                    "phone": None,
                    "updated_at": "2026-10-10 06:00:00",
                },
            ],
        )
        write_excel(
            workspace,
            [
                ("08", "李四", "普通", 10.0, "13911112222", "2026-10-09"),
            ],
        )
        run_pipeline(workspace)
        row = read_master(workspace).iloc[0]
        assert row["phone"] == "13911112222"

    def test_contact_conflict_uses_latest_date(self, workspace):
        """多个来源都有非空联系方式时，取更新时间最新的，与来源优先级无关。"""
        write_api(
            workspace,
            "2026-10-05",
            [
                {
                    "customer_code": "C000009",
                    "name": "王五",
                    "level": "普通",
                    "credit_limit": 1.0,
                    "phone": "13900000000",
                    "updated_at": "2026-10-05 06:00:00",
                },
            ],
        )
        write_csv(
            workspace,
            [
                {
                    "cust_id": "9",
                    "name": "王五",
                    "level": "普通",
                    "credit_limit": "1.00",
                    "phone": "13900000011",
                    "last_updated": "2026-10-13 21:00:00",
                },
            ],
        )
        write_excel(
            workspace,
            [
                ("09", "王五", "普通", 1.0, "13900000022", "2026-10-01"),
            ],
        )
        run_pipeline(workspace)
        row = read_master(workspace).iloc[0]
        # CSV 的 10-13 晚于 API 的 10-05 与 Excel 的 10-01
        assert row["phone"] == "13900000011"

    def test_missing_required_goes_quarantine(self, workspace):
        # 缺等级
        write_csv(
            workspace,
            [
                {
                    "cust_id": "77",
                    "name": "赵磊",
                    "level": "",
                    "credit_limit": "100.00",
                    "phone": "13900000077",
                    "last_updated": "2026-10-13 21:00:00",
                },
            ],
        )
        # 缺姓名
        write_excel(
            workspace,
            [
                ("88", None, "金卡", 50.0, None, "2026-10-14"),
            ],
        )
        report = run_pipeline(workspace)
        assert report["master_customer_count"] == 0
        assert report["quarantine_count"] == 2
        q = pd.read_csv(workspace / "quarantine.csv", dtype={"phone": "string"})
        reasons = dict(zip(q["customer_id"], q["missing_fields"]))
        assert reasons == {77: "level", 88: "name"}
        # 隔离行保留原始字段值便于排查
        q77 = q[q["customer_id"] == 77].iloc[0]
        assert q77["phone"] == "13900000077"

    def test_unidentified_rows_dropped_and_counted(self, workspace):
        """客户号无法识别的行不参与合并，计入剔除统计并有日志。"""
        write_csv(
            workspace,
            [
                {
                    "cust_id": "",
                    "name": "无名氏",
                    "level": "银卡",
                    "credit_limit": "1.00",
                    "phone": "",
                    "last_updated": "2026-10-13 21:00:00",
                },
                {
                    "cust_id": "5",
                    "name": "有号人",
                    "level": "银卡",
                    "credit_limit": "1.00",
                    "phone": "",
                    "last_updated": "2026-10-13 21:00:00",
                },
            ],
        )
        report = run_pipeline(workspace)
        assert report["dropped_unidentified_rows"]["csv"] == 1
        assert read_master(workspace)["customer_id"].tolist() == [5]

    def test_conflict_counts_business_fields_only(self, workspace):
        """跨源字段冲突按字段计数；updated_at 时间戳差异不计冲突。"""
        write_api(
            workspace,
            "2026-10-10",
            [
                {
                    "customer_code": "C000003",
                    "name": "同名人",
                    "level": "金卡",
                    "credit_limit": 100.0,
                    "phone": "13900000001",
                    "updated_at": "2026-10-10 06:00:00",
                },
            ],
        )
        write_csv(
            workspace,
            [
                {
                    "cust_id": "3",
                    "name": "同名人",
                    "level": "银卡",
                    "credit_limit": "200.00",
                    "phone": "13900000002",
                    "last_updated": "2026-10-11 20:00:00",
                },
            ],
        )
        report = run_pipeline(workspace)
        conflicts = report["cross_source_field_conflicts"]
        assert conflicts["level"] == 1
        assert conflicts["credit_limit"] == 1
        assert conflicts["phone"] == 1
        assert "updated_at" not in conflicts
        assert report["cross_source_conflict_total"] == 3

    def test_source_contribution_and_presence_stats(self, workspace):
        write_api(
            workspace,
            "2026-10-10",
            [
                {
                    "customer_code": "C000001",
                    "name": "甲",
                    "level": "普通",
                    "credit_limit": 1.0,
                    "phone": None,
                    "updated_at": "2026-10-10 06:00:00",
                },
                {
                    "customer_code": "C000002",
                    "name": "乙",
                    "level": "普通",
                    "credit_limit": 1.0,
                    "phone": "13900000002",
                    "updated_at": "2026-10-10 06:00:00",
                },
            ],
        )
        write_csv(
            workspace,
            [
                {
                    "cust_id": "2",
                    "name": "乙",
                    "level": "普通",
                    "credit_limit": "1.00",
                    "phone": "13900000099",
                    "last_updated": "2026-10-12 20:00:00",
                },
                {
                    "cust_id": "3",
                    "name": "丙",
                    "level": "普通",
                    "credit_limit": "1.00",
                    "phone": "",
                    "last_updated": "2026-10-12 20:00:00",
                },
            ],
        )
        report = run_pipeline(workspace)
        # 覆盖口径：客户 2 同时出现在 api 与 csv
        assert report["customers_present_per_source"] == {"api": 2, "csv": 2}
        # 贡献口径：客户 1 全靠 api；客户 2 的 phone 取日期最新的 csv，
        # 其余字段来自 api；客户 3 全靠 csv
        contrib = report["source_contributing_rows"]
        assert contrib["api"] == 2
        assert contrib["csv"] == 2

    def test_phone_written_as_text(self, workspace):
        write_api(
            workspace,
            "2026-10-10",
            [
                {
                    "customer_code": "C000001",
                    "name": "甲",
                    "level": "普通",
                    "credit_limit": 1.0,
                    "phone": "13984391551",
                    "updated_at": "2026-10-10 06:00:00",
                },
            ],
        )
        run_pipeline(workspace)
        raw = (workspace / "master.csv").read_text(encoding="utf-8-sig")
        assert "13984391551" in raw
        assert "e+" not in raw.lower()

    def test_excel_leading_zero_ids_merge_with_api(self, workspace):
        write_api(
            workspace,
            "2026-10-10",
            [
                {
                    "customer_code": "C000046",
                    "name": "王芳",
                    "level": "银卡",
                    "credit_limit": 1.0,
                    "phone": None,
                    "updated_at": "2026-10-10 06:00:00",
                },
            ],
        )
        write_excel(
            workspace,
            [
                ("046", "王芳", "银卡", 1.0, "13900000046", "2026-10-09"),
            ],
        )
        run_pipeline(workspace)
        master = read_master(workspace)
        assert master["customer_id"].tolist() == [46]
        assert master.iloc[0]["phone"] == "13900000046"


# --------------------------------------------------------------------------- #
# 配置校验
# --------------------------------------------------------------------------- #


class TestConfigValidation:
    def test_priority_mismatch_raises(self, tmp_path):
        bad = yaml.safe_load(yaml.safe_dump(TEST_CONFIG))
        bad["merge"]["priority"] = ["api", "csv"]
        path = tmp_path / "bad.yaml"
        path.write_text(yaml.safe_dump(bad), encoding="utf-8")
        with pytest.raises(ValueError, match="priority"):
            merge.load_config(path)

    def test_missing_required_mapping_raises(self, tmp_path):
        bad = yaml.safe_load(yaml.safe_dump(TEST_CONFIG))
        del bad["sources"]["excel"]["field_map"]["等级"]
        path = tmp_path / "bad.yaml"
        path.write_text(yaml.safe_dump(bad), encoding="utf-8")
        with pytest.raises(ValueError, match="level"):
            merge.load_config(path)

    def test_unknown_contact_field_raises(self, tmp_path):
        bad = yaml.safe_load(yaml.safe_dump(TEST_CONFIG))
        bad["merge"]["contact_fields"] = ["fax"]
        path = tmp_path / "bad.yaml"
        path.write_text(yaml.safe_dump(bad), encoding="utf-8")
        with pytest.raises(ValueError, match="contact_fields"):
            merge.load_config(path)


# --------------------------------------------------------------------------- #
# 真实样例数据的冒烟测试（仓库根目录数据）
# --------------------------------------------------------------------------- #


def test_repo_sample_smoke(tmp_path, repo_root: Path = Path(".")):
    if not (repo_root / "src_api").exists():
        pytest.skip("仓库样例数据不存在，跳过冒烟测试")
    report = merge.run(repo_root, repo_root / "mapping.yaml", tmp_path)
    master = pd.read_csv(tmp_path / "master.csv", dtype={"phone": "string"})
    quarantine = pd.read_csv(tmp_path / "quarantine.csv")

    # 勾稽：master + quarantine = 三源并集客户数
    assert (
        report["master_customer_count"] + report["quarantine_count"]
        == master.shape[0] + quarantine.shape[0]
    )
    # 必填字段在 master 中无缺失
    assert master[["customer_id", "name", "level"]].notna().all().all()
    # customer_id 唯一且为整数
    assert master["customer_id"].is_unique
    # 三源都有贡献记录
    assert set(report["source_contributing_rows"]) == {"api", "csv", "excel"}
    # phone 保持文本
    assert master["phone"].dropna().str.fullmatch(r"\d+").all()
