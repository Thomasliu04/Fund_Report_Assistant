"""配置加载与数据导入"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import yaml

from briefing.models import PeriodType, ReportConfig


def load_config(path: str | Path) -> ReportConfig:
    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    report = raw["report"]
    dates = raw["dates"]
    labels = raw.get("column_labels") or {}
    return ReportConfig(
        title=report["title"].format(period_label=report["period_label"]),
        period_label=report["period_label"],
        period_type=PeriodType(report["period_type"]),
        data_source_note=report["data_source_note"],
        current_date=str(dates["current"]),
        previous_quarter_date=str(dates["previous_quarter"]),
        year_start_date=str(dates["year_start"]),
        focus_company=raw["focus_company"],
        focus_company_short=raw["focus_company_short"],
        top_n=raw.get("ranking", {}).get("top_n", 30),
        categories=raw.get("categories", []),
        sections=raw.get("sections", []),
        ytd_column_label=str(labels.get("ytd") or ""),
        quarter_column_label=str(labels.get("quarter") or ""),
    )


def load_fund_data(data_dir: str | Path) -> pd.DataFrame:
    """加载标准化 fund_metrics.csv"""
    path = Path(data_dir) / "fund_metrics.csv"
    df = pd.read_csv(path, dtype={"date": str})
    required = {"company", "category", "date", "aum"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"fund_metrics.csv 缺少字段: {missing}")
    for col in ("new_issue", "nav_change", "holding_sales"):
        if col not in df.columns:
            df[col] = 0.0
    return df


def load_category_summary(data_dir: str | Path) -> pd.DataFrame:
    """加载一级大类汇总 category_summary.csv"""
    path = Path(data_dir) / "category_summary.csv"
    if not path.exists():
        return pd.DataFrame()
    return pd.read_csv(path, dtype={"date": str})
