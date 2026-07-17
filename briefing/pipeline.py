"""报告构建流水线"""

from __future__ import annotations

from datetime import datetime

import pandas as pd

from briefing.analytics.engine import (
    compute_category_summary,
    compute_company_ranking,
    compute_increment_breakdown,
)
from briefing.data_loader import load_config, load_fund_data
from briefing.models import BriefingReport, CategoryMetrics, ReportConfig, ReportSection
from briefing.narrative.generator import (
    generate_business_ranking_narrative,
    generate_increment_narrative,
    generate_non_money_ranking_narrative,
    generate_overview_narrative,
    generate_total_ranking_narrative,
)
from briefing.render.html_renderer import (
    attach_increment_table,
    attach_ranking_table,
    render_html,
)


CATEGORY_LABELS = {
    "active_equity": "主动权益",
    "passive_equity": "被动权益",
    "passive_equity_etf_with_link": "权益ETF（含联接）",
    "passive_equity_etf_no_link": "权益ETF（不含联接）",
    "money": "货币",
    "fixed_income": "固收",
    "fixed_income_plus": "固收+",
    "fof": "FOF",
}


def _category_table(metrics: list[CategoryMetrics], config: ReportConfig) -> tuple[list[str], list[list[str]]]:
    headers = [
        "类型",
        config.current_date,
        config.previous_quarter_date,
        config.year_start_date,
        f"{config.period_label}增量",
        f"{config.period_label}增速%",
        f"{config.period_label}新发",
        f"{config.period_label}净值变化",
        f"{config.period_label}持营",
    ]
    rows = []
    for m in metrics:
        rows.append([
            m.label,
            str(round(m.aum_current)),
            str(round(m.aum_prev_quarter)),
            str(round(m.aum_year_start)),
            str(round(m.increment)),
            f"{round(m.growth_pct * 100)}",
            str(round(m.new_issue)),
            str(round(m.nav_change)),
            str(round(m.holding_sales)),
        ])
    return headers, rows


def build_report(config: ReportConfig, df: pd.DataFrame) -> BriefingReport:
    sections: list[ReportSection] = []

    # 一级大类：使用 industry 汇总行
    industry_df = df[df["company"] == "__industry__"]
    if industry_df.empty:
        industry_df = df.groupby(["category", "date"], as_index=False).agg({
            "aum": "sum", "new_issue": "sum", "nav_change": "sum", "holding_sales": "sum",
        })
        industry_df["company"] = "__industry__"

    cat_metrics = compute_category_summary(industry_df, config)

    # Section 1: 概况
    s1 = ReportSection(
        id="overview",
        title=f"{config.period_label}公募规模整体情况",
        narrative=generate_overview_narrative(cat_metrics, config),
        table_headers=[],
        table_rows=[],
        footnotes=["注：REITs和另类基金未在表中注明"],
    )
    h, r = _category_table(cat_metrics, config)
    s1.table_headers, s1.table_rows = h, r
    sections.append(s1)

    sub_cats = ["money", "non_money", "active_equity", "passive_equity", "fixed_income_plus", "fixed_income", "fof"]

    # Section 2: 总规模排名
    total_rankings = compute_company_ranking(df, config, "total", sub_categories=sub_cats)
    s2 = ReportSection(
        id="total_ranking",
        title=f"{config.period_label}公募基金总规模（含货币）排名",
        narrative=generate_total_ranking_narrative(total_rankings, config),
        table_headers=[], table_rows=[],
        footnotes=["注：被动权益=ETF+联接+场外普通指数+场外指数增强"],
    )
    attach_ranking_table(s2, total_rankings)
    sections.append(s2)

    # Section 3: 非货排名
    nm_rankings = compute_company_ranking(
        df, config, "non_money",
        sub_categories=["active_equity", "passive_equity", "fixed_income_plus", "fixed_income", "fof"],
    )
    s3 = ReportSection(
        id="non_money_ranking",
        title=f"{config.period_label}公募基金非货规模排名",
        narrative=generate_non_money_ranking_narrative(nm_rankings, config),
        table_headers=[], table_rows=[],
        footnotes=["注：被动权益=ETF+联接+场外普通指数+场外指数增强"],
    )
    attach_ranking_table(s3, nm_rankings)
    sections.append(s3)

    # Section 4: 非货增量
    increments = compute_increment_breakdown(df, config)
    s4 = ReportSection(
        id="increment_overview",
        title=f"{config.period_label}非货增量情况概览",
        narrative=generate_increment_narrative(increments, config),
        table_headers=[], table_rows=[],
        footnotes=["注：被动权益=ETF+联接+场外普通指数+场外指数增强"],
    )
    attach_increment_table(s4, increments)
    sections.append(s4)

    # Section 5: 各项业务
    s5 = ReportSection(
        id="business_rankings",
        title=f"{config.period_label}各项业务排名情况",
        narrative=generate_business_ranking_narrative([], config, ""),
        table_headers=[], table_rows=[],
    )
    s5.subsections = []
    for cat_key, cat_label in [
        ("active_equity", "主动权益"),
        ("passive_equity_etf_with_link", "权益ETF（含联接）"),
        ("passive_equity_etf_no_link", "权益ETF（不含联接）"),
        ("money", "货币"),
        ("fixed_income", "固收"),
        ("fixed_income_plus", "固收+"),
        ("fof", "FOF"),
    ]:
        actual_key = "passive_equity" if "etf" in cat_key else cat_key
        if actual_key not in df["category"].unique():
            continue
        rankings = compute_company_ranking(df, config, actual_key)
        sub = ReportSection(
            id=f"business_{cat_key}",
            title=f"{cat_label}排名" if cat_key != "money" else "货币排名情况",
            narrative=generate_business_ranking_narrative(rankings, config, cat_label),
            table_headers=[], table_rows=[],
        )
        attach_ranking_table(sub, rankings, extended=True)
        s5.subsections.append(sub)
    sections.append(s5)

    return BriefingReport(
        config=config,
        sections=sections,
        generated_at=datetime.now().isoformat(timespec="seconds"),
    )


def generate_briefing(config_path: str, data_dir: str, output_path: str) -> BriefingReport:
    config = load_config(config_path)
    df = load_fund_data(data_dir)
    report = build_report(config, df)
    render_html(report, output_path)
    return report
