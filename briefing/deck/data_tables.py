"""各页表格行数据构建。"""

from __future__ import annotations

import pandas as pd

from briefing.analytics.engine import (
    compute_category_summary,
    compute_company_ranking,
    compute_increment_breakdown,
)
from briefing.excel.styled_table import ColumnSpec
from briefing.models import ReportConfig


def _rank_desc(s: pd.Series) -> pd.Series:
    return s.rank(ascending=False, method="min").astype(int)


def _aum_at(pivot: pd.DataFrame, company: str, date: str) -> float:
    if company not in pivot.index or date not in pivot.columns:
        return 0.0
    v = pivot.loc[company, date]
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def company_df(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["company"] != "__industry__"].copy()


def industry_df(df: pd.DataFrame) -> pd.DataFrame:
    ind = df[df["company"] == "__industry__"].copy()
    if ind.empty:
        ind = df.groupby(["category", "date"], as_index=False).agg(
            {"aum": "sum", "new_issue": "sum", "nav_change": "sum", "holding_sales": "sum"}
        )
        ind["company"] = "__industry__"
    return ind


# ---- column specs（按配置生成年/季双口径表头）----


def cols_category(config: ReportConfig) -> list[ColumnSpec]:
    y, q = config.ytd_tag(), config.quarter_tag()
    return [
        ColumnSpec("label", "类型", 70, "center", "text"),
        ColumnSpec("aum_current", "期末", 72, "center", "int"),
        ColumnSpec("aum_prev_quarter", "上季末", 72, "center", "int"),
        ColumnSpec("aum_year_start", "年初", 72, "center", "int"),
        ColumnSpec("increment", f"{y}增量", 68, "right", "int", bar="bidirectional"),
        ColumnSpec("growth_pct", f"{y}增速%", 52, "center", "pct", color_scale="growth"),
        ColumnSpec("new_issue", f"{y}新发", 56, "center", "int"),
        ColumnSpec("nav_change", f"{y}净值变化", 64, "center", "int"),
        ColumnSpec("holding_sales", f"{y}持营", 64, "center", "int"),
        ColumnSpec("q_increment", f"{q}增量", 64, "right", "int", bar="bidirectional"),
        ColumnSpec("q_growth_pct", f"{q}增速%", 52, "center", "pct", color_scale="growth"),
    ]


def cols_total(config: ReportConfig) -> list[ColumnSpec]:
    y, q = config.ytd_tag(), config.quarter_tag()
    return [
        ColumnSpec("rank", "排名", 34),
        ColumnSpec("rank_change", f"{y}排名变化", 48, "right", "int", bar="bidirectional"),
        ColumnSpec("rank_change_q", f"{q}排名变化", 48, "right", "int", bar="bidirectional"),
        ColumnSpec("company", "公司", 90, "center", "text"),
        ColumnSpec("aum", "总规模", 56, "center", "int"),
        ColumnSpec("increment", f"{y}增量", 62, "right", "int", bar="bidirectional"),
        ColumnSpec("growth_pct", f"{y}增速%", 44, "center", "pct"),
        ColumnSpec("q_increment", f"{q}增量", 58, "right", "int", bar="bidirectional"),
        ColumnSpec("q_growth_pct", f"{q}增速%", 44, "center", "pct"),
        ColumnSpec("money", "货币", 52, "center", "int"),
        ColumnSpec("non_money", "非货", 52, "center", "int"),
        ColumnSpec("money_rank", "货币排名", 48, "center", "int", color_scale="rank"),
        ColumnSpec("non_money_rank", "非货排名", 48, "center", "int", color_scale="rank"),
        ColumnSpec("active_rank", "主动权益排名", 52, "center", "int", color_scale="rank"),
        ColumnSpec("passive_rank", "被动权益排名", 52, "center", "int", color_scale="rank"),
        ColumnSpec("fi_plus_rank", "固收+排名", 48, "center", "int", color_scale="rank"),
        ColumnSpec("fi_rank", "固收排名", 48, "center", "int", color_scale="rank"),
        ColumnSpec("fof_rank", "FOF排名", 44, "center", "int", color_scale="rank"),
    ]


def cols_non_money(config: ReportConfig) -> list[ColumnSpec]:
    y, q = config.ytd_tag(), config.quarter_tag()
    return [
        ColumnSpec("rank", "排名", 34),
        ColumnSpec("rank_change", f"{y}排名变化", 48, "right", "int", bar="bidirectional"),
        ColumnSpec("rank_change_q", f"{q}排名变化", 48, "right", "int", bar="bidirectional"),
        ColumnSpec("company", "公司", 90, "center", "text"),
        ColumnSpec("aum", "非货总计", 56, "center", "int"),
        ColumnSpec("increment", f"{y}增量", 58, "right", "int", bar="bidirectional"),
        ColumnSpec("growth_pct", f"{y}增速%", 44, "center", "pct"),
        ColumnSpec("q_increment", f"{q}增量", 56, "right", "int", bar="bidirectional"),
        ColumnSpec("q_growth_pct", f"{q}增速%", 44, "center", "pct"),
        ColumnSpec(
            "active_equity", "主动权益规模", 52, "right", "int", bar="positive", bar_color="blue"
        ),
        ColumnSpec(
            "passive_equity", "被动权益规模", 52, "right", "int", bar="positive", bar_color="blue"
        ),
        ColumnSpec(
            "fixed_income_plus", "固收+规模", 48, "right", "int", bar="positive", bar_color="blue"
        ),
        ColumnSpec(
            "fixed_income", "固收规模", 48, "right", "int", bar="positive", bar_color="blue"
        ),
        ColumnSpec("fof", "FOF规模", 40, "right", "int", bar="positive", bar_color="blue"),
        ColumnSpec("active_rank", "主动权益排名", 52, "center", "int", color_scale="rank"),
        ColumnSpec("passive_rank", "被动权益排名", 52, "center", "int", color_scale="rank"),
        ColumnSpec("fi_rank", "固收排名", 48, "center", "int", color_scale="rank"),
        ColumnSpec("fi_plus_rank", "固收+排名", 48, "center", "int", color_scale="rank"),
        ColumnSpec("fof_rank", "FOF排名", 44, "center", "int", color_scale="rank"),
    ]


def cols_increment(config: ReportConfig) -> list[ColumnSpec]:
    y = config.ytd_tag()
    return [
        ColumnSpec("rank", "排名", 36),
        ColumnSpec("rank_change", f"{y}排名变化", 48),
        ColumnSpec("company", "公司", 100, "center", "text"),
        ColumnSpec("increment", f"{y}非货增量", 70, "right", "int", bar="bidirectional"),
        ColumnSpec("increment_rank", "增量排名", 52, "center", "int", color_scale="rank"),
        ColumnSpec("growth_pct", f"{y}增速%", 48, "center", "pct"),
        ColumnSpec("active_inc_rank", "主动权益增量排名", 64, "center", "int", color_scale="rank"),
        ColumnSpec("passive_inc_rank", "被动权益增量排名", 64, "center", "int", color_scale="rank"),
        ColumnSpec("fi_plus_inc_rank", "固收+增量排名", 58, "center", "int", color_scale="rank"),
        ColumnSpec("fi_inc_rank", "固收增量排名", 58, "center", "int", color_scale="rank"),
        ColumnSpec("active_inc", "主动权益增量", 70, "right", "int", bar="bidirectional"),
        ColumnSpec("passive_inc", "被动权益增量", 70, "right", "int", bar="bidirectional"),
        ColumnSpec("fi_plus_inc", "固收+增量", 64, "right", "int", bar="bidirectional"),
        ColumnSpec("fi_inc", "固收增量", 64, "right", "int", bar="bidirectional"),
    ]


def cols_business(config: ReportConfig) -> list[ColumnSpec]:
    y, q = config.ytd_tag(), config.quarter_tag()
    return [
        ColumnSpec("rank", "排名", 40),
        ColumnSpec("company", "公司", 110, "center", "text"),
        ColumnSpec("aum", "规模", 64, "center", "int"),
        ColumnSpec("increment", f"{y}规模增量", 72, "right", "int", bar="bidirectional"),
        ColumnSpec("growth_pct", f"{y}规模增幅%", 64, "center", "pct"),
        ColumnSpec("new_issue", "新发", 56, "center", "int"),
        ColumnSpec("nav_change", "净值变化", 64, "right", "int", bar="bidirectional"),
        ColumnSpec("holding_sales", "持营", 64, "right", "int", bar="bidirectional"),
        ColumnSpec("rank_change", f"{y}排名变化", 52),
        ColumnSpec("q_increment", f"{q}增量", 60, "right", "int", bar="bidirectional"),
        ColumnSpec("q_growth_pct", f"{q}增幅%", 52, "center", "pct"),
        ColumnSpec("rank_change_q", f"{q}排名变化", 52),
    ]


# 兼容旧引用：无配置时用中性表头
def _default_config_for_legacy() -> ReportConfig:
    from briefing.models import PeriodType

    return ReportConfig(
        title="",
        period_label="本期",
        period_type=PeriodType.QUARTER,
        data_source_note="",
        current_date="20251231",
        previous_quarter_date="20250930",
        year_start_date="20241231",
        focus_company="",
        focus_company_short="",
        ytd_column_label="年",
        quarter_column_label="当季",
    )


_legacy = _default_config_for_legacy()
COLS_CATEGORY = cols_category(_legacy)
COLS_TOTAL = cols_total(_legacy)
COLS_NON_MONEY = cols_non_money(_legacy)
COLS_INCREMENT = cols_increment(_legacy)
COLS_BUSINESS = cols_business(_legacy)


def build_category_rows(df: pd.DataFrame, config: ReportConfig) -> list[dict]:
    metrics = compute_category_summary(industry_df(df), config)
    rows = []
    for m in metrics:
        rows.append({
            "label": m.label,
            "aum_current": m.aum_current,
            "aum_prev_quarter": m.aum_prev_quarter,
            "aum_year_start": m.aum_year_start,
            "increment": m.increment,
            "growth_pct": m.growth_pct,
            "new_issue": m.new_issue,
            "nav_change": m.nav_change,
            "holding_sales": m.holding_sales,
            "q_increment": m.q_increment,
            "q_growth_pct": m.q_growth_pct,
            "company": m.label,
        })
    return rows


def _category_ranks(cos: pd.DataFrame, date: str, categories: list[str]) -> dict[str, pd.Series]:
    """全行业排名（数值越小越好），用于宽表色阶列。"""
    out: dict[str, pd.Series] = {}
    for cat in categories:
        sub = cos[cos["category"] == cat]
        if sub.empty:
            continue
        pivot = sub.pivot_table(index="company", columns="date", values="aum", aggfunc="sum")
        if date not in pivot.columns:
            continue
        out[cat] = _rank_desc(pivot[date])
    return out


def build_total_rows(df: pd.DataFrame, config: ReportConfig) -> list[dict]:
    cos = company_df(df)
    total = compute_company_ranking(cos, config, "total")
    cur = config.current_date
    money = cos[cos["category"] == "money"].pivot_table(index="company", columns="date", values="aum", aggfunc="sum")
    non = cos[cos["category"] == "non_money"].pivot_table(index="company", columns="date", values="aum", aggfunc="sum")
    ranks = _category_ranks(
        cos,
        cur,
        ["money", "non_money", "active_equity", "passive_equity", "fixed_income_plus", "fixed_income", "fof"],
    )

    def rnk(cat: str, company: str) -> int:
        s = ranks.get(cat)
        if s is None or company not in s.index:
            return 0
        return int(s.loc[company])

    rows = []
    for r in total:
        rows.append({
            "rank": r.rank,
            "rank_change": r.rank_change,
            "rank_change_q": r.rank_change_q,
            "company": r.company,
            "aum": r.aum,
            "increment": r.increment,
            "growth_pct": r.growth_pct,
            "q_increment": r.q_increment,
            "q_growth_pct": r.q_growth_pct,
            "money": _aum_at(money, r.company, cur),
            "non_money": _aum_at(non, r.company, cur),
            "money_rank": rnk("money", r.company),
            "non_money_rank": rnk("non_money", r.company),
            "active_rank": rnk("active_equity", r.company),
            "passive_rank": rnk("passive_equity", r.company),
            "fi_plus_rank": rnk("fixed_income_plus", r.company),
            "fi_rank": rnk("fixed_income", r.company),
            "fof_rank": rnk("fof", r.company),
        })
    return rows


def build_non_money_rows(df: pd.DataFrame, config: ReportConfig) -> list[dict]:
    cos = company_df(df)
    nm = compute_company_ranking(cos, config, "non_money")
    cur = config.current_date
    cats = ["active_equity", "passive_equity", "fixed_income_plus", "fixed_income", "fof"]
    pivots = {}
    for cat in cats:
        sub = cos[cos["category"] == cat]
        pivots[cat] = (
            sub.pivot_table(index="company", columns="date", values="aum", aggfunc="sum")
            if not sub.empty
            else pd.DataFrame()
        )
    ranks = _category_ranks(cos, cur, cats)

    def rnk(cat: str, company: str) -> int:
        s = ranks.get(cat)
        if s is None or company not in s.index:
            return 0
        return int(s.loc[company])

    rows = []
    for r in nm:
        rows.append({
            "rank": r.rank,
            "rank_change": r.rank_change,
            "rank_change_q": r.rank_change_q,
            "company": r.company,
            "aum": r.aum,
            "increment": r.increment,
            "growth_pct": r.growth_pct,
            "q_increment": r.q_increment,
            "q_growth_pct": r.q_growth_pct,
            "active_equity": _aum_at(pivots["active_equity"], r.company, cur),
            "passive_equity": _aum_at(pivots["passive_equity"], r.company, cur),
            "fixed_income_plus": _aum_at(pivots["fixed_income_plus"], r.company, cur),
            "fixed_income": _aum_at(pivots["fixed_income"], r.company, cur),
            "fof": _aum_at(pivots["fof"], r.company, cur),
            "active_rank": rnk("active_equity", r.company),
            "passive_rank": rnk("passive_equity", r.company),
            "fi_rank": rnk("fixed_income", r.company),
            "fi_plus_rank": rnk("fixed_income_plus", r.company),
            "fof_rank": rnk("fof", r.company),
        })
    return rows


def build_increment_rows(df: pd.DataFrame, config: ReportConfig) -> list[dict]:
    cos = company_df(df)
    breakdowns = compute_increment_breakdown(cos, config, "non_money")
    rows = []
    for b in breakdowns:
        rows.append({
            "rank": b.rank,
            "rank_change": b.rank_change,
            "company": b.company,
            "increment": b.increment,
            "increment_rank": b.increment_rank,
            "growth_pct": b.growth_pct,
            "active_inc_rank": b.category_increment_ranks.get("active_equity", 0),
            "passive_inc_rank": b.category_increment_ranks.get("passive_equity", 0),
            "fi_plus_inc_rank": b.category_increment_ranks.get("fixed_income_plus", 0),
            "fi_inc_rank": b.category_increment_ranks.get("fixed_income", 0),
            "active_inc": b.category_increments.get("active_equity", 0),
            "passive_inc": b.category_increments.get("passive_equity", 0),
            "fi_plus_inc": b.category_increments.get("fixed_income_plus", 0),
            "fi_inc": b.category_increments.get("fixed_income", 0),
        })
    return rows


def build_business_rows(df: pd.DataFrame, config: ReportConfig, category: str) -> list[dict]:
    cos = company_df(df)
    actual = category
    if category.startswith("passive_equity_etf") and cos[cos["category"] == category].empty:
        actual = "passive_equity"
    rankings = compute_company_ranking(cos, config, actual)
    rows = []
    for r in rankings:
        rows.append({
            "rank": r.rank,
            "company": r.company,
            "aum": r.aum,
            "increment": r.increment,
            "growth_pct": r.growth_pct,
            "new_issue": r.new_issue,
            "nav_change": r.nav_change,
            "holding_sales": r.holding_sales,
            "rank_change": r.rank_change,
            "q_increment": r.q_increment,
            "q_growth_pct": r.q_growth_pct,
            "rank_change_q": r.rank_change_q,
        })
    return rows
