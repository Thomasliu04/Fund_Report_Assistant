"""分析引擎：排名、增量、拆解"""

from __future__ import annotations

import pandas as pd

from briefing.models import CategoryMetrics, CompanyRanking, IncrementBreakdown, ReportConfig


def _fmt_pct(value: float) -> float:
    return round(value * 100, 0) if abs(value) < 5 else round(value * 100)


def _increment(row_current: pd.Series, row_start: pd.Series) -> tuple[float, float]:
    inc = row_current["aum"] - row_start["aum"]
    base = row_start["aum"]
    growth = inc / base if base else 0
    return inc, growth


def compute_category_summary(
    df: pd.DataFrame, config: ReportConfig
) -> list[CategoryMetrics]:
    """一级大类汇总（排除合计行，合计单独处理）"""
    dates = {
        "current": config.current_date,
        "prev_q": config.previous_quarter_date,
        "year_start": config.year_start_date,
    }
    results = []
    for cat in config.categories:
        key = cat["key"]
        label = cat["label"]
        subset = df[df["category"] == key]
        if subset.empty:
            continue
        by_date = {d: subset[subset["date"] == dt].iloc[0] for d, dt in dates.items() if not subset[subset["date"] == dt].empty}
        if "current" not in by_date or "year_start" not in by_date:
            continue
        cur, start = by_date["current"], by_date["year_start"]
        prev = by_date.get("prev_q", start)
        inc, growth = _increment(cur, start)
        q_inc, q_growth = _increment(cur, prev)

        def _flow(field: str) -> float:
            # 期末行存放本期/YTD 贡献
            return float(cur.get(field, 0) or 0)

        results.append(
            CategoryMetrics(
                category=key,
                label=label,
                aum_current=cur["aum"],
                aum_prev_quarter=prev["aum"],
                aum_year_start=start["aum"],
                increment=inc,
                growth_pct=growth,
                new_issue=_flow("new_issue"),
                nav_change=_flow("nav_change"),
                holding_sales=_flow("holding_sales"),
                q_increment=q_inc,
                q_growth_pct=q_growth,
            )
        )
    return results


def _rank_series(series: pd.Series, ascending: bool = False) -> pd.Series:
    return series.rank(ascending=ascending, method="min").astype(int)


def compute_company_ranking(
    df: pd.DataFrame,
    config: ReportConfig,
    category: str,
    sub_categories: list[str] | None = None,
) -> list[CompanyRanking]:
    """某公司品类的 TopN 排名"""
    dates = [config.current_date, config.previous_quarter_date, config.year_start_date]
    cat_df = df[df["category"] == category].copy()
    if cat_df.empty:
        return []

    pivot = cat_df.pivot_table(index="company", columns="date", values="aum", aggfunc="sum")
    for d in dates:
        if d not in pivot.columns:
            pivot[d] = 0

    cur_d, prev_d, ys_d = config.current_date, config.previous_quarter_date, config.year_start_date
    pivot["increment"] = pivot[cur_d] - pivot[ys_d]
    pivot["growth"] = pivot.apply(
        lambda r: r["increment"] / r[ys_d] if r[ys_d] else 0,
        axis=1,
    )
    pivot["q_increment"] = pivot[cur_d] - pivot[prev_d]
    pivot["q_growth"] = pivot.apply(
        lambda r: r["q_increment"] / r[prev_d] if r[prev_d] else 0,
        axis=1,
    )

    rank_cur = _rank_series(pivot[cur_d])
    rank_start = _rank_series(pivot[ys_d])
    rank_prev = _rank_series(pivot[prev_d])

    # 增量分解：CSV 约定——报告期末行存放本期/YTD 贡献（新发/净值/持营），直接取期末值
    for col in ("new_issue", "nav_change", "holding_sales"):
        if col in cat_df.columns:
            p = cat_df.pivot_table(index="company", columns="date", values=col, aggfunc="sum")
            if cur_d not in p.columns:
                p[cur_d] = 0
            pivot[col] = p[cur_d].fillna(0)
        else:
            pivot[col] = 0

    sub_categories = sub_categories or []
    sub_ranks: dict[str, pd.Series] = {}
    sub_aums: dict[str, pd.DataFrame] = {}
    for sc in sub_categories:
        sc_df = df[df["category"] == sc]
        if sc_df.empty:
            continue
        sp = sc_df.pivot_table(index="company", columns="date", values="aum", aggfunc="sum")
        if cur_d in sp.columns:
            sub_ranks[sc] = _rank_series(sp[cur_d])
            sub_aums[sc] = sp

    top = pivot.sort_values(cur_d, ascending=False).head(config.top_n)
    results = []
    for company, row in top.iterrows():
        sub_rank_dict = {sc: int(sub_ranks[sc].get(company, 0)) for sc in sub_ranks}
        sub_aum_dict = {
            sc: float(sub_aums[sc].loc[company, cur_d])
            if company in sub_aums[sc].index and cur_d in sub_aums[sc].columns
            else 0
            for sc in sub_aums
        }
        results.append(
            CompanyRanking(
                rank=int(rank_cur[company]),
                company=company,
                aum=row[cur_d],
                increment=row["increment"],
                growth_pct=row["growth"],
                rank_change=int(rank_start[company] - rank_cur[company]),
                rank_change_q=int(rank_prev[company] - rank_cur[company]),
                sub_rankings=sub_rank_dict,
                sub_aums=sub_aum_dict,
                new_issue=row.get("new_issue", 0),
                nav_change=row.get("nav_change", 0),
                holding_sales=row.get("holding_sales", 0),
                q_increment=row["q_increment"],
                q_growth_pct=row["q_growth"],
            )
        )
    results.sort(key=lambda x: x.rank)
    return results


def compute_increment_breakdown(
    df: pd.DataFrame,
    config: ReportConfig,
    category: str = "non_money",
    breakdown_categories: list[str] | None = None,
) -> list[IncrementBreakdown]:
    """非货增量拆解"""
    breakdown_categories = breakdown_categories or [
        "active_equity", "passive_equity", "fixed_income_plus", "fixed_income", "fof"
    ]
    cat_df = df[df["category"] == category]
    if cat_df.empty:
        return []

    cur_d, prev_d, ys_d = config.current_date, config.previous_quarter_date, config.year_start_date
    pivot = cat_df.pivot_table(index="company", columns="date", values="aum", aggfunc="sum")
    for d in [cur_d, ys_d, prev_d]:
        if d not in pivot.columns:
            pivot[d] = 0

    pivot["increment"] = pivot[cur_d] - pivot[ys_d]
    pivot["growth"] = pivot.apply(
        lambda r: r["increment"] / r[ys_d] if r[ys_d] else 0,
        axis=1,
    )
    pivot["q_increment"] = pivot[cur_d] - pivot[prev_d]
    pivot["q_growth"] = pivot.apply(
        lambda r: r["q_increment"] / r[prev_d] if r[prev_d] else 0,
        axis=1,
    )
    inc_rank = _rank_series(pivot["increment"])
    rank_cur = _rank_series(pivot[cur_d])
    rank_start = _rank_series(pivot[ys_d])
    rank_prev = _rank_series(pivot[prev_d])

    cat_increments: dict[str, pd.Series] = {}
    cat_inc_ranks: dict[str, pd.Series] = {}
    for bc in breakdown_categories:
        bc_df = df[df["category"] == bc]
        if bc_df.empty:
            continue
        bp = bc_df.pivot_table(index="company", columns="date", values="aum", aggfunc="sum")
        for d in [cur_d, ys_d]:
            if d not in bp.columns:
                bp[d] = 0
        inc = bp[cur_d] - bp[ys_d]
        cat_increments[bc] = inc
        cat_inc_ranks[bc] = _rank_series(inc)

    top = pivot.sort_values("increment", ascending=False).head(config.top_n)
    results = []
    for company in top.sort_values(cur_d, ascending=False).index[: config.top_n]:
        row = pivot.loc[company]
        results.append(
            IncrementBreakdown(
                rank=int(rank_cur[company]),
                company=company,
                increment=row["increment"],
                increment_rank=int(inc_rank[company]),
                growth_pct=row["growth"],
                rank_change=int(rank_start[company] - rank_cur[company]),
                category_increments={bc: float(cat_increments[bc].get(company, 0)) for bc in cat_increments},
                category_increment_ranks={
                    bc: int(cat_inc_ranks[bc].get(company, 0)) for bc in cat_inc_ranks
                },
                q_increment=row["q_increment"],
                q_growth_pct=row["q_growth"],
                rank_change_q=int(rank_prev[company] - rank_cur[company]),
            )
        )
    results.sort(key=lambda x: x.rank)
    return results


def find_notable_rank_changes(
    rankings: list[CompanyRanking], threshold: int = 3
) -> tuple[list[CompanyRanking], list[CompanyRanking]]:
    """找出排名变化显著的公司"""
    up = [r for r in rankings if r.rank_change >= threshold]
    down = [r for r in rankings if r.rank_change <= -threshold]
    return up, down
