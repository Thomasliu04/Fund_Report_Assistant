"""全册叙事 payload。"""

from __future__ import annotations

from briefing.analytics.engine import (
    compute_category_summary,
    compute_company_ranking,
    compute_increment_breakdown,
    find_notable_rank_changes,
)
from briefing.deck.data_tables import company_df, industry_df
from briefing.models import CompanyRanking, ReportConfig
from briefing.phase1 import build_overview_payload, build_total_ranking_payload


def _fmt_yi(v: float) -> str:
    return str(round(v))


def _fmt_pct(v: float) -> str:
    return f"{round(v * 100)}"


def _rank_change_text(change: int) -> str:
    if change > 0:
        return f"上升{change}名"
    if change < 0:
        return f"下降{abs(change)}名"
    return "不变"


def _focus(rankings: list[CompanyRanking], config: ReportConfig):
    return next((r for r in rankings if r.company == config.focus_company), None)


def build_non_money_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    focus = _focus(rankings, config)
    if not focus:
        return {
            "section_title": f"三、{config.period_label}公募基金非货规模排名",
            "p1": "暂无银华非货数据。",
            "p2": "",
            "p3": "",
            "footnote_passive": "注：被动权益=ETF+联接+场外普通指数+场外指数增强",
        }
    by_inc = sorted(rankings, key=lambda x: x.increment, reverse=True)
    by_g = sorted(rankings, key=lambda x: x.growth_pct, reverse=True)
    inc_pos = next(i + 1 for i, r in enumerate(by_inc) if r.company == config.focus_company)
    g_pos = next(i + 1 for i, r in enumerate(by_g) if r.company == config.focus_company)
    avg_g = sum(r.growth_pct for r in rankings) / max(len(rankings), 1)
    return {
        "section_title": f"三、{config.period_label}公募基金非货规模排名",
        "p1": (
            f"非货规模方面，{config.focus_company_short}位列第{focus.rank}位，"
            f"相较年初{_rank_change_text(focus.rank_change)}，"
            f"较上季度{_rank_change_text(focus.rank_change_q)}。"
        ),
        "p2": (
            f"{config.focus_company_short}非货，{config.period_label}增量{_fmt_yi(focus.increment)}亿，"
            f"增速{_fmt_pct(focus.growth_pct)}%，"
            f"增量和增速排名在Top{config.top_n}中分别是第{inc_pos}名和第{g_pos}名；"
            f"非货Top{config.top_n}公司作为一个整体，{config.period_label}增速为{_fmt_pct(avg_g)}%。"
        ),
        "p3": f"{config.focus_company_short}非货各品类表现请结合下表审阅，观点句可人工补充。",
        "footnote_passive": "注：被动权益=ETF+联接+场外普通指数+场外指数增强",
    }


def build_increment_payload(breakdowns, config: ReportConfig) -> dict:
    focus = next((b for b in breakdowns if b.company == config.focus_company), None)
    paras = []
    if focus:
        pe = focus.category_increment_ranks.get("passive_equity", 0)
        fi = focus.category_increment_ranks.get("fixed_income", 0)
        paras.append(
            f"{config.focus_company_short}{config.period_label}非货增量为{_fmt_yi(focus.increment)}亿，"
            f"增量排名第{focus.increment_rank}名；"
            f"其中被动权益增量和固收增量排名，分别在第{pe}名和第{fi}名。"
        )
    highlights = [
        b for b in breakdowns if b.increment_rank <= 10 and b.company != config.focus_company
    ][:2]
    for h in highlights:
        short = h.company.replace("基金", "")
        paras.append(
            f"值得关注的是{short}，非货增量{_fmt_yi(h.increment)}亿，排名第{h.increment_rank}；"
            f"本季度非货排名至第{h.rank}名。"
        )
    while len(paras) < 3:
        paras.append("")
    return {
        "section_title": f"四、{config.period_label}非货增量情况概览",
        "p1": paras[0],
        "p2": paras[1],
        "p3": paras[2],
        "footnote_passive": "注：被动权益=ETF+联接+场外普通指数+场外指数增强",
    }


def build_active_equity_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    focus = _focus(rankings, config)
    up, _ = find_notable_rank_changes(rankings, threshold=5)
    notable = [r for r in up if r.company != config.focus_company][:2]
    p1 = "主动权益Top30公司，规模与增速多数小幅波动。"
    p2 = ""
    if notable:
        parts = [
            f"{r.company.replace('基金', '')}增量{_fmt_yi(r.increment)}亿、增速{_fmt_pct(r.growth_pct)}%、"
            f"排名{_rank_change_text(r.rank_change)}"
            for r in notable
        ]
        p2 = "值得注意的是" + "；".join(parts) + "。"
    p3 = ""
    p4 = ""
    if focus:
        p4 = (
            f"{config.focus_company_short}主动权益排名，位列第{focus.rank}名，"
            f"排名较年初{_rank_change_text(focus.rank_change)}；"
            f"规模较年初{'增长' if focus.increment >= 0 else '下降'}{_fmt_yi(abs(focus.increment))}亿。"
        )
    return {
        "section_title": f"五、{config.period_label}各项业务排名情况",
        "p1": p1,
        "p2": p2,
        "p3": p3,
        "p4": p4,
        "subtitle": "1 主动权益排名",
    }


def build_etf_payload(with_link: list[CompanyRanking], no_link: list[CompanyRanking], config: ReportConfig) -> dict:
    f1 = _focus(with_link, config)
    f2 = _focus(no_link, config)
    p0 = (
        f"{config.focus_company_short}权益ETF（含联接）排名{f1.rank if f1 else '—'}名，"
        f"规模较年初{'增长' if f1 and f1.increment >= 0 else '下滑'}"
        f"{_fmt_yi(abs(f1.increment)) if f1 else '—'}亿。"
        if f1
        else f"{config.focus_company_short}权益ETF（含联接）暂无数据。"
    )
    p1 = (
        f"{config.focus_company_short}权益ETF（不含联接）排名{f2.rank if f2 else '—'}名，"
        f"规模较年初{'增长' if f2 and f2.increment >= 0 else '下滑'}"
        f"{_fmt_yi(abs(f2.increment)) if f2 else '—'}亿。"
        if f2
        else f"{config.focus_company_short}权益ETF（不含联接）暂无数据。"
    )
    return {
        "p0": p0,
        "p1": p1,
        "sub_with": "2.1权益ETF（含联接）排名",
        "sub_without": "2.2权益ETF（不含联接）排名",
    }


def build_simple_business_payload(
    rankings: list[CompanyRanking],
    config: ReportConfig,
    *,
    title: str,
    label: str,
) -> dict:
    focus = _focus(rankings, config)
    up, down = find_notable_rank_changes(rankings, threshold=3)
    notable = [r for r in (up + down) if r.company != config.focus_company][:2]
    peer = ""
    if notable:
        peer = "；".join(
            f"{r.company.replace('基金', '')}{_rank_change_text(r.rank_change)}、"
            f"增量{_fmt_yi(r.increment)}亿"
            for r in notable
        )
        peer = f"{label}方面，{peer}。"
    focus_line = ""
    if focus:
        focus_line = (
            f"{config.focus_company_short}{label}，位列第{focus.rank}名，"
            f"较年初{_rank_change_text(focus.rank_change)}；"
            f"规模较年初{'增长' if focus.increment >= 0 else '下滑'}"
            f"{_fmt_yi(abs(focus.increment))}亿，增速{_fmt_pct(focus.growth_pct)}%。"
        )
    return {
        "title": title,
        "p0": peer or focus_line or f"{label}排名见下表。",
        "p1": focus_line if peer else "",
    }


def build_all_narratives(df, config: ReportConfig) -> dict:
    cos = company_df(df)
    ind = industry_df(df)
    metrics = compute_category_summary(ind, config)

    total_r = compute_company_ranking(cos, config, "total")
    nm_r = compute_company_ranking(cos, config, "non_money")
    inc = compute_increment_breakdown(cos, config, "non_money")
    active = compute_company_ranking(cos, config, "active_equity")

    etf_key_w = "passive_equity_etf_with_link"
    etf_key_n = "passive_equity_etf_no_link"
    etf_w = compute_company_ranking(cos, config, etf_key_w) or compute_company_ranking(cos, config, "passive_equity")
    etf_n = compute_company_ranking(cos, config, etf_key_n) or compute_company_ranking(cos, config, "passive_equity")

    money = compute_company_ranking(cos, config, "money")
    from dataclasses import replace

    cfg_wide = replace(config, top_n=max(config.top_n, 50))
    fi_wide = compute_company_ranking(cos, cfg_wide, "fixed_income")
    fip_wide = compute_company_ranking(cos, cfg_wide, "fixed_income_plus")
    fof = compute_company_ranking(cos, config, "fof")

    # 固收+：Top30 表 + 若银华不在 Top30 则追加一行
    fip_table = list(fip_wide[:30])
    focus_fip = next((r for r in fip_wide if r.company == config.focus_company), None)
    if focus_fip and focus_fip.rank > 30:
        fip_table.append(focus_fip)

    fi_table = list(fi_wide[:30])
    focus_fi = next((r for r in fi_wide if r.company == config.focus_company), None)
    if focus_fi and focus_fi.rank > 30:
        fi_table.append(focus_fi)

    return {
        "overview": build_overview_payload(metrics, config),
        "total_ranking": build_total_ranking_payload(total_r, config),
        "non_money_ranking": build_non_money_payload(nm_r, config),
        "increment_overview": build_increment_payload(inc, config),
        "active_equity": build_active_equity_payload(active, config),
        "etf_dual": build_etf_payload(etf_w, etf_n, config),
        "money": build_simple_business_payload(
            money, config, title="3.货币排名情况", label="货币业务"
        ),
        "fixed_income": build_simple_business_payload(
            fi_wide, config, title="4. 固收排名情况", label=f"{config.period_label}固收"
        ),
        "fixed_income_plus": build_simple_business_payload(
            fip_wide, config, title="5. 固收+排名情况", label="固收+"
        ),
        "fof": build_simple_business_payload(
            fof, config, title="6. FOF 排名情况", label="FOF"
        ),
        "_fi_table_rankings": fi_table,
        "_fip_table_rankings": fip_table,
    }
