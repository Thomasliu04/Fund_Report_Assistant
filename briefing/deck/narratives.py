"""全册叙事 payload。

优先使用 `sheet_narratives.build_narratives_from_tables`（终表驱动）。
本模块保留 CSV 路径下的旧构建函数，供无底稿时回退。
"""

from __future__ import annotations

from statistics import median

from briefing.analytics.engine import (
    compute_category_summary,
    compute_company_ranking,
    compute_increment_breakdown,
)
from briefing.deck.data_tables import company_df, industry_df
from briefing.models import CompanyRanking, ReportConfig
from briefing.narrative.generator import (
    generate_increment_narrative,
    generate_non_money_ranking_narrative,
    generate_overview_narrative,
)
from briefing.phase1 import build_overview_payload, build_total_ranking_payload


_FI_CLASSIFICATION_NOTE = (
    "注：固收产品分类调整——原归属于固收分类的2只产品示例产品A和示例产品B，"
    "分别于24Q4和25Q1调整为固收+产品，因为分类调整对固收业务规模和排名会造成部分影响，"
    "这2只产品25Q4合计规模44亿"
)


def _fmt_yi(v: float) -> str:
    return str(round(v))


def _fmt_pct(v: float) -> str:
    return f"{round(v * 100)}"


def _rank_change_text(change: int, *, unit: str = "名") -> str:
    if change > 0:
        return f"上升{change}{unit}"
    if change < 0:
        return f"下降{abs(change)}{unit}"
    return "不变" if unit == "名" else "持平"


def _focus(rankings: list[CompanyRanking], config: ReportConfig):
    short = config.focus_company_short or config.focus_company.replace("基金", "")
    return next(
        (
            r
            for r in rankings
            if short in r.company.replace("基金", "") or config.focus_company in r.company
        ),
        None,
    )


def build_non_money_payload(
    rankings: list[CompanyRanking],
    config: ReportConfig,
    *,
    industry_growth: float | None = None,
    industry_q_growth: float | None = None,
) -> dict:
    narr = generate_non_money_ranking_narrative(rankings, config)
    focus = _focus(rankings, config)
    short = config.focus_company_short
    if not focus:
        return {
            "section_title": f"三、{config.period_label}公募基金非货规模排名",
            "p1": f"暂无{short}非货数据。",
            "p2": "",
            "p3": "",
            "footnote_passive": "",
        }
    # 案例「高于Top30公司（18%）」用行业非货增速，而非算术平均
    avg_g = industry_growth if industry_growth is not None else (
        sum(r.growth_pct for r in rankings) / max(len(rankings), 1)
    )
    avg_qg = industry_q_growth if industry_q_growth is not None else (
        sum(r.q_growth_pct for r in rankings) / max(len(rankings), 1)
    )
    q_tag = config.quarter_tag()
    vs_y = "高于" if focus.growth_pct >= avg_g else "低于"
    vs_q = "高于" if focus.q_growth_pct >= avg_qg else "低于"
    p1 = narr.paragraphs[0] if narr.paragraphs else ""
    p2 = (
        f"{short}非货{config.ytd_tag()}全年增长{_fmt_yi(focus.increment)}亿，"
        f"增速{_fmt_pct(focus.growth_pct)}%，"
        f"{vs_y}Top{config.top_n}公司（{_fmt_pct(avg_g)}%）；"
        f"{q_tag}增量{_fmt_yi(focus.q_increment)}亿，增速{_fmt_pct(focus.q_growth_pct)}%，"
        f"{vs_q}Top{config.top_n}公司（{_fmt_pct(avg_qg)}%）。"
    )
    return {
        "section_title": f"三、{config.period_label}公募基金非货规模排名",
        "p1": p1,
        "p2": p2,
        "p3": "",
        "footnote_passive": "",
    }


def build_increment_payload(breakdowns, config: ReportConfig) -> dict:
    narr = generate_increment_narrative(breakdowns, config)
    paras = list(narr.paragraphs)
    while len(paras) < 3:
        paras.append("")
    return {
        "section_title": f"四、{config.period_label}非货增量情况概览",
        "p1": paras[0],
        "p2": paras[1],
        "p3": paras[2],
        "footnote_passive": "",
    }


def build_active_equity_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    focus = _focus(rankings, config)
    short = config.focus_company_short
    ye = config.year_end_ref()
    p1 = "主动权益："
    # 案例：突出增速/增量亮点 + 300亿+公司
    star = max(rankings, key=lambda r: r.growth_pct) if rankings else None
    big = [r for r in rankings if r.increment >= 300 and (not star or r.company != star.company)]
    big = sorted(big, key=lambda r: r.increment, reverse=True)[:3]
    if star and star.growth_pct >= 1:
        hold_note = "，主要源于持营" if (star.holding_sales or 0) > max(star.new_issue, star.nav_change, 0) else ""
        p2 = (
            f"主动权益Top{config.top_n}公司中，"
            f"{star.company.replace('基金', '')}全年增长{_fmt_yi(star.increment)}亿，"
            f"增速达{_fmt_pct(star.growth_pct)}%{hold_note}"
        )
        if big:
            names = "、".join(r.company.replace("基金", "") for r in big)
            p2 += f"；{names}也实现了300亿+的规模增长"
        p2 += "。"
    else:
        p2 = f"主动权益Top{config.top_n}公司中，多数公司规模与增速小幅波动。"
    p3 = ""
    if focus:
        hold = getattr(focus, "holding_sales", 0) or 0
        if hold < -1:
            hold_txt = f"，但全年净赎回{_fmt_yi(abs(hold))}亿元，对规模影响显著"
        elif abs(hold) > 1:
            hold_txt = f"，持营{_fmt_yi(hold)}亿元"
        else:
            hold_txt = ""
        nav = getattr(focus, "nav_change", 0) or 0
        nav_txt = f"，其中净值贡献{_fmt_yi(nav)}亿元" if abs(nav) > 1 else ""
        p3 = (
            f"{short}主动权益排名第{focus.rank}名，较{ye}{_rank_change_text(focus.rank_change)}。"
            f"全年规模{'增长' if focus.increment >= 0 else '下降'}{_fmt_yi(abs(focus.increment))}亿"
            f"至{_fmt_yi(focus.aum)}亿{nav_txt}{hold_txt}。"
        )
    return {
        "section_title": f"五、{config.period_label}各项业务排名情况",
        "p1": p1,
        "p2": p2,
        "p3": p3,
        "p4": "",
        "subtitle": "",
    }


def build_etf_payload(with_link: list[CompanyRanking], no_link: list[CompanyRanking], config: ReportConfig) -> dict:
    f1 = _focus(with_link, config)
    f2 = _focus(no_link, config)
    short = config.focus_company_short
    ye = config.year_end_ref()

    def _etf_line(label: str, f: CompanyRanking | None, *, require_data: bool = False) -> str:
        if not f:
            return f"{label}：本季数据暂缺，请人工补充。" if require_data else f"{label}：{short}暂无数据。"
        # 无有效增量且排名像是回退品类时，视为数据缺失
        if require_data and abs(f.increment) < 0.5 and abs(f.growth_pct) < 0.001:
            return f"{label}：本季数据暂缺，请人工补充。"
        bits = [
            f"{label}：{short}排名{f.rank}名，较{ye}{_rank_change_text(f.rank_change)}",
            f"全年合计{'增长' if f.increment >= 0 else '下滑'}{_fmt_yi(abs(f.increment))}亿",
        ]
        if abs(f.growth_pct) > 0.001:
            bits.append(f"增速{_fmt_pct(f.growth_pct)}%")
            if f.growth_pct >= 1:
                bits.append("实现规模翻倍")
        hold = getattr(f, "holding_sales", 0) or 0
        nav = getattr(f, "nav_change", 0) or 0
        if abs(hold) > 1:
            bits.append(f"持营贡献{_fmt_yi(hold)}亿元")
        elif abs(nav) > 1:
            bits.append(f"其中净值贡献{_fmt_yi(nav)}亿元")
        return "，".join(bits) + "。"

    # 权益ETF 仅在有独立口径时输出；否则明确暂缺（避免误用被动权益回退）
    has_equity_etf = bool(no_link) and any(
        abs((r.increment or 0)) > 0.5 or abs((r.growth_pct or 0)) > 0.001 for r in no_link
    )
    return {
        "p0": _etf_line("非货ETF（含联接）", f1),
        "p1": _etf_line("权益ETF（含联接）", f2 if has_equity_etf else None, require_data=True),
        "sub_with": "",
        "sub_without": "",
    }


def build_money_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    focus = _focus(rankings, config)
    short = config.focus_company_short
    ye = config.year_end_ref()
    pq = config.prev_quarter_ref()
    avg_g = sum(r.growth_pct for r in rankings) / max(len(rankings), 1)
    p1 = ""
    if focus:
        vs = "高于" if focus.growth_pct >= avg_g else "低于"
        q_chg = _rank_change_text(focus.rank_change_q)
        if focus.rank_change_q == 0:
            q_chg = "持平"
        p1 = (
            f"{short}位列{focus.rank}名，排名较{ye}{_rank_change_text(focus.rank_change)}，"
            f"较{pq}{q_chg}。全年规模增长{_fmt_yi(focus.increment)}亿元，"
            f"增速{_fmt_pct(focus.growth_pct)}%，"
            f"{vs}货币Top{config.top_n}平均增幅（{_fmt_pct(avg_g)}%）。"
        )
    return {
        "title": "",
        "label_line": "货币：",
        "p1": p1,
    }


def build_fixed_income_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    focus = _focus(rankings, config)
    short = config.focus_company_short
    ye = config.year_end_ref()
    top = rankings[: config.top_n]
    n_up = sum(1 for r in top if r.increment > 0)
    med_inc = median([r.increment for r in top]) if top else 0
    med_g = median([r.growth_pct for r in top]) if top else 0
    p1 = (
        f"固收Top{config.top_n}公司中仅{n_up}家实现规模增长。"
        f"Top{config.top_n}公司规模增量中位数{_fmt_yi(med_inc)}亿，"
        f"增速中位数{_fmt_pct(med_g)}%。"
    )
    p2 = ""
    p3 = ""
    if focus:
        hold = getattr(focus, "holding_sales", 0) or 0
        p2 = (
            f"{short}固收{config.ytd_tag()}规模增量{_fmt_yi(focus.increment)}亿，"
            f"增速{_fmt_pct(focus.growth_pct)}%，持营增量{_fmt_yi(hold)}亿；"
            f"{short}固收位列第{focus.rank}名，较{ye}排名{_rank_change_text(focus.rank_change)}，"
            f"较Q{(int(config.previous_quarter_date[4:6]) - 1) // 3 + 1}"
            f"{_rank_change_text(focus.rank_change_q)}。"
        )
        # 季度位次
        qi = sorted(top, key=lambda r: r.q_increment, reverse=True)
        qg = sorted(top, key=lambda r: r.q_growth_pct, reverse=True)
        qi_pos = next((i for i, r in enumerate(qi, 1) if r.company == focus.company), None)
        qg_pos = next((i for i, r in enumerate(qg, 1) if r.company == focus.company), None)
        if qi_pos and qg_pos:
            p3 = (
                f"就{config.quarter_tag()}而言，{short}固收的规模增量达{_fmt_yi(focus.q_increment)}亿，"
                f"在固收Top{config.top_n}公司位列第{qi_pos}；"
                f"增速{_fmt_pct(focus.q_growth_pct)}%，在固收Top{config.top_n}公司位列第{qg_pos}。"
            )
    return {
        "title": "",
        "label_line": "固收：",
        "p1": p1,
        "p2": p2,
        "p3": p3,
        "footnote_fi": _FI_CLASSIFICATION_NOTE,
    }


def build_fixed_income_plus_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    focus = _focus(rankings, config)
    short = config.focus_company_short
    ye = config.year_end_ref().replace("底", "末")  # 较24年末
    pq = config.prev_quarter_ref()
    p1 = ""
    p2 = ""
    if focus:
        hold = getattr(focus, "holding_sales", 0) or 0
        driver = ""
        if hold > 0 and hold >= max(focus.new_issue, focus.nav_change, 0):
            driver = f"，主要源于持营（+{_fmt_yi(hold)}亿）"
        p1 = (
            f"{short}固收+规模{_fmt_yi(focus.aum)}亿，较年初增长{_fmt_yi(focus.increment)}亿，"
            f"增速高达{_fmt_pct(focus.growth_pct)}%{driver}。"
        )
        p2 = (
            f"{short}固收+排名第{focus.rank}名，较{ye}"
            f"{_rank_change_text(focus.rank_change, unit='位')}，"
            f"较{pq}{_rank_change_text(focus.rank_change_q, unit='位')}。"
        )
    return {
        "title": "",
        "label_line": "固收+:",
        "p1": p1,
        "p2": p2,
        "footnote_fi": _FI_CLASSIFICATION_NOTE,
    }


def build_fof_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    focus = _focus(rankings, config)
    short = config.focus_company_short
    ye = config.year_end_ref()
    pq = config.prev_quarter_ref()
    p1 = ""
    if focus:
        driver = ""
        if focus.new_issue > 0 and focus.new_issue >= max(focus.nav_change, focus.holding_sales, 0):
            driver = f"，主要来自新发（+{_fmt_yi(focus.new_issue)}亿）"
        elif focus.holding_sales > 0:
            driver = f"，主要来自持营（+{_fmt_yi(focus.holding_sales)}亿）"
        p1 = (
            f"{short}FOF位列第{focus.rank}名，较{ye}{_rank_change_text(focus.rank_change)}，"
            f"较{pq}{_rank_change_text(focus.rank_change_q)}。"
            f"全年规模增长{_fmt_yi(focus.increment)}亿，增速{_fmt_pct(focus.growth_pct)}%{driver}。"
        )
    return {
        "title": "",
        "label_line": "FOF：",
        "p1": p1,
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
    etf_w = compute_company_ranking(cos, config, etf_key_w)
    etf_n = compute_company_ranking(cos, config, etf_key_n)
    # 不回退到被动权益，避免权益ETF口径串数

    money = compute_company_ranking(cos, config, "money")
    from dataclasses import replace

    cfg_wide = replace(config, top_n=max(config.top_n, 50))
    fi_wide = compute_company_ranking(cos, cfg_wide, "fixed_income")
    fip_wide = compute_company_ranking(cos, cfg_wide, "fixed_income_plus")
    fof = compute_company_ranking(cos, config, "fof")

    fip_table = list(fip_wide[:30])
    focus_fip = _focus(fip_wide, config)
    if focus_fip and focus_fip.rank > 30:
        fip_table.append(focus_fip)

    fi_table = list(fi_wide[:30])
    focus_fi = _focus(fi_wide, config)
    if focus_fi and focus_fi.rank > 30:
        fi_table.append(focus_fi)

    focus_total = _focus(total_r, config)
    focus_nm = _focus(nm_r, config)
    nm_ind = next((m for m in metrics if m.category == "non_money"), None)

    return {
        "overview": build_overview_payload(
            metrics, config, focus_total=focus_total, focus_non_money=focus_nm
        ),
        "total_ranking": build_total_ranking_payload(total_r, config),
        "non_money_ranking": build_non_money_payload(
            nm_r,
            config,
            industry_growth=nm_ind.growth_pct if nm_ind else None,
            industry_q_growth=nm_ind.q_growth_pct if nm_ind else None,
        ),
        "increment_overview": build_increment_payload(inc, config),
        "active_equity": build_active_equity_payload(active, config),
        "etf_dual": build_etf_payload(etf_w, etf_n, config),
        "money": build_money_payload(money, config),
        "fixed_income": build_fixed_income_payload(fi_wide, config),
        "fixed_income_plus": build_fixed_income_plus_payload(fip_wide, config),
        "fof": build_fof_payload(fof, config),
        "_fi_table_rankings": fi_table,
        "_fip_table_rankings": fip_table,
    }
