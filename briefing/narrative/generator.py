"""叙事文本生成：规则 + 模板，对齐案例简报句式。"""

from __future__ import annotations

from briefing.analytics.engine import find_notable_rank_changes
from briefing.models import (
    CategoryMetrics,
    CompanyRanking,
    IncrementBreakdown,
    NarrativeBlock,
    ReportConfig,
)


def _fmt_aum(value: float, unit: str = "亿") -> str:
    v = round(value)
    return f"{v}{unit}"


def _fmt_pct(value: float) -> str:
    pct = round(value * 100)
    return f"{pct}%"


def _fmt_rank_change(change: int) -> str:
    if change > 0:
        return f"上升{change}名"
    if change < 0:
        return f"下降{abs(change)}名"
    return "不变"


def _date_label(date_str: str) -> str:
    """20250630 -> 25年6月30日"""
    y, m, d = date_str[:4], int(date_str[4:6]), int(date_str[6:8])
    return f"{y[2:]}年{m}月{d}日"


def _fmt_inc_amount(v: float) -> str:
    """增量口语：过万亿用万亿（1位小数）。"""
    av = abs(v)
    if av >= 10000:
        return f"{av / 10000:.1f}万亿".replace(".0万亿", "万亿")
    return f"{round(v)}亿"


def _overview_driver(m: CategoryMetrics) -> str:
    """品类增量驱动句，对齐案例简报（不含主观行情注释）。"""
    inc = m.increment
    if abs(inc) < 1:
        return ""
    new_i, nav, hold = m.new_issue, m.nav_change, m.holding_sales
    share = lambda x: round(abs(x) / abs(inc) * 100) if abs(inc) >= 1 else 0

    # 净值为主且持营大幅净赎回
    if nav > 0 and hold < -1000 and abs(nav) >= abs(new_i):
        return (
            f"其中净值增长{_fmt_inc_amount(nav)}，"
            f"但客户净赎回近{round(abs(hold))}亿元"
        )
    # 新发占增量大头
    if new_i > 0 and new_i >= max(nav, hold, 0) and share(new_i) >= 40:
        return f"新发增量{round(new_i)}亿，占增量{share(new_i)}%"
    # 净值为主
    if nav > 0 and abs(nav) >= abs(hold) and abs(nav) >= abs(new_i) and share(nav) >= 40:
        return f"其中净值增长{_fmt_inc_amount(nav)}，占增量的{share(nav)}%"
    # 持营为主
    if hold > 0 and hold >= max(new_i, nav, 0):
        return f"持营增长{'近' if hold >= 5000 else ''}{round(hold)}亿"
    if new_i > 0:
        return f"新发增量{round(new_i)}亿，占增量{share(new_i)}%"
    if abs(nav) > 1:
        return f"其中净值增长{_fmt_inc_amount(nav)}"
    if abs(hold) > 1:
        return f"持营增长{round(hold)}亿"
    return ""


def generate_overview_narrative(
    metrics: list[CategoryMetrics], config: ReportConfig
) -> NarrativeBlock:
    """第一章：公募规模整体情况"""
    paragraphs: list[str] = []
    bullets: list[str] = []

    total = next((m for m in metrics if m.category == "total"), None)
    non_money = next((m for m in metrics if m.category == "non_money"), None)

    if total and non_money:
        paragraphs.append(
            f"截至{_date_label(config.current_date)}，"
            f"公募行业总规模{_fmt_aum(total.aum_current)}（增速{_fmt_pct(total.growth_pct)}），"
            f"非货{_fmt_aum(non_money.aum_current)}（增速{_fmt_pct(non_money.growth_pct)}）。"
        )

    cats = [m for m in metrics if m.category not in ("total", "non_money")]
    # 增速较快的品类：正增长按增速排序，最多 5 条；货币可进入榜单
    positive = sorted(
        [m for m in cats if m.growth_pct > 0 and m.category != "fixed_income"],
        key=lambda x: x.growth_pct,
        reverse=True,
    )[:5]
    for i, m in enumerate(positive, 1):
        base = f"{i}、{m.label}增速{_fmt_pct(m.growth_pct)}，增量{_fmt_inc_amount(m.increment)}"
        # 货币案例通常只写增速+增量；驱动拆解留给其它品类
        if m.category == "money":
            bullets.append(f"{base}。")
            continue
        driver = _overview_driver(m)
        bullets.append(f"{base}，{driver}。" if driver else f"{base}。")

    # 固收若负增长，单独作为 highlight（案例放在 bullet 后）
    fi = next((m for m in metrics if m.category == "fixed_income"), None)
    if fi and fi.increment < 0:
        paragraphs.append(
            f"{config.ytd_tag()}全年，固收规模减少{round(abs(fi.increment))}亿，"
            f"增速{_fmt_pct(fi.growth_pct)}。"
        )

    return NarrativeBlock(section_id="overview", paragraphs=paragraphs, bullets=bullets)


def generate_total_ranking_narrative(
    rankings: list[CompanyRanking], config: ReportConfig
) -> NarrativeBlock:
    """第二章：总规模排名"""
    paragraphs: list[str] = []
    up, down = find_notable_rank_changes(rankings, threshold=2)

    if up:
        # 案例句式：景顺、中欧各上升3名，富国上升2名
        by_delta: dict[int, list[str]] = {}
        for r in up[:5]:
            short = r.company.replace("基金", "")
            if short.startswith("景顺"):
                short = "景顺"
            by_delta.setdefault(r.rank_change, []).append(short)
        parts = []
        for delta, names in sorted(by_delta.items(), key=lambda kv: -kv[0]):
            if len(names) >= 2:
                parts.append(f"{'、'.join(names)}各上升{delta}名")
            else:
                parts.append(f"{names[0]}上升{delta}名")
        if parts:
            paragraphs.append(
                f"总规模Top{config.top_n}公司，{'，'.join(parts)}，其余公司位次变化并不明显。"
            )

    focus = next((r for r in rankings if r.company == config.focus_company), None)
    if not focus:
        focus = next(
            (
                r
                for r in rankings
                if config.focus_company_short in r.company.replace("基金", "")
            ),
            None,
        )
    if focus:
        ye = config.year_end_ref()
        pq = config.prev_quarter_ref()
        paragraphs.append(
            f"{config.focus_company_short}总规模排名{focus.rank}，"
            f"较{ye}排名{_fmt_rank_change(focus.rank_change)}，"
            f"较{pq}季度{_fmt_rank_change(focus.rank_change_q)}。"
        )
        g_rank = sorted(rankings, key=lambda x: x.growth_pct, reverse=True)
        qg_rank = sorted(rankings, key=lambda x: x.q_growth_pct, reverse=True)
        g_pos = next(i for i, r in enumerate(g_rank, 1) if r.company == focus.company)
        qg_pos = 1
        last_pct = None
        for i, r in enumerate(qg_rank):
            if last_pct is None or abs(r.q_growth_pct - last_pct) > 1e-9:
                qg_pos = i + 1
                last_pct = r.q_growth_pct
            if r.company == focus.company:
                break
        q_line = (
            f"{config.quarter_tag()}单季度增速达{_fmt_pct(focus.q_growth_pct)}，"
            f"在Top{config.top_n}公司中位列第{qg_pos}名"
        )
        if qg_pos == 2 and qg_rank and qg_rank[0].company != focus.company:
            leader = qg_rank[0]
            q_line += (
                f"，仅次于{leader.company.replace('基金', '')}"
                f"（{_fmt_pct(leader.q_growth_pct)}）"
            )
        paragraphs.append(
            f"{config.focus_company_short}总规模全年增速为{_fmt_pct(focus.growth_pct)}，"
            f"在Top{config.top_n}公司中位列第{g_pos}；{q_line}。"
        )

    return NarrativeBlock(section_id="total_ranking", paragraphs=paragraphs)


def generate_non_money_ranking_narrative(
    rankings: list[CompanyRanking], config: ReportConfig
) -> NarrativeBlock:
    """第三章：非货规模排名"""
    paragraphs: list[str] = []
    focus = next((r for r in rankings if r.company == config.focus_company), None)
    if not focus:
        focus = next(
            (
                r
                for r in rankings
                if config.focus_company_short in r.company.replace("基金", "")
            ),
            None,
        )

    if focus:
        ye = config.year_end_ref()
        pq = config.prev_quarter_ref()
        lift = ""
        if focus.rank_change >= 2 and focus.rank <= 20:
            lift = "，非货排名显著提升，重回Top20行列"
        elif focus.rank_change >= 2:
            lift = "，非货排名显著提升"
        paragraphs.append(
            f"非货规模方面，{config.focus_company_short}位列第{focus.rank}位，"
            f"相较{ye}{_fmt_rank_change(focus.rank_change)}，"
            f"较{pq}{_fmt_rank_change(focus.rank_change_q)}{lift}。"
        )

    return NarrativeBlock(section_id="non_money_ranking", paragraphs=paragraphs)


def generate_increment_narrative(
    breakdowns: list[IncrementBreakdown], config: ReportConfig
) -> NarrativeBlock:
    """第四章：非货增量概览"""
    paragraphs: list[str] = []
    focus = next((b for b in breakdowns if b.company == config.focus_company), None)
    if not focus:
        focus = next(
            (
                b
                for b in breakdowns
                if config.focus_company_short in b.company.replace("基金", "")
            ),
            None,
        )

    if focus:
        cat_names = {
            "active_equity": "主动权益",
            "passive_equity": "被动权益",
            "fixed_income_plus": "固收+",
            "fixed_income": "固收",
            "fof": "FOF",
        }
        # 案例句式优先：固收+、被动权益、FOF；不足再补其它靠前品类
        preferred = ["fixed_income_plus", "passive_equity", "fof", "fixed_income", "active_equity"]
        ranked: list[tuple[str, int]] = []
        for k in preferred:
            v = focus.category_increment_ranks.get(k)
            if v and v <= 20:
                ranked.append((k, v))
            if len(ranked) >= 3:
                break
        # 保持 preferred 顺序（不为按名次重排，以贴合案例列举习惯）
        ranked = ranked[:3]

        q_ordered = sorted(breakdowns, key=lambda b: b.q_increment, reverse=True)
        q_rank = next(
            (i for i, b in enumerate(q_ordered, 1) if b.company == focus.company),
            None,
        )
        q_month = int(config.current_date[4:6])
        q_num = (q_month - 1) // 3 + 1
        if q_rank:
            paragraphs.append(
                f"{config.focus_company_short}{config.ytd_tag()}非货增量为{_fmt_aum(focus.increment)}，"
                f"在全行业排名第{focus.increment_rank}名；"
                f"{q_num}季度，非货增量为{_fmt_aum(focus.q_increment)}，"
                f"在全行业排名第{q_rank}名。"
            )
        else:
            paragraphs.append(
                f"{config.focus_company_short}{config.ytd_tag()}非货增量为{_fmt_aum(focus.increment)}，"
                f"在全行业排名第{focus.increment_rank}名。"
            )
        if ranked:
            names = "、".join(cat_names.get(k, k) for k, _ in ranked)
            ranks = "、".join(f"第{v}名" for _, v in ranked)
            if len(ranked) == 3:
                ranks = f"第{ranked[0][1]}名、第{ranked[1][1]}名和第{ranked[2][1]}名"
            paragraphs.append(f"{names}的全年增量排名靠前，分别为{ranks}。")

    return NarrativeBlock(section_id="increment_overview", paragraphs=paragraphs)


def generate_business_ranking_narrative(
    rankings: list[CompanyRanking],
    config: ReportConfig,
    category_label: str,
) -> NarrativeBlock:
    """分项业务排名叙事（通用兜底）"""
    paragraphs: list[str] = []
    focus = next((r for r in rankings if r.company == config.focus_company), None)
    if not focus:
        focus = next(
            (
                r
                for r in rankings
                if config.focus_company_short in r.company.replace("基金", "")
            ),
            None,
        )

    if focus:
        ye = config.year_end_ref()
        paragraphs.append(
            f"{config.focus_company_short}{category_label}，位列第{focus.rank}名，"
            f"排名较{ye}{_fmt_rank_change(focus.rank_change)}；"
            f"规模较年初{'增长' if focus.increment >= 0 else '下降'}"
            f"{_fmt_aum(abs(focus.increment))}，"
            f"{'主要由净值增长带动' if focus.nav_change > focus.new_issue else ''}。"
        )

    return NarrativeBlock(
        section_id=f"business_{category_label}",
        paragraphs=paragraphs,
    )
