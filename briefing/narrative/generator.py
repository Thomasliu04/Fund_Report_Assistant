"""叙事文本生成：规则 + 模板，对齐 PDF 简报风格"""

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

    # 按增速排序品类（排除合计、非货）
    cats = [m for m in metrics if m.category not in ("total", "non_money")]
    cats_sorted = sorted(cats, key=lambda x: x.growth_pct, reverse=True)

    positive = [m for m in cats if m.increment > 0]
    if len(positive) == len(cats):
        period_word = "上半年" if config.period_type.value == "half_year" else "本季度"
        bullets.append(f"{period_word}，所有品类规模增速均为正：")

    for i, m in enumerate(cats_sorted[:4], 1):
        detail_parts = [f"规模增长{_fmt_aum(m.increment)}，增速{_fmt_pct(m.growth_pct)}"]
        if abs(m.new_issue) > abs(m.holding_sales) and m.new_issue > 0:
            detail_parts.append(f"其中新发贡献{_fmt_aum(m.new_issue)}，是增量大头")
        elif m.holding_sales > 0 and m.holding_sales > m.new_issue:
            detail_parts.append(
                f"增量{_fmt_aum(m.increment)}中，其中有{_fmt_aum(m.holding_sales)}来自于持营贡献"
            )
        prefix = f"增速最快的品类是{m.label}" if i == 1 else m.label
        bullets.append(f"{i}、{prefix}，{'；'.join(detail_parts)}。")

    return NarrativeBlock(section_id="overview", paragraphs=paragraphs, bullets=bullets)


def generate_total_ranking_narrative(
    rankings: list[CompanyRanking], config: ReportConfig
) -> NarrativeBlock:
    """第二章：总规模排名"""
    paragraphs: list[str] = []
    up, down = find_notable_rank_changes(rankings, threshold=3)

    if up or down:
        parts = []
        for r in up[:2]:
            parts.append(f"{r.company.replace('基金', '')}{_fmt_rank_change(r.rank_change)}")
        for r in down[:2]:
            parts.append(f"{r.company.replace('基金', '')}{_fmt_rank_change(r.rank_change)}")
        if parts:
            paragraphs.append(
                f"总规模Top{config.top_n}公司，{'、'.join(parts)}，其他公司位次变化不大。"
            )

    focus = next((r for r in rankings if r.company == config.focus_company), None)
    if focus:
        paragraphs.append(
            f"{config.focus_company_short}总规模排名{focus.rank}，"
            f"较年初和{config.period_label}上期末排名{_fmt_rank_change(focus.rank_change_q) if focus.rank_change_q else '不变'}。"
        )
        paragraphs.append(
            f"{config.focus_company_short}总规模较年初增长{_fmt_aum(focus.increment)}，"
            f"增速为{_fmt_pct(focus.growth_pct)}；"
            f"Top{config.top_n}公司中，"
            f"有{sum(1 for r in rankings if r.increment > 0) * 100 // len(rankings)}%的公司增量均为正。"
        )

    return NarrativeBlock(section_id="total_ranking", paragraphs=paragraphs)


def generate_non_money_ranking_narrative(
    rankings: list[CompanyRanking], config: ReportConfig
) -> NarrativeBlock:
    """第三章：非货规模排名"""
    paragraphs: list[str] = []
    focus = next((r for r in rankings if r.company == config.focus_company), None)

    if focus:
        inc_rank = sorted(rankings, key=lambda x: x.increment, reverse=True)
        inc_rank_pos = next(i + 1 for i, r in enumerate(inc_rank) if r.company == config.focus_company)
        growth_rank = sorted(rankings, key=lambda x: x.growth_pct, reverse=True)
        growth_rank_pos = next(
            i + 1 for i, r in enumerate(growth_rank) if r.company == config.focus_company
        )

        paragraphs.append(
            f"非货规模方面，{config.focus_company_short}位列第{focus.rank}位，"
            f"相较年初{_fmt_rank_change(focus.rank_change)}，"
            f"较上季度{_fmt_rank_change(focus.rank_change_q)}。"
        )
        paragraphs.append(
            f"{config.focus_company_short}非货，{config.period_label}增量{_fmt_aum(focus.increment)}，"
            f"增速{_fmt_pct(focus.growth_pct)}，"
            f"增量和增速排名在Top{config.top_n}中分别是第{inc_rank_pos}名和第{growth_rank_pos}名；"
            f"非货Top{config.top_n}公司作为一个整体，{config.period_label}增速为"
            f"{_fmt_pct(sum(r.growth_pct for r in rankings) / len(rankings))}。"
        )

    return NarrativeBlock(section_id="non_money_ranking", paragraphs=paragraphs)


def generate_increment_narrative(
    breakdowns: list[IncrementBreakdown], config: ReportConfig
) -> NarrativeBlock:
    """第四章：非货增量概览"""
    paragraphs: list[str] = []
    focus = next((b for b in breakdowns if b.company == config.focus_company), None)

    if focus:
        pe_rank = focus.category_increment_ranks.get("passive_equity", 0)
        fi_rank = focus.category_increment_ranks.get("fixed_income", 0)
        paragraphs.append(
            f"{config.focus_company_short}{config.period_label}非货增量为{_fmt_aum(focus.increment)}，"
            f"在全行业第{focus.increment_rank}名；"
            f"其中被动权益增量和固收增量排名，分别在第{pe_rank}名和第{fi_rank}名。"
        )

    # 竞品亮点：增量排名前10中非银华的公司
    highlights = [
        b for b in breakdowns if b.increment_rank <= 10 and b.company != config.focus_company
    ][:2]
    for h in highlights:
        parts = []
        for cat, label in [
            ("fixed_income", "固收"),
            ("fixed_income_plus", "固收+"),
            ("active_equity", "主动权益"),
        ]:
            r = h.category_increment_ranks.get(cat)
            if r and r <= 10:
                parts.append(f"{label}增量位列行业第{r}名")
        if parts:
            short = h.company.replace("基金", "")
            paragraphs.append(
                f"值得关注的是{short}，非货增量{_fmt_aum(h.increment)}，"
                f"排名第{h.increment_rank}：{'、'.join(parts)}。"
                f"本季度非货排名{'提升' if h.rank_change > 0 else '变化'}"
                f"至第{h.rank}名；"
            )

    return NarrativeBlock(section_id="increment_overview", paragraphs=paragraphs)


def generate_business_ranking_narrative(
    rankings: list[CompanyRanking],
    config: ReportConfig,
    category_label: str,
) -> NarrativeBlock:
    """分项业务排名叙事"""
    paragraphs: list[str] = []
    focus = next((r for r in rankings if r.company == config.focus_company), None)

    if focus:
        paragraphs.append(
            f"{config.focus_company_short}{category_label}，位列第{focus.rank}名，"
            f"排名较年初{_fmt_rank_change(focus.rank_change)}；"
            f"规模较年初{'增长' if focus.increment >= 0 else '下降'}"
            f"{_fmt_aum(abs(focus.increment))}，"
            f"{'主要由净值增长带动' if focus.nav_change > focus.new_issue else ''}。"
        )

    # 行业异常亮点：增速超过50%或排名变化超过5
    for r in rankings:
        if r.company == config.focus_company:
            continue
        if r.rank_change >= 5 or r.growth_pct > 0.5:
            short = r.company.replace("基金", "")
            paragraphs.append(
                f"值得注意的是{short}，"
                f"{'排名上升' + str(r.rank_change) + '名' if r.rank_change >= 5 else ''}"
                f"规模增量{_fmt_aum(r.increment)}，增速{_fmt_pct(r.growth_pct)}。"
            )
            break

    return NarrativeBlock(
        section_id=f"business_{category_label}",
        paragraphs=paragraphs,
    )
