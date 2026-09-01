"""从终表事实生成各页 PPT 文字 payload（句式对齐 25Q4 案例；数字与表一致）。"""

from __future__ import annotations

from statistics import median
from typing import Any

from briefing.deck.table_facts import (
    CATEGORY_LABEL_TO_KEY,
    FactRow,
    TableBook,
    find_focus,
    load_table_book_from_final_sheets,
    load_table_book_from_xlsx,
    require_focus,
)
from briefing.models import ReportConfig
from briefing.period_profile import PeriodMode, PeriodProfile, resolve_period_profile

PAGE_INSUFFICIENT = "【本页数据不足，请人工补充】"


def _fmt_yi(v: float | None) -> str:
    if v is None:
        return "—"
    return str(round(v))


def _fmt_wan_yi(aum_yi: float) -> str:
    return f"{aum_yi / 10000:.1f}"


def _fmt_pct(v: float | None) -> str:
    if v is None:
        return "—"
    return f"{round(v * 100)}"


def _fmt_inc_amount(v: float) -> str:
    av = abs(v)
    if av >= 10000:
        return f"{av / 10000:.1f}万亿".replace(".0万亿", "万亿")
    return f"{round(v)}亿"


def _rank_change_text(change: int | None, *, unit: str = "名") -> str:
    if change is None:
        return "—"
    if change > 0:
        return f"上升{change}{unit}"
    if change < 0:
        return f"下降{abs(change)}{unit}"
    return "不变" if unit == "名" else "持平"


def _rank_vs_clause(
    change: int | None,
    ref: str,
    *,
    unit: str = "名",
    prefix: str = "较",
    rank_word: str = "",
) -> str:
    """有对照变化才拼接；缺数据时省略，避免「较25Q4—」。"""
    if change is None:
        return ""
    mid = f"排名{rank_word}" if rank_word else ""
    return f"{prefix}{ref}{mid}{_rank_change_text(change, unit=unit)}"


def _dual_rank_anchor(config: ReportConfig) -> bool:
    """是否需要同时写「较年底 / 较上季」。Q1 两锚点相同则只写年底。"""
    if config.previous_quarter_date == config.year_start_date:
        return False
    return _profile(config).dual_metrics


def _profile(config: ReportConfig) -> PeriodProfile:
    """优先尊重 yaml/网页写入的 column_labels，再补全 mode/lead/span。"""
    base = resolve_period_profile(
        config.period_label,
        config.period_type.value if hasattr(config.period_type, "value") else str(config.period_type),
        config.current_date,
    )
    ytd = (config.ytd_column_label or "").strip() or base.ytd_tag
    q = (config.quarter_column_label or "").strip() or base.quarter_tag
    if ytd == base.ytd_tag and q == base.quarter_tag:
        return base
    # 标签被手改时仍保留 mode/lead，只覆盖展示标签
    return PeriodProfile(
        mode=base.mode,
        period_type=base.period_type,
        period_label=base.period_label,
        ytd_tag=ytd,
        quarter_tag=q,
        span=base.span,
        ytd_phrase=ytd if base.mode != PeriodMode.PERIOD_LEAD or "年" in ytd or "H" in ytd.upper() else base.ytd_phrase,
        dual_metrics=base.dual_metrics,
        lead=base.lead,
    )


def _period_span(config: ReportConfig) -> str:
    return _profile(config).span


def _growth_intro(config: ReportConfig) -> str:
    y = config.current_date[2:4]
    p = _profile(config)
    if p.mode == PeriodMode.SINGLE_QUARTER:
        return f"{y}年{p.quarter_tag}增速为正的品类："
    if p.mode == PeriodMode.QUARTER_LEAD:
        return f"{y}年{p.ytd_phrase}规模增速较快的品类（以{p.quarter_tag}为主）："
    if p.span in {"上半年", "下半年"}:
        return f"{y}年{p.span}规模增速较快的品类："
    if p.span == "全年" or (p.ytd_tag.endswith("年") and "H" not in p.ytd_tag.upper()):
        return f"{y}年规模增速较快的品类："
    return f"{y}年{p.span}规模增速较快的品类："


def _ytd_phrase(config: ReportConfig) -> str:
    """增速/增量前的期别用语。"""
    return _profile(config).ytd_phrase


def _growth_label(config: ReportConfig) -> str:
    """开篇「行业增速」前的口径词：Q1 用季度增速，其余可空（直接写增速）。"""
    p = _profile(config)
    if p.mode == PeriodMode.SINGLE_QUARTER:
        return "季度"
    if p.mode == PeriodMode.QUARTER_LEAD:
        return "YTD"
    return ""


def _short_co(name: str) -> str:
    s = name.replace("基金", "")
    aliases = (
        ("华泰柏瑞", "华柏"),
        ("交银施罗德", "交银"),
        ("景顺长城", "景顺"),
        ("兴证全球", "兴证全球"),
    )
    for full, short in aliases:
        if full in s:
            return short
    if s.startswith("景顺"):
        return "景顺"
    return s


def _active_peer_sentence(top: list[FactRow], config: ReportConfig) -> str:
    """主动权益同行概括：只陈述谁增谁减、增幅/增量，不做观点判断。"""
    span = _period_span(config)
    n = config.top_n
    star = max(top, key=lambda r: r.growth_pct if r.growth_pct is not None else -999.0)

    def _hold_note(r: FactRow) -> str:
        hold = r.holding_sales or 0
        if hold > max(r.new_issue or 0, r.nav_change or 0, 0) and hold > 1:
            return f"，持营{_fmt_yi(hold)}亿"
        return ""

    # 翻倍级增速：明星 + 大额增量同行（与 25Q3/Q4 句式一致）
    if star and (star.growth_pct or 0) >= 1.0:
        big = [
            r
            for r in top
            if (r.increment or 0) >= 300 and r.name != star.name
        ]
        big = sorted(big, key=lambda r: r.increment or 0, reverse=True)[:3]
        p2 = (
            f"主动权益Top{n}公司中，"
            f"{_short_co(star.name)}{span}增长{_fmt_yi(star.increment)}亿，"
            f"增速达{_fmt_pct(star.growth_pct)}%{_hold_note(star)}"
        )
        if big:
            p2 += f"；{'、'.join(_short_co(r.name) for r in big)}也实现了300亿+的规模增长"
        return p2 + "。"

    neg_n = sum(1 for r in top if (r.increment or 0) < 0)
    pos_n = sum(1 for r in top if (r.increment or 0) > 0)
    high = sorted(
        [r for r in top if (r.growth_pct or 0) >= 0.25],
        key=lambda r: r.increment or 0,
        reverse=True,
    )
    down = sorted(
        [r for r in top if (r.growth_pct or 0) <= -0.10],
        key=lambda r: r.growth_pct or 0,
    )

    # 分化季度：家数 + 高增名单（增速/增量）+ 降幅超10%名单
    if neg_n >= max(8, n // 3) and (len(high) >= 2 or len(down) >= 2):
        bits = [f"主动权益Top{n}公司中，{neg_n}家规模下滑、{pos_n}家增长"]
        if high:
            cohort = high[:4]
            detail = "；".join(
                f"{_short_co(r.name)}增长{_fmt_yi(r.increment)}亿、增速{_fmt_pct(r.growth_pct)}%"
                f"{_hold_note(r)}"
                for r in cohort
            )
            bits.append(f"。{detail}")
            rise = sorted(
                [r for r in cohort if (r.rank_change or 0) >= 2],
                key=lambda r: r.increment or 0,
                reverse=True,
            )[:2]
            if rise:
                ye = config.year_end_ref()
                rise_txt = "；".join(
                    f"{_short_co(r.name)}排名第{r.rank}名、较{ye}{_rank_change_text(r.rank_change)}"
                    for r in rise
                )
                bits.append(f"。{rise_txt}")
        if down:
            dnames = "、".join(_short_co(r.name) for r in down[:3])
            dpcts = "、".join(f"{_fmt_pct(abs(r.growth_pct or 0))}%" for r in down[:3])
            bits.append(f"。{dnames}规模降幅分别为{dpcts}")
        return "".join(bits) + "。"

    # 中等增速明星
    if star and (star.growth_pct or 0) >= 0.20:
        big = [
            r
            for r in top
            if (r.increment or 0) >= 100 and r.name != star.name
        ]
        big = sorted(big, key=lambda r: r.increment or 0, reverse=True)[:3]
        p2 = (
            f"主动权益Top{n}公司中，"
            f"{_short_co(star.name)}{span}增长{_fmt_yi(star.increment)}亿，"
            f"增速达{_fmt_pct(star.growth_pct)}%{_hold_note(star)}"
        )
        if big:
            p2 += (
                f"；{'、'.join(_short_co(r.name) for r in big)}"
                f"分别增长{'、'.join(_fmt_yi(r.increment) for r in big)}亿"
            )
        return p2 + "。"

    if neg_n >= n // 2:
        return f"主动权益Top{n}公司中，{neg_n}家规模下滑、{pos_n}家增长。"
    return f"主动权益Top{n}公司中，{pos_n}家增长、{neg_n}家下滑。"


def _active_focus_attribution(focus: FactRow, span: str) -> str:
    """示例主动权益规模拆解：按新发/净值/持营陈述数字，不做因果判断。"""
    new_i = focus.new_issue or 0.0
    nav = focus.nav_change or 0.0
    hold = focus.holding_sales or 0.0
    parts: list[str] = []
    if abs(new_i) > 1:
        parts.append(f"新发{_fmt_yi(new_i)}亿")
    if abs(nav) > 1:
        parts.append(f"净值贡献{_fmt_yi(nav)}亿")
    if abs(hold) > 1:
        if hold < 0:
            parts.append(f"{span}净赎回{_fmt_yi(abs(hold))}亿")
        else:
            parts.append(f"持营{_fmt_yi(hold)}亿")
    if not parts:
        return ""
    return "，其中" + "，".join(parts)


def _overview_driver(r: FactRow) -> str:
    """品类增速拆解：只陈述新发/净值/持营数字，不做因果判断。"""
    inc = r.increment or 0.0
    if abs(inc) < 1:
        return ""
    new_i = r.new_issue or 0.0
    nav = r.nav_change or 0.0
    hold = r.holding_sales or 0.0
    parts: list[str] = []
    # 分量绝对值明显大于增量时（正负对冲），三项一并写出，避免只报净值误导
    if abs(new_i) > 1 and abs(nav) > 1 and (abs(new_i) + abs(nav) > abs(inc) * 1.5):
        parts.append(f"新发{_fmt_yi(new_i)}亿")
        parts.append(f"净值贡献{_fmt_yi(nav)}亿")
        if abs(hold) > 1:
            parts.append(f"持营{_fmt_yi(hold)}亿")
        return "，".join(parts)

    def share(x: float) -> int:
        return round(abs(x) / abs(inc) * 100) if abs(inc) >= 1 else 0

    sh_new = share(new_i)
    if new_i > 0 and new_i >= max(nav, hold, 0) and sh_new >= 40:
        return f"新发增量{round(new_i)}亿，占增量{min(sh_new, 100)}%"
    sh_nav = share(nav)
    if nav > 0 and abs(nav) >= abs(hold) and abs(nav) >= abs(new_i) and 40 <= sh_nav <= 100:
        return f"净值增长{_fmt_inc_amount(nav)}，占增量的{sh_nav}%"
    if hold > 0 and hold >= max(new_i, nav, 0):
        return f"持营增长{'近' if hold >= 5000 else ''}{round(hold)}亿"
    if new_i > 0 and sh_new <= 100:
        return f"新发增量{round(new_i)}亿，占增量{sh_new}%"
    if abs(nav) > 1:
        return f"净值贡献{_fmt_yi(nav)}亿"
    if abs(hold) > 1:
        return f"持营{_fmt_yi(hold)}亿"
    return ""


def _build_overview(book: TableBook, config: ReportConfig) -> dict:
    cat = book.get("category")
    total_sh = book.get("total")
    nm_sh = book.get("non_money")
    short = config.focus_company_short
    if not cat:
        return _fail_overview(config)

    by_name = {r.name: r for r in cat.rows}
    total = by_name.get("合计")
    non_money = by_name.get("非货")
    focus_t = find_focus(total_sh.rows, config.focus_company, short) if total_sh else None
    focus_n = find_focus(nm_sh.rows, config.focus_company, short) if nm_sh else None

    y, m, d = config.current_date[2:4], int(config.current_date[4:6]), int(config.current_date[6:8])
    opening = PAGE_INSUFFICIENT
    g_lab = _growth_label(config)
    g_prefix = f"{g_lab}增速" if g_lab else "增速"
    if total and non_money and total.aum is not None and non_money.aum is not None:
        # 开篇以行业数字为准；示例对比放后续业务页，避免与模版口径混写
        opening = (
            f"截至{y}年{m}月{d}日，公募行业总规模{_fmt_wan_yi(total.aum)}万亿"
            f"（{g_prefix}{_fmt_pct(total.growth_pct)}%），"
            f"非货{_fmt_wan_yi(non_money.aum)}万亿"
            f"（增速{_fmt_pct(non_money.growth_pct)}%）。"
        )
        if focus_t is None and total_sh:
            require_focus(book, "total", config.focus_company, short)
        if focus_n is None and nm_sh:
            require_focus(book, "non_money", config.focus_company, short)

    cats = [
        r
        for r in cat.rows
        if r.name not in {"合计", "非货"} and r.name in CATEGORY_LABEL_TO_KEY
    ]
    positive = sorted(
        [r for r in cats if (r.growth_pct or 0) > 0 and r.name != "固收"],
        key=lambda x: x.growth_pct or 0,
        reverse=True,
    )[:5]
    bullets = [""] * 5
    for i, r in enumerate(positive):
        base = f"{i + 1}、{r.name}增速{_fmt_pct(r.growth_pct)}%，增量{_fmt_inc_amount(r.increment or 0)}"
        driver = _overview_driver(r)
        bullets[i] = f"{base}，{driver}。" if driver else f"{base}。"

    # 显著收缩品类：被动权益 / 固收等，只报规模与增速
    highlight_parts: list[str] = []
    for name in ("被动权益", "固收"):
        row = by_name.get(name)
        if row and (row.increment or 0) < -100:
            highlight_parts.append(
                f"{name}规模减少{round(abs(row.increment or 0))}亿，增速{_fmt_pct(row.growth_pct)}%"
            )
    highlight = ""
    if highlight_parts:
        highlight = f"{_ytd_phrase(config)}，" + "；".join(highlight_parts) + "。"

    title = config.title
    if " " not in title and title.endswith("简报"):
        title = f"{config.period_label} 公募行业数据简报"

    return {
        "title": title,
        "data_source_note": config.data_source_note,
        "section_title": f"一、{config.period_label}公募规模整体情况",
        "opening": opening,
        "all_positive_intro": _growth_intro(config),
        "bullet_1": bullets[0],
        "bullet_2": bullets[1],
        "bullet_3": bullets[2],
        "bullet_4": bullets[3],
        "bullet_5": bullets[4],
        "highlight_1": highlight,
        "highlight_2": "",
        "footnote": "",
    }


def _fail_overview(config: ReportConfig) -> dict:
    return {
        "title": f"{config.period_label} 公募行业数据简报",
        "data_source_note": config.data_source_note,
        "section_title": f"一、{config.period_label}公募规模整体情况",
        "opening": PAGE_INSUFFICIENT,
        "all_positive_intro": "",
        "bullet_1": "",
        "bullet_2": "",
        "bullet_3": "",
        "bullet_4": "",
        "bullet_5": "",
        "highlight_1": "",
        "highlight_2": "",
        "footnote": "",
    }


def _build_total(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("total")
    title = f"二、{config.period_label}公募基金总规模（含货币）排名"
    empty = {
        "section_title": title,
        "peer_moves": PAGE_INSUFFICIENT,
        "focus_rank": "",
        "focus_growth": "",
        "footnote_passive": "",
    }
    if not sh or not sh.rows:
        return empty

    top = sh.rows[: config.top_n]
    up = [r for r in top if (r.rank_change or 0) >= 2]
    down = [r for r in top if (r.rank_change or 0) <= -2]
    peer = f"总规模Top{config.top_n}公司，位次变化并不明显。"
    parts: list[str] = []
    if up:
        by_delta: dict[int, list[str]] = {}
        for r in sorted(up, key=lambda x: -(x.rank_change or 0))[:8]:
            by_delta.setdefault(r.rank_change or 0, []).append(_short_co(r.name))
        for delta, names in sorted(by_delta.items(), key=lambda kv: -kv[0]):
            names = names[:5]
            if len(names) >= 2:
                parts.append(f"{'、'.join(names)}各上升{delta}名")
            else:
                parts.append(f"{names[0]}上升{delta}名")
    if down:
        by_delta_d: dict[int, list[str]] = {}
        for r in sorted(down, key=lambda x: x.rank_change or 0)[:6]:
            by_delta_d.setdefault(abs(r.rank_change or 0), []).append(_short_co(r.name))
        for delta, names in sorted(by_delta_d.items(), key=lambda kv: -kv[0]):
            names = names[:4]
            if len(names) >= 2:
                parts.append(f"{'、'.join(names)}各下降{delta}名")
            else:
                parts.append(f"{names[0]}下降{delta}名")
    if parts:
        peer = f"总规模Top{config.top_n}公司，{'，'.join(parts)}，其余公司位次变化并不明显。"

    focus = require_focus(book, "total", config.focus_company, config.focus_company_short)
    if not focus:
        return {**empty, "peer_moves": peer, "focus_rank": PAGE_INSUFFICIENT}

    ye, pq = config.year_end_ref(), config.prev_quarter_ref()
    rank_bits = [f"{config.focus_company_short}总规模排名{focus.rank}"]
    ye_clause = _rank_vs_clause(focus.rank_change, ye, rank_word="")
    if ye_clause:
        # 「较25年底排名不变」
        rank_bits.append(
            f"较{ye}排名{_rank_change_text(focus.rank_change)}"
        )
    if _dual_rank_anchor(config):
        q_clause = _rank_vs_clause(focus.rank_change_q, pq)
        if q_clause:
            rank_bits.append(f"较{pq}季度{_rank_change_text(focus.rank_change_q)}")
    focus_rank = "，".join(rank_bits) + "。"

    g_rank = sorted(top, key=lambda x: x.growth_pct or -999, reverse=True)
    qg_rank = sorted(top, key=lambda x: x.q_growth_pct or -999, reverse=True)
    g_pos = next((i for i, r in enumerate(g_rank, 1) if r.name == focus.name), None)
    qg_pos = next((i for i, r in enumerate(qg_rank, 1) if r.name == focus.name), None)
    p = _profile(config)
    span = p.span
    q_line = (
        f"{config.quarter_tag()}单季度增速达{_fmt_pct(focus.q_growth_pct)}%，"
        f"在Top{config.top_n}公司中位列第{qg_pos}名"
        if qg_pos
        else f"{config.quarter_tag()}单季度增速达{_fmt_pct(focus.q_growth_pct)}%"
    )
    if qg_pos == 2 and qg_rank and qg_rank[0].name != focus.name:
        leader = qg_rank[0]
        q_line += f"，仅次于{_short_co(leader.name)}（{_fmt_pct(leader.q_growth_pct)}%）"

    if not p.dual_metrics:
        # Q1：增量 + 增速位次
        focus_growth = (
            f"{config.focus_company_short}总规模{config.quarter_tag()}增{_fmt_yi(focus.increment)}亿，"
            f"增速{_fmt_pct(focus.growth_pct)}%，"
            f"在Top{config.top_n}公司中位列第{g_pos}。"
        )
    elif p.lead == "quarter":
        ytd_line = (
            f"{_ytd_phrase(config)}增速为{_fmt_pct(focus.growth_pct)}%，"
            f"在Top{config.top_n}公司中位列第{g_pos}"
        )
        focus_growth = (
            f"{config.focus_company_short}总规模{q_line}；{ytd_line}。"
        )
    else:
        focus_growth = (
            f"{config.focus_company_short}总规模{span}增速为{_fmt_pct(focus.growth_pct)}%，"
            f"在Top{config.top_n}公司中位列第{g_pos}；{q_line}。"
        )
    return {
        "section_title": title,
        "peer_moves": peer,
        "focus_rank": focus_rank,
        "focus_growth": focus_growth,
        "footnote_passive": "",
    }


def _build_non_money(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("non_money")
    cat = book.get("category")
    title = f"三、{config.period_label}公募基金非货规模排名"
    foot = ""
    if not sh:
        return {"section_title": title, "p1": PAGE_INSUFFICIENT, "p2": "", "p3": "", "footnote_passive": foot}

    focus = require_focus(book, "non_money", config.focus_company, config.focus_company_short)
    if not focus:
        return {"section_title": title, "p1": PAGE_INSUFFICIENT, "p2": "", "p3": "", "footnote_passive": foot}

    ye, pq = config.year_end_ref(), config.prev_quarter_ref()
    short = config.focus_company_short
    rank_bits = [f"非货规模方面，{short}位列第{focus.rank}位"]
    ye_c = _rank_vs_clause(focus.rank_change, ye, prefix="相较")
    if ye_c:
        rank_bits.append(ye_c)
    if _dual_rank_anchor(config):
        q_c = _rank_vs_clause(focus.rank_change_q, pq)
        if q_c:
            rank_bits.append(q_c)
    p1 = "，".join(rank_bits) + "。"

    ind_nm = next((r for r in (cat.rows if cat else []) if r.name == "非货"), None)
    avg_g = ind_nm.growth_pct if ind_nm and ind_nm.growth_pct is not None else (
        sum(r.growth_pct or 0 for r in sh.rows[: config.top_n]) / max(min(len(sh.rows), config.top_n), 1)
    )
    avg_qg = ind_nm.q_growth_pct if ind_nm and ind_nm.q_growth_pct is not None else (
        sum(r.q_growth_pct or 0 for r in sh.rows[: config.top_n]) / max(min(len(sh.rows), config.top_n), 1)
    )
    vs_y = "高于" if (focus.growth_pct or 0) >= avg_g else "低于"
    vs_q = "高于" if (focus.q_growth_pct or 0) >= avg_qg else "低于"
    p = _profile(config)
    short_ytd = _ytd_phrase(config)
    top = sh.rows[: config.top_n]
    g_pos = next(
        (
            i
            for i, r in enumerate(
                sorted(top, key=lambda x: x.growth_pct or -999, reverse=True), 1
            )
            if r.name == focus.name
        ),
        None,
    )
    i_pos = next(
        (
            i
            for i, r in enumerate(
                sorted(top, key=lambda x: x.increment or -1e18, reverse=True), 1
            )
            if r.name == focus.name
        ),
        None,
    )
    if not p.dual_metrics:
        pos_txt = ""
        if i_pos and g_pos:
            pos_txt = f"，增量和增速在Top{config.top_n}中分别位列第{i_pos}和第{g_pos}"
        p2 = (
            f"{short}非货{config.quarter_tag()}增长{_fmt_yi(focus.increment)}亿，"
            f"增速{_fmt_pct(focus.growth_pct)}%"
            f"{pos_txt}，"
            f"{vs_y}Top{config.top_n}公司（{_fmt_pct(avg_g)}%）。"
        )
    elif p.lead == "quarter":
        p2 = (
            f"{short}非货{config.quarter_tag()}增量{_fmt_yi(focus.q_increment)}亿，增速{_fmt_pct(focus.q_growth_pct)}%，"
            f"{vs_q}Top{config.top_n}公司（{_fmt_pct(avg_qg)}%）；"
            f"{short_ytd}增长{_fmt_yi(focus.increment)}亿，增速{_fmt_pct(focus.growth_pct)}%，"
            f"{vs_y}Top{config.top_n}公司（{_fmt_pct(avg_g)}%）。"
        )
    else:
        p2 = (
            f"{short}非货{short_ytd}增长{_fmt_yi(focus.increment)}亿，"
            f"增速{_fmt_pct(focus.growth_pct)}%，"
            f"{vs_y}Top{config.top_n}公司（{_fmt_pct(avg_g)}%）；"
            f"{config.quarter_tag()}增量{_fmt_yi(focus.q_increment)}亿，增速{_fmt_pct(focus.q_growth_pct)}%，"
            f"{vs_q}Top{config.top_n}公司（{_fmt_pct(avg_qg)}%）。"
        )
    return {"section_title": title, "p1": p1, "p2": p2, "p3": "", "footnote_passive": foot}


def _build_increment(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("increment")
    nm = book.get("non_money")
    title = f"四、{config.period_label}非货增量情况概览"
    foot = ""
    if not sh:
        return {"section_title": title, "p1": PAGE_INSUFFICIENT, "p2": "", "p3": "", "footnote_passive": foot}

    focus = require_focus(book, "increment", config.focus_company, config.focus_company_short)
    if not focus:
        return {"section_title": title, "p1": PAGE_INSUFFICIENT, "p2": "", "p3": "", "footnote_passive": foot}

    short = config.focus_company_short
    inc_rank = focus.increment_rank or focus.rank
    q_rank = None
    q_inc = focus.q_increment
    ytd_inc = focus.increment
    ytd_rank = inc_rank
    focus_nm = find_focus(nm.rows, config.focus_company, short) if nm else None
    if focus_nm:
        if ytd_inc is None and focus_nm.increment is not None:
            ytd_inc = focus_nm.increment
            ordered_y = sorted(
                [r for r in nm.rows if r.increment is not None],
                key=lambda r: r.increment or 0,
                reverse=True,
            )
            ytd_rank = next(
                (i for i, r in enumerate(ordered_y, 1) if r.name == focus_nm.name),
                inc_rank,
            )
        if focus_nm.q_increment is not None:
            q_inc = focus_nm.q_increment
            ordered = sorted(
                [r for r in nm.rows if r.q_increment is not None],
                key=lambda r: r.q_increment or 0,
                reverse=True,
            )
            q_rank = next((i for i, r in enumerate(ordered, 1) if r.name == focus_nm.name), None)
    if q_rank is None and focus.q_increment is not None:
        ordered_q = sorted(
            [r for r in sh.rows if r.q_increment is not None],
            key=lambda r: r.q_increment or 0,
            reverse=True,
        )
        q_rank = next((i for i, r in enumerate(ordered_q, 1) if r.name == focus.name), None)
        if focus.increment_rank:
            # 增量表只有季度列时，表内「增量排名」即单季排名
            pass

    q_tag = config.quarter_tag()
    p = _profile(config)
    span = p.span
    ytd_p = _ytd_phrase(config)
    # 增量表仅季度列：用表内排名当 q_rank
    if q_rank is None and focus.increment_rank and focus.q_increment is not None and focus.increment is None:
        q_rank = focus.increment_rank

    if not p.dual_metrics:
        # Q1：YTD≡单季；全行业名次必须用增量表「非货增量排名」列（如第132），
        # 不可用非货 Top 表内按增量重排的位次。
        use_inc = focus.increment if focus.increment is not None else q_inc
        use_rank = focus.increment_rank or inc_rank
        p1 = (
            f"{short}{q_tag}非货增量为{_fmt_yi(use_inc)}亿，"
            f"在全行业排名第{use_rank}名。"
        )
    elif p.lead == "quarter" and q_rank and q_inc is not None:
        p1 = (
            f"{short}{q_tag}非货增量为{_fmt_yi(q_inc)}亿，在全行业排名第{q_rank}名；"
            f"{ytd_p}非货增量为{_fmt_yi(ytd_inc)}亿，在全行业排名第{ytd_rank}名。"
        )
    elif q_rank and q_inc is not None:
        p1 = (
            f"{short}{ytd_p}非货增量为{_fmt_yi(ytd_inc if ytd_inc is not None else focus.increment)}亿，"
            f"在全行业排名第{ytd_rank}名；"
            f"{q_tag}，非货增量为{_fmt_yi(q_inc)}亿，在全行业排名第{q_rank}名。"
        )
    else:
        p1 = (
            f"{short}{ytd_p}非货增量为{_fmt_yi(ytd_inc if ytd_inc is not None else focus.increment)}亿，"
            f"在全行业排名第{ytd_rank}名。"
        )

    cat_names = {
        "active_equity": "主动权益",
        "passive_equity": "被动权益",
        "fixed_income_plus": "固收+",
        "fixed_income": "固收",
        "fof": "FOF",
    }
    ranked = sorted(
        [(k, v) for k, v in focus.category_increment_ranks.items() if v and v <= 30],
        key=lambda kv: kv[1],
    )[:3]
    p2 = ""
    if ranked:
        names = "、".join(cat_names.get(k, k) for k, _ in ranked)
        if len(ranked) == 3:
            ranks = f"第{ranked[0][1]}名、第{ranked[1][1]}名和第{ranked[2][1]}名"
        else:
            ranks = "、".join(f"第{v}名" for _, v in ranked)
        p2 = f"{names}的{(p.quarter_tag if p.lead == 'quarter' else span)}增量排名靠前，分别为{ranks}。"

    return {"section_title": title, "p1": p1, "p2": p2, "p3": "", "footnote_passive": foot}


def _build_active(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("active_equity")
    title = f"五、{config.period_label}各项业务排名情况"
    base = {
        "section_title": title,
        "p1": "主动权益：",
        "p2": PAGE_INSUFFICIENT,
        "p3": "",
        "p4": "",
        "subtitle": "",
    }
    if not sh or not sh.rows:
        return base

    top = sh.rows[: config.top_n]
    span = _period_span(config)
    p2 = _active_peer_sentence(top, config)

    focus = require_focus(book, "active_equity", config.focus_company, config.focus_company_short)
    p3 = ""
    if focus:
        ye = config.year_end_ref()
        verb = "增长" if (focus.increment or 0) >= 0 else "下降"
        attr = _active_focus_attribution(focus, span)
        p3 = (
            f"{config.focus_company_short}主动权益排名第{focus.rank}名，较{ye}"
            f"{_rank_change_text(focus.rank_change)}。"
            f"{span}规模{verb}{_fmt_yi(abs(focus.increment or 0))}亿"
            f"至{_fmt_yi(focus.aum)}亿{attr}。"
        )
    return {**base, "p2": p2, "p3": p3}


def _etf_line(label: str, focus: FactRow | None, config: ReportConfig) -> str:
    short = config.focus_company_short
    ye = config.year_end_ref()
    span = _period_span(config)
    if not focus:
        return f"{label}：{PAGE_INSUFFICIENT}"
    bits = [
        f"{label}：{short}排名第{focus.rank}名",
    ]
    ye_c = _rank_vs_clause(focus.rank_change, ye)
    if ye_c:
        bits.append(ye_c)
    bits.append(
        f"{span}规模{'增长' if (focus.increment or 0) >= 0 else '下滑'}{_fmt_yi(abs(focus.increment or 0))}亿"
    )
    if focus.growth_pct is not None and abs(focus.growth_pct) > 0.001:
        bits.append(f"增速{_fmt_pct(focus.growth_pct)}%")
    hold = focus.holding_sales or 0
    nav = focus.nav_change or 0
    new_i = focus.new_issue or 0
    # 拆解只报有实质变化的分项
    if abs(new_i) > 1:
        bits.append(f"新发{_fmt_yi(new_i)}亿")
    if abs(nav) > 1:
        bits.append(f"净值贡献{_fmt_yi(nav)}亿")
    if abs(hold) > 1:
        bits.append(f"持营{_fmt_yi(hold)}亿")
    return "，".join(bits) + "。"


def _build_etf(book: TableBook, config: ReportConfig) -> dict:
    w = book.get("etf_non_money")
    n = book.get("etf_equity")
    f1 = find_focus(w.rows, config.focus_company, config.focus_company_short) if w else None
    f2 = find_focus(n.rows, config.focus_company, config.focus_company_short) if n else None
    if w and not f1:
        require_focus(book, "etf_non_money", config.focus_company, config.focus_company_short)
    if n and not f2:
        require_focus(book, "etf_equity", config.focus_company, config.focus_company_short)
    if not w and not n:
        return {
            "p0": PAGE_INSUFFICIENT,
            "p1": PAGE_INSUFFICIENT,
            "sub_with": "",
            "sub_without": "",
        }
    return {
        "p0": _etf_line("非货ETF（含联接）", f1, config),
        "p1": _etf_line("权益ETF（含联接）", f2, config),
        "sub_with": "",
        "sub_without": "",
    }


def _build_money(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("money")
    if not sh:
        return {"title": "", "label_line": "货币：", "p1": PAGE_INSUFFICIENT}
    focus = require_focus(book, "money", config.focus_company, config.focus_company_short)
    if not focus:
        return {"title": "", "label_line": "货币：", "p1": PAGE_INSUFFICIENT}
    top = sh.rows[: config.top_n]
    ye, pq = config.year_end_ref(), config.prev_quarter_ref()
    span = _period_span(config)
    g_pos = next(
        (
            i
            for i, r in enumerate(
                sorted(top, key=lambda x: x.growth_pct or -999, reverse=True), 1
            )
            if r.name == focus.name
        ),
        None,
    )
    bits = [
        f"{config.focus_company_short}位列第{focus.rank}名",
    ]
    ye_c = _rank_vs_clause(focus.rank_change, ye, prefix="排名较")
    if ye_c:
        bits.append(ye_c)
    if _dual_rank_anchor(config):
        q_c = _rank_vs_clause(focus.rank_change_q, pq)
        if q_c:
            bits.append(q_c)
    bits.append(
        f"{span}规模增长{_fmt_yi(focus.increment)}亿元，增速{_fmt_pct(focus.growth_pct)}%"
    )
    if g_pos:
        bits.append(f"增速在货币Top{config.top_n}中位列第{g_pos}")
    p1 = "，".join(bits) + "。"
    return {"title": "", "label_line": "货币：", "p1": p1}


def _build_fi(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("fixed_income")
    if not sh:
        return {
            "title": "",
            "label_line": "固收：",
            "p1": PAGE_INSUFFICIENT,
            "p2": "",
            "p3": "",
            "footnote_fi": "",
        }
    top = sh.rows[: config.top_n]
    n_up = sum(1 for r in top if (r.increment or 0) > 0)
    med_inc = median([r.increment or 0 for r in top]) if top else 0
    med_g = median([r.growth_pct or 0 for r in top]) if top else 0
    p1 = (
        f"固收Top{config.top_n}公司中{n_up}家实现规模增长，"
        f"规模增量中位数{_fmt_yi(med_inc)}亿，增速中位数{_fmt_pct(med_g)}%。"
    )
    focus = require_focus(book, "fixed_income", config.focus_company, config.focus_company_short)
    p2 = p3 = ""
    if focus:
        ye = config.year_end_ref()
        hold = focus.holding_sales or 0
        pq_q = config.quarter_tag()
        p = _profile(config)
        bits = [
            f"{config.focus_company_short}固收{_ytd_phrase(config)}规模增量{_fmt_yi(focus.increment)}亿",
            f"增速{_fmt_pct(focus.growth_pct)}%",
        ]
        if abs(hold) > 1:
            bits.append(f"持营{_fmt_yi(hold)}亿")
        bits.append(f"位列第{focus.rank}名")
        ye_c = _rank_vs_clause(focus.rank_change, ye)
        if ye_c:
            bits.append(ye_c)
        if _dual_rank_anchor(config):
            q_c = _rank_vs_clause(focus.rank_change_q, config.prev_quarter_ref())
            if q_c:
                bits.append(q_c)
        p2 = "，".join(bits) + "。"
        if not p.dual_metrics:
            p3 = ""
        else:
            qi = sorted(top, key=lambda r: r.q_increment or -1e18, reverse=True)
            qg = sorted(top, key=lambda r: r.q_growth_pct or -1e18, reverse=True)
            qi_pos = next((i for i, r in enumerate(qi, 1) if r.name == focus.name), None)
            qg_pos = next((i for i, r in enumerate(qg, 1) if r.name == focus.name), None)
            if qi_pos and qg_pos:
                if p.lead == "quarter":
                    p3 = (
                        f"就{pq_q}而言，{config.focus_company_short}固收的规模增量达{_fmt_yi(focus.q_increment)}亿，"
                        f"在固收Top{config.top_n}公司位列第{qi_pos}；"
                        f"增速{_fmt_pct(focus.q_growth_pct)}%，在固收Top{config.top_n}公司位列第{qg_pos}。"
                        f"（{_ytd_phrase(config)}增量{_fmt_yi(focus.increment)}亿。）"
                    )
                else:
                    p3 = (
                        f"就{pq_q}而言，{config.focus_company_short}固收的规模增量达{_fmt_yi(focus.q_increment)}亿，"
                        f"在固收Top{config.top_n}公司位列第{qi_pos}；"
                        f"增速{_fmt_pct(focus.q_growth_pct)}%，在固收Top{config.top_n}公司位列第{qg_pos}。"
                    )
    return {
        "title": "",
        "label_line": "固收：",
        "p1": p1,
        "p2": p2,
        "p3": p3,
        "footnote_fi": "",
    }


def _build_fip(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("fixed_income_plus")
    if not sh:
        return {
            "title": "",
            "label_line": "固收+:",
            "p1": PAGE_INSUFFICIENT,
            "p2": "",
            "footnote_fi": "",
        }
    focus = require_focus(book, "fixed_income_plus", config.focus_company, config.focus_company_short)
    if not focus:
        return {
            "title": "",
            "label_line": "固收+:",
            "p1": PAGE_INSUFFICIENT,
            "p2": "",
            "footnote_fi": "",
        }
    ye = config.year_end_ref()
    pq = config.prev_quarter_ref()
    hold = focus.holding_sales or 0
    new_i = focus.new_issue or 0
    nav = focus.nav_change or 0
    attr_parts: list[str] = []
    if abs(new_i) > 1:
        attr_parts.append(f"新发{_fmt_yi(new_i)}亿")
    if abs(nav) > 1:
        attr_parts.append(f"净值贡献{_fmt_yi(nav)}亿")
    if abs(hold) > 1:
        attr_parts.append(f"持营{_fmt_yi(hold)}亿")
    attr = f"，其中{'，'.join(attr_parts)}" if attr_parts else ""
    p1 = (
        f"{config.focus_company_short}固收+规模{_fmt_yi(focus.aum)}亿，"
        f"较年初增长{_fmt_yi(focus.increment)}亿，增速{_fmt_pct(focus.growth_pct)}%{attr}。"
    )
    rank_bits = [f"{config.focus_company_short}固收+排名第{focus.rank}名"]
    ye_c = _rank_vs_clause(focus.rank_change, ye, unit="位")
    if ye_c:
        rank_bits.append(ye_c)
    if _dual_rank_anchor(config):
        q_c = _rank_vs_clause(focus.rank_change_q, pq, unit="位")
        if q_c:
            rank_bits.append(q_c)
    p2 = "，".join(rank_bits) + "。"
    return {
        "title": "",
        "label_line": "固收+:",
        "p1": p1,
        "p2": p2,
        "footnote_fi": "",
    }


def _build_fof(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("fof")
    if not sh:
        return {"title": "", "label_line": "FOF：", "p1": PAGE_INSUFFICIENT}
    focus = require_focus(book, "fof", config.focus_company, config.focus_company_short)
    if not focus:
        return {"title": "", "label_line": "FOF：", "p1": PAGE_INSUFFICIENT}
    ye, pq = config.year_end_ref(), config.prev_quarter_ref()
    span = _period_span(config)
    attr_parts: list[str] = []
    if abs(focus.new_issue or 0) > 1:
        attr_parts.append(f"新发{_fmt_yi(focus.new_issue)}亿")
    if abs(focus.nav_change or 0) > 1:
        attr_parts.append(f"净值贡献{_fmt_yi(focus.nav_change)}亿")
    if abs(focus.holding_sales or 0) > 1:
        attr_parts.append(f"持营{_fmt_yi(focus.holding_sales)}亿")
    attr = f"，其中{'，'.join(attr_parts)}" if attr_parts else ""
    bits = [
        f"{config.focus_company_short}FOF位列第{focus.rank}名",
    ]
    ye_c = _rank_vs_clause(focus.rank_change, ye)
    if ye_c:
        bits.append(ye_c)
    if _dual_rank_anchor(config):
        q_c = _rank_vs_clause(focus.rank_change_q, pq)
        if q_c:
            bits.append(q_c)
    bits.append(
        f"{span}规模{_fmt_yi(focus.aum)}亿，增长{_fmt_yi(focus.increment)}亿，"
        f"增速{_fmt_pct(focus.growth_pct)}%{attr}"
    )
    p1 = "，".join(bits) + "。"
    return {"title": "", "label_line": "FOF：", "p1": p1}


def build_narratives_from_tables(
    config: ReportConfig,
    *,
    tables_xlsx: str | None = None,
    final_sheets: list[tuple[str, list, list[dict]]] | None = None,
) -> tuple[dict[str, dict], TableBook]:
    """返回 (slide_id -> payload, TableBook)。"""
    if tables_xlsx:
        book = load_table_book_from_xlsx(tables_xlsx)
    elif final_sheets is not None:
        book = load_table_book_from_final_sheets(final_sheets)
    else:
        raise ValueError("需提供 tables_xlsx 或 final_sheets")

    payloads: dict[str, Any] = {
        "overview": _build_overview(book, config),
        "total_ranking": _build_total(book, config),
        "non_money_ranking": _build_non_money(book, config),
        "increment_overview": _build_increment(book, config),
        "active_equity": _build_active(book, config),
        "etf_dual": _build_etf(book, config),
        "money": _build_money(book, config),
        "fixed_income": _build_fi(book, config),
        "fixed_income_plus": _build_fip(book, config),
        "fof": _build_fof(book, config),
    }
    return payloads, book
