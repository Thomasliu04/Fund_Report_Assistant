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
from briefing.models import PeriodType, ReportConfig

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


def _period_span(config: ReportConfig) -> str:
    if config.period_type == PeriodType.HALF_YEAR:
        return "上半年"
    return "全年"


def _growth_intro(config: ReportConfig) -> str:
    y = config.current_date[2:4]
    if config.period_type == PeriodType.HALF_YEAR:
        return f"{y}年上半年规模增速较快的品类："
    return f"{y}年规模增速较快的品类："


def _short_co(name: str) -> str:
    s = name.replace("基金", "")
    if s.startswith("景顺"):
        return "景顺"
    return s


def _overview_driver(r: FactRow) -> str:
    inc = r.increment or 0.0
    if abs(inc) < 1:
        return ""
    new_i = r.new_issue or 0.0
    nav = r.nav_change or 0.0
    hold = r.holding_sales or 0.0

    def share(x: float) -> int:
        return round(abs(x) / abs(inc) * 100) if abs(inc) >= 1 else 0

    if nav > 0 and hold < -1000 and abs(nav) >= abs(new_i):
        return (
            f"其中净值增长{_fmt_inc_amount(nav)}，"
            f"但客户净赎回近{round(abs(hold))}亿元"
        )
    sh_new = share(new_i)
    if new_i > 0 and new_i >= max(nav, hold, 0) and 40 <= sh_new <= 100:
        return f"新发增量{round(new_i)}亿，占增量{sh_new}%"
    sh_nav = share(nav)
    if nav > 0 and abs(nav) >= abs(hold) and abs(nav) >= abs(new_i) and 40 <= sh_nav <= 100:
        return f"其中净值增长{_fmt_inc_amount(nav)}，占增量的{sh_nav}%"
    if hold > 0 and hold >= max(new_i, nav, 0):
        return f"持营增长{'近' if hold >= 5000 else ''}{round(hold)}亿"
    if new_i > 0 and sh_new <= 100:
        return f"新发增量{round(new_i)}亿，占增量{sh_new}%"
    if abs(nav) > 1:
        return f"其中净值增长{_fmt_inc_amount(nav)}"
    if abs(hold) > 1:
        return f"持营增长{round(hold)}亿"
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
    if total and non_money and total.aum is not None and non_money.aum is not None:
        ft = _fmt_pct(focus_t.growth_pct) if focus_t and focus_t.growth_pct is not None else "—"
        fn = _fmt_pct(focus_n.growth_pct) if focus_n and focus_n.growth_pct is not None else "—"
        opening = (
            f"截至{y}年{m}月{d}日，公募行业总规模{_fmt_wan_yi(total.aum)}万亿"
            f"（行业增速{_fmt_pct(total.growth_pct)}% vs {short}增速{ft}%），"
            f"非货{_fmt_wan_yi(non_money.aum)}万亿"
            f"（行业增速{_fmt_pct(non_money.growth_pct)}% vs {short}增速{fn}%）。"
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
        if r.name == "货币":
            bullets[i] = f"{base}。"
        else:
            driver = _overview_driver(r)
            bullets[i] = f"{base}，{driver}。" if driver else f"{base}。"

    highlight = ""
    fi = by_name.get("固收")
    if fi and (fi.increment or 0) < 0:
        highlight = (
            f"{config.ytd_tag()}{_period_span(config)}，固收规模减少{round(abs(fi.increment or 0))}亿，"
            f"增速{_fmt_pct(fi.growth_pct)}%。"
        )

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
        "footnote": "注：REITs和另类基金未在表中注明",
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
        "footnote": "注：REITs和另类基金未在表中注明",
    }


def _build_total(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("total")
    title = f"二、{config.period_label}公募基金总规模（含货币）排名"
    empty = {
        "section_title": title,
        "peer_moves": PAGE_INSUFFICIENT,
        "focus_rank": "",
        "focus_growth": "",
        "footnote_passive": "注：被动权益=ETF+联接+场外普通指数+场外指数增强",
    }
    if not sh or not sh.rows:
        return empty

    top = sh.rows[: config.top_n]
    up = [r for r in top if (r.rank_change or 0) >= 2]
    peer = f"总规模Top{config.top_n}公司，位次变化并不明显。"
    if up:
        by_delta: dict[int, list[str]] = {}
        for r in sorted(up, key=lambda x: -(x.rank_change or 0))[:8]:
            by_delta.setdefault(r.rank_change or 0, []).append(_short_co(r.name))
        parts = []
        for delta, names in sorted(by_delta.items(), key=lambda kv: -kv[0]):
            names = names[:5]
            if len(names) >= 2:
                parts.append(f"{'、'.join(names)}各上升{delta}名")
            else:
                parts.append(f"{names[0]}上升{delta}名")
        if parts:
            peer = f"总规模Top{config.top_n}公司，{'，'.join(parts)}，其余公司位次变化并不明显。"

    focus = require_focus(book, "total", config.focus_company, config.focus_company_short)
    if not focus:
        return {**empty, "peer_moves": peer, "focus_rank": PAGE_INSUFFICIENT}

    ye, pq = config.year_end_ref(), config.prev_quarter_ref()
    focus_rank = (
        f"{config.focus_company_short}总规模排名{focus.rank}，"
        f"较{ye}排名{_rank_change_text(focus.rank_change)}，"
        f"较{pq}季度{_rank_change_text(focus.rank_change_q)}。"
    )
    g_rank = sorted(top, key=lambda x: x.growth_pct or -999, reverse=True)
    qg_rank = sorted(top, key=lambda x: x.q_growth_pct or -999, reverse=True)
    g_pos = next((i for i, r in enumerate(g_rank, 1) if r.name == focus.name), None)
    qg_pos = next((i for i, r in enumerate(qg_rank, 1) if r.name == focus.name), None)
    span = _period_span(config)
    q_line = (
        f"{config.quarter_tag()}单季度增速达{_fmt_pct(focus.q_growth_pct)}%，"
        f"在Top{config.top_n}公司中位列第{qg_pos}名"
        if qg_pos
        else f"{config.quarter_tag()}单季度增速达{_fmt_pct(focus.q_growth_pct)}%"
    )
    if qg_pos == 2 and qg_rank and qg_rank[0].name != focus.name:
        leader = qg_rank[0]
        q_line += f"，仅次于{_short_co(leader.name)}（{_fmt_pct(leader.q_growth_pct)}%）"
    focus_growth = (
        f"{config.focus_company_short}总规模{span}增速为{_fmt_pct(focus.growth_pct)}%，"
        f"在Top{config.top_n}公司中位列第{g_pos}；{q_line}。"
    )
    return {
        "section_title": title,
        "peer_moves": peer,
        "focus_rank": focus_rank,
        "focus_growth": focus_growth,
        "footnote_passive": "注：被动权益=ETF+联接+场外普通指数+场外指数增强",
    }


def _build_non_money(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("non_money")
    cat = book.get("category")
    title = f"三、{config.period_label}公募基金非货规模排名"
    foot = "注：被动权益=ETF+联接+场外普通指数+场外指数增强"
    if not sh:
        return {"section_title": title, "p1": PAGE_INSUFFICIENT, "p2": "", "p3": "", "footnote_passive": foot}

    focus = require_focus(book, "non_money", config.focus_company, config.focus_company_short)
    if not focus:
        return {"section_title": title, "p1": PAGE_INSUFFICIENT, "p2": "", "p3": "", "footnote_passive": foot}

    ye, pq = config.year_end_ref(), config.prev_quarter_ref()
    short = config.focus_company_short
    lift = ""
    if (focus.rank_change or 0) >= 2 and (focus.rank or 99) <= 20:
        lift = "，非货排名显著提升，重回Top20行列"
    elif (focus.rank_change or 0) >= 2:
        lift = "，非货排名显著提升"
    p1 = (
        f"非货规模方面，{short}位列第{focus.rank}位，"
        f"相较{ye}{_rank_change_text(focus.rank_change)}，"
        f"较{pq}{_rank_change_text(focus.rank_change_q)}{lift}。"
    )

    ind_nm = next((r for r in (cat.rows if cat else []) if r.name == "非货"), None)
    avg_g = ind_nm.growth_pct if ind_nm and ind_nm.growth_pct is not None else (
        sum(r.growth_pct or 0 for r in sh.rows[: config.top_n]) / max(min(len(sh.rows), config.top_n), 1)
    )
    avg_qg = ind_nm.q_growth_pct if ind_nm and ind_nm.q_growth_pct is not None else (
        sum(r.q_growth_pct or 0 for r in sh.rows[: config.top_n]) / max(min(len(sh.rows), config.top_n), 1)
    )
    vs_y = "高于" if (focus.growth_pct or 0) >= avg_g else "低于"
    vs_q = "高于" if (focus.q_growth_pct or 0) >= avg_qg else "低于"
    span = _period_span(config)
    p2 = (
        f"{short}非货{config.ytd_tag()}{span}增长{_fmt_yi(focus.increment)}亿，"
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
    foot = "注：被动权益=ETF+联接+场外普通指数+场外指数增强"
    if not sh:
        return {"section_title": title, "p1": PAGE_INSUFFICIENT, "p2": "", "p3": "", "footnote_passive": foot}

    focus = require_focus(book, "increment", config.focus_company, config.focus_company_short)
    if not focus:
        return {"section_title": title, "p1": PAGE_INSUFFICIENT, "p2": "", "p3": "", "footnote_passive": foot}

    short = config.focus_company_short
    inc_rank = focus.increment_rank or focus.rank
    # Q 增量排名：优先用增量表自身；否则用非货表 Q2增量排序
    q_rank = None
    q_inc = focus.q_increment
    if nm:
        focus_nm = find_focus(nm.rows, config.focus_company, short)
        if focus_nm and focus_nm.q_increment is not None:
            q_inc = focus_nm.q_increment
            ordered = sorted(
                [r for r in nm.rows if r.q_increment is not None],
                key=lambda r: r.q_increment or 0,
                reverse=True,
            )
            q_rank = next((i for i, r in enumerate(ordered, 1) if r.name == focus_nm.name), None)

    q_tag = config.quarter_tag()
    span = _period_span(config)
    if q_rank and q_inc is not None:
        p1 = (
            f"{short}{config.ytd_tag()}非货增量为{_fmt_yi(focus.increment)}亿，"
            f"在全行业排名第{inc_rank}名；"
            f"{q_tag}，非货增量为{_fmt_yi(q_inc)}亿，在全行业排名第{q_rank}名。"
        )
    else:
        p1 = (
            f"{short}{config.ytd_tag()}非货增量为{_fmt_yi(focus.increment)}亿，"
            f"在全行业排名第{inc_rank}名。"
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
        p2 = f"{names}的{span}增量排名靠前，分别为{ranks}。"

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
        "subtitle": "1 主动权益排名",
    }
    if not sh or not sh.rows:
        return base

    top = sh.rows[: config.top_n]
    star = max(top, key=lambda r: r.growth_pct or -999)
    big = [
        r
        for r in top
        if (r.increment or 0) >= 300 and r.name != star.name
    ]
    big = sorted(big, key=lambda r: r.increment or 0, reverse=True)[:3]
    span = _period_span(config)
    if star and (star.growth_pct or 0) >= 1:
        hold_note = ""
        if (star.holding_sales or 0) > max(star.new_issue or 0, star.nav_change or 0, 0):
            hold_note = "，主要源于持营"
        p2 = (
            f"主动权益Top{config.top_n}公司中，"
            f"{_short_co(star.name)}{span}增长{_fmt_yi(star.increment)}亿，"
            f"增速达{_fmt_pct(star.growth_pct)}%{hold_note}"
        )
        if big:
            p2 += f"；{'、'.join(_short_co(r.name) for r in big)}也实现了300亿+的规模增长"
        p2 += "。"
    else:
        p2 = f"主动权益Top{config.top_n}公司中，多数公司规模与增速小幅波动。"

    focus = require_focus(book, "active_equity", config.focus_company, config.focus_company_short)
    p3 = ""
    if focus:
        ye = config.year_end_ref()
        hold = focus.holding_sales or 0
        if hold < -1:
            hold_txt = f"，但{span}净赎回{_fmt_yi(abs(hold))}亿元，对规模影响显著"
        elif abs(hold) > 1:
            hold_txt = f"，持营{_fmt_yi(hold)}亿元"
        else:
            hold_txt = ""
        nav = focus.nav_change or 0
        nav_txt = f"，其中净值贡献{_fmt_yi(nav)}亿元" if abs(nav) > 1 else ""
        verb = "增长" if (focus.increment or 0) >= 0 else "下降"
        p3 = (
            f"{config.focus_company_short}主动权益排名第{focus.rank}名，较{ye}"
            f"{_rank_change_text(focus.rank_change)}。"
            f"{span}规模{verb}{_fmt_yi(abs(focus.increment or 0))}亿"
            f"至{_fmt_yi(focus.aum)}亿{nav_txt}{hold_txt}。"
        )
    return {**base, "p2": p2, "p3": p3}


def _etf_line(label: str, focus: FactRow | None, config: ReportConfig) -> str:
    short = config.focus_company_short
    ye = config.year_end_ref()
    span = _period_span(config)
    if not focus:
        return f"{label}：{PAGE_INSUFFICIENT}"
    bits = [
        f"{label}：{short}排名{focus.rank}名，较{ye}{_rank_change_text(focus.rank_change)}",
        f"{span}合计{'增长' if (focus.increment or 0) >= 0 else '下滑'}{_fmt_yi(abs(focus.increment or 0))}亿",
    ]
    if focus.growth_pct is not None and abs(focus.growth_pct) > 0.001:
        bits.append(f"增速{_fmt_pct(focus.growth_pct)}%")
        if focus.growth_pct >= 1:
            bits.append("实现规模翻倍")
    hold = focus.holding_sales or 0
    nav = focus.nav_change or 0
    if abs(hold) > 1:
        bits.append(f"持营贡献{_fmt_yi(hold)}亿元")
    elif abs(nav) > 1:
        bits.append(f"其中净值贡献{_fmt_yi(nav)}亿元")
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
            "sub_with": "2.1 非货ETF（含联接）排名",
            "sub_without": "2.2 权益ETF（含联接）排名",
        }
    return {
        "p0": _etf_line("非货ETF（含联接）", f1, config),
        "p1": _etf_line("权益ETF（含联接）", f2, config),
        "sub_with": "2.1 非货ETF（含联接）排名",
        "sub_without": "2.2 权益ETF（含联接）排名",
    }


def _build_money(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("money")
    if not sh:
        return {"title": "3.货币排名情况", "label_line": "货币：", "p1": PAGE_INSUFFICIENT}
    focus = require_focus(book, "money", config.focus_company, config.focus_company_short)
    if not focus:
        return {"title": "3.货币排名情况", "label_line": "货币：", "p1": PAGE_INSUFFICIENT}
    top = sh.rows[: config.top_n]
    avg_g = sum(r.growth_pct or 0 for r in top) / max(len(top), 1)
    ye, pq = config.year_end_ref(), config.prev_quarter_ref()
    vs = "高于" if (focus.growth_pct or 0) >= avg_g else "低于"
    q_chg = _rank_change_text(focus.rank_change_q)
    if focus.rank_change_q == 0:
        q_chg = "持平"
    span = _period_span(config)
    p1 = (
        f"{config.focus_company_short}位列{focus.rank}名，排名较{ye}"
        f"{_rank_change_text(focus.rank_change)}，较{pq}{q_chg}。"
        f"{span}规模增长{_fmt_yi(focus.increment)}亿元，增速{_fmt_pct(focus.growth_pct)}%，"
        f"{vs}货币Top{config.top_n}平均增幅（{_fmt_pct(avg_g)}%）。"
    )
    return {"title": "3.货币排名情况", "label_line": "货币：", "p1": p1}


def _build_fi(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("fixed_income")
    if not sh:
        return {
            "title": "4. 固收排名情况",
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
        f"固收Top{config.top_n}公司中仅{n_up}家实现规模增长。"
        f"Top{config.top_n}公司规模增量中位数{_fmt_yi(med_inc)}亿，"
        f"增速中位数{_fmt_pct(med_g)}%。"
    )
    focus = require_focus(book, "fixed_income", config.focus_company, config.focus_company_short)
    p2 = p3 = ""
    if focus:
        ye = config.year_end_ref()
        hold = focus.holding_sales or 0
        pq_q = config.quarter_tag()
        # 上一季口语：prev_quarter_ref
        p2 = (
            f"{config.focus_company_short}固收{config.ytd_tag()}规模增量{_fmt_yi(focus.increment)}亿，"
            f"增速{_fmt_pct(focus.growth_pct)}%，持营增量{_fmt_yi(hold)}亿；"
            f"{config.focus_company_short}固收位列第{focus.rank}名，较{ye}排名"
            f"{_rank_change_text(focus.rank_change)}，"
            f"较{config.prev_quarter_ref()}{_rank_change_text(focus.rank_change_q)}。"
        )
        qi = sorted(top, key=lambda r: r.q_increment or -1e18, reverse=True)
        qg = sorted(top, key=lambda r: r.q_growth_pct or -1e18, reverse=True)
        qi_pos = next((i for i, r in enumerate(qi, 1) if r.name == focus.name), None)
        qg_pos = next((i for i, r in enumerate(qg, 1) if r.name == focus.name), None)
        if qi_pos and qg_pos:
            p3 = (
                f"就{pq_q}而言，{config.focus_company_short}固收的规模增量达{_fmt_yi(focus.q_increment)}亿，"
                f"在固收Top{config.top_n}公司位列第{qi_pos}；"
                f"增速{_fmt_pct(focus.q_growth_pct)}%，在固收Top{config.top_n}公司位列第{qg_pos}。"
            )
    return {
        "title": "4. 固收排名情况",
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
            "title": "5. 固收+排名情况",
            "label_line": "固收+:",
            "p1": PAGE_INSUFFICIENT,
            "p2": "",
            "footnote_fi": "",
        }
    focus = require_focus(book, "fixed_income_plus", config.focus_company, config.focus_company_short)
    if not focus:
        return {
            "title": "5. 固收+排名情况",
            "label_line": "固收+:",
            "p1": PAGE_INSUFFICIENT,
            "p2": "",
            "footnote_fi": "",
        }
    ye = config.year_end_ref().replace("底", "末")
    pq = config.prev_quarter_ref()
    hold = focus.holding_sales or 0
    driver = ""
    if hold > 0 and hold >= max(focus.new_issue or 0, focus.nav_change or 0, 0):
        driver = f"，主要源于持营（+{_fmt_yi(hold)}亿）"
    g_txt = _fmt_pct(focus.growth_pct)
    high = "高达" if (focus.growth_pct or 0) >= 1 else "为"
    p1 = (
        f"{config.focus_company_short}固收+规模{_fmt_yi(focus.aum)}亿，较年初增长{_fmt_yi(focus.increment)}亿，"
        f"增速{high}{g_txt}%{driver}。"
    )
    p2 = (
        f"{config.focus_company_short}固收+排名第{focus.rank}名，较{ye}"
        f"{_rank_change_text(focus.rank_change, unit='位')}，"
        f"较{pq}{_rank_change_text(focus.rank_change_q, unit='位')}。"
    )
    return {
        "title": "5. 固收+排名情况",
        "label_line": "固收+:",
        "p1": p1,
        "p2": p2,
        "footnote_fi": "",
    }


def _build_fof(book: TableBook, config: ReportConfig) -> dict:
    sh = book.get("fof")
    if not sh:
        return {"title": "6. FOF 排名情况", "label_line": "FOF：", "p1": PAGE_INSUFFICIENT}
    focus = require_focus(book, "fof", config.focus_company, config.focus_company_short)
    if not focus:
        return {"title": "6. FOF 排名情况", "label_line": "FOF：", "p1": PAGE_INSUFFICIENT}
    ye, pq = config.year_end_ref(), config.prev_quarter_ref()
    span = _period_span(config)
    driver = ""
    if (focus.new_issue or 0) > 0 and (focus.new_issue or 0) >= max(
        focus.nav_change or 0, focus.holding_sales or 0, 0
    ):
        driver = f"，主要来自新发（+{_fmt_yi(focus.new_issue)}亿）"
    elif (focus.holding_sales or 0) > 0:
        driver = f"，主要来自持营（+{_fmt_yi(focus.holding_sales)}亿）"
    p1 = (
        f"{config.focus_company_short}FOF位列第{focus.rank}名，较{ye}"
        f"{_rank_change_text(focus.rank_change)}，较{pq}{_rank_change_text(focus.rank_change_q)}。"
        f"{span}规模增长{_fmt_yi(focus.increment)}亿，增速{_fmt_pct(focus.growth_pct)}%{driver}。"
    )
    return {"title": "6. FOF 排名情况", "label_line": "FOF：", "p1": p1}


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
