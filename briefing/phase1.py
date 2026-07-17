"""Phase 1：第 1–2 页文字 + 表图自动灌入 PPT。"""

from __future__ import annotations

from pathlib import Path

from briefing.analytics.engine import (
    compute_category_summary,
    compute_company_ranking,
    find_notable_rank_changes,
)
from briefing.data_loader import load_config, load_fund_data
from briefing.models import CategoryMetrics, CompanyRanking, ReportConfig
from briefing.narrative.generator import (
    generate_overview_narrative,
    generate_total_ranking_narrative,
)
from briefing.render.pptx_filler import fill_pptx, load_slide_map
from briefing.render.table_image import render_table_image


def _fmt_wan_yi(aum_yi: float) -> str:
    """亿元 → 万亿，保留 1 位小数。"""
    return f"{aum_yi / 10000:.1f}"


def _fmt_yi(v: float) -> str:
    return str(round(v))


def _fmt_pct(v: float) -> str:
    return f"{round(v * 100)}"


def build_overview_payload(metrics: list[CategoryMetrics], config: ReportConfig) -> dict:
    narr = generate_overview_narrative(metrics, config)
    total = next((m for m in metrics if m.category == "total"), None)
    non_money = next((m for m in metrics if m.category == "non_money"), None)

    opening = narr.paragraphs[0] if narr.paragraphs else ""
    # 叙事里用「亿」，第 1 页模版习惯写「万亿」——对齐模版口径
    if total and non_money:
        y = config.current_date[2:4]
        m = int(config.current_date[4:6])
        d = int(config.current_date[6:8])
        opening = (
            f"截至{y}年{m}月{d}日，公募行业总规模{_fmt_wan_yi(total.aum_current)}万亿"
            f"（增速{_fmt_pct(total.growth_pct)}%），"
            f"非货{_fmt_wan_yi(non_money.aum_current)}万亿（增速{_fmt_pct(non_money.growth_pct)}%）。"
        )

    bullets = narr.bullets[:4]
    while len(bullets) < 4:
        bullets.append("")

    period_word = "上半年" if config.period_type.value == "half_year" else "本季度"
    y_short = config.current_date[2:4]

    return {
        "title": f"{config.period_label}公募行业数据简报",
        "data_source_note": config.data_source_note,
        "section_title": f"一、{config.period_label}公募规模整体情况",
        "opening": opening,
        "all_positive_intro": f"{y_short}年{period_word}，所有品类规模增速均为正：",
        "bullet_1": bullets[0] if bullets else "",
        "bullet_2": bullets[1] if len(bullets) > 1 else "",
        "bullet_3": bullets[2] if len(bullets) > 2 else "",
        "bullet_4": bullets[3] if len(bullets) > 3 else "",
        "highlight_1": "值得关注的是，请人工补充季度间拐点判断；",
        "highlight_2": "系统已填入规模与增速数字，观点句请审稿时确认。",
        "footnote": "注：REITs和另类基金未在表中注明",
    }


def build_total_ranking_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    narr = generate_total_ranking_narrative(rankings, config)
    paras = narr.paragraphs
    focus = next((r for r in rankings if r.company == config.focus_company), None)

    peer_moves = paras[0] if paras else ""
    focus_rank = paras[1] if len(paras) > 1 else ""
    focus_growth = paras[2] if len(paras) > 2 else ""

    if focus:
        focus_rank = (
            f"{config.focus_company_short}总规模排名{focus.rank}，"
            f"较年初排名{'上升' + str(focus.rank_change) + '名' if focus.rank_change > 0 else '下降' + str(abs(focus.rank_change)) + '名' if focus.rank_change < 0 else '不变'}。"
        )
        pos_pct = sum(1 for r in rankings if r.increment > 0) * 100 // max(len(rankings), 1)
        focus_growth = (
            f"{config.focus_company_short}总规模较年初增长{_fmt_yi(focus.increment)}亿，"
            f"增速为{_fmt_pct(focus.growth_pct)}%；"
            f"Top{config.top_n}公司中，有{pos_pct}%的公司上半年增量均为正。"
        )

    up, down = find_notable_rank_changes(rankings, threshold=3)
    if up or down:
        parts = []
        for r in up[:2]:
            parts.append(f"{r.company.replace('基金', '')}上升{r.rank_change}名")
        for r in down[:2]:
            parts.append(f"{r.company.replace('基金', '')}下降{abs(r.rank_change)}名")
        peer_moves = (
            f"总规模Top{config.top_n}公司，{'、'.join(parts)}，其他公司位次变化不大。"
            if parts
            else peer_moves
        )

    return {
        "section_title": f"二、{config.period_label}公募基金总规模（含货币）排名",
        "peer_moves": peer_moves,
        "focus_rank": focus_rank,
        "focus_growth": focus_growth,
        "footnote_passive": "注：被动权益=ETF+联接+场外普通指数+场外指数增强",
    }


def _category_table(metrics: list[CategoryMetrics], config: ReportConfig) -> tuple[list[str], list[list[str]]]:
    headers = [
        "类型",
        config.current_date,
        config.previous_quarter_date,
        config.year_start_date,
        f"{config.period_label}增量",
        f"{config.period_label}增速%",
        "新发",
        "净值变化",
        "持营",
    ]
    rows = []
    for m in metrics:
        rows.append([
            m.label,
            _fmt_yi(m.aum_current),
            _fmt_yi(m.aum_prev_quarter),
            _fmt_yi(m.aum_year_start),
            _fmt_yi(m.increment),
            _fmt_pct(m.growth_pct),
            _fmt_yi(m.new_issue),
            _fmt_yi(m.nav_change),
            _fmt_yi(m.holding_sales),
        ])
    return headers, rows


def _ranking_table(rankings: list[CompanyRanking]) -> tuple[list[str], list[list[str]], list[int]]:
    headers = ["排名", "基金公司", "总规模", "增量", "增速%", "排名变化"]
    rows = []
    focus_idx = []
    for i, r in enumerate(rankings):
        rows.append([
            str(r.rank),
            r.company,
            _fmt_yi(r.aum),
            _fmt_yi(r.increment),
            f"{_fmt_pct(r.growth_pct)}%",
            str(r.rank_change),
        ])
        if "银华" in r.company:
            focus_idx.append(i)
    return headers, rows, focus_idx


def run_phase1(
    config_path: str | Path,
    data_dir: str | Path,
    slide_map_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    work_dir: str | Path | None = None,
) -> Path:
    config = load_config(config_path)
    df = load_fund_data(data_dir)
    slide_map = load_slide_map(slide_map_path)

    work_dir = Path(work_dir or Path(output_path).parent / "_phase1_assets")
    work_dir.mkdir(parents=True, exist_ok=True)

    industry_df = df[df["company"] == "__industry__"].copy()
    if industry_df.empty:
        industry_df = df.groupby(["category", "date"], as_index=False).agg(
            {"aum": "sum", "new_issue": "sum", "nav_change": "sum", "holding_sales": "sum"}
        )
        industry_df["company"] = "__industry__"

    metrics = compute_category_summary(industry_df, config)
    rankings = compute_company_ranking(df[df["company"] != "__industry__"], config, "total")

    payload = {
        "overview": build_overview_payload(metrics, config),
        "total_ranking": build_total_ranking_payload(rankings, config),
    }

    # 修一下 all_positive_intro（已在 build_overview_payload 内生成）

    h1, r1 = _category_table(metrics, config)
    img1 = render_table_image(h1, r1, work_dir / "category_summary.png")

    h2, r2, focus_idx = _ranking_table(rankings)
    img2 = render_table_image(h2, r2, work_dir / "total_ranking_top30.png", focus_row_indices=focus_idx)

    return fill_pptx(
        template_path=template_path,
        slide_map=slide_map,
        payload=payload,
        output_path=output_path,
        image_paths={
            "category_summary": img1,
            "total_ranking_top30": img2,
        },
    )
