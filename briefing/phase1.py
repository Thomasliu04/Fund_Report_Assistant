"""【遗留】Phase 1：第 1–2 页文字 + 表图自动灌入 PPT。

正式出报请用 full-deck / serve（终表驱动）。
"""

from __future__ import annotations

from pathlib import Path

from briefing.analytics.engine import (
    compute_category_summary,
    compute_company_ranking,
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


def build_overview_payload(
    metrics: list[CategoryMetrics],
    config: ReportConfig,
    *,
    focus_total: CompanyRanking | None = None,
    focus_non_money: CompanyRanking | None = None,
) -> dict:
    narr = generate_overview_narrative(metrics, config)
    total = next((m for m in metrics if m.category == "total"), None)
    non_money = next((m for m in metrics if m.category == "non_money"), None)
    short = config.focus_company_short

    opening = narr.paragraphs[0] if narr.paragraphs else ""
    # 对齐 Q4 模版：行业增速 vs 示例增速
    if total and non_money:
        y = config.current_date[2:4]
        m = int(config.current_date[4:6])
        d = int(config.current_date[6:8])
        ft_g = _fmt_pct(focus_total.growth_pct) if focus_total else "—"
        fn_g = _fmt_pct(focus_non_money.growth_pct) if focus_non_money else "—"
        opening = (
            f"截至{y}年{m}月{d}日，公募行业总规模{_fmt_wan_yi(total.aum_current)}万亿"
            f"（行业增速{_fmt_pct(total.growth_pct)}% vs {short}增速{ft_g}%），"
            f"非货{_fmt_wan_yi(non_money.aum_current)}万亿"
            f"（行业增速{_fmt_pct(non_money.growth_pct)}% vs {short}增速{fn_g}%）。"
        )

    bullets = narr.bullets[:5]
    while len(bullets) < 5:
        bullets.append("")

    y_short = config.current_date[2:4]
    if config.period_type.value == "half_year":
        growth_intro = f"{y_short}年上半年规模增速较快的品类："
    else:
        growth_intro = f"{y_short}年规模增速较快的品类："

    # 案例：固收负增长单独放在 highlight；行情/产品注释属特殊情况，不自动编造
    highlight = ""
    if len(narr.paragraphs) > 1:
        highlight = narr.paragraphs[1]

    title = config.title
    if " " not in title and title.endswith("简报"):
        # 对齐案例「2025Q4 公募行业数据简报」
        title = f"{config.period_label} 公募行业数据简报"

    return {
        "title": title,
        "data_source_note": config.data_source_note,
        "section_title": f"一、{config.period_label}公募规模整体情况",
        "opening": opening,
        "all_positive_intro": growth_intro,
        "bullet_1": bullets[0] if bullets else "",
        "bullet_2": bullets[1] if len(bullets) > 1 else "",
        "bullet_3": bullets[2] if len(bullets) > 2 else "",
        "bullet_4": bullets[3] if len(bullets) > 3 else "",
        "bullet_5": bullets[4] if len(bullets) > 4 else "",
        "highlight_1": highlight,
        "highlight_2": "",
        "footnote": "",
    }


def build_total_ranking_payload(rankings: list[CompanyRanking], config: ReportConfig) -> dict:
    narr = generate_total_ranking_narrative(rankings, config)
    paras = narr.paragraphs
    return {
        "section_title": f"二、{config.period_label}公募基金总规模（含货币）排名",
        "peer_moves": paras[0] if paras else "",
        "focus_rank": paras[1] if len(paras) > 1 else "",
        "focus_growth": paras[2] if len(paras) > 2 else "",
        "footnote_passive": "",
    }


def _rank_change_phrase(change: int) -> str:
    if change > 0:
        return f"上升{change}名"
    if change < 0:
        return f"下降{abs(change)}名"
    return "不变"



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
        if "示例" in r.company:
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
