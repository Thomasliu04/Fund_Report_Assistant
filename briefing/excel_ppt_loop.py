"""【遗留】最小闭环：pandas → Excel → 出图 → PPT 第2页。正式出报请用 full-deck / serve。"""

from __future__ import annotations

import shutil
from pathlib import Path

import pandas as pd

from briefing.analytics.engine import compute_company_ranking
from briefing.data_loader import load_config, load_fund_data
from briefing.excel.builder import build_total_ranking_workbook
from briefing.excel.preview_render import render_ranking_with_databars
from briefing.phase1 import build_total_ranking_payload
from briefing.render.pptx_filler import (
    _find_shape,
    replace_narrative_paragraphs,
    replace_picture_fit,
)
from pptx import Presentation


def _rank_series(s: pd.Series) -> pd.Series:
    return s.rank(ascending=False, method="min").astype(int)


def build_ranking_rows(df: pd.DataFrame, config) -> list[dict]:
    companies = df[df["company"] != "__industry__"].copy()
    total = compute_company_ranking(companies, config, "total")

    # 货币 / 非货规模与排名
    cur = config.current_date
    money = companies[companies["category"] == "money"].pivot_table(
        index="company", columns="date", values="aum", aggfunc="sum"
    )
    non = companies[companies["category"] == "non_money"].pivot_table(
        index="company", columns="date", values="aum", aggfunc="sum"
    )
    money_rank = _rank_series(money[cur]) if cur in money.columns else pd.Series(dtype=int)
    non_rank = _rank_series(non[cur]) if cur in non.columns else pd.Series(dtype=int)

    rows = []
    for r in total:
        rows.append({
            "rank": r.rank,
            "rank_change": r.rank_change,
            "company": r.company,
            "aum": r.aum,
            "increment": r.increment,
            "growth_pct": r.growth_pct,
            "money": float(money.loc[r.company, cur]) if r.company in money.index and cur in money.columns else 0,
            "non_money": float(non.loc[r.company, cur]) if r.company in non.index and cur in non.columns else 0,
            "money_rank": int(money_rank.get(r.company, 0)),
            "non_money_rank": int(non_rank.get(r.company, 0)),
        })
    return rows


def _export_table_image(xlsx_path: Path, png_path: Path, rows: list[dict], focus: str) -> tuple[Path, str]:
    """
    PPT 贴图固定用双向数据条预览（负左正右）。
    Excel 文件本身单独可打开；其原生导出图暂不用于 PPT，避免盖掉双向效果。
    """
    path = render_ranking_with_databars(rows, png_path, focus_company=focus)
    return path, "preview_databars"


def run_excel_ppt_loop(
    config_path: str | Path = "config/report_25h1.yaml",
    data_dir: str | Path = "data/sample_25h1",
    template_path: str | Path = "templates/25H1_template.pptx",
    output_pptx: str | Path = "output/25H1_excel_loop.pptx",
    work_dir: str | Path = "output/_excel_loop",
) -> dict:
    config = load_config(config_path)
    df = load_fund_data(data_dir)
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)

    rows = build_ranking_rows(df, config)
    xlsx_path = build_total_ranking_workbook(
        rows, work_dir / "total_ranking.xlsx", focus_company=config.focus_company
    )

    png_path, render_mode = _export_table_image(
        xlsx_path, work_dir / "total_ranking.png", rows, config.focus_company
    )

    # 叙事（复用 phase1）
    from briefing.analytics.engine import compute_company_ranking

    rankings = compute_company_ranking(df[df["company"] != "__industry__"], config, "total")
    narrative = build_total_ranking_payload(rankings, config)

    # 只改第 2 页
    output_pptx = Path(output_pptx)
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(template_path, output_pptx)
    prs = Presentation(str(output_pptx))
    slide = prs.slides[1]

    shape = _find_shape(slide, "矩形 2")
    replace_narrative_paragraphs(
        shape,
        {
            0: narrative["section_title"],
            1: narrative["peer_moves"],
            2: narrative["focus_rank"],
            3: narrative["focus_growth"],
        },
        font_size_pt=12,
    )
    fn = _find_shape(slide, "文本框 5")
    replace_narrative_paragraphs(fn, {0: narrative["footnote_passive"]}, font_size_pt=10)

    # 原图名为「图片 6」
    replace_picture_fit(slide, "图片 6", png_path)
    prs.save(str(output_pptx))

    return {
        "xlsx": str(xlsx_path),
        "png": str(png_path),
        "pptx": str(output_pptx),
        "render_mode": render_mode,
        "n_rows": len(rows),
    }
