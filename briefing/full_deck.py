"""全册 PPT 灌版：10 页文字 + 表图；并输出可编辑的全册 Excel 表图。"""

from __future__ import annotations

import shutil
from pathlib import Path

from pptx import Presentation

from briefing.data_loader import load_config, load_fund_data
from briefing.deck.data_tables import (
    build_business_rows,
    build_category_rows,
    build_increment_rows,
    build_non_money_rows,
    build_total_rows,
    cols_business,
    cols_category,
    cols_increment,
    cols_non_money,
    cols_total,
)
from briefing.deck.narratives import build_all_narratives
from briefing.excel.deck_workbook import build_deck_workbook
from briefing.excel.styled_table import render_styled_table
from briefing.render.pptx_filler import (
    _find_shape,
    load_slide_map,
    replace_narrative_paragraphs,
    replace_picture_fit,
)


def _fill_slide_text(slide, shape_name: str, para_map: dict[int, str], font_size: float = 12) -> None:
    shape = _find_shape(slide, shape_name)
    replace_narrative_paragraphs(shape, para_map, font_size_pt=font_size)


def _business_dict_rows(rankings) -> list[dict]:
    return [
        {
            "rank": r.rank,
            "company": r.company,
            "aum": r.aum,
            "increment": r.increment,
            "growth_pct": r.growth_pct,
            "new_issue": r.new_issue,
            "nav_change": r.nav_change,
            "holding_sales": r.holding_sales,
            "rank_change": r.rank_change,
            "q_increment": getattr(r, "q_increment", 0),
            "q_growth_pct": getattr(r, "q_growth_pct", 0),
            "rank_change_q": getattr(r, "rank_change_q", 0),
        }
        for r in rankings
    ]


def _default_xlsx_path(output_pptx: Path) -> Path:
    stem = output_pptx.stem
    for suffix in ("_briefing", "_full_deck"):
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)]
            break
    return output_pptx.with_name(f"{stem}_tables.xlsx")


def run_full_deck(
    config_path: str | Path = "config/report_25h1.yaml",
    data_dir: str | Path = "data/sample_25h1",
    template_path: str | Path = "templates/25H1_template.pptx",
    slide_map_path: str | Path = "config/slide_map.yaml",
    output_pptx: str | Path = "output/25H1_full_deck.pptx",
    output_xlsx: str | Path | None = None,
    work_dir: str | Path = "output/_full_deck",
    draft_xlsx: str | Path | None = None,
) -> dict:
    config = load_config(config_path)
    df = load_fund_data(data_dir)
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    slide_map = load_slide_map(slide_map_path)
    narratives = build_all_narratives(df, config)
    focus = config.focus_company_short or config.focus_company

    images: dict[str, Path] = {}
    excel_sheets: list[tuple] = []

    draft_path = Path(draft_xlsx) if draft_xlsx else None
    if draft_path and draft_path.exists():
        from briefing.deck.final_tables import load_final_tables

        excel_sheets = load_final_tables(draft_path)
        image_key = {
            "01_行业变化": "category_summary",
            "02_总规模": "total_ranking_top30",
            "03_非货": "non_money_ranking",
            "04_非货增量": "increment_overview",
            "05_主动权益": "active_equity",
            "06_非货ETF含联接": "etf_with_link",
            "06_权益ETF含联接": "etf_no_link",
            "07_货币": "money",
            "08_固收": "fixed_income",
            "09_固收+": "fixed_income_plus",
            "10_FOF": "fof",
        }
        for title, cols, rows in excel_sheets:
            key = image_key.get(title)
            if not key:
                continue
            focus_for = "__none__" if title.startswith("01_") else focus
            safe = title.replace("+", "plus").replace("/", "_")
            images[key] = render_styled_table(
                cols, rows, work_dir / f"{safe}.png", focus_company=focus_for
            )
    else:
        c_cat, c_tot, c_nm, c_inc, c_biz = (
            cols_category(config),
            cols_total(config),
            cols_non_money(config),
            cols_increment(config),
            cols_business(config),
        )

        cat_rows = build_category_rows(df, config)
        images["category_summary"] = render_styled_table(
            c_cat, cat_rows, work_dir / "01_category.png", focus_company="__none__"
        )
        excel_sheets.append(("01_行业变化", c_cat, cat_rows))

        total_rows = build_total_rows(df, config)
        images["total_ranking_top30"] = render_styled_table(
            c_tot, total_rows, work_dir / "02_total.png", focus_company=focus
        )
        excel_sheets.append(("02_总规模", c_tot, total_rows))

        nm_rows = build_non_money_rows(df, config)
        images["non_money_ranking"] = render_styled_table(
            c_nm, nm_rows, work_dir / "03_non_money.png", focus_company=focus
        )
        excel_sheets.append(("03_非货", c_nm, nm_rows))

        inc_rows = build_increment_rows(df, config)
        images["increment_overview"] = render_styled_table(
            c_inc, inc_rows, work_dir / "04_increment.png", focus_company=focus
        )
        excel_sheets.append(("04_非货增量", c_inc, inc_rows))

        active_rows = build_business_rows(df, config, "active_equity")
        images["active_equity"] = render_styled_table(
            c_biz, active_rows, work_dir / "05_active.png", focus_company=focus
        )
        excel_sheets.append(("05_主动权益", c_biz, active_rows))

        etf_w_rows = build_business_rows(df, config, "passive_equity_etf_with_link")
        etf_n_rows = build_business_rows(df, config, "passive_equity_etf_no_link")
        images["etf_with_link"] = render_styled_table(
            c_biz, etf_w_rows, work_dir / "06_etf_with.png", focus_company=focus
        )
        images["etf_no_link"] = render_styled_table(
            c_biz, etf_n_rows, work_dir / "06_etf_no.png", focus_company=focus
        )
        excel_sheets.append(("06_ETF含联接", c_biz, etf_w_rows))
        excel_sheets.append(("06_ETF不含联接", c_biz, etf_n_rows))

        for key, cat, fname, sheet_title in [
            ("money", "money", "07_money.png", "07_货币"),
            ("fixed_income", "fixed_income", "08_fixed_income.png", "08_固收"),
            ("fixed_income_plus", "fixed_income_plus", "09_fixed_income_plus.png", "09_固收+"),
            ("fof", "fof", "10_fof.png", "10_FOF"),
        ]:
            rows = build_business_rows(df, config, cat)
            if key == "fixed_income" and narratives.get("_fi_table_rankings"):
                rows = _business_dict_rows(narratives["_fi_table_rankings"])
            if key == "fixed_income_plus" and narratives.get("_fip_table_rankings"):
                rows = _business_dict_rows(narratives["_fip_table_rankings"])
            images[key] = render_styled_table(
                c_biz, rows, work_dir / fname, focus_company=focus
            )
            excel_sheets.append((sheet_title, c_biz, rows))

    output_pptx = Path(output_pptx)
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    xlsx_path = Path(output_xlsx) if output_xlsx else _default_xlsx_path(output_pptx)
    xlsx_path = build_deck_workbook(
        excel_sheets,
        xlsx_path,
        period_label=config.period_label,
        focus_company=focus,
    )

    shutil.copy(template_path, output_pptx)
    prs = Presentation(str(output_pptx))

    for slide_cfg in slide_map.get("slides", []):
        idx = slide_cfg["index"]
        sid = slide_cfg["id"]
        slide = prs.slides[idx]
        payload = narratives.get(sid, {})

        if "narrative_shape" in slide_cfg and "paragraphs" in slide_cfg:
            para_map = {}
            for para_idx, key in slide_cfg["paragraphs"].items():
                if key in payload:
                    para_map[int(para_idx)] = payload[key]
            if para_map:
                _fill_slide_text(slide, slide_cfg["narrative_shape"], para_map, font_size=11)

        for extra in slide_cfg.get("extra_text", []):
            if extra["key"] in payload:
                _fill_slide_text(slide, extra["shape"], {0: payload[extra["key"]]}, font_size=11)

        if "footnote_shape" in slide_cfg:
            fk = slide_cfg.get("footnote_key", "footnote")
            if fk in payload:
                _fill_slide_text(slide, slide_cfg["footnote_shape"], {0: payload[fk]}, font_size=9)

        if "picture_shape" in slide_cfg and "table_image" in slide_cfg:
            img_key = slide_cfg["table_image"]
            if img_key in images:
                replace_picture_fit(slide, slide_cfg["picture_shape"], images[img_key])

        for pic in slide_cfg.get("pictures", []):
            if pic["table_image"] in images:
                replace_picture_fit(slide, pic["shape"], images[pic["table_image"]])

    prs.save(str(output_pptx))
    return {
        "pptx": str(output_pptx),
        "xlsx": str(xlsx_path),
        "work_dir": str(work_dir),
        "images": {k: str(v) for k, v in images.items()},
        "n_slides": len(slide_map.get("slides", [])),
    }
