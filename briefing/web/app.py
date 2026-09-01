"""网页操作台：上传底稿 → 配置 → 生成全册 PPT。"""

from __future__ import annotations

import shutil
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

ROOT = Path(__file__).resolve().parents[2]
WEB_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = ROOT / "data" / "_web_uploads"
JOBS_DIR = ROOT / "data" / "_web_jobs"
OUTPUT_DIR = ROOT / "output"
TEMPLATE_H1 = ROOT / "templates" / "25H1_template.pptx"
TEMPLATE_Q4 = ROOT / "templates" / "25Q4_template.pptx"
SLIDE_MAP_H1 = ROOT / "config" / "slide_map.yaml"
SLIDE_MAP_Q4 = ROOT / "config" / "slide_map_q4.yaml"
REPORT_TEMPLATE = ROOT / "config" / "report_template.yaml"


def _template_and_map(period_label: str) -> tuple[Path, Path]:
    """统一使用 Q4 版式（与案例段落槽位一致）；缺文件时回退 H1。"""
    if TEMPLATE_Q4.exists() and SLIDE_MAP_Q4.exists():
        return TEMPLATE_Q4, SLIDE_MAP_Q4
    return TEMPLATE_H1, SLIDE_MAP_H1


app = FastAPI(title="季度简报辅助系统", version="1.0")
app.mount("/static", StaticFiles(directory=WEB_DIR / "static"), name="static")


class GenerateRequest(BaseModel):
    job_id: str
    period_label: str = Field(..., min_length=1)
    period_type: str = "half_year"
    current: str
    previous_quarter: str
    year_start: str
    focus_company: str = "示例基金"
    focus_company_short: str = "示例"


def _ensure_dirs() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    JOBS_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def _job_dir(job_id: str) -> Path:
    return JOBS_DIR / job_id


def _safe_job_id(job_id: str) -> str:
    if not job_id or ".." in job_id or "/" in job_id or "\\" in job_id:
        raise HTTPException(400, "无效的任务 ID")
    return job_id


def _normalize_period(period_label: str, period_type: str, current: str) -> tuple[str, str, str]:
    """返回 (period_type, ytd_tag, quarter_tag)，走 PeriodProfile。"""
    from briefing.period_profile import resolve_period_profile

    profile = resolve_period_profile(period_label, period_type, current)
    return profile.period_type, profile.ytd_tag, profile.quarter_tag


def _write_config(path: Path, req: GenerateRequest) -> None:
    from briefing.period_profile import resolve_period_profile

    raw = yaml.safe_load(REPORT_TEMPLATE.read_text(encoding="utf-8"))
    profile = resolve_period_profile(req.period_label, req.period_type, req.current)
    raw["report"]["period_label"] = req.period_label.strip() or profile.period_label
    raw["report"]["period_type"] = profile.period_type
    raw["dates"] = {
        "current": req.current,
        "previous_quarter": req.previous_quarter,
        "year_start": req.year_start,
    }
    raw["focus_company"] = req.focus_company
    raw["focus_company_short"] = req.focus_company_short
    raw["column_labels"] = {"ytd": profile.ytd_tag, "quarter": profile.quarter_tag}
    path.write_text(
        yaml.safe_dump(raw, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )


def _preview_focus(data_dir: Path, config_path: Path, tables_xlsx: Path | None = None) -> dict[str, Any]:
    """生成后快速核对示例关键指标（优先读终表 Excel，与 PPT 文字同口径）。"""
    from briefing.data_loader import load_config
    from briefing.deck.table_facts import find_focus, load_table_book_from_xlsx

    config = load_config(config_path)
    focus = config.focus_company_short or config.focus_company
    items: list[dict[str, str]] = []

    if tables_xlsx and tables_xlsx.exists():
        book = load_table_book_from_xlsx(tables_xlsx)
        cat = book.get("category")
        ind = next((r for r in (cat.rows if cat else []) if r.name == "合计"), None)
        industry = (
            f"{(ind.aum or 0)/10000:.1f}万亿 · 增速{round((ind.growth_pct or 0)*100)}%"
            if ind and ind.aum is not None
            else "—"
        )

        def add(label: str, sheet_id: str) -> None:
            sh = book.get(sheet_id)
            row = find_focus(sh.rows, config.focus_company, focus) if sh else None
            if not row:
                items.append({"label": label, "text": "未找到"})
                return
            items.append(
                {
                    "label": label,
                    "text": f"#{row.rank} · {row.aum:.0f}亿 · 增量{(row.increment or 0):+.0f}",
                }
            )

        add("总规模", "total")
        add("非货", "non_money")
        add("主动权益", "active_equity")
        add("货币", "money")
        add("固收", "fixed_income")
        add("固收+", "fixed_income_plus")
        add("FOF", "fof")
        return {"industry": industry, "focus": focus, "items": items}

    # 回退 CSV（无终表时）
    from dataclasses import replace

    from briefing.analytics.engine import compute_company_ranking
    from briefing.data_loader import load_fund_data
    from briefing.deck.data_tables import (
        build_business_rows,
        build_category_rows,
        build_non_money_rows,
        build_total_rows,
        company_df,
    )

    df = load_fund_data(data_dir)
    cats = {r["label"]: r for r in build_category_rows(df, config)}

    def find(rows):
        return next((r for r in rows if focus in str(r["company"])), None)

    cos = company_df(df)
    wide = replace(config, top_n=60)
    fi = next((r for r in compute_company_ranking(cos, wide, "fixed_income") if focus in r.company), None)
    fip = next(
        (r for r in compute_company_ranking(cos, wide, "fixed_income_plus") if focus in r.company),
        None,
    )
    total = find(build_total_rows(df, config))
    nm = find(build_non_money_rows(df, config))
    active = find(build_business_rows(df, config, "active_equity"))
    money = find(build_business_rows(df, config, "money"))
    fof = find(build_business_rows(df, config, "fof"))

    def add_row(label: str, row) -> None:
        if not row:
            items.append({"label": label, "text": "未找到"})
            return
        if hasattr(row, "rank"):
            items.append(
                {"label": label, "text": f"#{row.rank} · {row.aum:.0f}亿 · 增量{row.increment:+.0f}"}
            )
        else:
            items.append(
                {
                    "label": label,
                    "text": f"#{row['rank']} · {row['aum']:.0f}亿 · 增量{row['increment']:+.0f}",
                }
            )

    ind = cats.get("合计")
    industry = (
        f"{ind['aum_current']/10000:.1f}万亿 · 增速{ind['growth_pct']*100:.0f}%"
        if ind
        else "—"
    )
    add_row("总规模", total)
    add_row("非货", nm)
    add_row("主动权益", active)
    add_row("货币", money)
    add_row("固收", fi)
    add_row("固收+", fip)
    add_row("FOF", fof)
    return {"industry": industry, "focus": focus, "items": items}


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    html = (WEB_DIR / "templates" / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(html)


_DOC_FILES = {
    "guide": ("出报指引", ROOT / "出报指引.md"),
    "readme": ("README", ROOT / "README.md"),
    "maintain": ("需求与维护说明", ROOT / "docs" / "需求与维护说明.md"),
}


def _render_markdown(path: Path) -> str:
    import mistune

    text = path.read_text(encoding="utf-8")
    md = mistune.create_markdown(plugins=["table", "strikethrough", "url"])
    return md(text)


@app.get("/api/docs/{doc_id}")
def get_doc(doc_id: str) -> dict:
    """返回出报指引 / README 的 HTML，供网页侧栏拉窗展示。"""
    meta = _DOC_FILES.get(doc_id)
    if not meta:
        raise HTTPException(404, f"未知文档：{doc_id}（可选：guide / readme / maintain）")
    title, path = meta
    if not path.exists():
        raise HTTPException(404, f"文件不存在：{path.name}")
    try:
        html = _render_markdown(path)
    except Exception as exc:
        raise HTTPException(500, f"渲染失败：{exc}") from exc
    return {"id": doc_id, "title": title, "filename": path.name, "html": html}


@app.get("/api/health")
def health() -> dict:
    return {"ok": True, "root": str(ROOT)}


@app.post("/api/upload")
async def upload_draft(file: UploadFile = File(...)) -> dict:
    _ensure_dirs()
    name = file.filename or "draft.xlsx"
    if not name.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "请上传 .xlsx 底稿文件")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    job_id = f"job_{stamp}"
    jdir = _job_dir(job_id)
    jdir.mkdir(parents=True, exist_ok=True)

    xlsx_path = jdir / "draft.xlsx"
    content = await file.read()
    xlsx_path.write_bytes(content)

    from briefing.importers.xlsx_draft import detect_dates, import_draft_xlsx
    from briefing.importers.normalize_draft import normalize_draft_xlsx
    from briefing.importers.validate_draft import validate_draft_xlsx

    # 先对原始上传做健全性检查（规范化前），便于作者对照底稿改数
    draft_report = validate_draft_xlsx(xlsx_path)
    (jdir / "draft_validation.yaml").write_text(
        yaml.safe_dump(draft_report.to_dict(), allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )

    try:
        # 上传后先规范化为案例口径，再导入指标（保留 TopN 外的默认关注公司）
        normalize_draft_xlsx(
            xlsx_path,
            xlsx_path,
            focus_company="示例基金",
            focus_short="示例",
        )
        df = import_draft_xlsx(xlsx_path, output_csv=jdir / "fund_metrics.csv")
        dates = detect_dates(xlsx_path)
    except Exception as exc:
        raise HTTPException(400, f"底稿解析失败：{exc}") from exc

    # 期别：优先文件名；否则按检测到的日期推断（如 20260630 → 26H1）
    period_guess = ""
    stem = Path(name).stem
    for token in (
        "26H1",
        "26H2",
        "26Q1",
        "26Q2",
        "26Q3",
        "26Q4",
        "25H1",
        "25H2",
        "25Q1",
        "25Q2",
        "25Q3",
        "25Q4",
        "24H1",
        "24H2",
    ):
        if token in stem.upper().replace(" ", ""):
            period_guess = token
            break
    if not period_guess:
        from briefing.period_profile import guess_period_from_date

        cur = dates.get("current", "")
        if len(cur) >= 6 and cur.isdigit():
            period_guess, period_type = guess_period_from_date(cur)
        else:
            period_guess, period_type = "25H1", "half_year"
    else:
        from briefing.period_profile import resolve_period_profile

        profile = resolve_period_profile(
            period_guess, "", dates.get("current", "")
        )
        period_type = profile.period_type

    return {
        "job_id": job_id,
        "filename": name,
        "rows": int(len(df)),
        "companies": int(df[df["company"] != "__industry__"]["company"].nunique()),
        "categories": sorted(df["category"].unique().tolist()),
        "dates": dates,
        "period_label": period_guess,
        "period_type": period_type,
        "draft_validation": draft_report.to_dict(),
        "message": (
            (
                f"底稿有 {len(draft_report.errors)} 个校验 error，请先改底稿后重新上传；"
                f"当前禁止生成（另有 {len(draft_report.warnings)} 个 warning）"
                if draft_report.errors
                else "底稿导入成功，请确认期别与日期后生成 PPT"
                + (
                    f"；校验有 {len(draft_report.warnings)} 个 warning（可不拦生成）"
                    if draft_report.warnings
                    else ""
                )
            )
        ),
    }


def _require_draft_ok(jdir: Path) -> None:
    """出报前强制：上传时校验无 error，否则禁止生成。"""
    cached = jdir / "draft_validation.yaml"
    if not cached.exists():
        raise HTTPException(400, "缺少底稿校验结果，请重新上传底稿")
    try:
        report_dict = yaml.safe_load(cached.read_text(encoding="utf-8")) or {}
    except Exception as exc:
        raise HTTPException(400, f"无法读取底稿校验结果，请重新上传：{exc}") from exc

    err_n = int(report_dict.get("error_count") or 0)
    if err_n <= 0:
        return

    findings = report_dict.get("findings") or []
    lines = [
        f"底稿仍有 {err_n} 个校验 error，已禁止生成。请按提示改底稿后重新上传。",
    ]
    for f in findings:
        if f.get("severity") != "error":
            continue
        loc = f"[{f.get('sheet')}] " if f.get("sheet") else ""
        lines.append(f"- {loc}{f.get('message', '')}")
        if len(lines) >= 9:
            lines.append("- …（完整列表见 validate-draft）")
            break
    raise HTTPException(400, "\n".join(lines))


@app.post("/api/generate")
def generate(req: GenerateRequest) -> dict:
    _ensure_dirs()
    job_id = _safe_job_id(req.job_id)
    jdir = _job_dir(job_id)
    csv_path = jdir / "fund_metrics.csv"
    if not csv_path.exists():
        raise HTTPException(400, "请先上传并导入底稿")

    _require_draft_ok(jdir)

    template_pptx, slide_map = _template_and_map(req.period_label)
    if not template_pptx.exists():
        raise HTTPException(500, f"缺少 PPT 模版：{template_pptx}")

    config_path = jdir / "report.yaml"
    _write_config(config_path, req)

    out_name = f"{req.period_label}_briefing.pptx"
    xlsx_name = f"{req.period_label}_tables.xlsx"
    out_pptx = OUTPUT_DIR / out_name
    out_xlsx = OUTPUT_DIR / xlsx_name
    work_dir = OUTPUT_DIR / f"_web_{job_id}"

    from briefing.full_deck import run_full_deck

    try:
        info = run_full_deck(
            config_path=config_path,
            data_dir=jdir,
            template_path=template_pptx,
            slide_map_path=slide_map,
            output_pptx=out_pptx,
            output_xlsx=out_xlsx,
            work_dir=work_dir,
            draft_xlsx=jdir / "draft.xlsx",
            paste_tables=False,
            strict=True,
        )
        preview = _preview_focus(jdir, config_path, tables_xlsx=Path(info["xlsx"]) if info.get("xlsx") else out_xlsx)
    except Exception as exc:
        # 网页端只展示可读原因；完整栈写日志式短摘要
        raise HTTPException(500, f"生成失败：{exc}") from exc

    # 中间 PNG 可删；xlsx/pptx 留在 output/
    if work_dir.exists():
        shutil.rmtree(work_dir, ignore_errors=True)

    meta = {
        "job_id": job_id,
        "pptx": str(out_pptx.relative_to(ROOT)),
        "pptx_name": out_name,
        "xlsx": str(Path(info["xlsx"]).relative_to(ROOT)) if info.get("xlsx") else str(out_xlsx.relative_to(ROOT)),
        "xlsx_name": xlsx_name,
        "n_slides": info.get("n_slides"),
        "table_images_from": info.get("table_images_from", "skipped_manual_paste"),
        "narrative_source": info.get("narrative_source"),
        "narrative_report": info.get("narrative_report", ""),
        "preview": preview,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    (jdir / "last_result.yaml").write_text(
        yaml.safe_dump(meta, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    return meta


@app.get("/api/download/{job_id}")
def download(job_id: str, kind: str = "pptx"):
    job_id = _safe_job_id(job_id)
    meta_path = _job_dir(job_id) / "last_result.yaml"
    if not meta_path.exists():
        raise HTTPException(404, "尚未生成文件")
    meta = yaml.safe_load(meta_path.read_text(encoding="utf-8"))
    if kind == "xlsx":
        rel = meta.get("xlsx")
        name = meta.get("xlsx_name", "tables.xlsx")
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        rel = meta.get("pptx")
        name = meta.get("pptx_name", "briefing.pptx")
        media = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    if not rel:
        raise HTTPException(404, f"无 {kind} 产物")
    path = ROOT / rel
    if not path.exists():
        raise HTTPException(404, f"{kind} 文件不存在，请重新生成")
    return FileResponse(path, filename=name, media_type=media)


@app.get("/api/outputs")
def list_outputs() -> dict:
    _ensure_dirs()
    items = []
    pairs: dict[str, dict] = {}
    for p in OUTPUT_DIR.glob("*.pptx"):
        key = p.stem.replace("_briefing", "").replace("_full_deck", "")
        pairs.setdefault(key, {})["pptx"] = p
    for p in OUTPUT_DIR.glob("*_tables.xlsx"):
        key = p.stem.replace("_tables", "")
        pairs.setdefault(key, {})["xlsx"] = p
    # 按最新时间排序
    ranked = []
    for key, files in pairs.items():
        mtimes = [f.stat().st_mtime for f in files.values()]
        ranked.append((max(mtimes), key, files))
    ranked.sort(reverse=True)
    for _, key, files in ranked[:12]:
        pptx = files.get("pptx")
        xlsx = files.get("xlsx")
        st = (pptx or xlsx).stat()
        items.append(
            {
                "key": key,
                "pptx": pptx.name if pptx else None,
                "xlsx": xlsx.name if xlsx else None,
                "size_mb": round(((pptx.stat().st_size if pptx else 0) + (xlsx.stat().st_size if xlsx else 0)) / 1e6, 2),
                "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"),
            }
        )
    return {"items": items}


@app.get("/api/download-file/{name}")
def download_named(name: str):
    if ".." in name or "/" in name or "\\" in name:
        raise HTTPException(400, "无效文件名")
    if not (name.endswith(".pptx") or name.endswith(".xlsx")):
        raise HTTPException(400, "仅支持 pptx/xlsx")
    path = OUTPUT_DIR / name
    if not path.exists():
        raise HTTPException(404, "文件不存在")
    if name.endswith(".xlsx"):
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        media = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    return FileResponse(path, filename=name, media_type=media)


def create_app() -> FastAPI:
    _ensure_dirs()
    return app
