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
TEMPLATE_PPTX = ROOT / "templates" / "25H1_template.pptx"
SLIDE_MAP = ROOT / "config" / "slide_map.yaml"
REPORT_TEMPLATE = ROOT / "config" / "report_template.yaml"

app = FastAPI(title="季度简报辅助系统", version="1.0")
app.mount("/static", StaticFiles(directory=WEB_DIR / "static"), name="static")


class GenerateRequest(BaseModel):
    job_id: str
    period_label: str = Field(..., min_length=1)
    period_type: str = "half_year"
    current: str
    previous_quarter: str
    year_start: str
    focus_company: str = "银华基金"
    focus_company_short: str = "银华"


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


def _write_config(path: Path, req: GenerateRequest) -> None:
    raw = yaml.safe_load(REPORT_TEMPLATE.read_text(encoding="utf-8"))
    raw["report"]["period_label"] = req.period_label
    raw["report"]["period_type"] = req.period_type
    raw["dates"] = {
        "current": req.current,
        "previous_quarter": req.previous_quarter,
        "year_start": req.year_start,
    }
    raw["focus_company"] = req.focus_company
    raw["focus_company_short"] = req.focus_company_short
    yy = req.current[2:4] if len(req.current) >= 4 else "25"
    month = int(req.current[4:6]) if len(req.current) >= 6 else 12
    q_tag = f"Q{(month - 1) // 3 + 1}"
    ytd_tag = req.period_label if req.period_type == "half_year" else f"{yy}年"
    raw["column_labels"] = {"ytd": ytd_tag, "quarter": q_tag}
    path.write_text(
        yaml.safe_dump(raw, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )


def _preview_focus(data_dir: Path, config_path: Path) -> dict[str, Any]:
    """生成后快速核对银华关键指标。"""
    from dataclasses import replace

    from briefing.analytics.engine import compute_company_ranking
    from briefing.data_loader import load_config, load_fund_data
    from briefing.deck.data_tables import (
        build_business_rows,
        build_category_rows,
        build_non_money_rows,
        build_total_rows,
        company_df,
    )

    config = load_config(config_path)
    df = load_fund_data(data_dir)
    cats = {r["label"]: r for r in build_category_rows(df, config)}
    focus = config.focus_company_short or config.focus_company

    def find(rows):
        return next((r for r in rows if focus in str(r["company"])), None)

    cos = company_df(df)
    wide = replace(config, top_n=60)
    fi = next((r for r in compute_company_ranking(cos, wide, "fixed_income") if focus in r.company), None)
    fip = next(
        (r for r in compute_company_ranking(cos, wide, "fixed_income_plus") if focus in r.company),
        None,
    )

    items = []
    total = find(build_total_rows(df, config))
    nm = find(build_non_money_rows(df, config))
    active = find(build_business_rows(df, config, "active_equity"))
    money = find(build_business_rows(df, config, "money"))
    fof = find(build_business_rows(df, config, "fof"))

    def add(label: str, row, aum=True):
        if not row:
            items.append({"label": label, "text": "未找到"})
            return
        if hasattr(row, "rank"):
            items.append(
                {
                    "label": label,
                    "text": f"#{row.rank} · {row.aum:.0f}亿 · 增量{row.increment:+.0f}",
                }
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
    add("总规模", total)
    add("非货", nm)
    add("主动权益", active)
    add("货币", money)
    add("固收", fi)
    add("固收+", fip)
    add("FOF", fof)
    return {"industry": industry, "focus": focus, "items": items}


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    html = (WEB_DIR / "templates" / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(html)


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

    try:
        df = import_draft_xlsx(xlsx_path, output_csv=jdir / "fund_metrics.csv")
        dates = detect_dates(xlsx_path)
    except Exception as exc:
        raise HTTPException(400, f"底稿解析失败：{exc}") from exc

    # 从文件名猜测期别
    period_guess = "25H1"
    stem = Path(name).stem
    for token in ("25H1", "25H2", "25Q1", "25Q2", "25Q3", "25Q4", "24H1", "24H2"):
        if token in stem.upper().replace(" ", ""):
            period_guess = token
            break

    period_type = "half_year" if "H" in period_guess else "quarter"

    return {
        "job_id": job_id,
        "filename": name,
        "rows": int(len(df)),
        "companies": int(df[df["company"] != "__industry__"]["company"].nunique()),
        "categories": sorted(df["category"].unique().tolist()),
        "dates": dates,
        "period_label": period_guess,
        "period_type": period_type,
        "message": "底稿导入成功，请确认期别与日期后生成 PPT",
    }


@app.post("/api/generate")
def generate(req: GenerateRequest) -> dict:
    _ensure_dirs()
    job_id = _safe_job_id(req.job_id)
    jdir = _job_dir(job_id)
    csv_path = jdir / "fund_metrics.csv"
    if not csv_path.exists():
        raise HTTPException(400, "请先上传并导入底稿")

    if not TEMPLATE_PPTX.exists():
        raise HTTPException(500, f"缺少 PPT 模版：{TEMPLATE_PPTX}")

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
            template_path=TEMPLATE_PPTX,
            slide_map_path=SLIDE_MAP,
            output_pptx=out_pptx,
            output_xlsx=out_xlsx,
            work_dir=work_dir,
            draft_xlsx=jdir / "draft.xlsx",
        )
        preview = _preview_focus(jdir, config_path)
    except Exception as exc:
        tb = traceback.format_exc(limit=5)
        raise HTTPException(500, f"生成失败：{exc}\n{tb}") from exc

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
