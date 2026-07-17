"""报告渲染"""

from __future__ import annotations

from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape

from briefing.models import BriefingReport, CompanyRanking, IncrementBreakdown, ReportSection


def _fmt_num(v: float) -> str:
    return str(round(v))


def _fmt_pct(v: float) -> str:
    return f"{round(v * 100)}%"


def _ranking_to_table(rankings: list[CompanyRanking], extended: bool = False) -> tuple[list[str], list[list[str]]]:
    headers = ["排名", "基金公司", "规模", "增量", "增速%", "排名变化"]
    if extended:
        headers += ["新发", "净值变化", "持营"]
    rows = []
    for r in rankings:
        row = [
            str(r.rank),
            r.company,
            _fmt_num(r.aum),
            _fmt_num(r.increment),
            _fmt_pct(r.growth_pct),
            str(r.rank_change),
        ]
        if extended:
            row += [_fmt_num(r.new_issue), _fmt_num(r.nav_change), _fmt_num(r.holding_sales)]
        rows.append(row)
    return headers, rows


def _increment_to_table(breakdowns: list[IncrementBreakdown]) -> tuple[list[str], list[list[str]]]:
    headers = [
        "排名", "基金公司", "非货增量", "增量排名", "增速%",
        "主动权益增量", "被动权益增量", "固收+增量", "固收增量",
    ]
    rows = []
    for b in breakdowns:
        rows.append([
            str(b.rank),
            b.company,
            _fmt_num(b.increment),
            str(b.increment_rank),
            _fmt_pct(b.growth_pct),
            _fmt_num(b.category_increments.get("active_equity", 0)),
            _fmt_num(b.category_increments.get("passive_equity", 0)),
            _fmt_num(b.category_increments.get("fixed_income_plus", 0)),
            _fmt_num(b.category_increments.get("fixed_income", 0)),
        ])
    return headers, rows


def attach_ranking_table(section: ReportSection, rankings: list[CompanyRanking], extended: bool = False) -> None:
    headers, rows = _ranking_to_table(rankings, extended=extended)
    section.table_headers = headers
    section.table_rows = rows


def attach_increment_table(section: ReportSection, breakdowns: list[IncrementBreakdown]) -> None:
    headers, rows = _increment_to_table(breakdowns)
    section.table_headers = headers
    section.table_rows = rows


def render_html(report: BriefingReport, output_path: str | Path) -> Path:
    template_dir = Path(__file__).parent / "templates"
    env = Environment(
        loader=FileSystemLoader(template_dir),
        autoescape=select_autoescape(["html"]),
    )
    template = env.get_template("briefing.html")
    html = template.render(report=report)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    return out
