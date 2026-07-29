"""数据模型定义"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class PeriodType(str, Enum):
    QUARTER = "quarter"
    HALF_YEAR = "half_year"
    YEAR = "year"


@dataclass
class ReportConfig:
    title: str
    period_label: str
    period_type: PeriodType
    data_source_note: str
    current_date: str
    previous_quarter_date: str
    year_start_date: str
    focus_company: str
    focus_company_short: str
    top_n: int = 30
    categories: list[dict] = field(default_factory=list)
    sections: list[dict] = field(default_factory=list)
    # 表头前缀，如 ytd="25年" / quarter="Q4"；空则按 period_type/日期推导
    ytd_column_label: str = ""
    quarter_column_label: str = ""

    def period_profile(self):
        """统一期别档案（Q1 单口径 / Q3 单季主导 / H1·Q4 期别主导）。"""
        from briefing.period_profile import resolve_period_profile

        return resolve_period_profile(
            self.period_label,
            self.period_type.value if isinstance(self.period_type, PeriodType) else str(self.period_type),
            self.current_date,
        )

    def ytd_tag(self) -> str:
        if self.ytd_column_label:
            return self.ytd_column_label
        return self.period_profile().ytd_tag

    def quarter_tag(self) -> str:
        if self.quarter_column_label:
            return self.quarter_column_label
        return self.period_profile().quarter_tag

    def year_end_ref(self) -> str:
        """相对年初口径的口语参照，如 24年底。"""
        y = int(self.year_start_date[:4])
        # year_start 若为年末时点（如 20241231），本身即对照年
        if self.year_start_date[4:8] == "1231":
            return f"{str(y)[2:]}年底"
        return f"{str(y - 1)[2:]}年底"

    def prev_quarter_ref(self) -> str:
        """上季度口语参照，如 25Q3。"""
        y = self.previous_quarter_date[2:4]
        m = int(self.previous_quarter_date[4:6])
        return f"{y}Q{(m - 1) // 3 + 1}"


@dataclass
class CategoryMetrics:
    """一级大类指标"""

    category: str
    label: str
    aum_current: float
    aum_prev_quarter: float
    aum_year_start: float
    increment: float
    growth_pct: float
    new_issue: float
    nav_change: float
    holding_sales: float
    q_increment: float = 0.0
    q_growth_pct: float = 0.0


@dataclass
class CompanyRanking:
    """公司排名行"""

    rank: int
    company: str
    aum: float
    increment: float
    growth_pct: float
    rank_change: int  # 相对年初
    rank_change_q: int  # 相对上季度
    sub_rankings: dict[str, int] = field(default_factory=dict)
    sub_aums: dict[str, float] = field(default_factory=dict)
    new_issue: float = 0
    nav_change: float = 0
    holding_sales: float = 0
    q_increment: float = 0.0
    q_growth_pct: float = 0.0


@dataclass
class IncrementBreakdown:
    """非货增量拆解"""

    rank: int
    company: str
    increment: float
    increment_rank: int
    growth_pct: float
    rank_change: int
    category_increments: dict[str, float] = field(default_factory=dict)
    category_increment_ranks: dict[str, int] = field(default_factory=dict)
    q_increment: float = 0.0
    q_growth_pct: float = 0.0
    rank_change_q: int = 0


@dataclass
class NarrativeBlock:
    """叙事文本块"""

    section_id: str
    paragraphs: list[str]
    bullets: list[str] = field(default_factory=list)


@dataclass
class ReportSection:
    """报告章节"""

    id: str
    title: str
    narrative: NarrativeBlock
    table_headers: list[str]
    table_rows: list[list[str]]
    footnotes: list[str] = field(default_factory=list)
    subsections: list[ReportSection] = field(default_factory=list)


@dataclass
class BriefingReport:
    """完整简报"""

    config: ReportConfig
    sections: list[ReportSection]
    generated_at: Optional[str] = None
