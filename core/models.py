"""核心数据模型。

本项目面向中文负荷预测业务，因此模型字段尽量保留业务含义，
所有 UI 和导出说明也以中文为主。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Optional


class SheetRole(str, Enum):
    """工作表业务角色。"""

    PROJECT_LIBRARY_110_35 = "110/35kV项目库"
    CHECK = "校核表"
    CITY_RATIO = "地市容载比表"
    COUNTY_RATIO = "区县容载比表"
    DIRECT_LOAD = "直供负荷明细"
    SUBSTATION_CAPACITY = "变电容量明细"
    POWER_SOURCE = "电源装机明细"
    STORAGE = "储能装机明细"
    UNKNOWN = "未识别"


@dataclass(slots=True)
class YearColumn:
    """Excel 年份列。"""

    year: int
    column_index: int
    header_row: int
    raw_header: str
    is_actual: bool = False


@dataclass(slots=True)
class SheetInfo:
    """工作表识别结果。"""

    name: str
    role: SheetRole = SheetRole.UNKNOWN
    area_name: str | None = None
    voltage_level: str | None = None
    year_columns: list[YearColumn] = field(default_factory=list)
    confidence: float = 0.0
    notes: list[str] = field(default_factory=list)


@dataclass(slots=True)
class WorkbookInfo:
    """工作簿解析信息。"""

    path: Path
    repaired_path: Path | None
    sheet_infos: dict[str, SheetInfo]
    latest_actual_year: int = 2025
    forecast_start_year: int = 2026


@dataclass(slots=True)
class MetricRecord:
    """长表形式的业务指标。

    一条记录代表某区域、某电压等级、某指标在某一年的值及其来源单元格。
    """

    area: str
    voltage_level: str
    metric: str
    year: int
    value: float | None
    source_sheet: str
    source_cell: str
    source_row: int
    source_col: int
    is_formula: bool = False
    is_actual: bool = False


@dataclass(slots=True)
class ForecastTarget:
    """用户设置的反推目标。"""

    target_type: str  # 容载比 / 增长率 / 负荷 / 容量 / 配变平均负载率
    area: str
    voltage_level: str
    year: int | None = None
    period_start: int | None = None
    period_end: int | None = None
    metric: str | None = None
    target_value: float | None = None
    min_value: float | None = None
    max_value: float | None = None
    enabled: bool = True
    remark: str = ""


@dataclass(slots=True)
class ForecastRules:
    """默认业务规则，可由 UI 编辑覆盖。"""

    latest_actual_year: int = 2025
    forecast_start_year: int = 2026
    ratio_min: float = 1.3
    ratio_max: float = 2.5
    transformer_load_rate_min: float = 0.0
    transformer_load_rate_max: float = 1.0
    transformer_load_rate_soft_target: float = 0.5
    coincidence_factor_min: float = 0.83
    coincidence_factor_max: float = 0.99
    coincidence_factor_max_abs_change: float = 0.1
    annual_growth_min: float = -0.02
    annual_growth_max: float = 0.12
    compound_growth_min: float | None = None
    compound_growth_max: float | None = None
    allow_change_commission_year: bool = False
    allow_external_exchange: bool = True
    external_exchange_priority: str = "低"


@dataclass(slots=True)
class Adjustment:
    """一条建议调整。"""

    object_type: str
    area: str
    voltage_level: str
    metric: str
    year: int
    old_value: float | None
    new_value: float
    delta: float | None
    delta_pct: float | None
    reason: str
    priority: str
    risk_level: str
    source_sheet: str | None = None
    source_cell: str | None = None
    write_back: bool = True
    note: str = ""


@dataclass(slots=True)
class TargetResult:
    """目标达成结果。"""

    target: ForecastTarget
    actual_value: float | None
    achieved: bool
    deviation: float | None
    message: str


@dataclass(slots=True)
class ScenarioResult:
    """一套反推方案。"""

    name: str
    description: str
    success: bool
    score: float
    target_results: list[TargetResult] = field(default_factory=list)
    adjustments: list[Adjustment] = field(default_factory=list)
    forecast_records: list[MetricRecord] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class CompatibilityIssue:
    """兼容性检测问题。"""

    level: str  # error / warning / info
    sheet: str
    location: str
    issue_type: str
    content: str
    suggestion: str
