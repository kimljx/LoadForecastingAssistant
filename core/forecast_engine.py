"""负荷预测与容载比反推引擎。

第一版定位：生成可解释的业务方案，不追求一次性替代规划人员决策。
核心实现覆盖：
- 动态年份预测；
- 容载比目标反推网供负荷；
- 增长率约束；
- 负荷优先、均衡、区外送受电兜底三类方案；
- 2025及以前现状数据锁定。
"""
from __future__ import annotations

from copy import deepcopy
from math import isfinite
from statistics import mean
from typing import Iterable

from .models import Adjustment, ForecastRules, ForecastTarget, MetricRecord, ScenarioResult, TargetResult
from .template_parser import records_to_key_map

LOAD_METRIC = "网供负荷"
CAPACITY_METRIC = "变电容量"
CAPACITY_NEED_METRIC = "变电容量需求"
RATIO_METRIC = "容载比"
TRANSFORMER_RATE_METRIC = "配变平均负载率"


def _key(r: MetricRecord) -> tuple[str, str, str, int]:
    return (r.area, r.voltage_level, r.metric, r.year)


def _get(records: dict[tuple[str, str, str, int], MetricRecord], area: str, voltage: str, metric: str, year: int) -> MetricRecord | None:
    return records.get((area, voltage, metric, year))


def _clone_record(record: MetricRecord, year: int, value: float | None) -> MetricRecord:
    return MetricRecord(
        area=record.area,
        voltage_level=record.voltage_level,
        metric=record.metric,
        year=year,
        value=value,
        source_sheet=record.source_sheet,
        source_cell=record.source_cell,
        source_row=record.source_row,
        source_col=record.source_col,
        is_formula=record.is_formula,
        is_actual=False,
    )


def _compound_growth(start_value: float, end_value: float, periods: int) -> float | None:
    if start_value <= 0 or end_value < 0 or periods <= 0:
        return None
    return (end_value / start_value) ** (1 / periods) - 1


def _safe_delta_pct(old: float | None, new: float) -> float | None:
    if old in (None, 0):
        return None
    return (new - old) / old


def forecast_series_by_growth(
    records_map: dict[tuple[str, str, str, int], MetricRecord],
    area: str,
    voltage: str,
    metric: str,
    start_actual_year: int,
    forecast_end_year: int,
    default_growth: float,
) -> list[MetricRecord]:
    """按增长率生成动态年份记录。"""
    base = _get(records_map, area, voltage, metric, start_actual_year)
    if not base or base.value is None:
        # 尝试找小于等于现状年的最后一个可用值。
        candidates = [
            r for k, r in records_map.items()
            if r.area == area and r.voltage_level == voltage and r.metric == metric and r.year <= start_actual_year and r.value is not None
        ]
        if not candidates:
            return []
        base = max(candidates, key=lambda r: r.year)
    value = float(base.value or 0)
    out: list[MetricRecord] = []
    for year in range(base.year + 1, forecast_end_year + 1):
        existing = _get(records_map, area, voltage, metric, year)
        if existing and existing.value is not None:
            value = float(existing.value)
            out.append(existing)
        else:
            value = value * (1 + default_growth)
            out.append(_clone_record(base, year, value))
    return out


def _select_ratio_target_value(target: ForecastTarget, rules: ForecastRules) -> float:
    if target.target_value is not None:
        return float(target.target_value)
    if target.min_value is not None and target.max_value is not None:
        return (float(target.min_value) + float(target.max_value)) / 2
    if target.min_value is not None:
        return float(target.min_value)
    if target.max_value is not None:
        return float(target.max_value)
    return (rules.ratio_min + rules.ratio_max) / 2


def _find_capacity_metric(voltage: str) -> str:
    if voltage == "10kV":
        return CAPACITY_NEED_METRIC
    return CAPACITY_METRIC


def _ratio_value(capacity: float | None, load: float | None, voltage: str) -> float | None:
    if load in (None, 0) or capacity is None:
        return None
    if voltage == "10kV":
        # 10kV 表中关注的是配变平均负载率，通常是负荷 / 容量需求。
        return load / capacity if capacity else None
    return capacity / load


def _evaluate_target(records_map: dict[tuple[str, str, str, int], MetricRecord], target: ForecastTarget) -> TargetResult:
    if not target.enabled:
        return TargetResult(target, None, True, None, "目标未启用")
    if target.target_type == "容载比":
        if target.year is None:
            return TargetResult(target, None, False, None, "容载比目标缺少年份")
        load = _get(records_map, target.area, target.voltage_level, LOAD_METRIC, target.year)
        cap = _get(records_map, target.area, target.voltage_level, _find_capacity_metric(target.voltage_level), target.year)
        actual = _ratio_value(cap.value if cap else None, load.value if load else None, target.voltage_level)
    elif target.target_type == "负荷":
        if target.year is None:
            return TargetResult(target, None, False, None, "负荷目标缺少年份")
        rec = _get(records_map, target.area, target.voltage_level, target.metric or LOAD_METRIC, target.year)
        actual = rec.value if rec else None
    elif target.target_type == "增长率":
        if target.period_start is None or target.period_end is None:
            return TargetResult(target, None, False, None, "增长率目标缺少起止年份")
        start = _get(records_map, target.area, target.voltage_level, target.metric or LOAD_METRIC, target.period_start)
        end = _get(records_map, target.area, target.voltage_level, target.metric or LOAD_METRIC, target.period_end)
        actual = _compound_growth(float(start.value), float(end.value), target.period_end - target.period_start) if start and end and start.value and end.value else None
    else:
        rec = _get(records_map, target.area, target.voltage_level, target.metric or "", target.year or 0)
        actual = rec.value if rec else None

    if actual is None:
        return TargetResult(target, None, False, None, "无法计算目标结果")
    achieved = True
    deviation = None
    if target.target_value is not None:
        deviation = actual - target.target_value
        achieved = abs(deviation) <= max(abs(target.target_value) * 0.002, 0.001)
    if target.min_value is not None and actual < target.min_value:
        achieved = False
        deviation = actual - target.min_value
    if target.max_value is not None and actual > target.max_value:
        achieved = False
        deviation = actual - target.max_value
    return TargetResult(target, actual, achieved, deviation, "达标" if achieved else "未达标")


def _apply_load_adjustment_for_ratio(
    records_map: dict[tuple[str, str, str, int], MetricRecord],
    target: ForecastTarget,
    rules: ForecastRules,
    scenario: str,
) -> list[Adjustment]:
    """根据容载比目标反推目标年份网供负荷。"""
    if target.year is None:
        return []
    voltage = target.voltage_level
    capacity_metric = _find_capacity_metric(voltage)
    cap = _get(records_map, target.area, voltage, capacity_metric, target.year)
    load = _get(records_map, target.area, voltage, LOAD_METRIC, target.year)
    if not cap or cap.value is None or not load or load.value is None:
        return []

    desired_ratio = _select_ratio_target_value(target, rules)
    if voltage == "10kV":
        # 配变平均负载率目标可视为 load / capacity。
        desired_load = float(cap.value) * desired_ratio
    else:
        desired_load = float(cap.value) / desired_ratio

    old_load = float(load.value)
    delta = desired_load - old_load
    if abs(delta) < max(0.001, abs(old_load) * 0.0005):
        return []

    adjustments: list[Adjustment] = []
    if scenario == "优先调整负荷":
        new_load = desired_load
        reason = f"根据{target.year}年{voltage}容载比目标 {desired_ratio:.3f} 反推网供负荷"
        priority = "高"
        risk = "低"
    elif scenario == "负荷容量均衡":
        new_load = old_load + delta * 0.5
        reason = f"均衡方案：先承担50%负荷调整，其余由容量侧或人工校核"
        priority = "中"
        risk = "中"
    else:
        # 兜底方案中先少改负荷，剩余通过区外送受电提示。
        new_load = old_load + delta * 0.25
        reason = "兜底方案：少量调整负荷，剩余建议通过区外送受电平衡"
        priority = "中"
        risk = "中"

    load.value = new_load
    adjustments.append(
        Adjustment(
            object_type="负荷",
            area=target.area,
            voltage_level=voltage,
            metric=LOAD_METRIC,
            year=target.year,
            old_value=old_load,
            new_value=new_load,
            delta=new_load - old_load,
            delta_pct=_safe_delta_pct(old_load, new_load),
            reason=reason,
            priority=priority,
            risk_level=risk,
            source_sheet=load.source_sheet,
            source_cell=load.source_cell,
            write_back=target.year > rules.latest_actual_year,
        )
    )

    if scenario == "区外送受电兜底" and rules.allow_external_exchange:
        remaining = desired_load - new_load
        if abs(remaining) > 0.01:
            adjustments.append(
                Adjustment(
                    object_type="区外送受电",
                    area=target.area,
                    voltage_level=voltage,
                    metric="区外送受电",
                    year=target.year,
                    old_value=0.0,
                    new_value=-remaining,
                    delta=-remaining,
                    delta_pct=None,
                    reason="负荷/容量正常调整不足，新增区外送(+)/受(-)电作为低优先级兜底变量",
                    priority="很低",
                    risk_level="高",
                    source_sheet=None,
                    source_cell=None,
                    write_back=False,
                    note="需要人工指定站点后才能写回原表。",
                )
            )
            # 评价目标达成时，区外送受电被视为净负荷平衡项。
            # 但原表写回仍只写已确认的负荷调整；兜底项需人工指定站点。
            load.value = desired_load
    return adjustments


def _smooth_series_to_target(
    records_map: dict[tuple[str, str, str, int], MetricRecord],
    area: str,
    voltage: str,
    metric: str,
    target_year: int,
    target_value: float,
    rules: ForecastRules,
) -> list[Adjustment]:
    """将现状年到目标年的负荷按 CAGR 平滑调整。"""
    base = _get(records_map, area, voltage, metric, rules.latest_actual_year)
    if not base or base.value is None or base.value <= 0:
        return []
    periods = target_year - rules.latest_actual_year
    if periods <= 0:
        return []
    cagr = _compound_growth(float(base.value), target_value, periods)
    if cagr is None:
        return []
    clamped = min(max(cagr, rules.annual_growth_min), rules.annual_growth_max)
    adjustments: list[Adjustment] = []
    prev_value = float(base.value)
    for year in range(rules.forecast_start_year, target_year + 1):
        rec = _get(records_map, area, voltage, metric, year)
        new_value = prev_value * (1 + clamped)
        if rec is None:
            rec = MetricRecord(area, voltage, metric, year, new_value, "", "", 0, 0, False, False)
            records_map[_key(rec)] = rec
            old = None
        else:
            old = rec.value
            rec.value = new_value
        if old is not None and abs(new_value - float(old)) < 0.001:
            prev_value = new_value
            continue
        adjustments.append(
            Adjustment(
                object_type="增长率平滑",
                area=area,
                voltage_level=voltage,
                metric=metric,
                year=year,
                old_value=float(old) if old is not None else None,
                new_value=new_value,
                delta=None if old is None else new_value - float(old),
                delta_pct=_safe_delta_pct(float(old), new_value) if old is not None else None,
                reason=f"按{rules.latest_actual_year}-{target_year}年复合增长率 {clamped:.2%} 平滑生成",
                priority="高",
                risk_level="低" if cagr == clamped else "中",
                source_sheet=rec.source_sheet or None,
                source_cell=rec.source_cell or None,
                write_back=year > rules.latest_actual_year,
            )
        )
        prev_value = new_value
    return adjustments


def _apply_growth_target(
    records_map: dict[tuple[str, str, str, int], MetricRecord],
    target: ForecastTarget,
    rules: ForecastRules,
) -> list[Adjustment]:
    if target.period_start is None or target.period_end is None:
        return []
    metric = target.metric or LOAD_METRIC
    start = _get(records_map, target.area, target.voltage_level, metric, target.period_start)
    if not start or start.value is None or start.value <= 0:
        return []
    if target.target_value is not None:
        g = target.target_value
    elif target.min_value is not None and target.max_value is not None:
        g = (target.min_value + target.max_value) / 2
    elif target.min_value is not None:
        g = target.min_value
    elif target.max_value is not None:
        g = target.max_value
    else:
        return []
    target_value = float(start.value) * ((1 + g) ** (target.period_end - target.period_start))
    return _smooth_series_to_target(records_map, target.area, target.voltage_level, metric, target.period_end, target_value, rules)


def _ensure_future_years(records_map: dict[tuple[str, str, str, int], MetricRecord], rules: ForecastRules, forecast_end_year: int) -> None:
    """对缺失未来年份做基础预测，便于 2031+ 目标计算。"""
    combos = sorted({(r.area, r.voltage_level, r.metric) for r in records_map.values()})
    for area, voltage, metric in combos:
        if metric not in {LOAD_METRIC, CAPACITY_METRIC, CAPACITY_NEED_METRIC}:
            continue
        available = [r for r in records_map.values() if r.area == area and r.voltage_level == voltage and r.metric == metric and r.value is not None]
        if not available:
            continue
        available.sort(key=lambda r: r.year)
        # 用最近两年估算默认增长，容量默认沿用最后一年。
        if metric in {CAPACITY_METRIC, CAPACITY_NEED_METRIC}:
            default_growth = 0.0
        elif len(available) >= 2 and available[-2].value not in (None, 0):
            default_growth = (available[-1].value / available[-2].value) - 1
            default_growth = min(max(default_growth, rules.annual_growth_min), rules.annual_growth_max)
        else:
            default_growth = 0.03
        last = max(available, key=lambda r: r.year)
        value = float(last.value or 0)
        for year in range(last.year + 1, forecast_end_year + 1):
            if (area, voltage, metric, year) in records_map:
                continue
            value = value * (1 + default_growth)
            records_map[(area, voltage, metric, year)] = _clone_record(last, year, value)


def solve_forecast(
    records: Iterable[MetricRecord],
    targets: list[ForecastTarget],
    rules: ForecastRules,
    forecast_end_year: int,
) -> list[ScenarioResult]:
    """生成三类反推方案。"""
    base_map = records_to_key_map(records)
    _ensure_future_years(base_map, rules, forecast_end_year)
    scenarios: list[ScenarioResult] = []
    for scenario_name, desc in [
        ("优先调整负荷", "主要通过未来负荷增长率和网供负荷调整满足目标。"),
        ("负荷容量均衡", "负荷和容量侧共同承担调整压力，适合目标较紧时参考。"),
        ("区外送受电兜底", "正常调整不足时，提示新增区外送(+)/受(-)电作为低优先级兜底。"),
    ]:
        records_map = deepcopy(base_map)
        adjustments: list[Adjustment] = []
        warnings: list[str] = []
        for t in targets:
            if not t.enabled:
                continue
            if t.target_type == "容载比":
                adjustments.extend(_apply_load_adjustment_for_ratio(records_map, t, rules, scenario_name))
                # 对已得到的目标年份负荷，再向前平滑分年，避免只改目标年。
                load = _get(records_map, t.area, t.voltage_level, LOAD_METRIC, t.year or 0)
                if load and load.value is not None and t.year and t.year > rules.latest_actual_year:
                    adjustments.extend(_smooth_series_to_target(records_map, t.area, t.voltage_level, LOAD_METRIC, t.year, float(load.value), rules))
            elif t.target_type == "增长率":
                adjustments.extend(_apply_growth_target(records_map, t, rules))
            elif t.target_type == "负荷" and t.year is not None and t.target_value is not None:
                adjustments.extend(_smooth_series_to_target(records_map, t.area, t.voltage_level, t.metric or LOAD_METRIC, t.year, t.target_value, rules))

        target_results = [_evaluate_target(records_map, t) for t in targets if t.enabled]
        success = all(tr.achieved for tr in target_results) if target_results else True
        score = sum(abs(a.delta or 0) for a in adjustments)
        if any(a.object_type == "区外送受电" for a in adjustments):
            warnings.append("本方案使用区外送受电兜底变量，请人工确认站点和送受电方向。")
        for a in adjustments:
            if a.year <= rules.latest_actual_year and a.write_back:
                a.write_back = False
                a.note += "现状年锁定，未写回。"
        scenarios.append(
            ScenarioResult(
                name=scenario_name,
                description=desc,
                success=success,
                score=score,
                target_results=target_results,
                adjustments=adjustments,
                forecast_records=list(records_map.values()),
                warnings=warnings,
            )
        )
    return scenarios
