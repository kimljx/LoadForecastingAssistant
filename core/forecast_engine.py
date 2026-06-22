"""负荷预测与容载比反推引擎。

第二版增强：
- 支持规则项对单个“区域/电压/指标/年份”的可改性、上下限和最大变化比例约束；
- 地市目标可按区县现有占比分摊到区县负荷，避免只改地市表；
- 显式校核地市—区县同时率，范围 0.83-0.99，调整幅度不超过 0.1；
- 区外送(+)/受(-)电作为低优先级兜底变量，并记录待指定站点；
- 生成 2031+ 等新增年份时，保留来源行号，便于导出时扩展年份列写回副本。
"""
from __future__ import annotations

from copy import deepcopy
from typing import Iterable

from .models import Adjustment, ForecastRules, ForecastTarget, MetricRecord, RuleItem, ScenarioResult, TargetResult
from .rule_persistence import metric_rule_target
from .template_parser import records_to_key_map

LOAD_METRIC = "网供负荷"
CAPACITY_METRIC = "变电容量"
CAPACITY_NEED_METRIC = "变电容量需求"
RATIO_METRIC = "容载比"
TRANSFORMER_RATE_METRIC = "配变平均负载率"


def _key(r: MetricRecord) -> tuple[str, str, str, int]:
    """是什么：把指标记录转换为统一索引键。

    为什么：反推过程中需要高频按区域、电压、指标、年份查找记录，统一键可降低查错风险。
    """
    return (r.area, r.voltage_level, r.metric, r.year)


def _get(records: dict[tuple[str, str, str, int], MetricRecord], area: str, voltage: str, metric: str, year: int) -> MetricRecord | None:
    """是什么：从指标字典中安全获取一条记录。

    为什么：很多模板可能缺某年或某指标，统一获取可以避免到处写 try/except。
    """
    return records.get((area, voltage, metric, year))


def _clone_record(record: MetricRecord, year: int, value: float | None) -> MetricRecord:
    # 新增年份没有原始单元格，但保留来源行号，导出时可以写到“同一行 + 新年份列”。
    """是什么：基于已有记录生成新增年份记录。

    为什么：2031 等新增预测年没有原始单元格，但需要继承来源行以便导出扩展列。
    """
    return MetricRecord(
        area=record.area,
        voltage_level=record.voltage_level,
        metric=record.metric,
        year=year,
        value=value,
        source_sheet=record.source_sheet,
        source_cell=record.source_cell if year == record.year else "",
        source_row=record.source_row,
        source_col=record.source_col,
        is_formula=record.is_formula if year == record.year else False,
        is_actual=False,
    )


def _compound_growth(start_value: float, end_value: float, periods: int) -> float | None:
    """是什么：计算阶段复合增长率。

    为什么：增长率目标通常按阶段控制，复合增长率比单年差值更符合预测口径。
    """
    if start_value <= 0 or end_value < 0 or periods <= 0:
        return None
    return (end_value / start_value) ** (1 / periods) - 1


def _safe_delta_pct(old: float | None, new: float) -> float | None:
    """是什么：计算变化比例并处理零值。

    为什么：导出说明要展示变化幅度，但零值不能直接做除法。
    """
    if old in (None, 0):
        return None
    return (new - old) / old


def _find_capacity_metric(voltage: str) -> str:
    """是什么：根据电压等级选择容量指标名称。

    为什么：10kV 关注容量需求，110/35kV 关注变电容量，必须分开处理。
    """
    if voltage == "10kV":
        return CAPACITY_NEED_METRIC
    return CAPACITY_METRIC


def _ratio_value(capacity: float | None, load: float | None, voltage: str) -> float | None:
    """是什么：计算容载比或配变负载率。

    为什么：不同电压等级的关键指标方向不同，集中处理避免公式散落。
    """
    if load in (None, 0) or capacity is None:
        return None
    if voltage == "10kV":
        # 10kV 表中关注配变平均负载率，通常为负荷 / 容量需求。
        return load / capacity if capacity else None
    return capacity / load


def _build_rule_index(rule_items: list[RuleItem] | None) -> dict[str, RuleItem]:
    """是什么：把规则项列表转成字典索引。

    为什么：求解器会频繁查规则，提前索引可保持逻辑清晰和速度稳定。
    """
    return {r.target: r for r in (rule_items or [])}


def _record_rule(rule_index: dict[str, RuleItem], area: str, voltage: str, metric: str, year: int) -> RuleItem | None:
    """是什么：查找某个业务指标对应的用户规则。

    为什么：规则应按业务键匹配而不是按 Excel 地址匹配。
    """
    return rule_index.get(metric_rule_target(area, voltage, metric, year))


def _apply_rule_bounds(old_value: float | None, desired: float, rule: RuleItem | None) -> tuple[float, str]:
    """是什么：按用户规则裁剪建议值，并返回规则说明。

为什么：负荷预测反推有明确业务口径，注释需要说明求解逻辑背后的业务原因。"""
    if rule is None:
        return desired, ""
    if not rule.editable:
        return float(old_value) if old_value is not None else desired, f"规则锁定：{rule.reason or '不可修改'}"
    low = rule.min_value
    high = rule.max_value
    if rule.max_change_pct is not None and old_value not in (None, 0):
        pct = abs(rule.max_change_pct)
        pct_low = float(old_value) * (1 - pct)
        pct_high = float(old_value) * (1 + pct)
        low = pct_low if low is None else max(low, pct_low)
        high = pct_high if high is None else min(high, pct_high)
    new = desired
    msg_parts = []
    if low is not None and new < low:
        new = low
        msg_parts.append(f"受最小值/最大变化比例约束，已抬升至 {low:.4f}")
    if high is not None and new > high:
        new = high
        msg_parts.append(f"受最大值/最大变化比例约束，已压降至 {high:.4f}")
    return new, "；".join(msg_parts)


def _make_adjustment(
    record: MetricRecord,
    old: float | None,
    new: float,
    object_type: str,
    reason: str,
    priority: str,
    risk_level: str,
    rules: ForecastRules,
    rule_msg: str = "",
) -> Adjustment:
    """是什么：把一次数值变化包装成调整项。

    为什么：界面展示、写回 Excel、结果表都依赖统一的调整项结构。
    """
    note = rule_msg
    return Adjustment(
        object_type=object_type,
        area=record.area,
        voltage_level=record.voltage_level,
        metric=record.metric,
        year=record.year,
        old_value=old,
        new_value=new,
        delta=None if old is None else new - float(old),
        delta_pct=_safe_delta_pct(float(old), new) if old is not None else None,
        reason=reason,
        priority=priority,
        risk_level=risk_level,
        source_sheet=record.source_sheet or None,
        source_cell=record.source_cell or None,
        write_back=record.year > rules.latest_actual_year,
        note=note,
        source_row=record.source_row or None,
        source_col=record.source_col or None,
    )


def _select_ratio_target_value(target: ForecastTarget, rules: ForecastRules) -> float:
    """是什么：从目标值或目标区间中取求解用容载比。

    为什么：上级可能给定单点或区间，求解器需要统一得到一个代表值。
    """
    if target.target_value is not None:
        return float(target.target_value)
    if target.min_value is not None and target.max_value is not None:
        return (float(target.min_value) + float(target.max_value)) / 2
    if target.min_value is not None:
        return float(target.min_value)
    if target.max_value is not None:
        return float(target.max_value)
    return (rules.ratio_min + rules.ratio_max) / 2


def _evaluate_target(records_map: dict[tuple[str, str, str, int], MetricRecord], target: ForecastTarget) -> TargetResult:
    """是什么：重新计算方案是否达到用户目标。

    为什么：方案生成后必须给出达标/未达标和偏差，不能只给修改项。
    """
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


def _is_city_like_area(records_map: dict[tuple[str, str, str, int], MetricRecord], area: str, voltage: str, metric: str, year: int) -> bool:
    """是什么：判断某区域是否像“地市上级”而不是普通区县。

    为什么：表格中的上下级关系不一定写在公式里。若把任意区县都当作上级，
    会错误地把其他区县当作它的下级并产生异常同时率。因此这里用“当前区域值 /
    其他同类区域合计”是否接近同时率区间来做保守判断。
    """
    current = _get(records_map, area, voltage, metric, year)
    if not current or current.value is None:
        return False
    candidates = [r for r in records_map.values() if r.voltage_level == voltage and r.metric == metric and r.year == year and r.area != area and r.value is not None]
    if len(candidates) < 2:
        return False
    total = sum(float(r.value or 0) for r in candidates)
    if total <= 0:
        return False
    factor = float(current.value) / total
    return 0.5 <= factor <= 1.5


def _county_records(records_map: dict[tuple[str, str, str, int], MetricRecord], city_area: str, voltage: str, metric: str, year: int) -> list[MetricRecord]:
    """是什么：获取某地市目标下可参与分摊的区县记录。

    为什么：地市负荷变化需要按区县现有占比拆解，方便替代人工分摊。
    """
    out = []
    for r in records_map.values():
        if r.area == city_area:
            continue
        if r.voltage_level == voltage and r.metric == metric and r.year == year and r.value is not None:
            out.append(r)
    return out


def _allocate_city_load_delta_to_counties(
    records_map: dict[tuple[str, str, str, int], MetricRecord],
    city_area: str,
    voltage: str,
    metric: str,
    year: int,
    city_old: float,
    city_new: float,
    rules: ForecastRules,
    rule_index: dict[str, RuleItem],
) -> tuple[list[Adjustment], list[str]]:
    """是什么：把地市目标负荷变化按区县现有占比分摊。

为什么：负荷预测反推有明确业务口径，注释需要说明求解逻辑背后的业务原因。"""
    warnings: list[str] = []
    counties = _county_records(records_map, city_area, voltage, metric, year)
    total = sum(float(r.value or 0) for r in counties)
    if not counties or total <= 0:
        return [], warnings
    old_factor = city_old / total if total else 1.0
    # 若原同时率异常，先用 1.0 分摊，避免人为放大。
    factor_for_delta = old_factor if rules.coincidence_factor_min <= old_factor <= rules.coincidence_factor_max else 1.0
    total_delta = (city_new - city_old) / factor_for_delta
    adjustments: list[Adjustment] = []
    new_total = 0.0
    for rec in counties:
        old = float(rec.value or 0)
        weight = old / total if total else 0
        desired = old + total_delta * weight
        rule = _record_rule(rule_index, rec.area, rec.voltage_level, rec.metric, rec.year)
        desired, rule_msg = _apply_rule_bounds(old, desired, rule)
        if abs(desired - old) < 0.001:
            new_total += old
            continue
        rec.value = desired
        new_total += desired
        adjustments.append(
            _make_adjustment(
                rec,
                old,
                desired,
                "区县负荷分摊",
                f"地市{year}年{voltage}{metric}变化按区县现有占比分摊",
                "高",
                "低" if not rule_msg else "中",
                rules,
                rule_msg,
            )
        )
    if new_total > 0:
        new_factor = city_new / new_total
        if new_factor < rules.coincidence_factor_min or new_factor > rules.coincidence_factor_max:
            warnings.append(f"{year}年{voltage}地市/区县同时率约为 {new_factor:.3f}，超出 {rules.coincidence_factor_min}-{rules.coincidence_factor_max}。")
        if abs(new_factor - old_factor) > rules.coincidence_factor_max_abs_change:
            warnings.append(f"{year}年{voltage}同时率由 {old_factor:.3f} 变为 {new_factor:.3f}，调整幅度超过 {rules.coincidence_factor_max_abs_change}。")
    return adjustments, warnings


def _apply_load_adjustment_for_ratio(
    records_map: dict[tuple[str, str, str, int], MetricRecord],
    target: ForecastTarget,
    rules: ForecastRules,
    scenario: str,
    rule_index: dict[str, RuleItem],
) -> tuple[list[Adjustment], list[str]]:
    """是什么：根据容载比目标反推目标年份网供负荷。

为什么：负荷预测反推有明确业务口径，注释需要说明求解逻辑背后的业务原因。"""
    warnings: list[str] = []
    if target.year is None:
        return [], warnings
    voltage = target.voltage_level
    capacity_metric = _find_capacity_metric(voltage)
    cap = _get(records_map, target.area, voltage, capacity_metric, target.year)
    load = _get(records_map, target.area, voltage, LOAD_METRIC, target.year)
    if not cap or cap.value is None or not load or load.value is None:
        return [], warnings

    desired_ratio = _select_ratio_target_value(target, rules)
    if voltage == "10kV":
        desired_load = float(cap.value) * desired_ratio
    else:
        desired_load = float(cap.value) / desired_ratio

    old_load = float(load.value)
    delta = desired_load - old_load
    if abs(delta) < max(0.001, abs(old_load) * 0.0005):
        return [], warnings

    if scenario == "优先调整负荷":
        raw_new_load = desired_load
        reason = f"根据{target.year}年{voltage}容载比目标 {desired_ratio:.3f} 反推网供负荷"
        priority = "高"
        risk = "低"
    elif scenario == "负荷容量均衡":
        raw_new_load = old_load + delta * 0.5
        reason = "均衡方案：先承担50%负荷调整，其余由容量侧或人工校核"
        priority = "中"
        risk = "中"
    else:
        raw_new_load = old_load + delta * 0.25
        reason = "兜底方案：少量调整负荷，剩余建议通过区外送受电平衡"
        priority = "中"
        risk = "中"

    rule = _record_rule(rule_index, load.area, load.voltage_level, load.metric, load.year)
    new_load, rule_msg = _apply_rule_bounds(old_load, raw_new_load, rule)
    if rule and not rule.editable:
        warnings.append(f"{load.area}-{voltage}-{target.year}年网供负荷被规则锁定，未按容载比目标直接调整。")
        new_load = old_load
    load.value = new_load

    adjustments: list[Adjustment] = []
    if abs(new_load - old_load) > 0.001:
        adjustments.append(
            _make_adjustment(load, old_load, new_load, "负荷", reason, priority, risk if not rule_msg else "中", rules, rule_msg)
        )
        if _is_city_like_area(records_map, target.area, voltage, LOAD_METRIC, target.year):
            county_adjustments, county_warnings = _allocate_city_load_delta_to_counties(
                records_map, target.area, voltage, LOAD_METRIC, target.year, old_load, new_load, rules, rule_index
            )
            adjustments.extend(county_adjustments)
            warnings.extend(county_warnings)

    if scenario == "区外送受电兜底" and rules.allow_external_exchange:
        remaining = desired_load - new_load
        if abs(remaining) > 0.01:
            station = rules.external_exchange_station_name or "待指定站点"
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
                    reason=f"负荷/容量正常调整不足，建议在【{station}】新增区外送(+)/受(-)电作为低优先级兜底变量",
                    priority="很低",
                    risk_level="高",
                    source_sheet=None,
                    source_cell=None,
                    write_back=False,
                    note="需要人工确认站点、送受电方向和调度口径后才能写回原表。",
                )
            )
            load.value = desired_load
    return adjustments, warnings


def _smooth_series_to_target(
    records_map: dict[tuple[str, str, str, int], MetricRecord],
    area: str,
    voltage: str,
    metric: str,
    target_year: int,
    target_value: float,
    rules: ForecastRules,
    rule_index: dict[str, RuleItem],
    object_type: str = "增长率平滑",
) -> list[Adjustment]:
    """是什么：将现状年到目标年的负荷按 CAGR 平滑调整。

为什么：负荷预测反推有明确业务口径，注释需要说明求解逻辑背后的业务原因。"""
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
            rec = MetricRecord(area, voltage, metric, year, new_value, base.source_sheet, "", base.source_row, base.source_col, False, False)
            records_map[_key(rec)] = rec
            old = None
        else:
            old = rec.value
        rule = _record_rule(rule_index, area, voltage, metric, year)
        if rule and not rule.editable:
            prev_value = float(rec.value or prev_value)
            continue
        bounded, rule_msg = _apply_rule_bounds(float(old) if old is not None else None, new_value, rule)
        rec.value = bounded
        if old is not None and abs(bounded - float(old)) < 0.001:
            prev_value = bounded
            continue
        adjustments.append(
            _make_adjustment(
                rec,
                float(old) if old is not None else None,
                bounded,
                object_type,
                f"按{rules.latest_actual_year}-{target_year}年复合增长率 {clamped:.2%} 平滑生成",
                "高",
                "低" if cagr == clamped and not rule_msg else "中",
                rules,
                rule_msg,
            )
        )
        prev_value = bounded
    return adjustments


def _apply_growth_target(
    records_map: dict[tuple[str, str, str, int], MetricRecord],
    target: ForecastTarget,
    rules: ForecastRules,
    rule_index: dict[str, RuleItem],
) -> list[Adjustment]:
    """是什么：按增长率目标生成平滑负荷调整。

    为什么：实际工作最常通过增长率反推未来负荷，需要独立逻辑。
    """
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
    return _smooth_series_to_target(records_map, target.area, target.voltage_level, metric, target.period_end, target_value, rules, rule_index)


def _ensure_future_years(records_map: dict[tuple[str, str, str, int], MetricRecord], rules: ForecastRules, forecast_end_year: int) -> None:
    """是什么：对缺失未来年份做基础预测，便于 2031+ 目标计算。

为什么：负荷预测反推有明确业务口径，注释需要说明求解逻辑背后的业务原因。"""
    combos = sorted({(r.area, r.voltage_level, r.metric) for r in records_map.values()})
    for area, voltage, metric in combos:
        if metric not in {LOAD_METRIC, CAPACITY_METRIC, CAPACITY_NEED_METRIC}:
            continue
        available = [r for r in records_map.values() if r.area == area and r.voltage_level == voltage and r.metric == metric and r.value is not None]
        if not available:
            continue
        available.sort(key=lambda r: r.year)
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


def _simultaneity_warnings(
    before_map: dict[tuple[str, str, str, int], MetricRecord],
    after_map: dict[tuple[str, str, str, int], MetricRecord],
    target_areas: set[str],
    rules: ForecastRules,
) -> list[str]:
    """是什么：检查地市与区县合计形成的同时率风险。

    为什么：很多上下级关系不在公式中体现，必须在业务模型中显式校核。
    """
    warnings: list[str] = []
    years = sorted({r.year for r in after_map.values() if r.metric == LOAD_METRIC})
    voltages = sorted({r.voltage_level for r in after_map.values() if r.metric == LOAD_METRIC})
    for city in target_areas:
        for voltage in voltages:
            if voltage == "总计":
                continue
            for year in years:
                city_rec = _get(after_map, city, voltage, LOAD_METRIC, year)
                if not city_rec or city_rec.value is None:
                    continue
                if not _is_city_like_area(after_map, city, voltage, LOAD_METRIC, year):
                    continue
                counties = _county_records(after_map, city, voltage, LOAD_METRIC, year)
                total = sum(float(r.value or 0) for r in counties)
                if total <= 0:
                    continue
                factor = float(city_rec.value) / total
                if factor < rules.coincidence_factor_min or factor > rules.coincidence_factor_max:
                    warnings.append(f"{city}-{voltage}-{year}年同时率约 {factor:.3f}，超出 {rules.coincidence_factor_min}-{rules.coincidence_factor_max}。")
                old_city = _get(before_map, city, voltage, LOAD_METRIC, year)
                old_counties = _county_records(before_map, city, voltage, LOAD_METRIC, year)
                old_total = sum(float(r.value or 0) for r in old_counties)
                if old_city and old_city.value is not None and old_total > 0:
                    old_factor = float(old_city.value) / old_total
                    if abs(factor - old_factor) > rules.coincidence_factor_max_abs_change:
                        warnings.append(f"{city}-{voltage}-{year}年同时率由 {old_factor:.3f} 调整至 {factor:.3f}，超过单次调整上限 {rules.coincidence_factor_max_abs_change}。")
    return sorted(set(warnings))


def solve_forecast(
    records: Iterable[MetricRecord],
    targets: list[ForecastTarget],
    rules: ForecastRules,
    forecast_end_year: int,
    rule_items: list[RuleItem] | None = None,
) -> list[ScenarioResult]:
    """是什么：生成三类反推方案。

为什么：负荷预测反推有明确业务口径，注释需要说明求解逻辑背后的业务原因。"""
    base_map = records_to_key_map(records)
    _ensure_future_years(base_map, rules, forecast_end_year)
    rule_index = _build_rule_index(rule_items)
    target_areas = {t.area for t in targets if t.enabled}
    scenarios: list[ScenarioResult] = []
    for scenario_name, desc in [
        ("优先调整负荷", "主要通过未来负荷增长率和网供负荷调整满足目标，并按区县占比分摊地市负荷变化。"),
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
                adjs, w = _apply_load_adjustment_for_ratio(records_map, t, rules, scenario_name, rule_index)
                adjustments.extend(adjs)
                warnings.extend(w)
                load = _get(records_map, t.area, t.voltage_level, LOAD_METRIC, t.year or 0)
                if load and load.value is not None and t.year and t.year > rules.latest_actual_year:
                    adjustments.extend(_smooth_series_to_target(records_map, t.area, t.voltage_level, LOAD_METRIC, t.year, float(load.value), rules, rule_index))
            elif t.target_type == "增长率":
                adjustments.extend(_apply_growth_target(records_map, t, rules, rule_index))
            elif t.target_type == "负荷" and t.year is not None and t.target_value is not None:
                adjustments.extend(_smooth_series_to_target(records_map, t.area, t.voltage_level, t.metric or LOAD_METRIC, t.year, t.target_value, rules, rule_index))

        target_results = [_evaluate_target(records_map, t) for t in targets if t.enabled]
        success = all(tr.achieved for tr in target_results) if target_results else True
        score = sum(abs(a.delta or 0) for a in adjustments)
        warnings.extend(_simultaneity_warnings(base_map, records_map, target_areas, rules))
        if any(a.object_type == "区外送受电" for a in adjustments):
            warnings.append("本方案使用区外送受电兜底变量，请人工确认站点、送受电方向和调度口径。")
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
                warnings=sorted(set(warnings)),
            )
        )
    return scenarios
