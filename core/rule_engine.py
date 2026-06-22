"""业务规则加载与校验。"""
from __future__ import annotations

from pathlib import Path
import yaml

from .models import ForecastRules, MetricRecord


def load_rules(path: str | Path = "config/default_rules.yaml") -> ForecastRules:
    """是什么：读取默认业务规则配置。

    为什么：业务边界需要可维护，不能硬编码在求解器里。
    """
    p = Path(path)
    if not p.exists():
        return ForecastRules()
    data = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    model = data.get("model", {})
    ratio = data.get("ratio_rules", {})
    tr = data.get("transformer_load_rate", {})
    cf = data.get("coincidence_factor", {})
    gr = data.get("growth_rules", {})
    pr = data.get("project_rules", {})
    ex = data.get("external_exchange", {})
    return ForecastRules(
        latest_actual_year=int(model.get("latest_actual_year", 2025)),
        forecast_start_year=int(model.get("forecast_start_year", 2026)),
        ratio_min=float(ratio.get("hard_min", 1.3)),
        ratio_max=float(ratio.get("hard_max", 2.5)),
        transformer_load_rate_min=float(tr.get("hard_min", 0.0)),
        transformer_load_rate_max=float(tr.get("hard_max", 1.0)),
        transformer_load_rate_soft_target=float(tr.get("soft_target", 0.5)),
        coincidence_factor_min=float(cf.get("hard_min", 0.83)),
        coincidence_factor_max=float(cf.get("hard_max", 0.99)),
        coincidence_factor_max_abs_change=float(cf.get("max_abs_change", 0.1)),
        annual_growth_min=float(gr.get("annual_min", -0.02)),
        annual_growth_max=float(gr.get("annual_max", 0.12)),
        compound_growth_min=gr.get("compound_min"),
        compound_growth_max=gr.get("compound_max"),
        allow_change_commission_year=bool(pr.get("allow_change_commission_year", False)),
        allow_external_exchange=bool(ex.get("enabled", True)),
        external_exchange_priority=str(ex.get("priority", "低")),
        external_exchange_station_name=str(ex.get("default_station_name", "待指定站点")),
    )


def is_record_editable(record: MetricRecord, rules: ForecastRules) -> bool:
    """是什么：判断指标记录是否可编辑。

当前口径：2025及以前现状数据不可改；项目投产年不在本模型中作为可调变量。

为什么：规则是业务边界入口，注释需要说明为什么这些值会影响是否可改。"""
    if record.year <= rules.latest_actual_year:
        return False
    if record.is_actual:
        return False
    return True


def validate_ratio(value: float | None, rules: ForecastRules) -> tuple[bool, str]:
    """是什么：校验容载比是否在硬边界内。

    为什么：容载比是关键指标，超出上级范围必须明确预警。
    """
    if value is None:
        return False, "无容载比数值"
    if value < rules.ratio_min or value > rules.ratio_max:
        return False, f"容载比 {value:.3f} 超出硬边界 {rules.ratio_min}-{rules.ratio_max}"
    return True, "容载比在合理区间内"


def coincidence_factor_bounds(original: float, rules: ForecastRules) -> tuple[float, float]:
    """是什么：计算同时率在原值和硬边界下的可调范围。

    为什么：同时率允许小幅调整，但不能超过业务边界和最大调整幅度。
    """
    low = max(rules.coincidence_factor_min, original - rules.coincidence_factor_max_abs_change)
    high = min(rules.coincidence_factor_max, original + rules.coincidence_factor_max_abs_change)
    return low, high
