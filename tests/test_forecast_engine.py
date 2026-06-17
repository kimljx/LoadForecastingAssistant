from core.forecast_engine import solve_forecast
from core.models import ForecastRules, ForecastTarget, MetricRecord


def rec(metric, year, value):
    return MetricRecord("测试市", "110kV", metric, year, value, "表1 测试市容载比计算", f"A{year}", 1, 1, False, year <= 2025)


def test_ratio_target_generates_load_adjustment():
    records = [
        rec("网供负荷", 2025, 100.0),
        rec("网供负荷", 2030, 150.0),
        rec("变电容量", 2025, 220.0),
        rec("变电容量", 2030, 260.0),
    ]
    target = ForecastTarget("容载比", "测试市", "110kV", year=2030, target_value=2.0)
    rules = ForecastRules(latest_actual_year=2025, forecast_start_year=2026)
    scenarios = solve_forecast(records, [target], rules, forecast_end_year=2030)
    first = scenarios[0]
    assert first.adjustments
    assert any(a.metric == "网供负荷" and a.year == 2030 for a in first.adjustments)
    final_load = [r.value for r in first.forecast_records if r.area == "测试市" and r.voltage_level == "110kV" and r.metric == "网供负荷" and r.year == 2030][0]
    assert abs(final_load - 130.0) < 1e-6
