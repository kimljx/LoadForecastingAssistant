"""规则档案持久化与模板匹配校验。

这个模块回答两个问题：
1. 是什么：把界面中的规则、目标预设、模板结构签名保存为 YAML 文件。
2. 为什么：负荷预测模板经常按年度、版本复制，业务人员不应每次重新配置
   容载比、增长率、同时率等规则；但规则也不能盲目套用到结构差异很大的表，
   因此保存时必须带上模板主名称和结构签名，加载时必须做匹配提示。
"""
from __future__ import annotations

import hashlib
import re
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import yaml

from .models import ForecastRules, ForecastTarget, MetricRecord, RuleItem, RuleItemValidation, RuleProfileMatchResult, WorkbookInfo

RULES_DIR = Path("rules")

# 常见预算模板命名会带日期、版本、定稿、反推结果等后缀。
# 规则文件要按“模型主名称”自动匹配，所以先把这些非业务后缀去掉。
_SUFFIX_PATTERNS = [
    r"[_\- ]?\d{8}$",
    r"[_\- ]?\d{4}[-_.年]\d{1,2}([-.月]\d{1,2}日?)?$",
    r"[_\- ]?[vV]\d+$",
    r"[_\- ]?版本\d+$",
    r"[_\- ]?(最终版|定稿|副本|复制|最新版|修改版|测算版)$",
    r"[_\- ]?反推结果[_\- ]?\d{8}[_\- ]?\d{6}$",
    r"[_\- ]?反推结果$",
]


def normalize_workbook_name(filename: str | Path) -> str:
    """是什么：从 Excel 文件名中提取稳定的模型主名称。

    为什么：同一个预测模型经常出现“_20240601”“_最终版”“_反推结果”
    等尾缀。如果直接按完整文件名找规则，会导致同一模型的规则无法复用。
    """
    stem = Path(filename).stem.strip()
    old = None
    while old != stem:
        old = stem
        for pattern in _SUFFIX_PATTERNS:
            stem = re.sub(pattern, "", stem).strip(" _-")
    return stem or Path(filename).stem


def default_rule_path(filename: str | Path, rules_dir: str | Path = RULES_DIR) -> Path:
    """是什么：根据工作簿主名称生成默认规则文件路径。

    为什么：用户打开同主名模板时应自动找到对应规则档案，减少重复配置。
    """
    base = normalize_workbook_name(filename)
    return Path(rules_dir) / f"{base}_规则.yaml"


# UI 第二版使用的名称，保留别名避免调用层关心历史命名。
def default_rule_path_for_workbook(filename: str | Path, rules_dir: str | Path = RULES_DIR) -> Path:
    """是什么：default_rule_path 的语义化别名。

    为什么：界面代码读起来更像“为当前工作簿找默认规则”，避免开发人员误解。
    """
    return default_rule_path(filename, rules_dir)


def metric_rule_target(area: str, voltage_level: str, metric: str, year: int) -> str:
    """是什么：生成单个业务指标规则的稳定键。

    为什么：规则不应绑定 Excel 单元格地址，因为模板新增年份列或移动列后，
    单元格地址会变化；区域、电压等级、指标、年份组成的业务键更稳定。
    """
    return f"指标|{area}|{voltage_level}|{metric}|{year}"


def rules_to_dict(rules: ForecastRules) -> dict[str, Any]:
    """是什么：把 ForecastRules 转成 YAML 可保存的字典结构。

    为什么：保持规则文件可读、可人工审查，也方便后续从桌面端迁移到数据库。
    """
    return {
        "model": {
            "latest_actual_year": rules.latest_actual_year,
            "forecast_start_year": rules.forecast_start_year,
            "dynamic_years": True,
        },
        "ratio_rules": {
            "hard_min": rules.ratio_min,
            "hard_max": rules.ratio_max,
        },
        "transformer_load_rate": {
            "hard_min": rules.transformer_load_rate_min,
            "hard_max": rules.transformer_load_rate_max,
            "soft_target": rules.transformer_load_rate_soft_target,
        },
        "coincidence_factor": {
            "hard_min": rules.coincidence_factor_min,
            "hard_max": rules.coincidence_factor_max,
            "max_abs_change": rules.coincidence_factor_max_abs_change,
            "editable": True,
            "priority": "low",
        },
        "growth_rules": {
            "annual_min": rules.annual_growth_min,
            "annual_max": rules.annual_growth_max,
            "compound_min": rules.compound_growth_min,
            "compound_max": rules.compound_growth_max,
        },
        "project_rules": {
            "allow_change_commission_year": rules.allow_change_commission_year,
            "allow_change_commission_month": False,
        },
        "external_exchange": {
            "enabled": rules.allow_external_exchange,
            "priority": rules.external_exchange_priority,
            "allow_new_item": True,
            "require_user_confirmation": True,
            "send_positive_receive_negative": True,
            "default_station_name": rules.external_exchange_station_name,
        },
    }


def forecast_rules_from_dict(data: dict[str, Any] | None) -> ForecastRules:
    """是什么：把规则 YAML 中的字典恢复为 ForecastRules。

    为什么：规则档案可能来自旧版本，字段不完整时要用默认值兜底，不能因为
    缺一个字段就导致规则无法加载。
    """
    data = data or {}
    model = data.get("model", {}) or {}
    ratio = data.get("ratio_rules", {}) or {}
    tr = data.get("transformer_load_rate", {}) or {}
    cf = data.get("coincidence_factor", {}) or {}
    gr = data.get("growth_rules", {}) or {}
    pr = data.get("project_rules", {}) or {}
    ex = data.get("external_exchange", {}) or {}
    return ForecastRules(
        latest_actual_year=int(model.get("latest_actual_year", 2025)),
        forecast_start_year=int(model.get("forecast_start_year", int(model.get("latest_actual_year", 2025)) + 1)),
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


def target_to_dict(target: ForecastTarget) -> dict[str, Any]:
    """是什么：序列化一个目标预设。

    为什么：很多反推任务会反复使用同一组目标，例如“2030 年容载比范围 +
    十五五增长率范围”，保存预设可以减少重复录入。
    """
    return asdict(target)


def target_from_dict(data: dict[str, Any]) -> ForecastTarget:
    """是什么：从 YAML 字典恢复一个 ForecastTarget。

    为什么：规则档案中的目标预设需要重新加载到桌面界面和求解引擎。
    """
    return ForecastTarget(
        target_type=str(data.get("target_type", "容载比")),
        area=str(data.get("area", "")),
        voltage_level=str(data.get("voltage_level", "")),
        year=data.get("year"),
        period_start=data.get("period_start"),
        period_end=data.get("period_end"),
        metric=data.get("metric"),
        target_value=data.get("target_value"),
        min_value=data.get("min_value"),
        max_value=data.get("max_value"),
        enabled=bool(data.get("enabled", True)),
        remark=str(data.get("remark", "")),
    )


def rule_item_to_dict(item: RuleItem) -> dict[str, Any]:
    """是什么：序列化单条指标规则。

    为什么：第二版的规则体系不仅保存全局边界，也要为后续“某区域某年不可改、
    某指标最大调整比例”等细粒度规则预留稳定文件格式。
    """
    return asdict(item)


def rule_item_from_dict(data: dict[str, Any]) -> RuleItem:
    """是什么：从 YAML 恢复单条指标规则。

    为什么：规则项级别匹配、标红、删除、编辑都需要结构化对象，而不是散乱字典。
    """
    return RuleItem(
        target=str(data.get("target") or metric_rule_target(data.get("area", ""), data.get("voltage_level", ""), data.get("metric", ""), int(data.get("year") or 0))),
        area=str(data.get("area", "")),
        voltage_level=str(data.get("voltage_level", "")),
        metric=str(data.get("metric", "")),
        year=data.get("year"),
        editable=bool(data.get("editable", True)),
        min_value=data.get("min_value"),
        max_value=data.get("max_value"),
        max_change_pct=data.get("max_change_pct"),
        reason=str(data.get("reason", "")),
        source=str(data.get("source", "用户规则")),
    )


def build_workbook_signature(workbook_info: WorkbookInfo, records: Iterable[MetricRecord] | None = None) -> dict[str, Any]:
    """是什么：生成当前工作簿的结构签名。

    为什么：只看文件名不安全；同名模板可能已经改了 sheet、年份列或关键指标。
    结构签名用于加载规则时判断“完全匹配 / 结构变化 / 不匹配”。
    """
    sheet_rows = []
    for sheet in sorted(workbook_info.sheet_infos.values(), key=lambda x: x.name):
        sheet_rows.append(
            {
                "sheet": sheet.name,
                "role": sheet.role.value,
                "area": sheet.area_name,
                "voltage": sheet.voltage_level,
                "years": [yc.year for yc in sheet.year_columns],
            }
        )
    metric_rows = []
    if records is not None:
        for r in sorted(records, key=lambda x: (x.area, x.voltage_level, x.metric, x.year, x.source_sheet, x.source_row, x.source_col)):
            metric_rows.append(
                {
                    "area": r.area,
                    "voltage": r.voltage_level,
                    "metric": r.metric,
                    "year": r.year,
                    "sheet": r.source_sheet,
                    "row": r.source_row,
                    "col": r.source_col,
                }
            )
    digest_source = {"sheets": sheet_rows, "metrics": metric_rows}
    digest = hashlib.sha256(yaml.safe_dump(digest_source, allow_unicode=True, sort_keys=True).encode("utf-8")).hexdigest()
    return {
        "latest_actual_year": workbook_info.latest_actual_year,
        "forecast_start_year": workbook_info.forecast_start_year,
        "sheets": sheet_rows,
        "metric_count": len(metric_rows),
        "structure_hash": digest,
    }


def _signature_from_workbook_info(info: WorkbookInfo | None) -> dict[str, Any]:
    """是什么：兼容旧 API 的轻量签名函数。

    为什么：旧代码只传 WorkbookInfo，不传指标记录；保留该函数可避免历史调用失效。
    """
    if info is None:
        return {}
    return build_workbook_signature(info, records=None)


def build_rule_file_data(
    rules: ForecastRules,
    workbook_filename: str | Path | None = None,
    workbook_info: WorkbookInfo | None = None,
    rule_name: str | None = None,
) -> dict[str, Any]:
    """是什么：构造旧版规则文件结构。

    为什么：第一版已经存在 save_rules API，第二版使用 build_rule_profile，
    但保留旧版函数可以降低迁移风险。
    """
    base_name = normalize_workbook_name(workbook_filename) if workbook_filename else "通用模板"
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data = rules_to_dict(rules)
    data["meta"] = {
        "rule_name": rule_name or f"{base_name}_规则",
        "workbook_base_name": base_name,
        "source_workbook_name": Path(workbook_filename).name if workbook_filename else "",
        "updated_at": now,
        "workbook_signature": _signature_from_workbook_info(workbook_info),
        "version": "2.0",
    }
    return data


def build_rule_profile(
    workbook_filename: str | Path,
    workbook_info: WorkbookInfo,
    records: Iterable[MetricRecord],
    rules: ForecastRules,
    targets: list[ForecastTarget] | None = None,
    rule_items: list[RuleItem] | None = None,
) -> dict[str, Any]:
    """是什么：构建第二版完整规则档案。

    为什么：一个可复用规则档案不仅要保存全局规则，还要保存目标预设、
    单条指标规则和工作簿签名，才能支持自动加载、匹配校验、失效项标红。
    """
    base_name = normalize_workbook_name(workbook_filename)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return {
        "meta": {
            "rule_name": f"{base_name}_规则",
            "workbook_base_name": base_name,
            "source_workbook_name": Path(workbook_filename).name,
            "created_at": now,
            "updated_at": now,
            "version": "2.0",
            "workbook_signature": build_workbook_signature(workbook_info, records),
        },
        "rules": rules_to_dict(rules),
        "target_presets": [target_to_dict(t) for t in (targets or [])],
        "rule_items": [rule_item_to_dict(r) for r in (rule_items or [])],
    }


def save_rules(
    rules: ForecastRules,
    path: str | Path,
    workbook_filename: str | Path | None = None,
    workbook_info: WorkbookInfo | None = None,
) -> Path:
    """是什么：保存旧版规则配置到 YAML。

    为什么：兼容早期调用；新功能推荐使用 save_rule_profile。
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    data = build_rule_file_data(rules, workbook_filename=workbook_filename, workbook_info=workbook_info, rule_name=p.stem)
    p.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return p


def save_rule_profile(path: str | Path, profile: dict[str, Any]) -> Path:
    """是什么：保存第二版规则档案。

    为什么：业务人员需要在不同模板版本间复用同一套规则，并保留目标预设。
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    if "meta" in profile:
        profile["meta"]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        profile["meta"]["rule_name"] = p.stem
    p.write_text(yaml.safe_dump(profile, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return p


def load_rule_profile(path: str | Path) -> dict[str, Any]:
    """是什么：从 YAML 文件读取规则档案。

    为什么：加载时要保留原始结构，便于匹配校验、规则项标红和用户编辑。
    """
    p = Path(path)
    data = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    # 兼容第一版规则文件：若没有 rules 包裹层，则把根节点视为 rules。
    if "rules" not in data and any(k in data for k in ["model", "ratio_rules", "growth_rules"]):
        data = {"meta": data.get("meta", {}), "rules": {k: v for k, v in data.items() if k != "meta"}, "target_presets": [], "rule_items": []}
    data.setdefault("target_presets", [])
    data.setdefault("rule_items", [])
    return data


def list_rule_files(rules_dir: str | Path = RULES_DIR) -> list[Path]:
    """是什么：列出本地已保存的规则档案。

    为什么：后续 UI 可以做规则管理列表，显示匹配状态并允许用户手动加载。
    """
    p = Path(rules_dir)
    if not p.exists():
        return []
    return sorted(p.glob("*_规则.yaml"))


def compare_rule_profile(rule_data: dict[str, Any], workbook_filename: str | Path, current_signature: dict[str, Any]) -> RuleProfileMatchResult:
    """是什么：比较规则档案与当前工作簿是否匹配。

    为什么：同名模板可能改过结构；不匹配规则不能自动套用，必须提醒用户。
    """
    meta = rule_data.get("meta", {}) or {}
    messages: list[str] = []
    expected_base = meta.get("workbook_base_name")
    current_base = normalize_workbook_name(workbook_filename)
    if expected_base and expected_base != current_base:
        messages.append(f"规则关联模型主名为“{expected_base}”，当前文件主名为“{current_base}”。")
        return RuleProfileMatchResult(level="error", matched=False, messages=messages)
    old_sig = meta.get("workbook_signature", {}) or {}
    old_hash = old_sig.get("structure_hash")
    current_hash = current_signature.get("structure_hash")
    if old_hash and current_hash and old_hash != current_hash:
        messages.append("当前模板结构与规则保存时不同，建议检查不匹配项后再使用。")
        old_sheets = {x.get("sheet") for x in old_sig.get("sheets", [])}
        new_sheets = {x.get("sheet") for x in current_signature.get("sheets", [])}
        missing = sorted(old_sheets - new_sheets)
        added = sorted(new_sheets - old_sheets)
        if missing:
            messages.append("当前模板缺少规则保存时的 Sheet：" + "、".join(missing[:5]))
        if added:
            messages.append("当前模板新增 Sheet：" + "、".join(added[:5]))
        return RuleProfileMatchResult(level="warning", matched=True, messages=messages)
    return RuleProfileMatchResult(level="ok", matched=True, messages=["规则与当前模板主名和结构匹配。"])


def compare_rule_to_workbook(rule_data: dict[str, Any], workbook_info: WorkbookInfo | None, workbook_filename: str | Path | None = None) -> tuple[str, list[str]]:
    """是什么：旧版匹配校验函数，返回元组格式。

    为什么：保留旧 API，避免第一版界面或测试代码因第二版重构而失效。
    """
    if workbook_info is None:
        return "warning", ["尚未打开工作簿，无法完整校验规则。"]
    if workbook_filename is None:
        return "warning", ["缺少当前文件名，无法校验模型主名。"]
    result = compare_rule_profile(rule_data, workbook_filename, build_workbook_signature(workbook_info, records=None))
    return result.level, result.messages


def _available_business_keys(records: Iterable[MetricRecord]) -> set[tuple[str, str, str]]:
    """是什么：收集当前模板可用的业务指标组合。

    为什么：校验规则项时需要判断目标是否还能在当前模板中找到。
    """
    return {(r.area, r.voltage_level, r.metric) for r in records}


def _available_years(records: Iterable[MetricRecord]) -> set[int]:
    """是什么：收集当前模板已有年份。

    为什么：规则校验时要区分模板缺失年份和可预测新增年份。
    """
    return {r.year for r in records}


def validate_target_presets(profile: dict[str, Any], records: Iterable[MetricRecord]) -> list[RuleItemValidation]:
    """是什么：逐条校验规则档案中的目标预设是否适用于当前模板。

    为什么：规则文件整体可能可用，但其中某些目标引用的区域、电压等级或指标
    已经消失。逐条标红后，用户可以只删除失效项，而不是废弃整个规则文件。
    """
    records_list = list(records)
    keys = _available_business_keys(records_list)
    years = _available_years(records_list)
    max_year = max(years) if years else None
    results: list[RuleItemValidation] = []
    for idx, raw in enumerate(profile.get("target_presets", []) or [], start=1):
        try:
            target = target_from_dict(raw)
        except Exception as exc:
            results.append(RuleItemValidation(item_id=str(idx), target=f"目标预设#{idx}", status="invalid", messages=[f"无法解析：{exc}"]))
            continue
        target_text = f"{target.target_type}|{target.area}|{target.voltage_level}|{target.year or f'{target.period_start}-{target.period_end}'}"
        messages: list[str] = []
        status = "matched"
        metric = target.metric or ("网供负荷" if target.target_type in {"负荷", "增长率"} else "容载比")
        if target.target_type == "容载比":
            # 容载比目标需要对应区域电压下同时存在负荷和容量口径，
            # 只要该区域电压在模板中出现，就允许后续自动生成新增年份。
            business_exists = any(area == target.area and voltage == target.voltage_level for area, voltage, _ in keys)
        else:
            business_exists = (target.area, target.voltage_level, metric) in keys
        if not target.area or not target.voltage_level:
            status = "invalid"
            messages.append("缺少区域或电压等级。")
        elif not business_exists:
            status = "missing"
            messages.append("当前模板中找不到该区域/电压等级/指标组合。")
        # 新增年份是允许的：2031+ 可由工具内部预测，但需要给 warning 提醒。
        check_year = target.year or target.period_end
        if check_year and max_year and check_year > max_year and status == "matched":
            status = "warning"
            messages.append(f"目标年份 {check_year} 超出当前模板最大年份 {max_year}，将由工具内部预测并扩展写回。")
        if not messages:
            messages.append("规则项可用于当前模板。")
        results.append(RuleItemValidation(item_id=str(idx), target=target_text, status=status, messages=messages))

    for idx, raw in enumerate(profile.get("rule_items", []) or [], start=1):
        try:
            item = rule_item_from_dict(raw)
        except Exception as exc:
            results.append(RuleItemValidation(item_id=f"rule-{idx}", target=f"指标规则#{idx}", status="invalid", messages=[f"无法解析：{exc}"]))
            continue
        messages = []
        status = "matched"
        if item.year is None or not item.area or not item.voltage_level or not item.metric:
            status = "invalid"
            messages.append("缺少区域、电压等级、指标或年份。")
        elif (item.area, item.voltage_level, item.metric) not in keys:
            status = "missing"
            messages.append("当前模板中找不到该指标规则对应的业务指标。")
        elif item.year not in years:
            status = "warning"
            messages.append("该年份当前模板中不存在，若为新增预测年会在导出时扩展写回。")
        if not messages:
            messages.append("规则项可用于当前模板。")
        results.append(RuleItemValidation(item_id=f"rule-{idx}", target=item.target, status=status, messages=messages))
    return results
