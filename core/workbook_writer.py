"""Excel 导出与写回模块。"""
from __future__ import annotations

from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .models import Adjustment, ForecastRules, MetricRecord, ScenarioResult, WorkbookInfo
from .template_parser import YEAR_RE
from .workbook_loader import LoadedWorkbook, copy_workbook_for_export


def build_output_path(input_path: str | Path, output_dir: str | Path | None = None) -> Path:
    p = Path(input_path)
    out_dir = Path(output_dir) if output_dir else p.parent
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return out_dir / f"{p.stem}_反推结果_{stamp}.xlsx"


def _write_cell(ws, cell: str, value, overwrite_formula: bool) -> tuple[bool, str]:
    c = ws[cell]
    if isinstance(c.value, str) and c.value.startswith("=") and not overwrite_formula:
        return False, "目标单元格为公式，按规则未覆盖"
    c.value = value
    return True, "已写入"


def _copy_column_style(ws, from_col: int, to_col: int) -> None:
    ws.column_dimensions[get_column_letter(to_col)].width = ws.column_dimensions[get_column_letter(from_col)].width
    for row in range(1, ws.max_row + 1):
        src = ws.cell(row, from_col)
        dst = ws.cell(row, to_col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)


def _find_year_columns(ws) -> dict[int, int]:
    result: dict[int, int] = {}
    for row in range(1, min(ws.max_row, 5) + 1):
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row, col).value
            if val is None:
                continue
            m = YEAR_RE.search(str(val))
            if m:
                result[int(m.group(1))] = col
    return result


def _ensure_year_column(ws, year: int) -> int:
    """确保某年列存在。若不存在，在最后一个年份列后插入新列。

    第一版仅复制样式和值写入，不尝试完整维护所有跨列公式。
    """
    cols = _find_year_columns(ws)
    if year in cols:
        return cols[year]
    if not cols:
        raise ValueError(f"工作表 {ws.title} 未找到年份列，无法扩展 {year} 年")
    last_year = max(cols)
    last_col = cols[last_year]
    insert_at = last_col + 1
    ws.insert_cols(insert_at)
    _copy_column_style(ws, last_col, insert_at)
    # 找到上一年表头所在行。
    header_row = None
    for row in range(1, min(ws.max_row, 5) + 1):
        if ws.cell(row, last_col).value and YEAR_RE.search(str(ws.cell(row, last_col).value)):
            header_row = row
            break
    if header_row is None:
        header_row = 3
    ws.cell(header_row, insert_at).value = f"{year}年"
    return insert_at


def _write_adjustments_to_original_sheets(
    wb: openpyxl.Workbook,
    scenario: ScenarioResult,
    rules: ForecastRules,
    overwrite_formula: bool,
) -> list[dict]:
    log: list[dict] = []
    for adj in scenario.adjustments:
        if not adj.write_back:
            log.append({"单元格": adj.source_cell or "", "状态": "未写回", "原因": adj.note or "该调整项仅作为方案提示"})
            continue
        if adj.year <= rules.latest_actual_year:
            log.append({"单元格": adj.source_cell or "", "状态": "跳过", "原因": "现状年不可修改"})
            continue
        if not adj.source_sheet or adj.source_sheet not in wb.sheetnames:
            log.append({"单元格": adj.source_cell or "", "状态": "未写回", "原因": "缺少来源工作表，通常是新增年份或兜底变量"})
            continue
        ws = wb[adj.source_sheet]
        target_cell = adj.source_cell
        # 新增年份可能没有原始 source_cell，此处第一版只写已有来源单元格。
        if not target_cell:
            log.append({"单元格": "", "状态": "未写回", "原因": "缺少来源单元格"})
            continue
        ok, msg = _write_cell(ws, target_cell, adj.new_value, overwrite_formula=overwrite_formula)
        log.append(
            {
                "工作表": adj.source_sheet,
                "单元格": target_cell,
                "年份": adj.year,
                "指标": adj.metric,
                "原值": adj.old_value,
                "写入值": adj.new_value,
                "状态": "已写回" if ok else "跳过",
                "原因": msg,
            }
        )
    return log


def _autosize(ws, min_width: int = 10, max_width: int = 38) -> None:
    # 合并单元格会产生 MergedCell，不能依赖 col[0].column_letter。
    for idx, col in enumerate(ws.columns, start=1):
        letter = get_column_letter(idx)
        max_len = 0
        for cell in col:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = max(min_width, min(max_width, max_len + 2))


def add_result_sheet(wb: openpyxl.Workbook, scenario: ScenarioResult, rules: ForecastRules, write_log: list[dict]) -> None:
    name = "预测结果表"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, 0)
    title_fill = PatternFill("solid", fgColor="1F4E78")
    section_fill = PatternFill("solid", fgColor="D9EAF7")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)

    row = 1
    ws.cell(row, 1).value = "负荷预测与容载比反推结果"
    ws.cell(row, 1).font = Font(size=16, bold=True, color="FFFFFF")
    ws.cell(row, 1).fill = title_fill
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    row += 2

    basic = [
        ("生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("方案名称", scenario.name),
        ("方案说明", scenario.description),
        ("最新现状年", rules.latest_actual_year),
        ("预测起始年", rules.forecast_start_year),
        ("容载比硬边界", f"{rules.ratio_min}-{rules.ratio_max}"),
        ("同时率边界", f"{rules.coincidence_factor_min}-{rules.coincidence_factor_max}，单次调整不超过{rules.coincidence_factor_max_abs_change}"),
    ]
    for k, v in basic:
        ws.cell(row, 1).value = k
        ws.cell(row, 1).font = bold_font
        ws.cell(row, 2).value = v
        row += 1
    row += 1

    def section(title: str):
        nonlocal row
        ws.cell(row, 1).value = title
        ws.cell(row, 1).font = bold_font
        ws.cell(row, 1).fill = section_fill
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
        row += 1

    def write_table(headers: list[str], rows: list[list]):
        nonlocal row
        for i, h in enumerate(headers, start=1):
            c = ws.cell(row, i)
            c.value = h
            c.font = bold_font
            c.fill = PatternFill("solid", fgColor="E2F0D9")
        row += 1
        for values in rows:
            for i, v in enumerate(values, start=1):
                ws.cell(row, i).value = v
            row += 1
        row += 1

    section("一、目标达成情况")
    target_rows = []
    for tr in scenario.target_results:
        t = tr.target
        target_rows.append([
            t.target_type,
            t.area,
            t.voltage_level,
            t.year or f"{t.period_start}-{t.period_end}",
            t.metric or "",
            t.target_value if t.target_value is not None else f"{t.min_value or ''}-{t.max_value or ''}",
            tr.actual_value,
            "达标" if tr.achieved else "未达标",
            tr.deviation,
            tr.message,
        ])
    write_table(["目标类型", "区域", "电压等级", "年份/阶段", "指标", "目标", "预测结果", "是否达标", "偏差", "说明"], target_rows)

    section("二、调整项明细")
    adj_rows = []
    for a in scenario.adjustments:
        adj_rows.append([
            a.object_type, a.area, a.voltage_level, a.metric, a.year, a.old_value, a.new_value, a.delta, a.delta_pct,
            a.reason, a.priority, a.risk_level, a.source_sheet or "", a.source_cell or "", a.note,
        ])
    write_table(["调整对象", "区域", "电压等级", "指标", "年份", "原值", "建议值", "变化量", "变化比例", "调整原因", "优先级", "风险", "来源表", "来源单元格", "备注"], adj_rows)

    section("三、预测结果明细")
    records = sorted(scenario.forecast_records, key=lambda r: (r.area, r.voltage_level, r.metric, r.year))
    pred_rows = [
        [r.area, r.voltage_level, r.metric, r.year, r.value, r.source_sheet, r.source_cell, "是" if r.is_actual else "否"]
        for r in records
    ]
    write_table(["区域", "电压等级", "指标", "年份", "预测/结果值", "来源表", "来源单元格", "是否现状年"], pred_rows)

    section("四、写回单元格清单")
    log_rows = []
    for item in write_log:
        log_rows.append([
            item.get("工作表", ""), item.get("单元格", ""), item.get("年份", ""), item.get("指标", ""), item.get("原值", ""),
            item.get("写入值", ""), item.get("状态", ""), item.get("原因", ""),
        ])
    write_table(["工作表", "单元格", "年份", "指标", "原值", "写入值", "状态", "原因"], log_rows)

    if scenario.warnings:
        section("五、风险提示")
        for w in scenario.warnings:
            ws.cell(row, 1).value = w
            row += 1

    for row_cells in ws.iter_rows():
        for cell in row_cells:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    _autosize(ws)


def export_scenario_to_workbook(
    loaded: LoadedWorkbook,
    scenario: ScenarioResult,
    rules: ForecastRules,
    output_path: str | Path,
    overwrite_formula: bool = True,
) -> Path:
    """复制原工作簿，写回方案，并新增“预测结果表”。

    原始传入文件绝不修改。
    """
    out = copy_workbook_for_export(loaded, output_path)
    wb = openpyxl.load_workbook(out)
    write_log = _write_adjustments_to_original_sheets(wb, scenario, rules, overwrite_formula=overwrite_formula)
    add_result_sheet(wb, scenario, rules, write_log)
    wb.save(out)
    return out
