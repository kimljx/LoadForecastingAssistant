"""兼容性检测模块。"""
from __future__ import annotations

import re
from pathlib import Path

from .models import CompatibilityIssue
from .workbook_loader import LoadedWorkbook


UNSUPPORTED_FUNCTIONS = {
    "INDIRECT",
    "OFFSET",
    "动态数组",
}
# 这些函数在模板中可能大量存在，工具不再尝试完整复刻，而是通过业务模型绕开。
COMPLEX_BUT_ALLOWED_AS_REFERENCE = {"INDEX", "MATCH", "SUMIF", "SUMIFS", "IFERROR", "YEAR", "MONTH", "MAX", "IFS", "SUBTOTAL"}
SUPPORTED_EXCEL_FUNCTIONS_REFERENCE = {"SUM", "IF", "VLOOKUP"} | COMPLEX_BUT_ALLOWED_AS_REFERENCE

_FUNCTION_RE = re.compile(r"\b([A-Z][A-Z0-9_]*)\s*\(", re.IGNORECASE)
_EXTERNAL_LINK_RE = re.compile(r"\[[^\]]+\]")


def check_workbook_compatibility(loaded: LoadedWorkbook) -> list[CompatibilityIssue]:
    """是什么：检测当前工具对工作簿的兼容风险。

注意：本工具的核心计算基于业务模型，不要求完整执行每一个 Excel 公式。
因此 INDEX/MATCH/SUMIFS 等在这里作为 warning，而不是 error。

为什么：兼容性风险必须在正式反推前暴露，避免业务人员在不适配模板上得到误导性结果。"""
    issues: list[CompatibilityIssue] = []
    if loaded.source_format == "xls":
        issues.append(
            CompatibilityIssue(
                level="info",
                sheet="工作簿",
                location="文件格式",
                issue_type="Excel 97-2003 格式",
                content=str(loaded.original_path.name),
                suggestion="系统会读取并导出为 .xlsx；建议后续升级模板格式。",
            )
        )
    if loaded.repaired:
        issues.append(
            CompatibilityIssue(
                level="info",
                sheet="工作簿",
                location="读取阶段",
                issue_type="已使用安全读取副本",
                content=str(loaded.original_path.name),
                suggestion="原始文件未被修改；导出会基于可读取副本生成。",
            )
        )

    for ws in loaded.formula_wb.worksheets:
        if ws.sheet_state != "visible":
            issues.append(
                CompatibilityIssue("warning", ws.title, "工作表", "隐藏Sheet", ws.sheet_state, "请确认隐藏表是否参与预测。")
            )
        if ws.protection.sheet:
            issues.append(
                CompatibilityIssue("warning", ws.title, "工作表", "Sheet保护", "已保护", "如需写回该表，可能需要解除保护。")
            )
        if ws.merged_cells.ranges:
            issues.append(
                CompatibilityIssue(
                    "warning", ws.title, "合并单元格", "合并单元格", f"{len(ws.merged_cells.ranges)}处", "写回时将尽量保留格式。"
                )
            )
        if ws.auto_filter and ws.auto_filter.ref:
            issues.append(
                CompatibilityIssue("info", ws.title, ws.auto_filter.ref, "筛选区域", ws.auto_filter.ref, "不影响业务模型计算。")
            )
        if ws.data_validations and ws.data_validations.count:
            issues.append(
                CompatibilityIssue("warning", ws.title, "数据验证", "数据验证", str(ws.data_validations.count), "写回值需符合原表验证规则。")
            )

        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                formula = v.upper()
                if _EXTERNAL_LINK_RE.search(formula):
                    issues.append(
                        CompatibilityIssue("error", ws.title, cell.coordinate, "外部工作簿引用", v[:120], "请取消外部链接后再使用。")
                    )
                for fn in _FUNCTION_RE.findall(formula):
                    fn_upper = fn.upper()
                    if fn_upper in UNSUPPORTED_FUNCTIONS:
                        issues.append(
                            CompatibilityIssue("error", ws.title, cell.coordinate, f"不支持函数 {fn_upper}", v[:120], "该函数难以静态解析，请改为明确数值或业务规则。")
                        )
                    elif fn_upper in COMPLEX_BUT_ALLOWED_AS_REFERENCE:
                        issues.append(
                            CompatibilityIssue("info", ws.title, cell.coordinate, f"复杂函数 {fn_upper}", v[:80], "工具将以业务模型计算关键指标，不完整执行该公式。")
                        )
    return issues


def has_blocking_errors(issues: list[CompatibilityIssue]) -> bool:
    """是什么：处理兼容性检测结果。

    为什么：导入模板前必须明确哪些问题会阻断反推。
    """
    return any(i.level == "error" for i in issues)
