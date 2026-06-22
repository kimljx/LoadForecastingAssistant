"""Qt 表格辅助函数。

是什么：封装 QTableWidget 的常见填充和着色逻辑。
为什么：主窗口中有大量中文表格，如果每处都手写填表代码，会造成重复且难维护。
"""
from __future__ import annotations

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QTableWidget, QTableWidgetItem
    from PySide6.QtGui import QColor
except Exception:  # pragma: no cover - 便于无 Qt 环境下运行核心测试
    Qt = object  # type: ignore
    QTableWidget = object  # type: ignore
    QTableWidgetItem = object  # type: ignore
    QColor = object  # type: ignore


def set_table_data(table: QTableWidget, headers: list[str], rows: list[list], editable: bool = False) -> None:
    """是什么：一次性设置表头和表格数据。

    为什么：业务界面需要展示模板识别、兼容性检测、调整项和预测结果等多张表，
    统一封装可以避免遗漏只读状态、列宽调整等细节。
    """
    table.clear()
    table.setColumnCount(len(headers))
    table.setRowCount(len(rows))
    table.setHorizontalHeaderLabels(headers)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            item = QTableWidgetItem("" if value is None else str(value))
            if not editable:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(r, c, item)
    table.resizeColumnsToContents()


def set_row_background(table: QTableWidget, row: int, color: str) -> None:
    """是什么：设置某一整行背景色。

    为什么：规则匹配状态需要绿色/橙色/红色区分，整行着色比单元格文字更直观。
    """
    qcolor = QColor(color)
    for c in range(table.columnCount()):
        item = table.item(row, c)
        if item is not None:
            item.setBackground(qcolor)
