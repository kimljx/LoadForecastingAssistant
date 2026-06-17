"""Qt 表格辅助函数。"""
from __future__ import annotations

try:
    from PySide6.QtWidgets import QTableWidget, QTableWidgetItem
    from PySide6.QtGui import QColor
except Exception:  # 便于无 Qt 环境下运行核心测试
    QTableWidget = object  # type: ignore
    QTableWidgetItem = object  # type: ignore
    QColor = object  # type: ignore


def set_table_data(table: QTableWidget, headers: list[str], rows: list[list]) -> None:
    table.clear()
    table.setColumnCount(len(headers))
    table.setRowCount(len(rows))
    table.setHorizontalHeaderLabels(headers)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            item = QTableWidgetItem("" if value is None else str(value))
            table.setItem(r, c, item)
    table.resizeColumnsToContents()
