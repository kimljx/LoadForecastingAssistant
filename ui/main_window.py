"""PySide6 桌面主窗口。"""
from __future__ import annotations

import traceback
from pathlib import Path

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QFormLayout,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QSpinBox,
        QSplitter,
        QTabWidget,
        QTableWidget,
        QTextEdit,
        QTreeWidget,
        QTreeWidgetItem,
        QVBoxLayout,
        QWidget,
    )
except Exception:  # pragma: no cover
    raise

from core.compatibility_checker import check_workbook_compatibility, has_blocking_errors
from core.forecast_engine import solve_forecast
from core.models import ForecastTarget
from core.rule_engine import load_rules
from core.template_parser import build_workbook_info, extract_all_metric_records, list_areas, list_voltage_levels, list_years
from core.workbook_loader import load_workbook_auto
from core.workbook_writer import build_output_path, export_scenario_to_workbook

from .table_utils import set_table_data


class MainWindow(QMainWindow):
    """主窗口，面向全中文用户。"""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("负荷预测与容载比反推助手")
        self.resize(1280, 820)
        self.loaded = None
        self.workbook_info = None
        self.records = []
        self.rules = load_rules()
        self.scenarios = []
        self.selected_scenario = None
        self.current_file: Path | None = None
        self._init_ui()

    def _init_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        top = QHBoxLayout()
        self.btn_open = QPushButton("打开Excel模板")
        self.btn_open.clicked.connect(self.open_file)
        self.btn_run = QPushButton("生成反推方案")
        self.btn_run.clicked.connect(self.run_solver)
        self.btn_export = QPushButton("导出新Excel")
        self.btn_export.clicked.connect(self.export_file)
        self.btn_export.setEnabled(False)
        self.lbl_status = QLabel("未打开文件")
        top.addWidget(self.btn_open)
        top.addWidget(self.btn_run)
        top.addWidget(self.btn_export)
        top.addWidget(self.lbl_status, 1)
        layout.addLayout(top)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(splitter, 1)

        self.nav = QTreeWidget()
        self.nav.setHeaderLabel("功能导航")
        for name in ["模板识别", "兼容性检测", "目标设置", "反推方案", "预测结果", "写回说明"]:
            QTreeWidgetItem(self.nav, [name])
        self.nav.itemClicked.connect(self._nav_clicked)
        splitter.addWidget(self.nav)

        self.tabs = QTabWidget()
        splitter.addWidget(self.tabs)
        splitter.setStretchFactor(1, 1)

        self.tab_template = QWidget()
        self.tab_compat = QWidget()
        self.tab_target = QWidget()
        self.tab_result = QWidget()
        self.tab_forecast = QWidget()
        self.tab_log = QWidget()
        self.tabs.addTab(self.tab_template, "模板识别")
        self.tabs.addTab(self.tab_compat, "兼容性检测")
        self.tabs.addTab(self.tab_target, "目标设置")
        self.tabs.addTab(self.tab_result, "反推方案")
        self.tabs.addTab(self.tab_forecast, "预测结果")
        self.tabs.addTab(self.tab_log, "写回说明")

        self.template_table = QTableWidget()
        lay = QVBoxLayout(self.tab_template)
        lay.addWidget(QLabel("系统会按字段和表头识别模板角色，项目库不依赖固定 sheet 名。"))
        lay.addWidget(self.template_table)

        self.compat_table = QTableWidget()
        lay = QVBoxLayout(self.tab_compat)
        lay.addWidget(self.compat_table)

        self._init_target_tab()
        self._init_result_tabs()
        self.statusBar().showMessage("就绪")

    def _init_target_tab(self) -> None:
        lay = QVBoxLayout(self.tab_target)
        form = QFormLayout()
        self.cmb_area = QComboBox()
        self.cmb_voltage = QComboBox()
        self.cmb_target_type = QComboBox()
        self.cmb_target_type.addItems(["容载比", "增长率", "负荷"])
        self.spin_year = QSpinBox()
        self.spin_year.setRange(2026, 2050)
        self.spin_year.setValue(2030)
        self.edit_target_value = QLineEdit("1.85")
        self.edit_min = QLineEdit("")
        self.edit_max = QLineEdit("")
        self.spin_forecast_end = QSpinBox()
        self.spin_forecast_end.setRange(2026, 2050)
        self.spin_forecast_end.setValue(2030)
        self.chk_overwrite_formula = QCheckBox("导出时允许覆盖公式单元格（用户确认后在副本中写入结果值）")
        self.chk_overwrite_formula.setChecked(True)
        form.addRow("区域", self.cmb_area)
        form.addRow("电压等级", self.cmb_voltage)
        form.addRow("目标类型", self.cmb_target_type)
        form.addRow("目标年份", self.spin_year)
        form.addRow("目标值", self.edit_target_value)
        form.addRow("目标下限", self.edit_min)
        form.addRow("目标上限", self.edit_max)
        form.addRow("预测截止年", self.spin_forecast_end)
        form.addRow("写回策略", self.chk_overwrite_formula)
        lay.addLayout(form)
        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setPlainText(
            "使用说明：\n"
            "1. 2025年及以前默认为现状数据，不参与修改。\n"
            "2. 容载比默认硬边界为1.3-2.5，上级指定目标优先。\n"
            "3. 同时率范围为0.83-0.99，单次调整不超过0.1。\n"
            "4. 区外送(+)/受(-)电为低优先级兜底变量，需要人工确认站点后写回。\n"
            "5. 导出的文件是新副本，不会修改传入的原始文件。"
        )
        lay.addWidget(help_text, 1)

    def _init_result_tabs(self) -> None:
        lay = QVBoxLayout(self.tab_result)
        self.cmb_scenario = QComboBox()
        self.cmb_scenario.currentIndexChanged.connect(self._scenario_changed)
        self.result_table = QTableWidget()
        lay.addWidget(QLabel("反推方案选择"))
        lay.addWidget(self.cmb_scenario)
        lay.addWidget(self.result_table)

        lay2 = QVBoxLayout(self.tab_forecast)
        self.forecast_table = QTableWidget()
        lay2.addWidget(self.forecast_table)

        lay3 = QVBoxLayout(self.tab_log)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        lay3.addWidget(self.log_text)

    def _nav_clicked(self, item, column):
        idx = ["模板识别", "兼容性检测", "目标设置", "反推方案", "预测结果", "写回说明"].index(item.text(0))
        self.tabs.setCurrentIndex(idx)

    def open_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "选择供需预测模板", "", "Excel文件 (*.xlsx *.xls)")
        if not file_path:
            return
        try:
            self.current_file = Path(file_path)
            self.loaded = load_workbook_auto(file_path)
            self.workbook_info = build_workbook_info(self.loaded, latest_actual_year=self.rules.latest_actual_year)
            self.rules.latest_actual_year = self.workbook_info.latest_actual_year
            self.rules.forecast_start_year = self.workbook_info.forecast_start_year
            self.records = extract_all_metric_records(self.loaded, self.workbook_info)
            self._refresh_template_table()
            self._refresh_compat_table()
            self._refresh_target_options()
            self.lbl_status.setText(f"当前文件：{self.current_file.name}；记录数：{len(self.records)}")
            self.statusBar().showMessage("模板读取完成")
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", f"读取文件失败：\n{exc}\n\n{traceback.format_exc()}")

    def _refresh_template_table(self) -> None:
        rows = []
        for info in self.workbook_info.sheet_infos.values():
            years = ", ".join(str(y.year) + ("现状" if y.is_actual else "") for y in info.year_columns)
            rows.append([info.name, info.role.value, info.area_name or "", info.voltage_level or "", years, f"{info.confidence:.2f}", ";".join(info.notes)])
        set_table_data(self.template_table, ["Sheet", "角色", "区域", "电压等级", "年份列", "置信度", "说明"], rows)

    def _refresh_compat_table(self) -> None:
        issues = check_workbook_compatibility(self.loaded)
        rows = [[i.level, i.sheet, i.location, i.issue_type, i.content, i.suggestion] for i in issues]
        set_table_data(self.compat_table, ["级别", "Sheet", "位置", "问题类型", "内容", "建议"], rows)
        if has_blocking_errors(issues):
            QMessageBox.warning(self, "兼容性警告", "检测到 Error 级问题，业务模型仍可查看，但请先处理后再正式导出。")

    def _refresh_target_options(self) -> None:
        self.cmb_area.clear()
        self.cmb_voltage.clear()
        self.cmb_area.addItems(list_areas(self.records))
        self.cmb_voltage.addItems([v for v in list_voltage_levels(self.records) if v != "总计"])
        years = list_years(self.records)
        if years:
            self.spin_year.setValue(max(y for y in years if y >= self.rules.forecast_start_year) if any(y >= self.rules.forecast_start_year for y in years) else max(years))
            self.spin_forecast_end.setValue(max(years))

    def _build_targets_from_ui(self) -> list[ForecastTarget]:
        def parse_float(text: str):
            text = text.strip()
            return None if not text else float(text)
        ttype = self.cmb_target_type.currentText()
        area = self.cmb_area.currentText()
        voltage = self.cmb_voltage.currentText()
        year = int(self.spin_year.value())
        target_value = parse_float(self.edit_target_value.text())
        min_value = parse_float(self.edit_min.text())
        max_value = parse_float(self.edit_max.text())
        if ttype == "增长率":
            return [ForecastTarget("增长率", area, voltage, period_start=self.rules.latest_actual_year, period_end=year, metric="网供负荷", target_value=target_value, min_value=min_value, max_value=max_value)]
        metric = "网供负荷" if ttype == "负荷" else None
        return [ForecastTarget(ttype, area, voltage, year=year, metric=metric, target_value=target_value, min_value=min_value, max_value=max_value)]

    def run_solver(self) -> None:
        if not self.records:
            QMessageBox.information(self, "提示", "请先打开Excel模板。")
            return
        try:
            targets = self._build_targets_from_ui()
            forecast_end = int(self.spin_forecast_end.value())
            self.scenarios = solve_forecast(self.records, targets, self.rules, forecast_end)
            self.cmb_scenario.clear()
            self.cmb_scenario.addItems([s.name for s in self.scenarios])
            self._scenario_changed(0)
            self.btn_export.setEnabled(True)
            self.tabs.setCurrentIndex(3)
        except Exception as exc:
            QMessageBox.critical(self, "反推失败", f"生成方案失败：\n{exc}\n\n{traceback.format_exc()}")

    def _scenario_changed(self, index: int) -> None:
        if not self.scenarios or index < 0:
            return
        self.selected_scenario = self.scenarios[index]
        adj_rows = [[a.object_type, a.area, a.voltage_level, a.metric, a.year, a.old_value, a.new_value, a.delta, a.delta_pct, a.reason, a.priority, a.risk_level, a.source_sheet or "", a.source_cell or ""] for a in self.selected_scenario.adjustments]
        set_table_data(self.result_table, ["调整对象", "区域", "电压等级", "指标", "年份", "原值", "建议值", "变化量", "变化比例", "原因", "优先级", "风险", "来源表", "来源单元格"], adj_rows)
        pred_rows = [[r.area, r.voltage_level, r.metric, r.year, r.value, r.source_sheet, r.source_cell, "是" if r.is_actual else "否"] for r in sorted(self.selected_scenario.forecast_records, key=lambda r: (r.area, r.voltage_level, r.metric, r.year))[:2000]]
        set_table_data(self.forecast_table, ["区域", "电压等级", "指标", "年份", "结果值", "来源表", "来源单元格", "是否现状"], pred_rows)
        text = [self.selected_scenario.description]
        text.extend(self.selected_scenario.warnings)
        self.log_text.setPlainText("\n".join(text))

    def export_file(self) -> None:
        if not self.selected_scenario or not self.loaded or not self.current_file:
            QMessageBox.information(self, "提示", "请先生成方案。")
            return
        default = build_output_path(self.current_file)
        out, _ = QFileDialog.getSaveFileName(self, "保存反推结果", str(default), "Excel文件 (*.xlsx)")
        if not out:
            return
        reply = QMessageBox.question(self, "确认导出", "将复制原始工作簿并在副本中写入调整值，同时新增“预测结果表”。\n原始传入文件不会被修改。是否继续？")
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            export_scenario_to_workbook(self.loaded, self.selected_scenario, self.rules, out, overwrite_formula=self.chk_overwrite_formula.isChecked())
            QMessageBox.information(self, "导出完成", f"已生成：\n{out}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", f"导出失败：\n{exc}\n\n{traceback.format_exc()}")
