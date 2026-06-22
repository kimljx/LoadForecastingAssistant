"""PySide6 桌面主窗口。

是什么：
    这是负荷预测与容载比反推助手的桌面端 UI 层，负责文件选择、规则编辑、
    目标设置、方案展示和导出确认。

为什么：
    用户主要是中文业务人员，典型使用场景是在 Windows 内网电脑上双击程序、
    选择 Excel 模板、输入上级目标并导出新表。因此 UI 需要尽量中文化，且不能
    把核心计算逻辑写死在界面里。
"""
from __future__ import annotations

import traceback
from dataclasses import asdict
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
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
    QTableWidgetItem,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.compatibility_checker import check_workbook_compatibility, has_blocking_errors
from core.forecast_engine import solve_forecast
from core.models import ForecastRules, ForecastTarget
from core.rule_engine import load_rules
from core.rule_persistence import (
    build_rule_profile,
    build_workbook_signature,
    compare_rule_profile,
    default_rule_path_for_workbook,
    forecast_rules_from_dict,
    load_rule_profile,
    save_rule_profile,
    target_from_dict,
    validate_target_presets,
)
from core.template_parser import build_workbook_info, extract_all_metric_records, list_areas, list_voltage_levels, list_years
from core.workbook_loader import load_workbook_auto
from core.workbook_writer import build_output_path, export_scenario_to_workbook

from .table_utils import set_table_data


NAV_NAMES = ["模板识别", "兼容性检测", "目标设置", "规则管理", "反推方案", "预测结果", "写回说明"]


class MainWindow(QMainWindow):
    """主窗口。

    是什么：整合所有业务页面的 Qt 主窗口。
    为什么：桌面端需要一个稳定入口承载多步骤流程，同时保持核心计算层与 UI 层分离。
    """

    def __init__(self) -> None:
        """是什么：处理桌面界面动作和状态更新。

        为什么：界面方法需要说明业务意图，便于后续维护人员区分 UI 与核心逻辑。
        """
        super().__init__()
        self.setWindowTitle("负荷预测与容载比反推助手")
        self.resize(1360, 860)
        self.loaded = None
        self.workbook_info = None
        self.records = []
        self.rules: ForecastRules = load_rules()
        self.scenarios = []
        self.selected_scenario = None
        self.current_file: Path | None = None
        self.current_rule_path: Path | None = None
        self.current_rule_profile: dict[str, Any] | None = None
        self.rules_dirty = False
        self._init_ui()
        self._apply_rules_to_widgets()

    def _init_ui(self) -> None:
        """是什么：初始化主界面布局和所有页签。

        为什么：把 UI 创建集中在一个方法中，便于后续拆分成独立面板；同时避免
        在业务方法中混杂控件构造逻辑。
        """
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
        for name in NAV_NAMES:
            QTreeWidgetItem(self.nav, [name])
        self.nav.itemClicked.connect(self._nav_clicked)
        splitter.addWidget(self.nav)

        self.tabs = QTabWidget()
        splitter.addWidget(self.tabs)
        splitter.setStretchFactor(1, 1)

        self.tab_template = QWidget()
        self.tab_compat = QWidget()
        self.tab_target = QWidget()
        self.tab_rules = QWidget()
        self.tab_result = QWidget()
        self.tab_forecast = QWidget()
        self.tab_log = QWidget()
        self.tabs.addTab(self.tab_template, "模板识别")
        self.tabs.addTab(self.tab_compat, "兼容性检测")
        self.tabs.addTab(self.tab_target, "目标设置")
        self.tabs.addTab(self.tab_rules, "规则管理")
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
        self._init_rule_tab()
        self._init_result_tabs()
        self.statusBar().showMessage("就绪")

    def _init_target_tab(self) -> None:
        """是什么：初始化目标设置页。

        为什么：实际工作经常同时受容载比和增长率约束，所以第二版支持多目标列表，
        不再只使用一个表单目标。
        """
        lay = QVBoxLayout(self.tab_target)
        form = QFormLayout()
        self.cmb_area = QComboBox()
        self.cmb_voltage = QComboBox()
        self.cmb_target_type = QComboBox()
        self.cmb_target_type.addItems(["容载比", "增长率", "负荷", "配变平均负载率"])
        self.cmb_target_type.currentTextChanged.connect(self._target_type_changed)
        self.spin_year = QSpinBox()
        self.spin_year.setRange(2026, 2050)
        self.spin_year.setValue(2030)
        self.edit_target_value = QLineEdit("1.85")
        self.edit_min = QLineEdit("")
        self.edit_max = QLineEdit("")
        self.spin_forecast_end = QSpinBox()
        self.spin_forecast_end.setRange(2026, 2050)
        self.spin_forecast_end.setValue(2030)
        self.chk_overwrite_formula = QCheckBox("导出时允许覆盖公式单元格（只影响导出副本）")
        self.chk_overwrite_formula.setChecked(True)
        form.addRow("区域", self.cmb_area)
        form.addRow("电压等级", self.cmb_voltage)
        form.addRow("目标类型", self.cmb_target_type)
        form.addRow("目标年份/阶段结束年", self.spin_year)
        form.addRow("目标值", self.edit_target_value)
        form.addRow("目标下限", self.edit_min)
        form.addRow("目标上限", self.edit_max)
        form.addRow("预测截止年", self.spin_forecast_end)
        form.addRow("写回策略", self.chk_overwrite_formula)
        lay.addLayout(form)

        btns = QHBoxLayout()
        self.btn_add_target = QPushButton("添加目标到列表")
        self.btn_add_target.clicked.connect(self.add_target_from_form)
        self.btn_delete_target = QPushButton("删除选中目标")
        self.btn_delete_target.clicked.connect(self.delete_selected_targets)
        self.btn_clear_targets = QPushButton("清空目标")
        self.btn_clear_targets.clicked.connect(self.clear_targets)
        btns.addWidget(self.btn_add_target)
        btns.addWidget(self.btn_delete_target)
        btns.addWidget(self.btn_clear_targets)
        btns.addStretch(1)
        lay.addLayout(btns)

        self.target_table = QTableWidget()
        self.target_table.setColumnCount(9)
        self.target_table.setHorizontalHeaderLabels(["启用", "目标类型", "区域", "电压等级", "年份", "阶段起", "阶段止", "目标值", "下限/上限"])
        lay.addWidget(QLabel("多目标列表：为空时，系统会使用上方表单当前目标。"))
        lay.addWidget(self.target_table, 1)

        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setMaximumHeight(140)
        help_text.setPlainText(
            "使用说明：\n"
            "1. 2025年及以前默认为现状数据，不参与修改。\n"
            "2. 容载比默认硬边界为1.3-2.5，上级指定目标优先。\n"
            "3. 同时率范围为0.83-0.99，单次调整不超过0.1。\n"
            "4. 增长率目标可输入 0.06 或 6% / 6。\n"
            "5. 区外送(+)/受(-)电为低优先级兜底变量，需要人工确认站点后写回。"
        )
        lay.addWidget(help_text)

    def _init_rule_tab(self) -> None:
        """是什么：初始化规则管理页。

        为什么：业务规则需要长期复用，并且模板版本变化时需要看到规则匹配状态；
        所以这里提供规则应用、保存、加载和目标预设校验。
        """
        lay = QVBoxLayout(self.tab_rules)
        form = QFormLayout()
        self.rule_latest_actual = QSpinBox(); self.rule_latest_actual.setRange(2000, 2050)
        self.rule_ratio_min = QDoubleSpinBox(); self.rule_ratio_min.setRange(0.1, 10); self.rule_ratio_min.setDecimals(3); self.rule_ratio_min.setSingleStep(0.05)
        self.rule_ratio_max = QDoubleSpinBox(); self.rule_ratio_max.setRange(0.1, 10); self.rule_ratio_max.setDecimals(3); self.rule_ratio_max.setSingleStep(0.05)
        self.rule_tr_min = QDoubleSpinBox(); self.rule_tr_min.setRange(0, 2); self.rule_tr_min.setDecimals(3)
        self.rule_tr_max = QDoubleSpinBox(); self.rule_tr_max.setRange(0, 2); self.rule_tr_max.setDecimals(3)
        self.rule_tr_target = QDoubleSpinBox(); self.rule_tr_target.setRange(0, 2); self.rule_tr_target.setDecimals(3)
        self.rule_cf_min = QDoubleSpinBox(); self.rule_cf_min.setRange(0, 1.5); self.rule_cf_min.setDecimals(3)
        self.rule_cf_max = QDoubleSpinBox(); self.rule_cf_max.setRange(0, 1.5); self.rule_cf_max.setDecimals(3)
        self.rule_cf_delta = QDoubleSpinBox(); self.rule_cf_delta.setRange(0, 1); self.rule_cf_delta.setDecimals(3)
        self.rule_growth_min = QDoubleSpinBox(); self.rule_growth_min.setRange(-1, 1); self.rule_growth_min.setDecimals(4); self.rule_growth_min.setSingleStep(0.01)
        self.rule_growth_max = QDoubleSpinBox(); self.rule_growth_max.setRange(-1, 1); self.rule_growth_max.setDecimals(4); self.rule_growth_max.setSingleStep(0.01)
        self.rule_external = QCheckBox("允许区外送受电兜底")
        widgets = [
            self.rule_latest_actual, self.rule_ratio_min, self.rule_ratio_max,
            self.rule_tr_min, self.rule_tr_max, self.rule_tr_target,
            self.rule_cf_min, self.rule_cf_max, self.rule_cf_delta,
            self.rule_growth_min, self.rule_growth_max,
        ]
        for w in widgets:
            w.valueChanged.connect(self._mark_rules_dirty)
        self.rule_external.stateChanged.connect(self._mark_rules_dirty)
        form.addRow("最新现状年", self.rule_latest_actual)
        form.addRow("容载比下限", self.rule_ratio_min)
        form.addRow("容载比上限", self.rule_ratio_max)
        form.addRow("配变负载率下限", self.rule_tr_min)
        form.addRow("配变负载率上限", self.rule_tr_max)
        form.addRow("配变负载率软目标", self.rule_tr_target)
        form.addRow("同时率下限", self.rule_cf_min)
        form.addRow("同时率上限", self.rule_cf_max)
        form.addRow("同时率最大调整", self.rule_cf_delta)
        form.addRow("单年增长率下限", self.rule_growth_min)
        form.addRow("单年增长率上限", self.rule_growth_max)
        form.addRow("区外送受电", self.rule_external)
        lay.addLayout(form)

        btns = QHBoxLayout()
        self.btn_apply_rules = QPushButton("应用规则")
        self.btn_apply_rules.clicked.connect(self.apply_rules_from_widgets)
        self.btn_save_rules = QPushButton("保存规则档案")
        self.btn_save_rules.clicked.connect(self.save_current_rules)
        self.btn_load_rules = QPushButton("加载规则档案")
        self.btn_load_rules.clicked.connect(self.load_rules_from_file)
        self.btn_delete_invalid_presets = QPushButton("删除失效目标预设")
        self.btn_delete_invalid_presets.clicked.connect(self.delete_invalid_target_presets)
        btns.addWidget(self.btn_apply_rules)
        btns.addWidget(self.btn_save_rules)
        btns.addWidget(self.btn_load_rules)
        btns.addWidget(self.btn_delete_invalid_presets)
        btns.addStretch(1)
        lay.addLayout(btns)

        self.rule_status = QLabel("规则状态：默认规则")
        lay.addWidget(self.rule_status)
        self.rule_validation_table = QTableWidget()
        lay.addWidget(QLabel("规则项/目标预设校验：不匹配项会标红，可在目标列表中删除或在规则档案中修正。"))
        lay.addWidget(self.rule_validation_table, 1)

    def _init_result_tabs(self) -> None:
        """是什么：初始化方案、预测明细和日志页。

        为什么：反推方案不仅要给出结果，还要展示风险、写回说明和预测明细，便于
        中文用户复核后再导出。
        """
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

    def _nav_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        """是什么：左侧导航点击事件。

        为什么：业务人员不一定熟悉页签结构，左侧导航能让流程更清晰。
        """
        idx = NAV_NAMES.index(item.text(0))
        self.tabs.setCurrentIndex(idx)

    def _target_type_changed(self, text: str) -> None:
        """是什么：根据目标类型填充默认输入值。

        为什么：容载比、增长率、配变负载率常用输入不同，自动填充可以减少误输。
        """
        if text == "容载比":
            self.edit_target_value.setText("1.85")
            self.edit_min.setText(""); self.edit_max.setText("")
        elif text == "增长率":
            self.edit_target_value.setText("")
            self.edit_min.setText("5%"); self.edit_max.setText("7%")
        elif text == "配变平均负载率":
            self.edit_target_value.setText("0.5")
            self.edit_min.setText(""); self.edit_max.setText("")
        else:
            self.edit_target_value.setText("")
            self.edit_min.setText(""); self.edit_max.setText("")

    def _mark_rules_dirty(self, *args: Any) -> None:
        """是什么：标记规则已修改但未保存。

        为什么：规则是可复用资产，用户关闭程序或换文件前需要知道是否需要保存。
        """
        self.rules_dirty = True
        if hasattr(self, "rule_status"):
            self.rule_status.setText("规则状态：已修改未保存")

    def _parse_float(self, text: str) -> float | None:
        """是什么：解析中文界面输入的数字或百分数。

        为什么：业务人员常输入“6%”或“6”，工具需要统一转为 0.06，避免增长率错算。
        """
        text = (text or "").strip()
        if not text:
            return None
        is_percent = text.endswith("%")
        text = text.rstrip("%")
        val = float(text)
        if is_percent or abs(val) > 1.0:
            return val / 100.0
        return val

    def _target_from_form(self) -> ForecastTarget:
        """是什么：从表单构造一个 ForecastTarget。

        为什么：核心求解器不应该认识 Qt 控件，所以需要在 UI 层转换为业务对象。
        """
        ttype = self.cmb_target_type.currentText()
        area = self.cmb_area.currentText()
        voltage = self.cmb_voltage.currentText()
        year = int(self.spin_year.value())
        target_value = self._parse_float(self.edit_target_value.text())
        min_value = self._parse_float(self.edit_min.text())
        max_value = self._parse_float(self.edit_max.text())
        if ttype == "增长率":
            return ForecastTarget("增长率", area, voltage, period_start=self.rules.latest_actual_year, period_end=year, metric="网供负荷", target_value=target_value, min_value=min_value, max_value=max_value)
        if ttype == "配变平均负载率":
            return ForecastTarget("容载比", area, "10kV", year=year, metric="配变平均负载率", target_value=target_value, min_value=min_value, max_value=max_value, remark="按10kV负荷/容量需求校核")
        metric = "网供负荷" if ttype == "负荷" else None
        return ForecastTarget(ttype, area, voltage, year=year, metric=metric, target_value=target_value, min_value=min_value, max_value=max_value)

    def add_target_from_form(self) -> None:
        """是什么：把当前表单目标加入多目标列表。

        为什么：实际业务经常同时受多个年份、多个区域目标约束，多目标列表便于批量求解。
        """
        t = self._target_from_form()
        row = self.target_table.rowCount()
        self.target_table.insertRow(row)
        values = ["是", t.target_type, t.area, t.voltage_level, t.year or "", t.period_start or "", t.period_end or "", t.target_value if t.target_value is not None else "", self._range_text(t.min_value, t.max_value)]
        for col, value in enumerate(values):
            self.target_table.setItem(row, col, QTableWidgetItem(str(value)))

    def _range_text(self, lo: float | None, hi: float | None) -> str:
        """是什么：把上下限显示为文本。

        为什么：目标表格用一列承载范围，能减少列宽，也方便规则档案保存。
        """
        if lo is None and hi is None:
            return ""
        if lo is None:
            return f"-{hi}"
        if hi is None:
            return f"{lo}-"
        return f"{lo}-{hi}"

    def delete_selected_targets(self) -> None:
        """是什么：删除目标列表中的选中行。

        为什么：规则或模板不匹配时，用户需要能快速删掉不适用目标，而不是重建整组目标。
        """
        rows = sorted({idx.row() for idx in self.target_table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.target_table.removeRow(row)
        if rows:
            self.rules_dirty = True

    def clear_targets(self) -> None:
        """是什么：清空多目标列表。

        为什么：用户切换测算口径时，需要快速重新输入目标。
        """
        self.target_table.setRowCount(0)
        self.rules_dirty = True

    def _build_targets_from_ui(self) -> list[ForecastTarget]:
        """是什么：读取 UI 中的所有目标。

        为什么：求解器需要纯业务对象列表，不能直接依赖 QTableWidget。
        """
        if self.target_table.rowCount() == 0:
            return [self._target_from_form()]
        targets: list[ForecastTarget] = []
        for row in range(self.target_table.rowCount()):
            vals = [self.target_table.item(row, col).text().strip() if self.target_table.item(row, col) else "" for col in range(self.target_table.columnCount())]
            if vals[0] not in {"是", "true", "True", "1", "启用"}:
                continue
            ttype, area, voltage = vals[1], vals[2], vals[3]
            year = int(vals[4]) if vals[4] else None
            ps = int(vals[5]) if vals[5] else None
            pe = int(vals[6]) if vals[6] else None
            target_value = self._parse_float(vals[7])
            lo = hi = None
            if vals[8]:
                parts = vals[8].split("-", 1)
                if len(parts) == 2:
                    lo = self._parse_float(parts[0]); hi = self._parse_float(parts[1])
            if ttype == "增长率":
                targets.append(ForecastTarget("增长率", area, voltage, period_start=ps or self.rules.latest_actual_year, period_end=pe or year, metric="网供负荷", target_value=target_value, min_value=lo, max_value=hi))
            else:
                metric = "网供负荷" if ttype == "负荷" else None
                targets.append(ForecastTarget(ttype, area, voltage, year=year, metric=metric, target_value=target_value, min_value=lo, max_value=hi))
        return targets

    def _apply_rules_to_widgets(self) -> None:
        """是什么：把规则对象显示到规则控件。

        为什么：加载规则档案或打开模板后，界面需要同步反映当前生效规则。
        """
        self.rule_latest_actual.blockSignals(True)
        self.rule_latest_actual.setValue(self.rules.latest_actual_year)
        self.rule_latest_actual.blockSignals(False)
        self.rule_ratio_min.setValue(self.rules.ratio_min)
        self.rule_ratio_max.setValue(self.rules.ratio_max)
        self.rule_tr_min.setValue(self.rules.transformer_load_rate_min)
        self.rule_tr_max.setValue(self.rules.transformer_load_rate_max)
        self.rule_tr_target.setValue(self.rules.transformer_load_rate_soft_target)
        self.rule_cf_min.setValue(self.rules.coincidence_factor_min)
        self.rule_cf_max.setValue(self.rules.coincidence_factor_max)
        self.rule_cf_delta.setValue(self.rules.coincidence_factor_max_abs_change)
        self.rule_growth_min.setValue(self.rules.annual_growth_min)
        self.rule_growth_max.setValue(self.rules.annual_growth_max)
        self.rule_external.setChecked(self.rules.allow_external_exchange)

    def _rules_from_widgets(self) -> ForecastRules:
        """是什么：从规则控件构造 ForecastRules。

        为什么：求解器和导出器只依赖核心规则对象，保证 UI 与计算逻辑解耦。
        """
        latest = self.rule_latest_actual.value()
        return ForecastRules(
            latest_actual_year=latest,
            forecast_start_year=latest + 1,
            ratio_min=self.rule_ratio_min.value(),
            ratio_max=self.rule_ratio_max.value(),
            transformer_load_rate_min=self.rule_tr_min.value(),
            transformer_load_rate_max=self.rule_tr_max.value(),
            transformer_load_rate_soft_target=self.rule_tr_target.value(),
            coincidence_factor_min=self.rule_cf_min.value(),
            coincidence_factor_max=self.rule_cf_max.value(),
            coincidence_factor_max_abs_change=self.rule_cf_delta.value(),
            annual_growth_min=self.rule_growth_min.value(),
            annual_growth_max=self.rule_growth_max.value(),
            allow_external_exchange=self.rule_external.isChecked(),
        )

    def apply_rules_from_widgets(self) -> None:
        """是什么：把界面规则应用到当前任务。

        为什么：用户可能只想临时改变规则做一次测算，不一定立即保存规则档案。
        """
        self.rules = self._rules_from_widgets()
        self.rules_dirty = True
        self.rule_status.setText("规则状态：已应用但未保存")
        QMessageBox.information(self, "规则已应用", "当前界面规则已应用到本次反推任务。")

    def _profile_targets(self) -> list[ForecastTarget]:
        """是什么：提取要保存到规则档案的目标预设。

        为什么：目标预设读取失败不应影响规则保存，所以这里做容错返回空列表。
        """
        try:
            return self._build_targets_from_ui()
        except Exception:
            return []

    def save_current_rules(self) -> None:
        """是什么：保存当前规则和目标预设为 YAML 档案。

        为什么：同一模板后续还会重复预测，保存规则能显著减少重复配置。
        """
        if not self.current_file or not self.workbook_info:
            QMessageBox.information(self, "提示", "请先打开一个模板，再保存规则档案。")
            return
        self.rules = self._rules_from_widgets()
        default = default_rule_path_for_workbook(self.current_file)
        out, _ = QFileDialog.getSaveFileName(self, "保存规则档案", str(default), "规则档案 (*.yaml)")
        if not out:
            return
        profile = build_rule_profile(self.current_file, self.workbook_info, self.records, self.rules, self._profile_targets())
        save_rule_profile(out, profile)
        self.current_rule_profile = profile
        self.current_rule_path = Path(out)
        self.rules_dirty = False
        self.rule_status.setText(f"规则状态：已保存 {self.current_rule_path.name}")
        self._refresh_rule_validation(profile)

    def load_rules_from_file(self) -> None:
        """是什么：手动加载规则档案。

        为什么：自动匹配规则并不总是用户想要的，手动加载可以支持多套测算口径。
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "加载规则档案", "rules", "规则档案 (*.yaml)")
        if not file_path:
            return
        self._load_rule_profile(Path(file_path), ask_when_warning=True)

    def _load_rule_profile(self, path: Path, ask_when_warning: bool = True) -> None:
        """是什么：加载并校验某个规则档案。

        为什么：规则可能不匹配当前模板，必须在应用前提示用户，避免错误套用。
        """
        profile = load_rule_profile(path)
        if self.current_file and self.workbook_info:
            sig = build_workbook_signature(self.workbook_info, self.records)
            match = compare_rule_profile(profile, self.current_file, sig)
            if match.level == "error":
                reply = QMessageBox.question(self, "规则不匹配", "\n".join(match.messages) + "\n\n是否强制加载？")
                if reply != QMessageBox.StandardButton.Yes:
                    return
            elif match.level == "warning" and ask_when_warning:
                reply = QMessageBox.question(self, "规则可能不完全匹配", "\n".join(match.messages) + "\n\n是否继续加载？")
                if reply != QMessageBox.StandardButton.Yes:
                    return
            self._refresh_rule_validation(profile)
        self.rules = forecast_rules_from_dict(profile.get("rules"))
        self.current_rule_profile = profile
        self.current_rule_path = path
        self.rules_dirty = False
        self._apply_rules_to_widgets()
        self.rule_status.setText(f"规则状态：已加载 {path.name}")
        self._load_target_presets(profile)

    def _load_target_presets(self, profile: dict[str, Any]) -> None:
        """是什么：把规则档案中的目标预设加载到目标表。

        为什么：常见测算往往复用一组目标，自动回填可以减少重复输入。
        """
        presets = profile.get("target_presets", []) or []
        self.target_table.setRowCount(0)
        if not presets:
            return
        for item in presets:
            t = target_from_dict(item)
            row = self.target_table.rowCount(); self.target_table.insertRow(row)
            vals = ["是", t.target_type, t.area, t.voltage_level, t.year or "", t.period_start or "", t.period_end or "", t.target_value if t.target_value is not None else "", self._range_text(t.min_value, t.max_value)]
            for col, value in enumerate(vals):
                self.target_table.setItem(row, col, QTableWidgetItem(str(value)))

    def _refresh_rule_validation(self, profile: dict[str, Any]) -> None:
        """是什么：展示规则项/目标预设与当前模板的匹配结果。

        为什么：不匹配项需要标红，让用户知道哪些历史目标需要删除或修正。
        """
        results = validate_target_presets(profile, self.records)
        rows = [[r.status, r.target, "；".join(r.messages)] for r in results]
        set_table_data(self.rule_validation_table, ["状态", "规则项", "说明"], rows)
        colors = {"matched": QColor("#D9EAD3"), "warning": QColor("#FCE5CD"), "missing": QColor("#F4CCCC"), "invalid": QColor("#E06666")}
        for row, r in enumerate(results):
            color = colors.get(r.status)
            if color:
                for col in range(self.rule_validation_table.columnCount()):
                    item = self.rule_validation_table.item(row, col)
                    if item:
                        item.setBackground(color)

    def delete_invalid_target_presets(self) -> None:
        """是什么：删除当前规则档案中的失效目标预设。

        为什么：模板升级后可能只有部分目标不匹配。删除 missing/invalid 项，
        可以保留仍然有效的规则和目标，避免整份规则重建。
        """
        if not self.current_rule_profile:
            QMessageBox.information(self, "提示", "请先加载或保存一个规则档案。")
            return
        results = validate_target_presets(self.current_rule_profile, self.records)
        bad_indexes = {int(r.item_id) - 1 for r in results if r.item_id.isdigit() and r.status in {"missing", "invalid"}}
        if not bad_indexes:
            QMessageBox.information(self, "提示", "当前没有 missing / invalid 状态的目标预设。")
            return
        reply = QMessageBox.question(self, "确认删除", f"将删除 {len(bad_indexes)} 条失效目标预设，是否继续？")
        if reply != QMessageBox.StandardButton.Yes:
            return
        presets = self.current_rule_profile.get("target_presets", []) or []
        self.current_rule_profile["target_presets"] = [p for i, p in enumerate(presets) if i not in bad_indexes]
        self._load_target_presets(self.current_rule_profile)
        self._refresh_rule_validation(self.current_rule_profile)
        self.rules_dirty = True
        self.rule_status.setText("规则状态：已删除失效预设，尚未保存")

    def _auto_load_rule_for_current_file(self) -> None:
        """是什么：打开工作簿后自动加载同主名规则。

        为什么：业务人员通常按同一模板持续测算，自动加载可以直接进入目标设置。
        """
        if not self.current_file:
            return
        path = default_rule_path_for_workbook(self.current_file)
        if path.exists():
            try:
                self._load_rule_profile(path, ask_when_warning=False)
                self.statusBar().showMessage(f"已自动加载规则：{path.name}")
            except Exception as exc:
                self.statusBar().showMessage(f"自动加载规则失败：{exc}")

    def open_file(self) -> None:
        """是什么：选择并读取 Excel 模板。

        为什么：读取阶段需要完成安全加载、模板识别、核心指标抽取、规则自动加载和界面刷新。
        """
        if self.rules_dirty:
            reply = QMessageBox.question(self, "规则未保存", "当前规则已修改但未保存，是否继续打开新文件？")
            if reply != QMessageBox.StandardButton.Yes:
                return
        file_path, _ = QFileDialog.getOpenFileName(self, "选择供需预测模板", "", "Excel文件 (*.xlsx *.xls)")
        if not file_path:
            return
        try:
            self.current_file = Path(file_path)
            self.loaded = load_workbook_auto(file_path)
            self.workbook_info = build_workbook_info(self.loaded, latest_actual_year=self.rules.latest_actual_year)
            self.records = extract_all_metric_records(self.loaded, self.workbook_info)
            self.rules.latest_actual_year = self.workbook_info.latest_actual_year
            self.rules.forecast_start_year = self.workbook_info.forecast_start_year
            self._apply_rules_to_widgets()
            self._auto_load_rule_for_current_file()
            self._refresh_template_table()
            self._refresh_compat_table()
            self._refresh_target_options()
            self.lbl_status.setText(f"当前文件：{self.current_file.name}；记录数：{len(self.records)}")
            self.statusBar().showMessage("模板读取完成")
        except Exception as exc:
            QMessageBox.critical(self, "读取失败", f"读取文件失败：\n{exc}\n\n{traceback.format_exc()}")

    def _refresh_template_table(self) -> None:
        """是什么：刷新模板识别结果。

        为什么：用户需要知道每个 sheet 被识别成什么业务角色，尤其是项目库动态识别结果。
        """
        rows = []
        for info in self.workbook_info.sheet_infos.values():
            years = ", ".join(str(y.year) + ("现状" if y.is_actual else "") for y in info.year_columns)
            rows.append([info.name, info.role.value, info.area_name or "", info.voltage_level or "", years, f"{info.confidence:.2f}", ";".join(info.notes)])
        set_table_data(self.template_table, ["Sheet", "角色", "区域", "电压等级", "年份列", "置信度", "说明"], rows)

    def _refresh_compat_table(self) -> None:
        """是什么：刷新兼容性检测结果。

        为什么：模板中可能存在隐藏表、保护表、复杂公式或外部链接，需要在导出前提醒。
        """
        issues = check_workbook_compatibility(self.loaded)
        rows = [[i.level, i.sheet, i.location, i.issue_type, i.content, i.suggestion] for i in issues]
        set_table_data(self.compat_table, ["级别", "Sheet", "位置", "问题类型", "内容", "建议"], rows)
        if has_blocking_errors(issues):
            QMessageBox.warning(self, "兼容性警告", "检测到 Error 级问题，业务模型仍可查看，但请先处理后再正式导出。")

    def _refresh_target_options(self) -> None:
        """是什么：根据当前模板刷新目标区域、电压等级和年份范围。

        为什么：不同模板区县和年份可能不同，目标设置不能写死。
        """
        self.cmb_area.clear(); self.cmb_voltage.clear()
        self.cmb_area.addItems(list_areas(self.records))
        self.cmb_voltage.addItems([v for v in list_voltage_levels(self.records) if v != "总计"])
        years = list_years(self.records)
        if years:
            future_years = [y for y in years if y >= self.rules.forecast_start_year]
            self.spin_year.setValue(max(future_years) if future_years else max(years))
            self.spin_forecast_end.setValue(max(years))
            self.spin_year.setMinimum(self.rules.forecast_start_year)
            self.spin_forecast_end.setMinimum(self.rules.forecast_start_year)

    def run_solver(self) -> None:
        """是什么：根据目标和规则生成三类反推方案。

        为什么：用户需要比较“优先改负荷”“负荷容量均衡”“区外送受电兜底”不同风险的方案。
        """
        if not self.records:
            QMessageBox.information(self, "提示", "请先打开Excel模板。")
            return
        try:
            self.rules = self._rules_from_widgets()
            targets = self._build_targets_from_ui()
            forecast_end = max(int(self.spin_forecast_end.value()), max([t.year or t.period_end or self.rules.forecast_start_year for t in targets]))
            self.scenarios = solve_forecast(self.records, targets, self.rules, forecast_end)
            self.cmb_scenario.clear()
            self.cmb_scenario.addItems([s.name for s in self.scenarios])
            self._scenario_changed(0)
            self.btn_export.setEnabled(True)
            self.tabs.setCurrentIndex(4)
        except Exception as exc:
            QMessageBox.critical(self, "反推失败", f"生成方案失败：\n{exc}\n\n{traceback.format_exc()}")

    def _scenario_changed(self, index: int) -> None:
        """是什么：切换当前展示的反推方案。

        为什么：同一目标可能有多种调整路径，方案切换需要同步刷新调整项、预测明细和风险提示。
        """
        if not self.scenarios or index < 0:
            return
        self.selected_scenario = self.scenarios[index]
        adj_rows = [
            [a.object_type, a.area, a.voltage_level, a.metric, a.year, a.old_value, a.new_value, a.delta, a.delta_pct, a.reason, a.priority, a.risk_level, a.source_sheet or "", a.source_cell or "", a.note]
            for a in self.selected_scenario.adjustments
        ]
        set_table_data(self.result_table, ["调整对象", "区域", "电压等级", "指标", "年份", "原值", "建议值", "变化量", "变化比例", "原因", "优先级", "风险", "来源表", "来源单元格", "备注"], adj_rows)
        pred_rows = [[r.area, r.voltage_level, r.metric, r.year, r.value, r.source_sheet, r.source_cell, "是" if r.is_actual else "否"] for r in sorted(self.selected_scenario.forecast_records, key=lambda r: (r.area, r.voltage_level, r.metric, r.year))[:3000]]
        set_table_data(self.forecast_table, ["区域", "电压等级", "指标", "年份", "结果值", "来源表", "来源单元格", "是否现状"], pred_rows)
        text = [self.selected_scenario.description, ""]
        text.append("风险提示：")
        text.extend(self.selected_scenario.warnings or ["暂无明显风险。"])
        self.log_text.setPlainText("\n".join(text))

    def export_file(self) -> None:
        """是什么：导出新 Excel 文件。

        为什么：必须保证原始传入文件不被修改；所有写回只发生在导出副本，并新增预测结果表留痕。
        """
        if not self.selected_scenario or not self.loaded or not self.current_file:
            QMessageBox.information(self, "提示", "请先生成方案。")
            return
        if self.rules_dirty:
            reply = QMessageBox.question(self, "规则未保存", "当前规则已修改但未保存，是否仍继续导出？")
            if reply != QMessageBox.StandardButton.Yes:
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

    def closeEvent(self, event):  # noqa: N802 - Qt API
        """是什么：关闭窗口前检查规则未保存状态。

        为什么：规则配置通常是长期资产，误关程序会导致重复配置。
        """
        if self.rules_dirty:
            reply = QMessageBox.question(self, "规则未保存", "当前规则已修改但未保存，确认退出？")
            if reply != QMessageBox.StandardButton.Yes:
                event.ignore(); return
        event.accept()
