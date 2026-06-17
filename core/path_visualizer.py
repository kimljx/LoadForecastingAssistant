"""公式/业务影响路径可视化数据构建。"""
from __future__ import annotations

from dataclasses import dataclass, field

from .models import Adjustment, MetricRecord


@dataclass(slots=True)
class TreeNode:
    name: str
    node_type: str
    value: float | None = None
    children: list["TreeNode"] = field(default_factory=list)
    color: str = "black"
    note: str = ""


def build_business_tree(records: list[MetricRecord], area: str, voltage: str, metric: str, year: int, adjustments: list[Adjustment] | None = None) -> TreeNode:
    """构建简化业务影响树。

    第一版以业务关系展示为主，不完整解析 Excel 公式链。
    """
    adjustments = adjustments or []
    changed_keys = {(a.area, a.voltage_level, a.metric, a.year) for a in adjustments}
    value_map = {(r.area, r.voltage_level, r.metric, r.year): r.value for r in records}

    root_key = (area, voltage, metric, year)
    root = TreeNode(f"{area} {voltage} {metric} {year}", "目标", value_map.get(root_key), color="blue")

    if metric == "容载比":
        cap_metric = "变电容量需求" if voltage == "10kV" else "变电容量"
        for m, typ in [(cap_metric, "容量"), ("网供负荷", "负荷")]:
            key = (area, voltage, m, year)
            color = "purple" if key in changed_keys else ("orange" if m == cap_metric else "green")
            root.children.append(TreeNode(f"{area} {voltage} {m} {year}", typ, value_map.get(key), color=color))
    elif metric == "网供负荷":
        root.children.append(TreeNode("区县负荷分摊", "业务关系", None, color="orange", note="地市指标可由区县合计和同时率共同影响"))
        root.children.append(TreeNode("区外送(+)/受(-)电", "兜底变量", None, color="red", note="低优先级，需人工确认站点"))
    return root


def render_tree_text(node: TreeNode, prefix: str = "") -> str:
    line = f"{prefix}{node.name}"
    if node.value is not None:
        line += f"：{node.value:.4g}"
    if node.note:
        line += f"（{node.note}）"
    lines = [line]
    for idx, child in enumerate(node.children):
        connector = "└─ " if idx == len(node.children) - 1 else "├─ "
        child_prefix = prefix + connector
        lines.append(render_tree_text(child, child_prefix))
    return "\n".join(lines)
