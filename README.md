# 负荷预测与容载比反推助手

面向供需预测及平衡分析模板的 Windows 桌面端工具。

本工具用于辅助负荷预测工作，重点解决：

- 根据上级指定的容载比目标反推未来负荷；
- 根据增长率目标生成平滑预测；
- 检查容载比、配变平均负载率、同时率等关键指标；
- 将用户确认后的调整值写入导出的 Excel 副本；
- 在导出文件中新增“预测结果表”，记录目标达成、调整项、预测明细和写回清单。

> 重要：程序不会修改用户传入的原始 Excel 文件。所有写回都发生在导出的新文件副本中。

---

## 当前版本

V2 MVP

第二版重点：

- 规则档案保存与自动加载；
- 规则/目标预设匹配校验；
- 不匹配目标预设标红与删除；
- 多目标配置；
- 新增年份动态预测和写回；
- 更完整的“预测结果表”；
- 关键方法增加“是什么、为什么”中文注释；
- 新增中文使用说明和开发文档。

---

## 项目定位

本项目不是“万能 Excel 公式引擎”。它是面向供需预测模板的业务反推工具：

```text
Excel模板导入
↓
模板角色识别
↓
动态年份识别
↓
业务指标抽取
↓
目标与规则设置
↓
反推方案生成
↓
写入导出副本
↓
新增预测结果表
```

核心思路：

```text
Excel 作为输入输出载体
业务模型作为计算核心
```

---

## 已支持范围

### 文件

- `.xlsx`
- `.xls`（读取后转换为内部临时结构，导出统一为 `.xlsx`）

### 模板识别

支持识别以下类型：

- 地市容载比表；
- 区县容载比表；
- 校核表；
- 直供负荷明细；
- 变电容量明细；
- 电源装机明细；
- 储能装机明细；
- 110/35kV 项目库。

项目库不依赖固定名称。以下名称均可通过字段识别：

- `十四五 十五五 110-35明细表`
- `十五五 十六五 110-35明细表`
- `十六五 十七五 110-35明细表`
- `110-35项目明细表`

### 动态年份

不限制年份到 2030。工具会自动识别：

```text
2025年（现状）
2026年
2030年
2031年
2035年
```

默认规则：

- 2025 年及以前为现状数据，不允许修改；
- 2026 年及以后为预测期，可参与反推。

### 关键规则

- 容载比硬边界：`1.3–2.5`；
- 同时率范围：`0.83–0.99`，单次调整不超过 `0.1`；
- 配变平均负载率：硬边界 `0–1`，软目标在 `0.5` 附近；
- 项目投产年不可修改；
- 区外送（+）受（-）电作为低优先级兜底变量。

---

## 安装与运行

建议使用 Python 3.11 或 3.12。

Windows：

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

Linux/macOS 开发环境：

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

---

## 文档

- [使用说明](docs/使用说明.md)
- [开发文档](docs/开发文档.md)
- [开发方案说明](docs/开发方案.md)

---

## 测试

```bash
pytest
```

---

## 打包 EXE

安装 PyInstaller：

```bash
pip install pyinstaller
```

开发阶段 onedir：

```bash
pyinstaller --noconfirm --onedir --windowed --name LoadForecastingAssistant main.py
```

发布阶段 onefile：

```bash
pyinstaller --noconfirm --onefile --windowed --name LoadForecastingAssistant main.py
```

---

## 目录结构

```text
LoadForecastingAssistant/
├── main.py
├── requirements.txt
├── config/default_rules.yaml
├── core/
│   ├── workbook_loader.py
│   ├── compatibility_checker.py
│   ├── template_parser.py
│   ├── rule_engine.py
│   ├── rule_persistence.py
│   ├── forecast_engine.py
│   ├── workbook_writer.py
│   └── path_visualizer.py
├── ui/
│   ├── main_window.py
│   └── table_utils.py
├── tests/
├── rules/
└── docs/
```

---

## 安全与注意事项

1. 原始 Excel 文件不会被修改。
2. 导出文件是原文件副本 + 写回调整值 + 新增“预测结果表”。
3. 若写回目标是公式单元格，界面提供“允许覆盖公式单元格”选项。
4. 当前版本仍属于 MVP，正式生产使用前建议用典型模板进行人工复核。
5. 区外送受电属于低优先级兜底变量，出现该方案时必须人工确认站点和调度口径。

---

## GitHub 推送提醒

请不要把 GitHub Token 写入代码或提交到仓库。
