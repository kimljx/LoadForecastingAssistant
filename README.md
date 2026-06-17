# 负荷预测与容载比反推助手

面向供需预测及平衡分析模板的 Windows 桌面端工具。

本工具用于辅助负荷预测工作，重点解决：

- 根据上级指定的容载比目标反推未来负荷；
- 根据增长率目标生成平滑预测；
- 检查容载比、配变平均负载率、同时率等关键指标；
- 将用户确认后的调整值写入导出的 Excel 副本；
- 在导出文件中新增“预测结果表”，记录目标达成、调整项和写回清单。

> 重要：程序不会修改用户传入的原始 Excel 文件。所有写回都发生在导出的新文件副本中。

## 当前定位

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

## 已支持范围

### 文件

- `.xlsx`
- `.xls`（读取后转换为内部临时 xlsx，导出统一为 xlsx）

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

## 安装与运行

建议使用 Python 3.11。

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

## 测试

```bash
pytest
```

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

## GitHub 推送

如果本地已经配置 GitHub 权限：

```bash
git init
git add .
git commit -m "初始化负荷预测与容载比反推助手"
git branch -M main
git remote add origin https://github.com/kimljx/LoadForecastingAssistant.git
git push -u origin main
```

如果使用 Token：

```bash
git remote set-url origin https://<YOUR_GITHUB_TOKEN>@github.com/kimljx/LoadForecastingAssistant.git
git push -u origin main
```

请不要把 Token 写入代码或提交到仓库。

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
│   ├── forecast_engine.py
│   ├── workbook_writer.py
│   └── path_visualizer.py
├── ui/
│   ├── main_window.py
│   └── table_utils.py
├── tests/
└── docs/
```

## 注意事项

第一版会把用户确认后的调整值写入导出的副本。若写回目标是公式单元格，界面提供“允许覆盖公式单元格”选项。实际业务中更推荐逐步增强为“优先写入底层输入项，让原模板公式自动重算”。
