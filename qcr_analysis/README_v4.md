# QCR数据分析系统 v4.0

## 🎯 系统概述

全新重构的质量控制记录分析系统，采用**三层架构**，支持**三大分析模式**，提供**Web界面**。

---

## 📂 目录结构

```
qcr_analysis/
├── main_v4.py                  # 统一入口
├── run_web.bat                 # 快速启动
├── config.py                   # 系统配置
├── prompts.py                  # AI提示词库
├── data/                       # 数据层
├── services/                   # 功能层
├── web/                        # Web界面层
└── modules/                    # 原有模块（保留）
```

---

## 🚀 快速启动

### Web模式（推荐）
```bash
python main_v4.py
```
或双击 `run_web.bat`

### 命令行模式
```bash
# Weekly Report
python main_v4.py --cli --mode weekly \
  --data "数据.xlsx" \
  --mtm "MTM.xlsx" \
  --output "output" \
  --start-date "2024-04-09" \
  --end-date "2025-11-23" \
  --batch-name "2024-2025" \
  --filter-unmapped \
  --generate-ppt

# Top Issue
python main_v4.py --cli --mode top-issue \
  --data "数据.xlsx" --mtm "MTM.xlsx" \
  --top-n 10 --generate-ppt

# Top Model  
python main_v4.py --cli --mode top-model \
  --data "数据.xlsx" --mtm "MTM.xlsx" \
  --top-n 15 --generate-ppt
```

---

## 🎮 三大分析模式

### 1. Weekly Report
- 7天无理由分析
- 非7天无理由分析
- 机型分布统计
- **输出**: `weekly_report_{batch_name}.pptx`

### 2. Top Issue
- Top N问题统计
- 机型分布分析
- **输出**: `top_issue_report_{batch_name}.pptx`

### 3. Top Model
- 基于**问题类别数量**的机型排名
- Top 15机型深度分析
- **输出**: `top_model_report_{batch_name}.pptx`

---

## 🏗️ 架构

```
Web UI Layer (Flask)
    ↓
Service Layer
    ├─ WeeklyAnalysisService
    ├─ TopIssueAnalysisService
    ├─ TopModelAnalysisService
    ├─ VisualizationService
    └─ ReportService
    ↓
Data Layer
    ├─ DataManager
    └─ MTMManager
```

---

## 📊 安装

```bash
pip install -r requirements.txt
```

---

## 🧪 测试

```bash
python test_services.py
```

---

版本：v4.0  
状态：生产就绪 ✅

