# -*- coding: utf-8 -*-
"""
=============================================================================
QCR 数据分析工具 - 配置文件
=============================================================================
集中管理所有配置参数，方便后续维护
=============================================================================
"""

import os

# -----------------------------
# 数据库配置
# -----------------------------
DB_CONFIG = {
    'host': 'localhost',
    'port': 3306,
    'user': 'root',
    'password': '09291',
    'database': 'local_qcr',
    'table_name': 'QCR_data'  # 修改为大写，与SQL定义一致
}

# -----------------------------
# API配置 - Kimi LLM
# -----------------------------
# 方式1: 直接在代码中配置（不推荐提交到版本控制）
DEFAULT_KIMI_API_KEY = "sk-z4mdCQLUIpPYoMwz7CMTonTHT8rgzgiaDOkkut5AJaHgU8wh"

# 方式2: 从环境变量读取（推荐，环境变量优先级高于代码配置）
KIMI_API_KEY = os.getenv("KIMI_API_KEY", DEFAULT_KIMI_API_KEY)
KIMI_API_URL = os.getenv("KIMI_API_URL", "https://api.moonshot.cn/v1/chat/completions")
KIMI_MODEL = os.getenv("KIMI_MODEL", "kimi-k2-0905-preview")
KIMI_TIMEOUT = int(os.getenv("KIMI_TIMEOUT", "60"))

# -----------------------------
# LLM参数配置
# -----------------------------
LLM_TOP_N = int(os.getenv("LLM_TOP_N", "3"))
LLM_COVERAGE_THRESHOLD = float(os.getenv("LLM_COVERAGE_THRESHOLD", "80"))
LLM_FOCUS_THRESHOLD = float(os.getenv("LLM_FOCUS_THRESHOLD", "10"))

# -----------------------------
# 文件路径配置
# -----------------------------
DEFAULT_OUTPUT_DIR = "output"
DEFAULT_MTM_FILE = "mtm.xlsx"
DEFAULT_PPT_PATH = "report.pptx"

# -----------------------------
# PPT样式配置
# -----------------------------
PPT_STYLE = {
    'font_name': '微软雅黑',
    'title_font_size': 28,
    'subtitle_font_size': 28,
    'body_font_size': 14,
}

# -----------------------------
# 图表样式配置
# -----------------------------
CHART_STYLE = {
    'pie_chart_size': (8, 8),
    'bar_chart_size': (12, 6),
    'reason_chart_size': (6, 6),
}

# -----------------------------
# 数据处理配置
# -----------------------------
# 审核原因类型
AUDIT_REASONS = ["15天质量换新", "180天只换不修", "7天无理由", "质量维修"]

# 分类后缀
CATEGORY_SUFFIXES = ["7天无理由", "非7天无理由"]

# 数据库字段映射
DB_COLUMN_MAPPING = {
    '服务单号': 'service_order_id',
    '日期': 'date',
    '订单号': 'order_id',
    '问题描述': 'issue_description',
    'SKU': 'sku',
    'SN编码': 'sn_code',
    '客户账号': 'customer_account',
    '客户账户': 'customer_account',
    '商品名称': 'product_name',
    '产品名称': 'product_name',  # 添加产品名称的映射
    'MTM': 'mtm',
    '审核原因': 'audit_reason',
    '问题分类': 'issue_category',
    '分类': 'category'
}

# 数据库字段类型
DB_NUMERIC_COLUMNS = ['service_order_id', 'order_id', 'sku']
DB_STRING_COLUMNS = [
    'issue_description', 'sn_code', 'customer_account',
    'product_name', 'mtm', 'audit_reason',
    'issue_category', 'category'
]

# 字符串字段最大长度
DB_STRING_MAX_LENGTHS = {
    'issue_description': 500,
    'sn_code': 100,
    'customer_account': 100,
    'product_name': 200,
    'mtm': 100,  # 修改为100，与SQL定义一致
    'audit_reason': 100,
    'issue_category': 100,
    'category': 100
}

# 必需的数据库字段
DB_REQUIRED_COLUMNS = [
    'service_order_id', 'date', 'order_id', 'issue_description',
    'sku', 'sn_code', 'customer_account', 'product_name',
    'mtm', 'audit_reason', 'issue_category', 'category'
]

# -----------------------------
# Matplotlib中文字体配置
# -----------------------------
MATPLOTLIB_FONTS = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']

