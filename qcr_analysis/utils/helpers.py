# -*- coding: utf-8 -*-
"""
=============================================================================
辅助工具函数
=============================================================================
通用的辅助函数
=============================================================================
"""

from datetime import datetime, date
from typing import Optional


def parse_date(date_str: str) -> date:
    """
    尝试解析多种日期格式
    
    Args:
        date_str: 日期字符串
        
    Returns:
        date对象
        
    Raises:
        ValueError: 无法解析日期时抛出
    """
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"无法解析日期: {date_str}，请使用 YYYY-MM-DD 或 YYYY/MM/DD 格式")


def format_percentage(value: float, decimals: int = 1) -> str:
    """
    格式化百分比
    
    Args:
        value: 数值
        decimals: 小数位数
        
    Returns:
        格式化后的百分比字符串
    """
    return f"{value:.{decimals}f}%"


def parse_percentage(value) -> Optional[float]:
    """
    解析百分比字符串为浮点数
    
    Args:
        value: 百分比值（可以是字符串或数字）
        
    Returns:
        浮点数，解析失败返回None
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    value_str = str(value).strip().replace('%', '')
    if not value_str:
        return None
    try:
        return float(value_str)
    except ValueError:
        return None

