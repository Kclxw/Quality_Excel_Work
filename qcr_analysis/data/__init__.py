# -*- coding: utf-8 -*-
"""
=============================================================================
数据层模块
=============================================================================
提供统一的数据访问接口
"""

from .data_manager import DataManager, load_data
from modules.mtm_manager import MTMManager

__all__ = [
    'DataManager',
    'MTMManager',
    'load_data'
]

