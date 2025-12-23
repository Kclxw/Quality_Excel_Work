# -*- coding: utf-8 -*-
"""
=============================================================================
MTM映射关系管理
=============================================================================
预定义的MTM到机型名称的映射关系
此文件由 sync_mtm_to_code.py 自动生成，以MTM.xlsx为准
优先级：预定义映射 > MTM.xlsx文件映射
=============================================================================
"""

# 预定义的MTM映射关系（已禁用，仅从MTM.xlsx文件加载）
# 所有映射关系现在都从MTM.xlsx文件中导入
PREDEFINED_MTM_MAPPINGS = {}


def get_mtm_mapping(mtm: str) -> str:
    """
    获取MTM对应的机型名称
    
    Args:
        mtm: MTM编码
        
    Returns:
        机型名称，如果未找到则返回原MTM
    """
    return PREDEFINED_MTM_MAPPINGS.get(mtm, mtm)


def has_predefined_mapping(mtm: str) -> bool:
    """
    检查MTM是否在预定义映射中
    
    Args:
        mtm: MTM编码
        
    Returns:
        True如果存在预定义映射，否则False
    """
    return mtm in PREDEFINED_MTM_MAPPINGS


def add_mapping(mtm: str, model_name: str):
    """
    添加新的映射关系到预定义映射（运行时）
    
    Args:
        mtm: MTM编码
        model_name: 机型名称
    """
    PREDEFINED_MTM_MAPPINGS[mtm] = model_name


def get_all_mappings() -> dict:
    """
    获取所有预定义映射
    
    Returns:
        字典形式的所有映射关系
    """
    return PREDEFINED_MTM_MAPPINGS.copy()


def get_mappings_count() -> int:
    """
    获取预定义映射数量
    
    Returns:
        映射关系数量
    """
    return len(PREDEFINED_MTM_MAPPINGS)
