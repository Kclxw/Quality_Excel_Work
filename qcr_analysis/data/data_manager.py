# -*- coding: utf-8 -*-
"""
=============================================================================
数据管理器 - 统一数据层
=============================================================================
整合Excel读写和数据库操作，提供统一的数据接口
"""

import pandas as pd
from pathlib import Path
from typing import Optional, Dict, List, Tuple
from datetime import date, datetime
import sys

# 导入已有的数据库管理器
sys.path.append(str(Path(__file__).parent.parent))
from modules.database import DatabaseManager as DBManager
from config import DB_CONFIG


class DataManager:
    """统一数据管理器：整合Excel和数据库操作"""
    
    def __init__(self, db_config: Optional[Dict] = None):
        """
        初始化数据管理器
        
        Args:
            db_config: 数据库配置字典，为None时使用默认配置
        """
        self.db_config = db_config or DB_CONFIG
        self.db_manager = None
        self._last_df = None  # 缓存最后加载的数据
    
    # ================================================================
    # Excel 操作
    # ================================================================
    
    def read_excel(self, file_path: str, sheet_name: int = 0) -> pd.DataFrame:
        """
        从Excel文件读取数据
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表索引，默认第一个
            
        Returns:
            DataFrame
        """
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            self._last_df = df.copy()
            return df
        except Exception as e:
            raise IOError(f"读取Excel文件失败: {e}")
    
    def write_excel(self, df: pd.DataFrame, file_path: str, sheet_name: str = "Sheet1"):
        """
        将数据写入Excel文件
        
        Args:
            df: 要写入的DataFrame
            file_path: 目标文件路径
            sheet_name: 工作表名称
        """
        try:
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
        except Exception as e:
            raise IOError(f"写入Excel文件失败: {e}")
    
    # ================================================================
    # 数据库操作
    # ================================================================
    
    def connect_database(self) -> bool:
        """
        连接数据库
        
        Returns:
            连接是否成功
        """
        try:
            self.db_manager = DBManager(self.db_config)
            return True
        except Exception as e:
            print(f"数据库连接失败: {e}")
            return False
    
    def read_from_database(
        self, 
        start_date: Optional[date] = None,
        end_date: Optional[date] = None,
        filters: Optional[Dict] = None
    ) -> pd.DataFrame:
        """
        从数据库读取数据
        
        Args:
            start_date: 开始日期
            end_date: 结束日期
            filters: 其他筛选条件
            
        Returns:
            DataFrame
        """
        if not self.db_manager:
            raise RuntimeError("数据库未连接，请先调用 connect_database()")
        
        try:
            df = self.db_manager.read_data_by_date_range(start_date, end_date)
            self._last_df = df.copy()
            return df
        except Exception as e:
            raise RuntimeError(f"从数据库读取数据失败: {e}")
    
    def write_to_database(self, df: pd.DataFrame, skip_duplicates: bool = True) -> int:
        """
        将数据写入数据库
        
        Args:
            df: 要写入的DataFrame
            skip_duplicates: 是否跳过重复数据
            
        Returns:
            写入的记录数
        """
        if not self.db_manager:
            raise RuntimeError("数据库未连接，请先调用 connect_database()")
        
        try:
            if skip_duplicates:
                df = self.db_manager.check_and_import_new_data(df)
            else:
                self.db_manager.import_data(df)
            return len(df)
        except Exception as e:
            raise RuntimeError(f"写入数据库失败: {e}")
    
    def update_mtm_in_database(self, mtm_file_path: str) -> bool:
        """
        更新数据库中的MTM映射
        
        Args:
            mtm_file_path: MTM映射文件路径
            
        Returns:
            更新是否成功
        """
        if not self.db_manager:
            raise RuntimeError("数据库未连接，请先调用 connect_database()")
        
        try:
            self.db_manager.update_mtm_mapping(mtm_file_path)
            return True
        except Exception as e:
            print(f"更新MTM映射失败: {e}")
            return False
    
    # ================================================================
    # 数据筛选和预处理
    # ================================================================
    
    def filter_by_date_range(
        self,
        df: pd.DataFrame,
        start_date: Optional[date] = None,
        end_date: Optional[date] = None,
        date_column: Optional[str] = None
    ) -> pd.DataFrame:
        """
        按日期范围筛选数据
        
        Args:
            df: 原始DataFrame
            start_date: 开始日期
            end_date: 结束日期
            date_column: 日期列名，为None时使用第一列
            
        Returns:
            筛选后的DataFrame
        """
        if df.empty:
            return df
        
        # 确定日期列
        if date_column is None:
            date_column = df.columns[0]
        
        # 转换日期列
        df[date_column] = pd.to_datetime(df[date_column]).dt.date
        
        # 应用筛选
        if start_date and end_date:
            mask = (df[date_column] >= start_date) & (df[date_column] <= end_date)
        elif start_date:
            mask = df[date_column] >= start_date
        elif end_date:
            mask = df[date_column] <= end_date
        else:
            return df
        
        return df[mask].copy()
    
    def filter_by_audit_reason(
        self,
        df: pd.DataFrame,
        reason_type: str = "7day"
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        按审核原因分类数据
        
        Args:
            df: 原始DataFrame
            reason_type: 分类类型，"7day"返回7天和非7天
            
        Returns:
            (7天无理由DataFrame, 非7天无理由DataFrame)
        """
        if reason_type == "7day":
            cond_7d = df["审核原因"] == "7天无理由"
            cond_non_7d = df["审核原因"].isin(["15天质量换新", "180天只换不修", "质量维修"])
            
            df_7d = df[cond_7d].copy()
            df_non_7d = df[cond_non_7d].copy()
            
            return df_7d, df_non_7d
        else:
            # 返回按审核原因分组的字典
            grouped = {}
            for reason in df["审核原因"].unique():
                grouped[reason] = df[df["审核原因"] == reason].copy()
            return grouped
    
    def filter_unmapped_mtm(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        过滤未映射的MTM记录
        
        Args:
            df: 包含"机型名称"和"MTM"列的DataFrame
            
        Returns:
            只包含已映射记录的DataFrame
        """
        if "机型名称" not in df.columns or "MTM" not in df.columns:
            print("警告：DataFrame中缺少'机型名称'或'MTM'列")
            return df
        
        # 机型名称 != MTM 表示已映射
        return df[df["机型名称"] != df["MTM"]].copy()
    
    # ================================================================
    # 数据统计
    # ================================================================
    
    def get_model_statistics(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        获取机型统计信息
        
        Args:
            df: 包含"机型名称"列的DataFrame
            
        Returns:
            机型统计DataFrame（机型名称、记录数、问题类别数）
        """
        if "机型名称" not in df.columns:
            raise ValueError("DataFrame中缺少'机型名称'列")
        
        stats = df.groupby("机型名称").agg({
            "机型名称": "count",  # 记录数
            "问题分类": "nunique"  # 问题类别数
        }).rename(columns={
            "机型名称": "记录数",
            "问题分类": "问题类别数"
        })
        
        stats = stats.reset_index()
        return stats.sort_values("记录数", ascending=False)
    
    def get_issue_statistics(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        获取问题统计信息
        
        Args:
            df: 包含"问题分类"列的DataFrame
            
        Returns:
            问题统计DataFrame
        """
        if "问题分类" not in df.columns:
            raise ValueError("DataFrame中缺少'问题分类'列")
        
        issue_stats = df["问题分类"].value_counts().reset_index()
        issue_stats.columns = ["问题分类", "数量"]
        issue_stats["占比(%)"] = (issue_stats["数量"] / len(df) * 100).round(2)
        
        return issue_stats
    
    # ================================================================
    # 工具方法
    # ================================================================
    
    def get_last_dataframe(self) -> Optional[pd.DataFrame]:
        """获取最后加载的DataFrame"""
        return self._last_df.copy() if self._last_df is not None else None
    
    def validate_columns(self, df: pd.DataFrame, required_columns: List[str]) -> bool:
        """
        验证DataFrame是否包含必需的列
        
        Args:
            df: 要验证的DataFrame
            required_columns: 必需的列名列表
            
        Returns:
            是否包含所有必需列
        """
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            print(f"警告：缺少以下列：{', '.join(missing)}")
            return False
        return True
    
    def close_database(self):
        """关闭数据库连接"""
        if self.db_manager:
            self.db_manager.close()
            self.db_manager = None


# ================================================================
# 便捷函数
# ================================================================

def load_data(
    source: str,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
    use_database: bool = False,
    db_config: Optional[Dict] = None
) -> pd.DataFrame:
    """
    便捷函数：从Excel或数据库加载数据
    
    Args:
        source: Excel文件路径 或 "database"
        start_date: 开始日期
        end_date: 结束日期
        use_database: 是否使用数据库
        db_config: 数据库配置
        
    Returns:
        DataFrame
    """
    manager = DataManager(db_config)
    
    if use_database or source.lower() == "database":
        manager.connect_database()
        df = manager.read_from_database(start_date, end_date)
    else:
        df = manager.read_excel(source)
        if start_date or end_date:
            df = manager.filter_by_date_range(df, start_date, end_date)
    
    return df

