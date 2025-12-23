# -*- coding: utf-8 -*-
"""
=============================================================================
MTM映射管理器
=============================================================================
负责MTM与机型名称的映射关系管理
执行逻辑：
仅从MTM.xlsx文件中加载映射关系，不使用预定义映射，不从数据中提取映射
=============================================================================
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Optional, Tuple

import sys
sys.path.append(str(Path(__file__).parent.parent))
from data.mtm_mappings import (
    get_mtm_mapping,
    has_predefined_mapping,
    add_mapping,
    get_all_mappings,
    get_mappings_count
)


class MTMManager:
    """MTM映射管理器"""
    
    def __init__(self, mtm_file_path: Optional[Path] = None):
        """
        初始化MTM管理器
        
        Args:
            mtm_file_path: MTM映射表文件路径（必需）
        """
        self.mtm_file_path = mtm_file_path
        self.file_mappings = {}     # 从文件加载的映射（唯一映射来源）
        
        # 加载文件映射（如果文件存在）
        if mtm_file_path and mtm_file_path.exists():
            self._load_file_mappings()
        else:
            print(f"⚠️  警告：MTM映射文件不存在，无法加载映射关系")
    
    def _load_file_mappings(self):
        """从MTM.xlsx文件加载映射关系"""
        try:
            # 先尝试使用header=0读取（假设有表头）
            mtm_df = pd.read_excel(self.mtm_file_path, sheet_name=0, header=0)
            
            # 检查第一行是否是表头
            if mtm_df.columns[0] == 'MTM' or 'MTM' in str(mtm_df.columns[0]).upper():
                # 有表头，直接使用
                if len(mtm_df.columns) >= 2:
                    mtm_df.columns = ['MTM', '机型名称']
                else:
                    print("警告：MTM文件格式不正确，至少需要两列")
                    self.file_mappings = {}
                    return
            else:
                # 没有表头，重新读取
                mtm_df = pd.read_excel(self.mtm_file_path, sheet_name=0, header=None)
                mtm_df.columns = ['MTM', '机型名称']
            
            # 过滤掉表头行（如果MTM列的值就是"MTM"）
            mtm_df = mtm_df[mtm_df['MTM'] != 'MTM']
            mtm_df = mtm_df[mtm_df['MTM'] != '机型名称']
            
            # 创建映射字典
            self.file_mappings = dict(zip(mtm_df['MTM'], mtm_df['机型名称']))
            print(f"✓ 从文件加载了 {len(self.file_mappings)} 条MTM映射关系")
        except Exception as e:
            print(f"警告：加载MTM文件失败: {e}")
            self.file_mappings = {}
    
    def get_model_name(self, mtm: str) -> str:
        """
        获取MTM对应的机型名称
        仅从MTM.xlsx文件映射中查找
        
        Args:
            mtm: MTM编码
            
        Returns:
            机型名称，如果未找到则返回原MTM
        """
        # 仅从文件映射中查找
        if mtm in self.file_mappings:
            return self.file_mappings[mtm]
        
        # 未找到映射，返回原MTM
        return mtm
    
    def map_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        为DataFrame添加机型名称列
        
        Args:
            df: 包含MTM列的DataFrame
            
        Returns:
            添加了"机型名称"列的DataFrame
        """
        if 'MTM' not in df.columns:
            print("警告：DataFrame中未找到'MTM'列")
            return df
        
        # 应用映射
        df['机型名称'] = df['MTM'].apply(self.get_model_name)
        
        # 统计映射情况
        unmapped_count = (df['机型名称'] == df['MTM']).sum()
        total_count = len(df)
        mapped_count = total_count - unmapped_count
        
        print(f"✓ MTM映射完成: {mapped_count}/{total_count} 条记录已映射")
        if unmapped_count > 0:
            print(f"  注意: {unmapped_count} 条记录未找到映射关系，使用原MTM值")
            print(f"  💡 提示: 使用 --filter-unmapped-mtm 参数可以只分析已映射的机型")
        
        return df
    
    def get_mapped_mtms(self) -> set:
        """
        获取所有已映射的MTM集合
        
        Returns:
            已映射的MTM集合
        """
        return set(self.file_mappings.keys())
    
    def update_mappings_from_data(self, df: pd.DataFrame, model_name_column: str = '商品名称') -> int:
        """
        此功能已禁用 - 不再从数据中提取映射关系
        所有映射关系仅从MTM.xlsx文件中加载
        
        Args:
            df: 原始数据DataFrame
            model_name_column: 机型名称所在列名
            
        Returns:
            始终返回0
        """
        # 已禁用此功能
        return 0
    
    def save_to_file(self, output_path: Optional[Path] = None) -> bool:
        """
        保存映射关系到Excel文件
        仅保存文件映射（MTM.xlsx的内容）
        
        Args:
            output_path: 输出文件路径，默认为原MTM文件路径
            
        Returns:
            保存是否成功
        """
        if output_path is None:
            output_path = self.mtm_file_path
        
        if output_path is None:
            print("错误：未指定输出文件路径")
            return False
        
        try:
            # 只保存文件映射
            all_mappings = self.file_mappings.copy()
            
            # 转换为DataFrame
            mapping_df = pd.DataFrame(
                list(all_mappings.items()),
                columns=['MTM', '机型名称']
            )
            
            # 按MTM排序
            mapping_df = mapping_df.sort_values('MTM').reset_index(drop=True)
            
            # 保存到文件
            mapping_df.to_excel(output_path, index=False, header=False)
            print(f"✓ MTM映射表已保存到: {output_path}")
            print(f"  总计 {len(all_mappings)} 条映射关系")
            
            return True
        except Exception as e:
            print(f"错误：保存MTM映射表失败: {e}")
            return False
    
    def print_statistics(self):
        """打印映射统计信息"""
        print("\n" + "="*60)
        print("MTM映射统计")
        print("="*60)
        print(f"MTM.xlsx映射数量: {len(self.file_mappings)}")
        print("="*60 + "\n")

