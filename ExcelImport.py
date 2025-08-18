import os
import pandas as pd
from pathlib import Path
from typing import List, Dict, Union
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 常量配置
REQUIRED_COLUMNS = {
    'date': '日期',
    'service_id': '服务单号',
    'order_id': '订单号',
    'problem_desc': '问题描述',
    'sku': 'SKU',
    'sn_code': 'SN编码',
    'customer_id': ['客户账号', '客户账户'],  # 支持多种可能的列名
    'product_series': ['商品名称','产品系列'],
    'audit_reason': '审核原因',
    'problem_type': ['问题分类', '问题分类一'],
    'category': ['分类', '问题分类二']
}

COLUMN_MAPPINGS = {
    '问题分类一': '问题分类',
    '问题分类二': '分类',
    '客户账户': '客户账号',
    '产品系列': '商品名称'
}

TARGET_COLUMNS = [
    '日期', '服务单号', '订单号', '问题描述', 'SKU', 
    'SN编码', '客户账号', '商品名称', '审核原因', 
    '问题分类', '分类'
]

class ExcelProcessor:
    def __init__(self, base_path: str):
        self.base_path = Path(base_path)
        self.combined_data = pd.DataFrame()
        
    def find_excel_files(self) -> List[Path]:
        """查找所有符合条件的Excel文件"""
        excel_files = []
        target_files = ['持续落入D等级 30天服务单明细.xlsx', '新增D等级服务单明细.xlsx']
        
        for root, _, files in os.walk(self.base_path):
            for file in files:
                if file in target_files:
                    excel_files.append(Path(root) / file)
        return excel_files

    def process_sheet(self, df: pd.DataFrame) -> pd.DataFrame:
        """处理单个sheet的数据"""
        # 重命名列
        df = df.rename(columns=COLUMN_MAPPINGS)
        
        # 选择所需的列
        available_columns = []
        for target_col in TARGET_COLUMNS:
            if target_col in df.columns:
                available_columns.append(target_col)
            else:
                df[target_col] = '-'  # 添加缺失的列并填充默认值
        
        # 填充空值
        df = df.fillna('-')
        
        # 处理日期格式
        try:
            df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
        except:
            df['日期'] = df['日期'].fillna('1900-01-01')
        
        return df[TARGET_COLUMNS]

    def process_file(self, file_path: Path) -> None:
        """处理单个Excel文件"""
        try:
            # 读取所有sheet
            excel_file = pd.ExcelFile(file_path)
            total_sheets = len(excel_file.sheet_names)
            
            logging.info(f"文件 {file_path.name} 包含 {total_sheets} 个工作表")
            
            for sheet_idx, sheet_name in enumerate(excel_file.sheet_names, 1):
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    original_rows = len(df)
                    
                    if original_rows == 0:
                        logging.info(f"  Sheet [{sheet_idx}/{total_sheets}] {sheet_name}: 空工作表，跳过处理")
                        continue
                    
                    logging.info(f"  开始处理 Sheet [{sheet_idx}/{total_sheets}] {sheet_name}: 原始数据 {original_rows} 行")
                    
                    processed_df = self.process_sheet(df)
                    processed_rows = len(processed_df)
                    
                    self.combined_data = pd.concat([self.combined_data, processed_df], ignore_index=True)
                    logging.info(f"  完成处理 Sheet {sheet_name}: 处理后 {processed_rows} 行数据")
                    
                except Exception as e:
                    logging.error(f"  处理文件 {file_path.name} 的 sheet {sheet_name} 时出错: {str(e)}")
                    
        except Exception as e:
            logging.error(f"读取文件 {file_path.name} 失败: {str(e)}")

    def process_all_files(self) -> None:
        """处理所有文件并生成最终结果"""
        excel_files = self.find_excel_files()
        total_files = len(excel_files)
        
        if total_files == 0:
            logging.warning("未找到任何符合条件的Excel文件")
            return
            
        logging.info(f"共找到 {total_files} 个Excel文件需要处理")
        
        for file_idx, file_path in enumerate(excel_files, 1):
            logging.info(f"\n处理第 {file_idx}/{total_files} 个文件: {file_path.name}")
            self.process_file(file_path)
        
        # 去重
        if not self.combined_data.empty:
            original_rows = len(self.combined_data)
            self.combined_data = self.combined_data.drop_duplicates(
                subset=['服务单号'], 
                keep='first'
            )
            final_rows = len(self.combined_data)
            
            # 保存结果
            output_path = self.base_path / 'Sumdata.xlsx'
            self.combined_data.to_excel(output_path, index=False)
            logging.info(f"\n处理完成总结:")
            logging.info(f"- 合并数据总行数: {original_rows}")
            logging.info(f"- 去重后最终行数: {final_rows}")
            logging.info(f"- 删除重复行数: {original_rows - final_rows}")
            logging.info(f"- 结果文件已保存: {output_path}")
        else:
            logging.warning("没有找到有效数据")

def main():
    # 设置数据文件夹路径
    base_path = Path(r"D:\WorkDocument\WeeklyReport\QCR\DataImport")
    
    # 创建处理器实例并执行处理
    processor = ExcelProcessor(base_path)
    processor.process_all_files()

if __name__ == "__main__":
    main()