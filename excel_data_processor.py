import os
import pandas as pd
from datetime import datetime

def process_excel_files(root_dir):
    """
    处理指定目录下所有Excel文件，合并数据后输出到Sumdata.xlsx
    
    参数:
    root_dir (str): 要处理的根目录路径
    
    返回:
    bool: 处理是否成功
    """
    # 最终合并的数据集
    all_data = []
    
    # 列名映射关系
    column_mapping = {
        "持续落入D等级 30天服务单明细.xlsx": {
            "日期": "日期",
            "服务单号": "服务单号",
            "订单号": "订单号",
            "问题描述": "问题描述",
            "SKU": "SKU",
            "SN编码": "SN编码",
            "客户账号": "客户账号",
            "产品系列": "产品系列",
            "审核原因": "审核原因",
            "问题分类": "问题分类",
            "分类": "分类"
        },
        "新增D等级服务单明细.xlsx": {
            "日期": "日期",
            "服务单号": "服务单号",
            "订单号": "订单号",
            "问题描述": "问题描述",
            "SKU": "SKU",
            "SN编码": "SN编码",
            "客户账户": "客户账号",  # 映射到标准列名
            "产品系列": "产品系列",
            "审核原因": "审核原因",
            "问题分类一": "问题分类",  # 映射到标准列名
            "问题分类二": "分类"     # 映射到标准列名
        }
    }
    
    # 标准输出列顺序
    output_columns = [
        "日期", "服务单号", "订单号", "问题描述", "SKU", 
        "SN编码", "客户账号", "产品系列", "审核原因", "问题分类", "分类"
    ]
    
    # 遍历目录结构
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename not in column_mapping:
                continue
                
            file_path = os.path.join(dirpath, filename)
            print(f"处理文件: {file_path}")
            
            try:
                # 读取Excel文件中的所有sheet
                xls = pd.ExcelFile(file_path)
                
                for sheet_name in xls.sheet_names:
                    try:
                        # 读取sheet数据
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        
                        # 跳过空sheet
                        if df.empty:
                            print(f"  ⚠️ 空工作表: {sheet_name}")
                            continue
                            
                        # 删除完全空白的行
                        df.dropna(how='all', inplace=True)
                        
                        # 列名映射和重命名
                        mapping = column_mapping[filename]
                        df.rename(columns=mapping, inplace=True)
                        
                        # 检查必要列是否存在
                        required_cols = set(mapping.values())
                        missing_cols = required_cols - set(df.columns)
                        if missing_cols:
                            print(f"  ❌ 缺少必要列: {', '.join(missing_cols)}")
                            continue
                            
                        # 选择需要的列
                        df = df[list(required_cols)]
                        
                        # 服务单号去重 (保留首次出现)
                        df.drop_duplicates(subset='服务单号', keep='first', inplace=True)
                        
                        # 日期格式处理
                        if '日期' in df.columns:
                            df['日期'] = pd.to_datetime(df['日期'], errors='coerce').dt.strftime('%Y-%m-%d')
                        
                        # 空值处理
                        df.fillna('-', inplace=True)
                        
                        # 添加到总数据集
                        all_data.append(df)
                        print(f"  ✅ 成功处理工作表: {sheet_name}, 数据行数: {len(df)}")
                        
                    except Exception as sheet_e:
                        print(f"  ❌ 处理工作表 {sheet_name} 错误: {str(sheet_e)}")
                        
            except Exception as e:
                print(f"❌ 文件读取失败: {file_path}, 错误: {str(e)}")
    
    # 合并所有数据
    if not all_data:
        print("⚠️ 未找到有效数据处理")
        return False
        
    final_df = pd.concat(all_data, ignore_index=True)
    
    # 最终去重
    final_df.drop_duplicates(subset='服务单号', keep='first', inplace=True)
    
    # 按标准列顺序输出
    final_df = final_df[output_columns]
    
    # 输出文件路径
    output_path = os.path.join(root_dir, "Sumdata.xlsx")
    final_df.to_excel(output_path, index=False)
    print(f"\n✅ 处理完成! 共处理 {len(final_df)} 条数据")
    print(f"📁 输出文件: {output_path}")
    return True

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("用法: python excel_data_processor.py <根目录路径>")
        print("示例: python excel_data_processor.py D:\\DataImport")
        sys.exit(1)
        
    root_directory = sys.argv[1]
    
    if not os.path.exists(root_directory):
        print(f"错误: 路径不存在 {root_directory}")
        sys.exit(1)
        
    process_excel_files(root_directory)
