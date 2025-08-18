import pandas as pd
import os

def extract_and_compare_sn_data(table1_path, table2_path, output_path):
    """
    从表格1中提取SN编码列，从表格2中提取主机编号列，
    比对数据并将匹配的行输出到新表格
    
    参数:
    table1_path: 表格1的文件路径（包含SN编码列）
    table2_path: 表格2的文件路径（包含主机编号列）
    output_path: 输出表格3的文件路径
    """
    
    try:
        # 读取表格1和表格2
        print(f"正在读取表格1: {table1_path}")
        df1 = pd.read_excel(table1_path)
        
        print(f"正在读取表格2: {table2_path}")
        df2 = pd.read_excel(table2_path)
        
        # 提取SN编码列和主机编号列
        sn_data1 = df1['SN编码'].dropna().astype(str).str.strip()
        sn_data2 = df2['主机编号'].dropna().astype(str).str.strip()
        
        print(f"表格1中提取到 {len(sn_data1)} 个SN编码")
        print(f"表格2中提取到 {len(sn_data2)} 个主机编号")
        
        # 找出sn_data2中与sn_data1相同的数据
        matching_sn = sn_data2[sn_data2.isin(sn_data1)]
        
        print(f"找到 {len(matching_sn)} 个匹配的主机编号")
        
        # 筛选表格2中匹配的行
        mask = df2['主机编号'].astype(str).str.strip().isin(sn_data1)
        df3 = df2[mask].copy()
        
        print(f"从表格2中提取出 {len(df3)} 行匹配数据")
        
        # 保存到新表格3
        df3.to_excel(output_path, index=False)
        print(f"匹配数据已保存到: {output_path}")
        
        return df3
        
    except KeyError as e:
        print(f"错误：找不到指定的列 - {e}")
        return None
    except Exception as e:
        print(f"处理过程中出现错误: {e}")
        return None

def main():
    """主函数 - 示例用法"""
    
    # 示例文件路径（请根据实际情况修改）
    table2_file = r"D:\工作文件\WeeklyReport\QCR\QCR_Data_Clean\TB维修机台汇总(.xlsx"  # 包含主机编号列的表格
    table1_file = r"D:\工作文件\WeeklyReport\QCR\QCR_Data_Clean\闪屏-0703-0724.xlsx"  # 包含SN编码列的表格
    output_file = r"D:\工作文件\WeeklyReport\QCR\QCR_Data_Clean\Cleandata.xlsx"  # 输出表格3
    
    # 检查文件是否存在
    if not os.path.exists(table1_file):
        print(f"错误：找不到文件 {table1_file}")
        return
    
    if not os.path.exists(table2_file):
        print(f"错误：找不到文件 {table2_file}")
        return
    
    # 执行数据提取和比对
    result = extract_and_compare_sn_data(table1_file, table2_file, output_file)
    
    if result is not None:
        print("\n处理完成！")
        print(f"匹配数据预览：")
        print(result.head())
    else:
        print("\n处理失败，请检查文件路径和列名是否正确")

if __name__ == "__main__":
    main()
