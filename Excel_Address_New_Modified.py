# -*- coding: utf-8 -*-
"""
处理逻辑
1. 读取Excel数据
2. 指定处理时间周期（基于日期列）
3. 去重MTM
4. 读取MTM表格并映射机型名称
5. 统计四种审核原因
6. 统计7天无理由/非7天无理由的机型分布
7. 按机型统计分类描述词频次
8. 为每个机型的所有分类生成详细数据文件
9. 生成表格+饼图/柱状图

使用方法:
python Excel_Address_New_Modified.py <输入文件路径> [MTM表格路径] [输出目录路径] [开始日期] [结束日期]
示例:
python Excel_Address_New_Modified.py "持续落入D等级 30天服务单明细.xlsx" "mtm.xlsx" "output" "2025-07-01" "2025-07-18"
"""

import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import sys
import os
import re
from datetime import datetime
import pymysql
from sqlalchemy import create_engine

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 设置中文字体（Windows 示例）
import matplotlib
matplotlib.rcParams['font.family'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False

# 处理字体警告 - 使用更兼容的字体设置
import warnings
warnings.filterwarnings("ignore", category=UserWarning, message=".*Glyph.*missing.*")

# -----------------------------
# 数据库配置
# -----------------------------
DB_CONFIG = {
    'host': 'localhost',
    'port': 3306,
    'user': 'root',
    'password': '0929',
    'database': 'local_qcr'
}

# -----------------------------
# 工具函数：清理文件名中的非法字符
# -----------------------------
def sanitize_filename(filename):
    """清理文件名中的非法字符"""
    # Windows非法字符：<>:"/\|?*
    illegal_chars = r'[<>:\"/\\|?*]'
    # 替换为空格
    filename = re.sub(illegal_chars, ' ', filename)
    # 去除前后空格
    filename = filename.strip()
    # 限制长度
    if len(filename) > 200:
        filename = filename[:200]
    return filename

# -----------------------------
# 数据库工具函数
# -----------------------------
def check_and_import_new_data(df):
    """检查数据库中不存在的服务单号并导入新数据"""
    try:
        print("开始连接数据库...")
        # 创建数据库连接
        connection_string = (
            f"mysql+pymysql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@"
            f"{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
        )
        engine = create_engine(connection_string)
        
        # 检查数据框中是否有服务单号列
        service_order_column = None
        for col in df.columns:
            if '服务单号' in str(col):
                service_order_column = col
                break
        
        if service_order_column is None:
            print("警告：未找到'服务单号'列，跳过数据库检查")
            return df
        
        # 获取当前数据中的服务单号
        current_service_orders = df[service_order_column].dropna().astype(str).tolist()
        print(f"当前数据包含 {len(current_service_orders)} 个服务单号")
        
        # 查询数据库中已存在的服务单号
        try:
            # 先检查表是否存在
            table_exists = pd.read_sql(
                "SELECT COUNT(*) as count FROM information_schema.tables WHERE table_schema = %s AND table_name = 'qcr_data'", 
                engine, 
                params=(DB_CONFIG['database'],)
            )['count'].iloc[0] > 0
            
            if table_exists:
                existing_service_orders = pd.read_sql(
                    "SELECT service_order_id FROM qcr_data", 
                    engine
                )['service_order_id'].astype(str).tolist()
                print(f"数据库中已存在 {len(existing_service_orders)} 个服务单号")
            else:
                print("数据库表qcr_data不存在，将创建新表")
                existing_service_orders = []
        except Exception as e:
            print(f"查询数据库失败，假设数据库为空: {e}")
            existing_service_orders = []
        
        # 筛选出数据库中不存在的新服务单号
        new_service_orders = [
            order for order in current_service_orders 
            if order not in existing_service_orders
        ]
        print(f"新服务单号数量: {len(new_service_orders)}")
        
        # 筛选新数据
        df_new = df[df[service_order_column].astype(str).isin(new_service_orders)].copy()
        
        if len(df_new) == 0:
            print("没有新数据需要导入和分析")
            return df_new
        
        # 准备导入数据库的数据
        df_to_import = df_new.copy()
        
        # 重命名列以匹配数据库字段
        column_mapping = {}
        for col in df_to_import.columns:
            if '服务单号' in str(col):
                column_mapping[col] = 'service_order_id'
            elif '日期' in str(col):
                column_mapping[col] = 'date'
            elif '订单号' in str(col):
                column_mapping[col] = 'order_id'
            elif '问题描述' in str(col):
                column_mapping[col] = 'issue_description'
            elif 'SKU' in str(col):
                column_mapping[col] = 'sku'
            elif 'SN编码' in str(col):
                column_mapping[col] = 'sn_code'
            elif '客户账号' in str(col) or '客户账户' in str(col):
                column_mapping[col] = 'customer_account'
            elif '商品名称' in str(col):
                column_mapping[col] = 'product_name'
            elif 'MTM' in str(col):
                column_mapping[col] = 'mtm'
            elif '审核原因' in str(col):
                column_mapping[col] = 'audit_reason'
            elif '问题分类' in str(col) and '一' not in str(col):
                column_mapping[col] = 'issue_category'
            elif '分类' in str(col) and '问题' not in str(col):
                column_mapping[col] = 'category'
        
        # 应用列映射
        df_to_import = df_to_import.rename(columns=column_mapping)
        
        # 确保必需的列存在
        required_db_columns = [
            'service_order_id', 'date', 'order_id', 'issue_description', 
            'sku', 'sn_code', 'customer_account', 'product_name', 
            'mtm', 'audit_reason', 'issue_category', 'category'
        ]
        
        for col in required_db_columns:
            if col not in df_to_import.columns:
                df_to_import[col] = ''
        
        # 数据类型转换和清洗
        # 处理日期
        if 'date' in df_to_import.columns:
            df_to_import['date'] = pd.to_datetime(df_to_import['date'], errors='coerce').dt.strftime('%Y-%m-%d')
            df_to_import['date'] = df_to_import['date'].fillna('1900-01-01')
        
        # 处理数值列
        numeric_columns = ['service_order_id', 'order_id', 'sku']
        for col in numeric_columns:
            if col in df_to_import.columns:
                df_to_import[col] = pd.to_numeric(df_to_import[col], errors='coerce').astype('Int64')
        
        # 处理字符串列
        string_columns = [
            'issue_description', 'sn_code', 'customer_account', 
            'product_name', 'mtm', 'audit_reason', 
            'issue_category', 'category'
        ]
        for col in string_columns:
            if col in df_to_import.columns:
                df_to_import[col] = df_to_import[col].fillna('').astype(str).str.strip()
                # 字符串长度限制
                max_lengths = {
                    'issue_description': 500,
                    'sn_code': 100,
                    'customer_account': 100,
                    'product_name': 200,
                    'mtm': 50,
                    'audit_reason': 100,
                    'issue_category': 100,
                    'category': 100
                }
                if col in max_lengths:
                    df_to_import[col] = df_to_import[col].str[:max_lengths[col]]
        
        # 删除必需字段为空的行
        if 'service_order_id' in df_to_import.columns:
            df_to_import = df_to_import.dropna(subset=['service_order_id'])
        
        # 只选择数据库需要的列
        df_to_import = df_to_import[required_db_columns]
        
        # 导入数据到数据库
        if len(df_to_import) > 0:
            try:
                df_to_import.to_sql(
                    'qcr_data', 
                    engine, 
                    if_exists='append', 
                    index=False,
                    method='multi'
                )
                print(f"成功导入 {len(df_to_import)} 条新记录到数据库")
            except Exception as e:
                print(f"导入数据到数据库失败: {e}")
                print("将继续分析当前数据，但新数据不会保存到数据库")
        
        # 返回新数据用于后续分析
        return df_new
        
    except Exception as e:
        print(f"数据库操作失败: {e}")
        print("将继续分析原始数据，跳过数据库检查和导入")
        return df

# -----------------------------
# 1. 命令行参数处理
# -----------------------------
if len(sys.argv) < 2:
    print("用法: python Excel_Address_New_Modified.py <输入文件路径> [MTM表格路径] [输出目录路径] [开始日期] [结束日期]")
    print("日期格式: YYYY-MM-DD")
    print("示例: python Excel_Address_New_Modified.py \"持续落入D等级 30天服务单明细.xlsx\" \"mtm.xlsx\" \"output\" \"2025-07-01\" \"2025-07-18\"")
    sys.exit(1)

# 获取输入文件路径
file_path = Path(sys.argv[1])
if not file_path.exists():
    print(f"错误: 文件 '{file_path}' 不存在")
    sys.exit(1)

# 获取MTM表格路径（默认为当前目录下的mtm.xlsx）
if len(sys.argv) >= 3:
    mtm_file_path = Path(sys.argv[2])
else:
    mtm_file_path = Path("mtm.xlsx")

if not mtm_file_path.exists():
    print(f"警告: MTM表格文件 '{mtm_file_path}' 不存在，将使用原始MTM值")
    use_mtm_mapping = False
else:
    use_mtm_mapping = True

# 获取输出目录路径（默认为当前目录下的output）
if len(sys.argv) >= 4:
    out_dir = Path(sys.argv[3])
else:
    out_dir = Path("output")

out_dir.mkdir(exist_ok=True)
sheet_name = 0  # 默认第一张表

# -----------------------------
# 日期解析函数
# -----------------------------
def parse_date(date_str):
    """尝试解析多种日期格式"""
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"无法解析日期: {date_str}，请使用 YYYY-MM-DD 或 YYYY/MM/DD 格式")

# -----------------------------
# 获取日期范围参数
# -----------------------------
start_date = None
end_date = None
if len(sys.argv) >= 5:
    try:
        start_date = parse_date(sys.argv[4])
    except ValueError as e:
        print(f"警告: {e}，将处理所有数据")
if len(sys.argv) >= 6:
    try:
        end_date = parse_date(sys.argv[5])
    except ValueError as e:
        print(f"警告: {e}，将处理所有数据")

# -----------------------------
# 2. 读取数据
# -----------------------------
df = pd.read_excel(file_path, sheet_name=sheet_name)

# 假设第一列是日期列，转换为日期格式
date_column = df.columns[0]
df[date_column] = pd.to_datetime(df[date_column]).dt.date

# 根据日期范围筛选数据
if start_date and end_date:
    mask = (df[date_column] >= start_date) & (df[date_column] <= end_date)
    df = df[mask]
    print(f"已筛选 {start_date} 到 {end_date} 的数据，共 {len(df)} 条记录")
elif start_date:
    mask = df[date_column] >= start_date
    df = df[mask]
    print(f"已筛选 {start_date} 之后的数据，共 {len(df)} 条记录")
elif end_date:
    mask = df[date_column] <= end_date
    df = df[mask]
    print(f"已筛选 {end_date} 之前的数据，共 {len(df)} 条记录")

# -----------------------------
# 数据库检查和导入新数据
# -----------------------------
print("\n开始检查数据库中已存在的服务单号...")
df = check_and_import_new_data(df)
print(f"数据库检查后，剩余 {len(df)} 条新记录需要分析\n")

# -----------------------------
# 3. 去重MTM
# -----------------------------
# original_count = len(df)
# df = df.drop_duplicates(subset=['MTM'])
# print(f"已去重MTM，从 {original_count} 条记录减少到 {len(df)} 条记录")

# -----------------------------
# 4. 读取MTM映射表
# -----------------------------
if use_mtm_mapping:
    mtm_df = pd.read_excel(mtm_file_path, sheet_name=sheet_name, header=None)
    mtm_df.columns = ['MTM', '机型名称']
    mtm_mapping = dict(zip(mtm_df['MTM'], mtm_df['机型名称']))
    
    # 映射MTM到机型名称
    df['机型名称'] = df['MTM'].map(mtm_mapping).fillna(df['MTM'])
else:
    # 如果没有MTM映射表，使用原始MTM值作为机型名称
    df['机型名称'] = df['MTM']

# -----------------------------
# 5. 预计算常用条件
# -----------------------------
cond_7d = df["审核原因"] == "7天无理由"
cond_non_7d = df["审核原因"].isin(["15天质量换新", "180天只换不修", "质量维修"])

# 缓存中间结果
df_7d = df[cond_7d].copy()
df_non_7d = df[cond_non_7d].copy()

# 创建文件夹结构
detailed_dir_7d = out_dir / "详细数据" / "7天无理由"
detailed_dir_non7d = out_dir / "详细数据" / "非7天无理由"
detailed_dir_7d.mkdir(parents=True, exist_ok=True)
detailed_dir_non7d.mkdir(parents=True, exist_ok=True)

# -----------------------------
# 6. 统计四种审核原因
# -----------------------------
reasons = ["15天质量换新", "180天只换不修", "7天无理由", "质量维修"]
counts = {r: int((df["审核原因"] == r).sum()) for r in reasons}

summary_df = pd.DataFrame(list(counts.items()), columns=["审核原因", "数量"])
summary_df.to_excel(out_dir / "审核原因统计.xlsx", index=False)

plt.figure(figsize=(6, 6))
plt.pie(summary_df["数量"], labels=summary_df["审核原因"], autopct="%1.1f%%")
plt.title("审核原因占比")
plt.tight_layout()
plt.savefig(out_dir / "审核原因占比.png")
plt.close()

# -----------------------------
# 7. 7天无理由机型分布
# -----------------------------
if len(df_7d) > 0:
    model_7d_dist = (
        df_7d["机型名称"]
        .value_counts()
        .rename_axis("机型名称")
        .reset_index(name="数量")
        .assign(占比=lambda x: (x["数量"] / x["数量"].sum() * 100).round(1))
    )
    model_7d_dist.to_excel(out_dir / "7天无理由_机型分布.xlsx", index=False)

    plt.figure(figsize=(8, 8))
    plt.pie(model_7d_dist["数量"], labels=model_7d_dist["机型名称"], autopct="%1.1f%%")
    plt.title("7天无理由 - 机型分布")
    plt.tight_layout()
    plt.savefig(out_dir / "7天无理由_机型分布.png")
    plt.close()
else:
    print("警告：7天无理由数据为空")

# -----------------------------
# 8. 非7天无理由机型分布
# -----------------------------
if len(df_non_7d) > 0:
    model_non_7d_dist = (
        df_non_7d["机型名称"]
        .value_counts()
        .rename_axis("机型名称")
        .reset_index(name="数量")
        .assign(占比=lambda x: (x["数量"] / x["数量"].sum() * 100).round(1))
    )
    model_non_7d_dist.to_excel(out_dir / "非7天无理由_机型分布.xlsx", index=False)

    plt.figure(figsize=(8, 8))
    plt.pie(model_non_7d_dist["数量"], labels=model_non_7d_dist["机型名称"], autopct="%1.1f%%")
    plt.title("非7天无理由 - 机型分布")
    plt.tight_layout()
    plt.savefig(out_dir / "非7天无理由_机型分布.png")
    plt.close()
else:
    print("警告：非7天无理由数据为空")

# -----------------------------
# 9. 按机型统计分类描述词频次
# -----------------------------
def build_model_issue_table(df_sub, suffix, detailed_dir):
    """为每个机型计算'分类'描述词频次，输出excel和柱状图，并生成详细数据文件"""
    if len(df_sub) == 0:
        print(f"警告：{suffix}数据为空，跳过机型分析")
        return
        
    # 非7天无理由数据：过滤掉问题描述为空的行
    if suffix == "非7天无理由":
        # 检查问题描述列是否存在
        if "问题描述" in df_sub.columns:
            # 过滤掉问题描述为空的行
            df_sub = df_sub[df_sub["问题描述"].notna() & (df_sub["问题描述"] != "")]
            print(f"已过滤空问题描述行，剩余 {len(df_sub)} 条记录")
        else:
            print("警告：未找到'问题描述'列，无法过滤空值")
    
    for model in df_sub["机型名称"].unique():
        # 清理机型名称用于文件夹和文件名
        clean_model = sanitize_filename(str(model))
        
        # 创建机型文件夹
        model_dir = detailed_dir / clean_model
        model_dir.mkdir(parents=True, exist_ok=True)
        
        # 获取该机型的所有数据
        model_data = df_sub[df_sub["机型名称"] == model].copy()
        
        # 统计分类频次
        sub = (
            model_data["分类"]
            .value_counts()
            .rename_axis("分类")
            .reset_index(name="次数")
        )
        
        # 保存频次统计
        freq_filename = f"{clean_model}_{suffix}_分类频次.xlsx"
        freq_path = model_dir / freq_filename
        sub.to_excel(freq_path, index=False)

        # 为每个机型的所有分类生成一个综合详细数据文件
        detailed_filename = f"{clean_model}_{suffix}_详细数据.xlsx"
        detailed_path = model_dir / detailed_filename
        model_data.to_excel(detailed_path, index=False)
        
        # 生成柱状图
        plt.figure(figsize=(12, 6))
        bars = plt.bar(sub["分类"], sub["次数"])
        plt.xticks(rotation=45, ha="right")
        plt.title(f"{model} - {suffix} - 分类频次")
        
        # 添加数量标签
        for bar in bars:
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height)}', ha='center', va='bottom')
        
        plt.tight_layout()
        
        chart_filename = f"{clean_model}_{suffix}_柱状图.png"
        plt.savefig(model_dir / chart_filename)
        plt.close()
        
        print(f"已生成 {model} 的 {suffix} 数据，共 {len(sub)} 个分类，{len(model_data)} 条记录")


# -----------------------------
# 修复 generate_analysis_report 函数中的列名问题
# -----------------------------
def generate_analysis_report(df, df_7d, df_non_7d, out_dir, start_date, end_date):
    """生成分析报告并保存到文本文件"""
    report_lines = []

    # 1. 落入D等级产品数据统计
    d_grade_data = df[df['审核原因'] == 'D等级']
    d_grade_count = len(d_grade_data)
    d_grade_models = d_grade_data['机型名称'].unique()
    d_grade_model_count = len(d_grade_models)
    report_lines.append(f"落入D等级产品数据：{d_grade_count} 条")
    report_lines.append(f"覆盖周期：{start_date} 至 {end_date}")
    report_lines.append(f"涉及机型：{', '.join(d_grade_models)}")
    report_lines.append(f"共计机型数量：{d_grade_model_count} 款\n")

    # 2. 审核原因占比
    reasons = ["7天无理由", "15天质量换新", "质量维修", "180天只换不修"]
    total_count = len(df)
    for reason in reasons:
        count = (df['审核原因'] == reason).sum()
        percentage = (count / total_count * 100) if total_count > 0 else 0
        report_lines.append(f"审核原因 - {reason}：{count} 条，占比 {percentage:.2f}%")
    report_lines.append("")

    # 3. 七天无理由机型占比
    if len(df_7d) > 0:
        model_7d_dist = (
            df_7d['机型名称']
            .value_counts()
            .rename_axis('机型名称')
            .reset_index(name='数量')
            .assign(占比=lambda x: (x['数量'] / x['数量'].sum() * 100).round(2))
        )
        report_lines.append("七天无理由机型占比：")
        for _, row in model_7d_dist.iterrows():
            report_lines.append(f"  {row['机型名称']}: {row['数量']} 条，占比 {row['占比']}%")
        report_lines.append("")

    # 4. 非七天无理由机型占比
    if len(df_non_7d) > 0:
        model_non_7d_dist = (
            df_non_7d['机型名称']
            .value_counts()
            .rename_axis('机型名称')
            .reset_index(name='数量')
            .assign(占比=lambda x: (x['数量'] / x['数量'].sum() * 100).round(2))
        )
        report_lines.append("非七天无理由机型占比：")
        for _, row in model_non_7d_dist.iterrows():
            report_lines.append(f"  {row['机型名称']}: {row['数量']} 条，占比 {row['占比']}%")
        report_lines.append("")

    # 5. 每个机型七天无理由的分类数据分析
    report_lines.append("每个机型七天无理由的分类数据分析：")
    for model in df_7d['机型名称'].unique():
        model_data = df_7d[df_7d['机型名称'] == model]
        total_comments = len(model_data)
        no_reason_count = (model_data['分类'] == '无理由退货').sum()
        no_reason_percentage = (no_reason_count / total_comments * 100) if total_comments > 0 else 0
        top_issues = (
            model_data['分类']
            .value_counts()
            .reset_index(name='次数')
            .rename(columns={'index': '分类'})  # 确保列名正确
        )

        # 调试信息：打印 top_issues 列名和数据
        print("七天无理由 - Top Issues:")
        print(top_issues.head())

        top_issues = top_issues[top_issues['次数'] >= 2].head(2)
        report_lines.append(f"  {model}:")
        report_lines.append(f"    评论总数：{total_comments}")
        report_lines.append(f"    无理由退货：{no_reason_count} 条，占比 {no_reason_percentage:.2f}%")
        for _, row in top_issues.iterrows():
            issue_percentage = (row['次数'] / total_comments * 100) if total_comments > 0 else 0
            report_lines.append(f"    Top问题：{row['分类']}，次数：{row['次数']}，占比：{issue_percentage:.2f}%")
    report_lines.append("")

    # 6. 每个机型非七天无理由的分类数据分析
    report_lines.append("每个机型非七天无理由的分类数据分析：")
    for model in df_non_7d['机型名称'].unique():
        model_data = df_non_7d[df_non_7d['机型名称'] == model]
        total_comments = len(model_data)
        top_issues = (
            model_data['分类']
            .value_counts()
            .reset_index(name='次数')
            .rename(columns={'index': '分类'})  # 确保列名正确
        )

        # 调试信息：打印 top_issues 列名和数据
        print("非七天无理由 - Top Issues:")
        print(top_issues.head())

        top_issues = top_issues[top_issues['次数'] >= 2].head(2)
        report_lines.append(f"  {model}:")
        report_lines.append(f"    有效评论总数：{total_comments}")
        for _, row in top_issues.iterrows():
            issue_percentage = (row['次数'] / total_comments * 100) if total_comments > 0 else 0
            report_lines.append(f"    Top问题：{row['分类']}，次数：{row['次数']}，占比：{issue_percentage:.2f}%")
    report_lines.append("")

    # 7. 总结
    report_lines.append("总结：")
    report_lines.append(f"本次报告时间覆盖：{start_date} 至 {end_date}")
    report_lines.append(f"覆盖机型：{', '.join(df['机型名称'].unique())}")
    report_lines.append("非七天无理由分类中，以下机型的问题较为突出：")
    for model in df_non_7d['机型名称'].unique():
        model_data = df_non_7d[df_non_7d['机型名称'] == model]
        top_issues = (
            model_data['分类']
            .value_counts()
            .reset_index(name='次数')
            .rename(columns={'index': '分类'})  # 确保列名正确
        )
        top_issues = top_issues[top_issues['次数'] >= 2].head(2)
        for _, row in top_issues.iterrows():
            report_lines.append(f"  {model} - {row['分类']}：{row['次数']} 次")

    # 保存报告到文件
    report_path = out_dir / "分析报告.txt"
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))

    print(f"分析报告已生成：{report_path}")

# 在主流程中调用生成分析报告的函数
generate_analysis_report(df, df_7d, df_non_7d, out_dir, start_date, end_date)


# 生成7天无理由数据
build_model_issue_table(df_7d, "7天无理由", detailed_dir_7d)

# 生成非7天无理由数据
build_model_issue_table(df_non_7d, "非7天无理由", detailed_dir_non7d)

print("✅ 所有处理完成，结果已保存到 output 目录！")
print("文件结构：")
print("output/")
print("├── 审核原因统计.xlsx")
print("├── 7天无理由_机型分布.xlsx")
print("├── 非7天无理由_机型分布.xlsx")
print("├── 审核原因占比.png")
print("├── 7天无理由_机型分布.png")
print("├── 非7天无理由_机型分布.png")
print("└── 详细数据/")
print("    ├── 7天无理由/")
print("    │   └── [机型名称]/")
print("    │       ├── [机型]_7天无理由_分类频次.xlsx")
print("    │       ├── [机型]_7天无理由_柱状图.png")
print("    │       └── [机型]_7天无理由_详细数据.xlsx")
print("    └── 非7天无理由/")
print("        └── [机型名称]/")
print("            ├── [机型]_非7天无理由_分类频次.xlsx")
print("            ├── [机型]_非7天无理由_柱状图.png")
print("            └── [机型]_非7天无理由_详细数据.xlsx")
