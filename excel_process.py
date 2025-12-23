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
