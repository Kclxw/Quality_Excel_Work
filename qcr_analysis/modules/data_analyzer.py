# -*- coding: utf-8 -*-
"""
=============================================================================
数据分析模块
=============================================================================
负责数据统计分析和图表生成
=============================================================================
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from datetime import date, datetime, timedelta
import re

import sys
sys.path.append(str(Path(__file__).parent.parent))
from config import (
    AUDIT_REASONS,
    MATPLOTLIB_FONTS,
    CHART_STYLE
)

# 设置中文字体
matplotlib.rcParams['font.family'] = MATPLOTLIB_FONTS
matplotlib.rcParams['axes.unicode_minus'] = False

# 处理字体警告
import warnings
warnings.filterwarnings("ignore", category=UserWarning, message=".*Glyph.*missing.*")


def sanitize_filename(filename: str) -> str:
    """
    清理文件名中的非法字符
    
    Args:
        filename: 原始文件名
        
    Returns:
        清理后的文件名
    """
    # Windows非法字符：<>:"/\|?*
    illegal_chars = r'[<>:\"/\\|?*]'
    filename = re.sub(illegal_chars, ' ', filename)
    filename = filename.strip()
    if len(filename) > 200:
        filename = filename[:200]
    return filename


def get_week_workday_range(reference_date: Optional[date] = None) -> Tuple[str, str]:
    """
    获取本周工作日范围（周一到周五）
    
    Args:
        reference_date: 参考日期，默认为今天
        
    Returns:
        (周一日期, 周五日期) 元组
    """
    today = reference_date if reference_date else datetime.today().date()
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)
    return monday.strftime("%Y/%m/%d"), friday.strftime("%Y/%m/%d")


def determine_coverage_range(df: pd.DataFrame, date_column: str,
                            start_date: Optional[date],
                            end_date: Optional[date]) -> Tuple[str, str]:
    """
    确定数据覆盖的日期范围
    
    Args:
        df: 数据DataFrame
        date_column: 日期列名
        start_date: 开始日期
        end_date: 结束日期
        
    Returns:
        (开始日期字符串, 结束日期字符串) 元组
    """
    if df.empty:
        return ("-", "-")
    
    actual_start = start_date if start_date else df[date_column].min()
    actual_end = end_date if end_date else df[date_column].max()
    
    # 将 pandas Timestamp 转换为 date 对象
    if hasattr(actual_start, 'date') and callable(actual_start.date):
        actual_start = actual_start.date()
    if hasattr(actual_end, 'date') and callable(actual_end.date):
        actual_end = actual_end.date()
    
    return actual_start.strftime("%Y/%m/%d"), actual_end.strftime("%Y/%m/%d")


class DataAnalyzer:
    """数据分析器"""
    
    def __init__(self, output_dir: Path):
        """
        初始化数据分析器
        
        Args:
            output_dir: 输出目录路径
        """
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # 创建详细数据目录
        self.detailed_dir_7d = output_dir / "详细数据" / "7天无理由"
        self.detailed_dir_non7d = output_dir / "详细数据" / "非7天无理由"
        self.detailed_dir_7d.mkdir(parents=True, exist_ok=True)
        self.detailed_dir_non7d.mkdir(parents=True, exist_ok=True)
    
    def analyze_audit_reasons(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, Path]:
        """
        统计审核原因
        
        Args:
            df: 数据DataFrame
            
        Returns:
            (统计结果DataFrame, 图表路径)
        """
        counts = {r: int((df["审核原因"] == r).sum()) for r in AUDIT_REASONS}
        
        summary_df = pd.DataFrame(list(counts.items()), columns=["审核原因", "数量"])
        total_count = summary_df["数量"].sum()
        summary_df["占比"] = (summary_df["数量"] / total_count * 100).round(2)
        
        # 保存Excel
        summary_df.to_excel(self.output_dir / "审核原因统计.xlsx", index=False)
        
        # 生成饼图
        plt.figure(figsize=CHART_STYLE['reason_chart_size'])
        plt.pie(summary_df["数量"], labels=summary_df["审核原因"], autopct="%1.1f%%")
        plt.title("审核原因占比")
        plt.tight_layout()
        chart_path = self.output_dir / "审核原因占比.png"
        plt.savefig(chart_path)
        plt.close()
        
        print(f"✓ 审核原因统计完成，共 {total_count} 条记录")
        
        return summary_df, chart_path
    
    def analyze_model_distribution(self, df: pd.DataFrame, suffix: str) -> Tuple[pd.DataFrame, Optional[Path]]:
        """
        统计机型分布
        
        Args:
            df: 数据DataFrame
            suffix: 分类后缀（7天无理由 或 非7天无理由）
            
        Returns:
            (统计结果DataFrame, 图表路径)
        """
        if len(df) == 0:
            print(f"警告：{suffix}数据为空")
            return pd.DataFrame(), None
        
        model_dist = (
            df["机型名称"]
            .value_counts()
            .rename_axis("机型名称")
            .reset_index(name="数量")
            .assign(占比=lambda x: (x["数量"] / x["数量"].sum() * 100).round(1))
        )
        
        # 保存Excel
        model_dist.to_excel(self.output_dir / f"{suffix}_机型分布.xlsx", index=False)
        
        # 生成饼图
        plt.figure(figsize=CHART_STYLE['pie_chart_size'])
        plt.pie(model_dist["数量"], labels=model_dist["机型名称"], autopct="%1.1f%%")
        plt.title(f"{suffix} - 机型分布")
        plt.tight_layout()
        chart_path = self.output_dir / f"{suffix}_机型分布.png"
        plt.savefig(chart_path)
        plt.close()
        
        print(f"✓ {suffix}机型分布统计完成，共 {len(df)} 条记录，{len(model_dist)} 个机型")
        
        return model_dist, chart_path
    
    def analyze_model_issues(self, df: pd.DataFrame, suffix: str) -> List[Dict]:
        """
        按机型分析问题分类
        
        Args:
            df: 数据DataFrame
            suffix: 分类后缀（7天无理由 或 非7天无理由）
            
        Returns:
            机型分析结果列表
        """
        if len(df) == 0:
            print(f"警告：{suffix}数据为空，跳过机型分析")
            return []
        
        # 选择详细数据目录
        detailed_dir = self.detailed_dir_7d if suffix == "7天无理由" else self.detailed_dir_non7d
        
        # 非7天无理由数据：过滤掉问题描述为空的行
        if suffix == "非7天无理由" and "问题描述" in df.columns:
            original_len = len(df)
            df = df[df["问题描述"].notna() & (df["问题描述"] != "")]
            print(f"已过滤空问题描述行，从 {original_len} 条减少到 {len(df)} 条记录")
        
        summaries = []
        
        for model in df["机型名称"].unique():
            # 清理机型名称
            clean_model = sanitize_filename(str(model))
            
            # 创建机型文件夹
            model_dir = detailed_dir / clean_model
            model_dir.mkdir(parents=True, exist_ok=True)
            
            # 获取该机型的所有数据
            model_data = df[df["机型名称"] == model].copy()
            
            # 统计分类频次
            category_stats = (
                model_data["分类"]
                .value_counts()
                .rename_axis("分类")
                .reset_index(name="次数")
            )
            
            if "次数" in category_stats.columns and category_stats["次数"].sum() > 0:
                category_stats["占比"] = (category_stats["次数"] / category_stats["次数"].sum() * 100).round(1)
            else:
                category_stats["占比"] = 0
            
            # 保存频次统计
            freq_filename = f"{clean_model}_{suffix}_分类频次.xlsx"
            freq_path = model_dir / freq_filename
            category_stats.to_excel(freq_path, index=False)
            
            # 保存详细数据
            detailed_filename = f"{clean_model}_{suffix}_详细数据.xlsx"
            detailed_path = model_dir / detailed_filename
            model_data.to_excel(detailed_path, index=False)
            
            # 生成柱状图
            plt.figure(figsize=CHART_STYLE['bar_chart_size'])
            bars = plt.bar(category_stats["分类"], category_stats["次数"])
            plt.xticks(rotation=45, ha="right")
            plt.title(f"{model} - {suffix} - 分类频次")
            
            # 添加数量标签
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height,
                        f'{int(height)}', ha='center', va='bottom')
            
            plt.tight_layout()
            
            chart_filename = f"{clean_model}_{suffix}_柱状图.png"
            chart_path = model_dir / chart_filename
            plt.savefig(chart_path)
            plt.close()
            
            print(f"  - {model}: {len(category_stats)} 个分类，{len(model_data)} 条记录")
            
            # 保存摘要信息
            model_summary = {
                "model": model,
                "clean_model": clean_model,
                "suffix": suffix,
                "category_df": category_stats,
                "chart_path": str(chart_path),
                "total_records": len(model_data)
            }
            summaries.append(model_summary)
        
        print(f"✓ {suffix}机型问题分析完成，共 {len(summaries)} 个机型")
        
        return summaries
    
    def generate_text_report(self, df: pd.DataFrame, df_7d: pd.DataFrame,
                           df_non_7d: pd.DataFrame, start_date: Optional[date],
                           end_date: Optional[date]):
        """
        生成文本分析报告
        
        Args:
            df: 完整数据DataFrame
            df_7d: 7天无理由数据
            df_non_7d: 非7天无理由数据
            start_date: 开始日期
            end_date: 结束日期
        """
        report_lines = []
        
        # 1. 基本统计
        report_lines.append("="*60)
        report_lines.append("QCR 数据分析报告")
        report_lines.append("="*60)
        report_lines.append(f"分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(f"数据范围: {start_date or '最早'} 至 {end_date or '最新'}")
        report_lines.append(f"数据总量: {len(df)} 条记录")
        report_lines.append("")
        
        # 2. 机型统计
        unique_models = df['机型名称'].unique()
        report_lines.append(f"涉及机型数: {len(unique_models)} 款")
        report_lines.append(f"机型列表: {', '.join(unique_models[:10])}")
        if len(unique_models) > 10:
            report_lines.append(f"          ... 等共 {len(unique_models)} 款")
        report_lines.append("")
        
        # 3. 审核原因统计
        report_lines.append("审核原因统计:")
        for reason in AUDIT_REASONS:
            count = (df['审核原因'] == reason).sum()
            percentage = (count / len(df) * 100) if len(df) > 0 else 0
            report_lines.append(f"  {reason}: {count} 条 ({percentage:.2f}%)")
        report_lines.append("")
        
        # 4. 7天无理由分析
        if len(df_7d) > 0:
            report_lines.append("七天无理由机型TOP5:")
            model_7d_dist = df_7d['机型名称'].value_counts().head(5)
            for model, count in model_7d_dist.items():
                percentage = (count / len(df_7d) * 100)
                report_lines.append(f"  {model}: {count} 条 ({percentage:.1f}%)")
            report_lines.append("")
        
        # 5. 非7天无理由分析
        if len(df_non_7d) > 0:
            report_lines.append("非七天无理由机型TOP5:")
            model_non_7d_dist = df_non_7d['机型名称'].value_counts().head(5)
            for model, count in model_non_7d_dist.items():
                percentage = (count / len(df_non_7d) * 100)
                report_lines.append(f"  {model}: {count} 条 ({percentage:.1f}%)")
            report_lines.append("")
        
        report_lines.append("="*60)
        report_lines.append("报告结束")
        report_lines.append("="*60)
        
        # 保存报告
        report_path = self.output_dir / "分析报告.txt"
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(report_lines))
        
        print(f"✓ 文本分析报告已生成：{report_path}")
    
    def analyze_top_issues(self, df: pd.DataFrame, top_n: int = 10) -> Dict:
        """
        分析Top N Issue及其机型分布
        
        Args:
            df: 数据DataFrame
            top_n: Top N数量
            
        Returns:
            分析结果字典
        """
        if len(df) == 0 or '分类' not in df.columns:
            print("警告：数据为空或缺少'分类'列")
            return {}
        
        print(f"\n📊 开始Top {top_n} Issue分析...")
        
        # 创建Top Issue分析目录
        top_issue_dir = self.output_dir / "Top_Issue分析"
        top_issue_dir.mkdir(parents=True, exist_ok=True)
        charts_dir = top_issue_dir / "charts"
        charts_dir.mkdir(parents=True, exist_ok=True)
        
        # 1. 统计Top N Issue
        issue_counts = df['分类'].value_counts().head(top_n)
        
        # 创建统计表
        issue_stats = pd.DataFrame({
            '排名': range(1, len(issue_counts) + 1),
            'Issue名称': issue_counts.index,
            '数量': issue_counts.values,
            '占比(%)': (issue_counts.values / len(df) * 100).round(2)
        })
        issue_stats['累计占比(%)'] = issue_stats['占比(%)'].cumsum().round(2)
        
        print(f"✓ 统计了Top {len(issue_stats)} Issue")
        
        # 2. 生成Top Issue总览图
        plt.figure(figsize=(14, 7))
        bars = plt.bar(range(len(issue_stats)), issue_stats['数量'],
                      color=plt.cm.Blues(range(len(issue_stats), 0, -1)))
        
        plt.xlabel("Issue分类", fontsize=12)
        plt.ylabel("数量", fontsize=12)
        plt.title(f"Top {top_n} Issue分布", fontsize=16, fontweight='bold')
        plt.xticks(range(len(issue_stats)), issue_stats['Issue名称'], rotation=45, ha='right')
        
        # 添加数量标签
        for i, bar in enumerate(bars):
            height = bar.get_height()
            percentage = issue_stats.iloc[i]['占比(%)']
            plt.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}\n({percentage:.1f}%)',
                   ha='center', va='bottom', fontsize=10)
        
        plt.tight_layout()
        summary_chart_path = charts_dir / "Top_Issue总览图.png"
        plt.savefig(summary_chart_path, dpi=150)
        plt.close()
        print(f"✓ 生成Top Issue总览图")
        
        # 3. 分析每个Issue的机型分布
        issue_details = []
        
        print(f"\n📈 分析每个Issue的机型分布...")
        for idx, row in issue_stats.iterrows():
            issue_name = row['Issue名称']
            issue_count = row['数量']
            
            # 筛选该Issue的数据
            issue_df = df[df['分类'] == issue_name]
            
            # 统计机型分布
            model_dist = (
                issue_df['机型名称']
                .value_counts()
                .rename_axis('机型名称')
                .reset_index(name='数量')
            )
            model_dist['占比(%)'] = (model_dist['数量'] / issue_count * 100).round(2)
            
            # 生成机型分布图
            if len(model_dist) >= 2:
                fig, ax = plt.subplots(figsize=(12, 6))
                
                # 只显示前15个机型
                display_data = model_dist.head(15)
                bars = ax.barh(range(len(display_data)), display_data['数量'])
                
                ax.set_yticks(range(len(display_data)))
                ax.set_yticklabels(display_data['机型名称'], fontsize=10)
                ax.set_xlabel("数量", fontsize=12)
                ax.set_title(f"Issue: {issue_name}\n机型分布 (共{len(model_dist)}款机型)", 
                           fontsize=12, fontweight='bold')
                
                # 添加数量标签
                for i, bar in enumerate(bars):
                    width = bar.get_width()
                    percentage = display_data.iloc[i]['占比(%)']
                    ax.text(width, bar.get_y() + bar.get_height()/2.,
                           f' {int(width)} ({percentage:.1f}%)',
                           ha='left', va='center', fontsize=9)
                
                plt.tight_layout()
                chart_filename = f"Issue{idx+1}_{sanitize_filename(issue_name)}_机型分布.png"
                chart_path = charts_dir / chart_filename
                plt.savefig(chart_path, dpi=150)
                plt.close()
            else:
                chart_path = None
            
            # 保存Issue详情
            issue_details.append({
                'rank': idx + 1,
                'issue_name': issue_name,
                'count': issue_count,
                'percentage': row['占比(%)'],
                'cumulative_percentage': row['累计占比(%)'],
                'model_count': len(model_dist),
                'model_dist': model_dist,
                'chart_path': str(chart_path) if chart_path else None
            })
            
            print(f"  - Issue #{idx+1}: {issue_name} ({issue_count}条) -> {len(model_dist)}款机型")
        
        # 4. 保存Excel（多sheet）
        excel_path = top_issue_dir / "Top_Issue统计汇总.xlsx"
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # Sheet1: Top N列表
            issue_stats.to_excel(writer, sheet_name='Top_Issue列表', index=False)
            
            # Sheet2~N+1: 各Issue机型分布
            for detail in issue_details:
                sheet_name = sanitize_filename(detail['issue_name'])[:31]
                detail['model_dist'].to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"✓ Excel汇总已保存：{excel_path.name}")
        
        # 5. 生成文本报告
        self._generate_top_issue_text_report(df, issue_stats, issue_details, top_issue_dir, top_n)
        
        # 返回结果
        result = {
            'top_n': top_n,
            'total_records': len(df),
            'issue_stats': issue_stats,
            'issue_details': issue_details,
            'excel_path': str(excel_path),
            'summary_chart_path': str(summary_chart_path),
            'charts_dir': str(charts_dir)
        }
        
        print(f"\n✅ Top Issue分析完成！")
        print(f"   总记录数: {len(df)}")
        print(f"   Top {top_n} 累计占比: {issue_stats['累计占比(%)'].iloc[-1]:.2f}%")
        print(f"   涉及机型数: {df['机型名称'].nunique()} 款")
        
        return result
    
    def _generate_top_issue_text_report(self, df: pd.DataFrame, issue_stats: pd.DataFrame,
                                        issue_details: List[Dict], output_dir: Path, top_n: int):
        """
        生成Top Issue文本分析报告
        
        Args:
            df: 原始数据
            issue_stats: Issue统计表
            issue_details: Issue详情列表
            output_dir: 输出目录
            top_n: Top N
        """
        report_lines = []
        
        report_lines.append("="*70)
        report_lines.append(f"Top {top_n} Issue 分析报告")
        report_lines.append("="*70)
        report_lines.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_lines.append(f"数据总量: {len(df)} 条记录")
        report_lines.append(f"涉及机型: {df['机型名称'].nunique()} 款")
        report_lines.append("")
        
        # Top Issue概览
        report_lines.append(f"一、Top {top_n} Issue分布概览")
        report_lines.append("-"*70)
        for idx, row in issue_stats.iterrows():
            report_lines.append(
                f"  {row['排名']}. {row['Issue名称']}: "
                f"{row['数量']}条 ({row['占比(%)']:.2f}%) "
                f"[累计: {row['累计占比(%)']:.2f}%]"
            )
        report_lines.append("")
        
        # 每个Issue的机型分布详情
        report_lines.append(f"二、Top Issue 机型分布详情")
        report_lines.append("-"*70)
        for detail in issue_details:
            report_lines.append(f"\n【Issue #{detail['rank']}】{detail['issue_name']}")
            report_lines.append(f"  总计: {detail['count']} 条记录 ({detail['percentage']:.2f}%)")
            report_lines.append(f"  涉及机型: {detail['model_count']} 款")
            report_lines.append(f"  主要机型分布（Top 5）:")
            
            model_dist = detail['model_dist'].head(5)
            for m_idx, m_row in model_dist.iterrows():
                report_lines.append(
                    f"    {m_idx+1}. {m_row['机型名称']}: "
                    f"{m_row['数量']}条 ({m_row['占比(%)']:.2f}%)"
                )
        
        report_lines.append("")
        report_lines.append("="*70)
        report_lines.append("报告结束")
        report_lines.append("="*70)
        
        # 保存报告
        report_path = output_dir / f"Top{top_n}_Issue分析报告.txt"
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(report_lines))
        
        print(f"✓ 文本报告已生成：{report_path.name}")

