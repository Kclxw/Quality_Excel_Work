# -*- coding: utf-8 -*-
"""
=============================================================================
Top Model Analysis Service - 热门机型分析服务
=============================================================================
基于分类数量分析Top N机型
核心指标：分类数 = df.groupby('机型名称')['分类'].nunique()
=============================================================================
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from pathlib import Path
from typing import Dict, List, Optional

import sys
sys.path.append(str(Path(__file__).parent.parent))

from config import MATPLOTLIB_FONTS
from modules.llm_service import LLMService
from prompts import TOP_MODEL_OVERVIEW_PROMPT

# 设置中文字体
matplotlib.rcParams['font.family'] = MATPLOTLIB_FONTS
matplotlib.rcParams['axes.unicode_minus'] = False


class TopModelAnalysisService:
    """Top Model分析服务 - 基于分类数量"""
    
    def __init__(self, output_dir: str or Path):
        """初始化Top Model分析服务"""
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.top_model_dir = self.output_dir / "Top_Model分析"
        self.top_model_dir.mkdir(parents=True, exist_ok=True)
        
        self.charts_dir = self.top_model_dir / "charts"
        self.charts_dir.mkdir(parents=True, exist_ok=True)
        
        self.results = {}
    
    def analyze(
        self,
        df: pd.DataFrame,
        top_n: int = 15,
        use_llm: bool = False,
        llm_config: Optional[Dict] = None
    ) -> Dict:
        """执行Top Model完整分析流程"""
        print("\n" + "="*70)
        print(f"🏆 Top {top_n} Model 分析（基于分类数量）")
        print("="*70)
        
        if len(df) == 0 or '机型名称' not in df.columns or '分类' not in df.columns:
            print("❌ 错误：数据为空或缺少必需列（需要'机型名称'和'分类'）")
            return {}
        
        # 1. 计算分类数（使用"分类"列）
        print(f"\n📊 统计所有机型的分类数...")
        model_stats = df.groupby('机型名称').agg({
            '分类': 'nunique',
            '机型名称': 'count'
        }).rename(columns={'分类': '分类数', '机型名称': '记录数'})
        
        model_stats = model_stats.reset_index()
        model_stats['平均每类记录数'] = (model_stats['记录数'] / model_stats['分类数']).round(1)
        model_stats = model_stats.sort_values('分类数', ascending=False)
        model_stats['排名'] = range(1, len(model_stats) + 1)
        model_stats = model_stats[['排名', '机型名称', '分类数', '记录数', '平均每类记录数']]
        
        print(f"✓ 共统计 {len(model_stats)} 个机型")
        
        # 2. 提取Top N
        top_models = model_stats.head(top_n)
        print(f"\n✓ Top {top_n} 机型:")
        for idx, row in top_models.iterrows():
            print(f"   {row['排名']}. {row['机型名称']}: {row['分类数']}个分类, {row['记录数']}条记录")

        # 保存 Top N 统计表
        top_stats_path = self.top_model_dir / f"Top{top_n}_Model统计.xlsx"
        top_models.to_excel(top_stats_path, index=False)
        
        # 3. 生成图表
        overall_chart = self._generate_overall_chart(model_stats)
        comparison_chart = self._generate_comparison_chart(top_models, top_n)
        
        # 4. 详细分析
        model_details = self._analyze_top_models(df, top_models)
        
        # 5. 生成报告
        report_path = self._generate_report(model_stats, top_models, model_details, top_n)
        
        self.results = {
            "model_stats": model_stats,
            "top_models": top_models,
            "model_details": model_details,
            "overall_chart": overall_chart,
            "comparison_chart": comparison_chart,
            "report_path": report_path,
            "total_records": len(df),
            "total_models": len(model_stats),
            "top_n": top_n
        }
        
        print("\n✅ Top Model分析完成")
        return self.results
    
    def _generate_overall_chart(self, model_stats):
        """生成整体分布图（带数据标签）"""
        display_data = model_stats.head(30)
        fig, ax = plt.subplots(figsize=(14, 10))
        bars = ax.barh(range(len(display_data)), display_data['分类数'])
        ax.set_yticks(range(len(display_data)))
        ax.set_yticklabels(display_data['机型名称'], fontsize=10)
        ax.set_xlabel("分类数", fontsize=13)
        ax.set_title("机型分类复杂度分布 (Top 30)", fontsize=14, fontweight='bold')
        
        # 添加数值标签
        for i, bar in enumerate(bars):
            width = bar.get_width()
            ax.text(
                width + max(display_data['分类数']) * 0.01,
                bar.get_y() + bar.get_height()/2.,
                f'{int(width)}',
                ha='left', va='center', fontsize=9, fontweight='bold'
            )
        
        plt.tight_layout()
        chart_path = self.charts_dir / "整体机型问题复杂度分布.png"
        plt.savefig(chart_path, dpi=150, bbox_inches='tight')
        plt.close()
        return chart_path
    
    def _generate_comparison_chart(self, top_models, top_n):
        """生成对比图（带数据标签）"""
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 10))
        
        # 分类数对比
        bars1 = ax1.barh(range(len(top_models)), top_models['分类数'])
        ax1.set_yticks(range(len(top_models)))
        ax1.set_yticklabels(top_models['机型名称'], fontsize=10)
        ax1.set_xlabel("分类数", fontsize=12)
        ax1.set_title(f"Top {top_n} 机型分类数对比", fontsize=13, fontweight='bold')
        for i, bar in enumerate(bars1):
            width = bar.get_width()
            ax1.text(width + max(top_models['分类数']) * 0.01, bar.get_y() + bar.get_height()/2.,
                    f'{int(width)}', ha='left', va='center', fontsize=9, fontweight='bold')
        
        # 记录数对比
        bars2 = ax2.barh(range(len(top_models)), top_models['记录数'])
        ax2.set_yticks(range(len(top_models)))
        ax2.set_yticklabels(top_models['机型名称'], fontsize=10)
        ax2.set_xlabel("记录数", fontsize=12)
        ax2.set_title(f"Top {top_n} 机型记录数对比", fontsize=13, fontweight='bold')
        for i, bar in enumerate(bars2):
            width = bar.get_width()
            ax2.text(width + max(top_models['记录数']) * 0.01, bar.get_y() + bar.get_height()/2.,
                    f'{int(width)}', ha='left', va='center', fontsize=9, fontweight='bold')
        
        plt.tight_layout()
        chart_path = self.charts_dir / f"Top{top_n}_机型对比图.png"
        plt.savefig(chart_path, dpi=150, bbox_inches='tight')
        plt.close()
        return chart_path
    
    def _analyze_top_models(self, df, top_models):
        """分析每个Top机型的详细情况"""
        model_details = []
        
        print(f"\n📊 分析每个Top机型的问题分布...")
        for idx, row in top_models.iterrows():
            model_name = row['机型名称']
            category_count = row['分类数']
            total_records = row['记录数']
            
            # 筛选该机型的数据
            model_df = df[df['机型名称'] == model_name]
            
            # 统计问题分类分布（使用"分类"列）
            category_dist = model_df['分类'].value_counts().reset_index()
            category_dist.columns = ['分类', '数量']
            category_dist['占比(%)'] = (category_dist['数量'] / total_records * 100).round(2)
            
            # 统计7天 vs 质量问题
            return_7day_count = (model_df['审核原因'] == '7天无理由').sum()
            quality_count = model_df['审核原因'].isin([
                '15天质量换新', '180天只换不修', '质量维修'
            ]).sum()
            
            return_7day_pct = (return_7day_count / total_records * 100).round(1) if total_records > 0 else 0
            quality_pct = (quality_count / total_records * 100).round(1) if total_records > 0 else 0
            
            # 保存详细数据
            safe_name = self._safe_filename(model_name)
            detail_path = self.top_model_dir / f"{idx+1:02d}_{safe_name}_详细数据.xlsx"
            
            with pd.ExcelWriter(detail_path, engine='openpyxl') as writer:
                category_dist.to_excel(writer, sheet_name='分类分布', index=False)
            
            # 生成单个机型的图表
            chart_path = self._generate_model_detail_chart(model_name, category_dist, idx+1)
            
            model_details.append({
                'rank': idx + 1,
                'model_name': model_name,
                'category_count': category_count,
                'total_records': total_records,
                'avg_per_category': row['平均每类记录数'],
                'category_distribution': category_dist,
                'return_7day_count': return_7day_count,
                'return_7day_pct': return_7day_pct,
                'quality_count': quality_count,
                'quality_pct': quality_pct,
                'detail_path': detail_path,
                'chart_path': chart_path
            })
            
            print(f"  - 机型 #{idx+1}: {model_name} ({category_count}个分类, {total_records}条记录)")
        
        print(f"✓ 完成 {len(model_details)} 个机型的详细分析")
        return model_details
    
    def _generate_model_detail_chart(self, model_name, category_dist, rank):
        """生成单个机型的详细图表（带数据标签）"""
        display_data = category_dist.head(20)
        
        fig, ax = plt.subplots(figsize=(14, 10))
        bars = ax.barh(range(len(display_data)), display_data['数量'])
        
        ax.set_yticks(range(len(display_data)))
        ax.set_yticklabels(display_data['分类'], fontsize=10)
        ax.set_xlabel("数量", fontsize=13)
        ax.set_title(
            f"机型: {model_name}\n分类分布 (Top 20, 共{len(category_dist)}类)",
            fontsize=14, fontweight='bold'
        )
        
        # 添加数值标签（显示数量和占比）
        for i, bar in enumerate(bars):
            width = bar.get_width()
            percentage = display_data.iloc[i]['占比(%)']
            ax.text(
                width + max(display_data['数量']) * 0.01,
                bar.get_y() + bar.get_height()/2.,
                f'{int(width)} ({percentage:.1f}%)',
                ha='left', va='center', fontsize=9, fontweight='bold'
            )
        
        plt.tight_layout()
        
        safe_name = self._safe_filename(model_name)
        chart_path = self.charts_dir / f"{rank:02d}_{safe_name}_分类分布.png"
        plt.savefig(chart_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return chart_path
    
    def _safe_filename(self, name, max_len=50):
        """清理文件名"""
        import re
        name = re.sub(r'[<>:"/\\|?*]', '_', name)
        if len(name) > max_len:
            name = name[:max_len]
        return name.strip()
    
    def _generate_report(self, model_stats, top_models, model_details, top_n):
        """生成报告"""
        lines = ["="*70, f"Top {top_n} Model 分析报告", "="*70]
        lines.append(f"总机型数: {len(model_stats)}")
        lines.append("")
        
        for idx, row in top_models.iterrows():
            lines.append(f"{row['排名']}. {row['机型名称']}: {row['分类数']}个分类")
        
        report_path = self.top_model_dir / f"Top{top_n}_Model分析报告.txt"
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))
        return report_path
    
    def get_ppt_payload(self):
        """获取PPT数据载荷"""
        return self.results


def run_top_model_analysis(df, output_dir, top_n=15, use_llm=False, llm_config=None):
    """便捷函数：运行Top Model分析"""
    service = TopModelAnalysisService(output_dir)
    return service.analyze(df, top_n, use_llm, llm_config)

