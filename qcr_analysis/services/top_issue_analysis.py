# -*- coding: utf-8 -*-
"""
=============================================================================
Top Issue Analysis Service - 热点问题分析服务
=============================================================================
负责Top N问题的统计分析和可视化
=============================================================================
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime

import sys
sys.path.append(str(Path(__file__).parent.parent))

from config import MATPLOTLIB_FONTS
from modules.llm_service import LLMService
from prompts import TOP_ISSUE_SUMMARY_PROMPT

# 设置中文字体
matplotlib.rcParams['font.family'] = MATPLOTLIB_FONTS
matplotlib.rcParams['axes.unicode_minus'] = False


class TopIssueAnalysisService:
    """Top Issue分析服务"""
    
    def __init__(self, output_dir: str or Path):
        """初始化Top Issue分析服务"""
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.top_issue_dir = self.output_dir / "Top_Issue分析"
        self.top_issue_dir.mkdir(parents=True, exist_ok=True)
        
        self.charts_dir = self.top_issue_dir / "charts"
        self.charts_dir.mkdir(parents=True, exist_ok=True)
        
        self.results = {}
    
    def analyze(
        self,
        df: pd.DataFrame,
        top_n: int = 10,
        use_llm: bool = False,
        llm_config: Optional[Dict] = None
    ) -> Dict:
        """执行Top Issue完整分析流程"""
        print("\n" + "="*70)
        print(f"🔥 Top {top_n} Issue 分析")
        print("="*70)
        
        if len(df) == 0 or '分类' not in df.columns:
            print("❌ 错误：数据为空或缺少'分类'列")
            return {}
        
        # 1. 统计Top N Issue
        print(f"\n📊 统计Top {top_n} Issue...")
        issue_counts = df['分类'].value_counts().head(top_n)
        
        issue_stats = pd.DataFrame({
            '排名': range(1, len(issue_counts) + 1),
            'Issue名称': issue_counts.index,
            '数量': issue_counts.values,
            '占比(%)': (issue_counts.values / len(df) * 100).round(2)
        })
        issue_stats['累计占比(%)'] = issue_stats['占比(%)'].cumsum().round(2)
        
        print(f"✓ 统计了Top {len(issue_stats)} Issue")
        
        # 保存统计表
        stats_path = self.top_issue_dir / f"Top{top_n}_Issue统计.xlsx"
        issue_stats.to_excel(stats_path, index=False)
        
        # 2. 生成总览图
        summary_chart = self._generate_summary_chart(issue_stats, top_n)
        
        # 3. 分析机型分布
        issue_details = self._analyze_issue_models(df, issue_stats)
        
        # 4. 生成报告
        report_path = self._generate_report(df, issue_stats, issue_details)
        
        self.results = {
            "issue_stats": issue_stats,
            "issue_details": issue_details,
            "summary_chart": summary_chart,
            "report_path": report_path,
            "total_records": len(df),
            "top_n": top_n
        }
        
        print("\n✅ Top Issue分析完成")
        return self.results
    
    def _generate_summary_chart(self, issue_stats, top_n):
        """生成总览图（带数据标签）"""
        plt.figure(figsize=(16, 8))
        bars = plt.bar(range(len(issue_stats)), issue_stats['数量'])
        plt.xlabel("Issue分类", fontsize=13)
        plt.ylabel("数量", fontsize=13)
        plt.title(f"Top {top_n} Issue分布", fontsize=14, fontweight='bold')
        plt.xticks(range(len(issue_stats)), issue_stats['Issue名称'], rotation=45, ha='right', fontsize=10)
        
        # 添加数值标签
        for i, bar in enumerate(bars):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + max(issue_stats['数量']) * 0.01,
                    f'{int(height)}\n({issue_stats.iloc[i]["占比(%)"]}%)',
                    ha='center', va='bottom', fontsize=9, fontweight='bold')
        
        plt.tight_layout()
        chart_path = self.charts_dir / "Top_Issue总览图.png"
        plt.savefig(chart_path, dpi=150, bbox_inches='tight')
        plt.close()
        return chart_path
    
    def _analyze_issue_models(self, df, issue_stats):
        """分析每个Issue的机型分布"""
        issue_details = []
        
        print(f"\n📊 分析每个Issue的机型分布...")
        for idx, row in issue_stats.iterrows():
            issue_name = row['Issue名称']
            issue_count = row['数量']
            
            # 筛选该Issue的数据
            issue_df = df[df['分类'] == issue_name]
            
            # 统计机型分布
            model_dist = issue_df['机型名称'].value_counts().reset_index()
            model_dist.columns = ['机型名称', '数量']
            model_dist['占比(%)'] = (model_dist['数量'] / issue_count * 100).round(2)
            
            # 保存机型分布Excel
            safe_name = self._safe_filename(issue_name)
            model_dist_path = self.top_issue_dir / f"{idx+1:02d}_{safe_name}_机型分布.xlsx"
            model_dist.to_excel(model_dist_path, index=False)
            
            # 生成机型分布图
            chart_path = None
            if len(model_dist) >= 2:
                chart_path = self._generate_model_chart(issue_name, model_dist, idx+1)
            
            issue_details.append({
                'rank': idx + 1,
                'issue_name': issue_name,
                'count': issue_count,
                'percentage': row['占比(%)'],
                'model_count': len(model_dist),
                'model_distribution': model_dist,
                'model_dist_path': model_dist_path,
                'chart_path': chart_path
            })
            
            print(f"  - Issue #{idx+1}: {issue_name} ({issue_count}条) -> {len(model_dist)}款机型")
        
        print(f"✓ 完成 {len(issue_details)} 个Issue的机型分布分析")
        return issue_details
    
    def _generate_model_chart(self, issue_name, model_dist, rank):
        """生成单个Issue的机型分布图（带数据标签）"""
        import re
        
        # 只显示前15个机型
        display_data = model_dist.head(15)
        
        fig, ax = plt.subplots(figsize=(14, 8))
        bars = ax.barh(range(len(display_data)), display_data['数量'])
        
        ax.set_yticks(range(len(display_data)))
        ax.set_yticklabels(display_data['机型名称'], fontsize=10)
        ax.set_xlabel("数量", fontsize=13)
        ax.set_title(
            f"Issue: {issue_name}\n机型分布 (共{len(model_dist)}款机型)",
            fontsize=14, fontweight='bold'
        )
        
        # 添加数值标签
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
        
        safe_name = self._safe_filename(issue_name)
        chart_path = self.charts_dir / f"{rank:02d}_{safe_name}_机型分布.png"
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
    
    def _generate_report(self, df, issue_stats, issue_details):
        """生成文本报告"""
        lines = ["="*70, "Top Issue 分析报告", "="*70]
        lines.append(f"总记录数: {len(df)}")
        lines.append(f"Top N: {len(issue_stats)}")
        lines.append("")
        
        for idx, row in issue_stats.iterrows():
            lines.append(f"{row['排名']}. {row['Issue名称']}: {row['数量']}条 ({row['占比(%)']}%)")
        
        report_path = self.top_issue_dir / "Top_Issue分析报告.txt"
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))
        return report_path
    
    def get_ppt_payload(self):
        """获取PPT数据载荷"""
        return self.results


def run_top_issue_analysis(df, output_dir, top_n=10, use_llm=False, llm_config=None):
    """便捷函数：运行Top Issue分析"""
    service = TopIssueAnalysisService(output_dir)
    return service.analyze(df, top_n, use_llm, llm_config)

