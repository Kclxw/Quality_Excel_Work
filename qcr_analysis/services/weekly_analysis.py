# -*- coding: utf-8 -*-
"""
=============================================================================
Weekly Analysis Service - 周报分析服务
=============================================================================
负责7天无理由和非7天无理由的质量分析
完全复制原有逻辑，不做任何修改
=============================================================================
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from datetime import date

import sys
sys.path.append(str(Path(__file__).parent.parent))

# 导入现有模块（完全复用）
from modules.data_analyzer import DataAnalyzer
from modules.llm_service import LLMService
from data import DataManager
from modules.mtm_manager import MTMManager


class WeeklyAnalysisService:
    """
    Weekly Report分析服务
    提供7天无理由和非7天无理由的完整分析流程
    """
    
    def __init__(self, output_dir: str or Path):
        """
        初始化Weekly分析服务
        
        Args:
            output_dir: 输出目录路径
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # 使用现有的DataAnalyzer（完全复用）
        self.analyzer = DataAnalyzer(self.output_dir)
        
        # 结果缓存
        self.results = {}
        self.summary_excel_path = self.output_dir / "weekly_summary.xlsx"
    
    def analyze(
        self,
        df: pd.DataFrame,
        start_date: Optional[date] = None,
        end_date: Optional[date] = None,
        use_llm: bool = False,
        llm_config: Optional[Dict] = None
    ) -> Dict:
        """
        执行Weekly Report完整分析流程
        
        Args:
            df: 输入数据DataFrame（必须包含'机型名称'列）
            start_date: 开始日期
            end_date: 结束日期
            use_llm: 是否使用LLM生成摘要
            llm_config: LLM配置参数
            
        Returns:
            分析结果字典
        """
        print("\n" + "="*70)
        print("📊 Weekly Report 分析")
        print("="*70)
        
        # 1. 数据分类（完全复制原有逻辑）
        print("\n📊 开始数据分析...")
        cond_7d = df["审核原因"] == "7天无理由"
        cond_non_7d = df["审核原因"].isin(["15天质量换新", "180天只换不修", "质量维修"])
        
        df_7d = df[cond_7d].copy()
        df_non_7d = df[cond_non_7d].copy()
        
        print(f"  7天无理由记录: {len(df_7d)} 条")
        print(f"  非7天无理由记录: {len(df_non_7d)} 条")
        
        # 2. 审核原因统计
        print("\n📈 统计审核原因...")
        reason_stats, reason_chart_path = self.analyzer.analyze_audit_reasons(df)
        
        # 3. 机型分布统计
        print("\n📈 统计机型分布...")
        model_7d_dist, model_7d_chart_path = self.analyzer.analyze_model_distribution(
            df_7d, "7天无理由"
        )
        model_non_7d_dist, model_non_7d_chart_path = self.analyzer.analyze_model_distribution(
            df_non_7d, "非7天无理由"
        )
        
        # 4. 机型问题分析
        print("\n📈 分析机型问题分类...")
        print("  7天无理由机型分析:")
        summaries_7d = self.analyzer.analyze_model_issues(df_7d, "7天无理由")
        print("\n  非7天无理由机型分析:")
        summaries_non7d = self.analyzer.analyze_model_issues(df_non_7d, "非7天无理由")
        
        # 5. 生成文本报告
        print("\n📝 生成文本报告...")
        self.analyzer.generate_text_report(df, df_7d, df_non_7d, start_date, end_date)
        
        # 6. 保存结果
        self.results = {
            "total_df": df,
            "df_7d": df_7d,
            "df_non_7d": df_non_7d,
            "reason_stats": reason_stats,
            "reason_chart": reason_chart_path,
            "model_7d_dist": model_7d_dist,
            "model_7d_chart": model_7d_chart_path,
            "model_non_7d_dist": model_non_7d_dist,
            "model_non_7d_chart": model_non_7d_chart_path,
            "summaries_7d": summaries_7d,
            "summaries_non7d": summaries_non7d,
            "start_date": start_date,
            "end_date": end_date,
        }

        # 导出关键数据到Excel，便于留档
        self._export_summary_excel()
        
        print("\n✅ Weekly Report分析完成")
        print("="*70)
        
        return self.results
    
    def _detect_date_column(self, df: pd.DataFrame) -> str:
        """智能选择日期列"""
        # 优先常用列名
        preferred = ['审核日期', '日期', 'date', 'Date']
        for col in preferred:
            if col in df.columns:
                return col

        # 尝试找可解析为日期的列
        for col in df.columns:
            try:
                pd.to_datetime(df[col])
                return col
            except Exception:
                continue

        # 回退第一列
        return df.columns[0]

    def get_ppt_payload(self, date_column: str = None) -> Dict:
        """
        获取PPT生成所需的数据载荷
        
        Args:
            date_column: 日期列名
            
        Returns:
            PPT数据字典
        """
        from modules.data_analyzer import get_week_workday_range, determine_coverage_range
        
        if not self.results:
            raise ValueError("请先调用 analyze() 方法")
        
        df = self.results["total_df"]
        start_date = self.results["start_date"]
        end_date = self.results["end_date"]
        
        # 确定日期列
        if date_column is None:
            date_column = self._detect_date_column(df)
        
        payload = {
            "start_date": start_date,
            "end_date": end_date,
            "week_range": get_week_workday_range(),
            "coverage_period": determine_coverage_range(df, date_column, start_date, end_date),
            "total_records": len(df),
            "reason_stats": self.results["reason_stats"],
            "model_7d_dist": self.results["model_7d_dist"],
            "model_non_7d_dist": self.results["model_non_7d_dist"],
            "summaries_7d": self.results["summaries_7d"],
            "summaries_non7d": self.results["summaries_non7d"],
            "reason_chart_path": self.results["reason_chart"],
            "model_7d_chart_path": self.results["model_7d_chart"],
            "model_non_7d_chart_path": self.results["model_non_7d_chart"],
        }
        
        return payload
    
    def print_model_list(self, df: pd.DataFrame):
        """输出所有分析的机型名称列表"""
        print("\n📋 本次分析涉及的机型列表:")
        print("="*60)
        unique_models = df['机型名称'].unique()
        print(f"共 {len(unique_models)} 个机型:\n")
        
        model_stats = df.groupby('机型名称').size().reset_index(name='记录数')
        model_stats = model_stats.sort_values('记录数', ascending=False)
        
        for idx, row in model_stats.iterrows():
            model_name = row['机型名称']
            count = row['记录数']
            sample_mtm = df[df['机型名称'] == model_name]['MTM'].iloc[0]
            is_mapped = (model_name != sample_mtm)
            status = "✓" if is_mapped else "⊗"
            print(f"  {status} {model_name[:60]:60s} - {count:5d} 条记录")
        
        print("="*60)
        print(f"说明: ✓=已映射机型  ⊗=未映射机型(显示原MTM)")
        print("="*60)
    
    def get_results(self) -> Dict:
        """获取分析结果"""
        return self.results

    def _export_summary_excel(self):
        """导出Weekly关键数据到Excel"""
        if not self.results:
            return
        try:
            with pd.ExcelWriter(self.summary_excel_path, engine="openpyxl") as writer:
                # 原始拆分数据
                self.results["df_7d"].to_excel(writer, sheet_name="7天无理由", index=False)
                self.results["df_non_7d"].to_excel(writer, sheet_name="非7天无理由", index=False)
                # 统计表
                self.results["reason_stats"].to_excel(writer, sheet_name="审核原因统计", index=False)
                self.results["model_7d_dist"].to_excel(writer, sheet_name="7天机型分布", index=False)
                self.results["model_non_7d_dist"].to_excel(writer, sheet_name="非7天机型分布", index=False)
        except Exception as e:
            print(f"导出Weekly汇总Excel失败: {e}")


# ================================================================
# 便捷函数
# ================================================================

def run_weekly_analysis(
    data_source: str,
    mtm_file: str,
    output_dir: str,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
    filter_unmapped: bool = False,
    use_database: bool = False,
    use_llm: bool = False,
    **kwargs
) -> Dict:
    """
    便捷函数：运行完整的Weekly Report分析
    
    Args:
        data_source: 数据源路径或"database"
        mtm_file: MTM映射文件路径
        output_dir: 输出目录
        start_date: 开始日期
        end_date: 结束日期
        filter_unmapped: 是否过滤未映射的MTM
        use_database: 是否使用数据库
        use_llm: 是否使用LLM
        **kwargs: 其他参数
        
    Returns:
        分析结果字典
    """
    # 1. 加载数据
    print("\n🔄 加载数据...")
    data_manager = DataManager()
    
    if use_database:
        data_manager.connect_database()
        df = data_manager.read_from_database(start_date, end_date)
    else:
        df = data_manager.read_excel(data_source)
        if start_date or end_date:
            df = data_manager.filter_by_date_range(df, start_date, end_date)
    
    print(f"✓ 成功读取 {len(df)} 条记录")
    
    # 2. MTM映射
    print("\n🔄 MTM映射处理...")
    mtm_manager = MTMManager(Path(mtm_file))
    df = mtm_manager.map_dataframe(df)
    mtm_manager.print_statistics()
    
    # 3. 过滤未映射（如果需要）
    if filter_unmapped:
        print("\n🔍 过滤未映射的MTM...")
        original_count = len(df)
        df = data_manager.filter_unmapped_mtm(df)
        print(f"✓ 已过滤 {original_count - len(df)} 条未映射的记录")
        print(f"✓ 剩余 {len(df)} 条已映射的记录用于分析")
    
    # 4. 执行分析
    service = WeeklyAnalysisService(output_dir)
    service.print_model_list(df)
    
    # 传递LLM配置
    llm_config = kwargs.get('llm_config')
    results = service.analyze(df, start_date, end_date, use_llm, llm_config)
    
    return results

