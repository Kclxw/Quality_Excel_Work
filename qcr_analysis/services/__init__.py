# -*- coding: utf-8 -*-
"""
=============================================================================
功能层服务模块
=============================================================================
提供三大分析服务、可视化和报告生成服务
"""

from .weekly_analysis import WeeklyAnalysisService, run_weekly_analysis
from .top_issue_analysis import TopIssueAnalysisService, run_top_issue_analysis
from .top_model_analysis import TopModelAnalysisService, run_top_model_analysis
from .visualization_service import VisualizationService, create_visualization_service
from .report_service import (
    ReportService,
    create_report_service,
    generate_weekly_report,
    generate_top_issue_report,
    generate_top_model_report
)

__all__ = [
    # 分析服务
    'WeeklyAnalysisService',
    'TopIssueAnalysisService',
    'TopModelAnalysisService',
    # 可视化服务
    'VisualizationService',
    'create_visualization_service',
    # 报告服务
    'ReportService',
    'create_report_service',
    # 便捷函数
    'run_weekly_analysis',
    'run_top_issue_analysis',
    'run_top_model_analysis',
    'generate_weekly_report',
    'generate_top_issue_report',
    'generate_top_model_report'
]

