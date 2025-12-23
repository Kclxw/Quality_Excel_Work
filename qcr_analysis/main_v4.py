# -*- coding: utf-8 -*-
"""
=============================================================================
QCR 数据分析系统 v4.0 - 主入口
=============================================================================
支持Web模式（默认）和命令行模式
"""

import argparse
import sys
import webbrowser
import threading
import time
from pathlib import Path
from datetime import date

sys.path.append(str(Path(__file__).parent))

from data import DataManager
from modules.mtm_manager import MTMManager
from services import (
    run_weekly_analysis,
    run_top_issue_analysis,
    run_top_model_analysis,
    generate_weekly_report,
    generate_top_issue_report,
    generate_top_model_report
)

def parse_arguments():
    parser = argparse.ArgumentParser(description="QCR v4.0")
    parser.add_argument("--cli", action="store_true", help="命令行模式")
    parser.add_argument("--mode", choices=['weekly', 'top-issue', 'top-model'], help="分析模式")
    parser.add_argument("--data", dest="data_file", help="数据文件")
    parser.add_argument("--mtm", dest="mtm_file", help="MTM文件")
    parser.add_argument("--output", dest="output_dir", default="output", help="输出目录")
    parser.add_argument("--start-date", dest="start_date", help="开始日期")
    parser.add_argument("--end-date", dest="end_date", help="结束日期")
    parser.add_argument("--batch-name", dest="batch_name", default="2024-2025", help="批次名称")
    parser.add_argument("--top-n", dest="top_n", type=int, default=10, help="Top N")
    parser.add_argument("--filter-unmapped", action="store_true", help="过滤未映射")
    parser.add_argument("--generate-ppt", action="store_true", help="生成PPT")
    parser.add_argument("--port", type=int, default=5000, help="Web端口")
    return parser.parse_args()

def parse_date(date_str):
    if not date_str:
        return None
    try:
        parts = date_str.split('-')
        return date(int(parts[0]), int(parts[1]), int(parts[2]))
    except:
        return None

def run_cli_mode(args):
    """命令行模式"""
    if not args.mode or not args.data_file:
        print("错误：命令行模式需要 --mode 和 --data 参数")
        sys.exit(1)
    
    data_manager = DataManager()
    df = data_manager.read_excel(args.data_file)
    
    start_date = parse_date(args.start_date)
    end_date = parse_date(args.end_date)
    if start_date or end_date:
        df = data_manager.filter_by_date_range(df, start_date, end_date)
    
    mtm_manager = MTMManager(Path(args.mtm_file))
    df = mtm_manager.map_dataframe(df)
    
    if args.filter_unmapped:
        df = data_manager.filter_unmapped_mtm(df)
    
    if args.mode == 'weekly':
        from services.weekly_analysis import WeeklyAnalysisService
        service = WeeklyAnalysisService(args.output_dir)
        service.print_model_list(df)
        results = service.analyze(df, start_date, end_date)
        if args.generate_ppt:
            payload = service.get_ppt_payload()
            ppt_path = generate_weekly_report(payload, args.output_dir, args.batch_name)
            print(f"✓ PPT: {ppt_path}")
    
    elif args.mode == 'top-issue':
        results = run_top_issue_analysis(df, args.output_dir, args.top_n)
        if args.generate_ppt:
            from services.top_issue_analysis import TopIssueAnalysisService
            service = TopIssueAnalysisService(args.output_dir)
            payload = service.get_ppt_payload()
            ppt_path = generate_top_issue_report(payload, args.output_dir, args.batch_name)
            print(f"✓ PPT: {ppt_path}")
    
    elif args.mode == 'top-model':
        results = run_top_model_analysis(df, args.output_dir, args.top_n)
        if args.generate_ppt:
            from services.top_model_analysis import TopModelAnalysisService
            service = TopModelAnalysisService(args.output_dir)
            payload = service.get_ppt_payload()
            ppt_path = generate_top_model_report(payload, args.output_dir, args.batch_name)
            print(f"✓ PPT: {ppt_path}")

def run_web_mode(args):
    """Web模式"""
    print("="*70)
    print("QCR 数据分析系统 v4.0 - Web模式")
    print("="*70)
    
    from web import create_app
    app = create_app()
    
    port = args.port
    print(f"\n✓ 服务器启动: http://localhost:{port}")
    print(f"✓ 浏览器将自动打开")
    print(f"✓ 按 Ctrl+C 停止\n")
    
    def open_browser():
        time.sleep(1.5)
        webbrowser.open(f'http://localhost:{port}')
    threading.Thread(target=open_browser, daemon=True).start()
    
    try:
        app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)
    except KeyboardInterrupt:
        print("\n服务器已停止")

def main():
    args = parse_arguments()
    if args.cli:
        run_cli_mode(args)
    else:
        run_web_mode(args)

if __name__ == '__main__':
    main()

