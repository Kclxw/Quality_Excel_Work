# -*- coding: utf-8 -*-
"""测试所有服务"""
from pathlib import Path
from datetime import date
from data import DataManager
from modules.mtm_manager import MTMManager
from services import run_weekly_analysis, run_top_issue_analysis, run_top_model_analysis

CONFIG = {
    "data": r"D:\QCR\Annual_QCR_2025\Annual_data.xlsx",
    "mtm": r"D:\QCR\Annual_QCR_2025\MTM.xlsx",
    "output": r"D:\QCR\Annual_QCR_2025\output_v4_test",
    "start_date": date(2024, 4, 9),
    "end_date": date(2025, 11, 23),
}

def test_weekly():
    print("\n" + "="*70)
    print("测试 Weekly Service")
    print("="*70)
    try:
        results = run_weekly_analysis(
            data_source=CONFIG["data"],
            mtm_file=CONFIG["mtm"],
            output_dir=CONFIG["output"] + "/weekly",
            start_date=CONFIG["start_date"],
            end_date=CONFIG["end_date"],
            filter_unmapped=True
        )
        print(f"✅ Weekly 测试成功: {len(results['total_df'])} 条记录")
        return True
    except Exception as e:
        print(f"❌ Weekly 测试失败: {e}")
        return False

def test_top_issue():
    print("\n" + "="*70)
    print("测试 Top Issue Service")
    print("="*70)
    try:
        dm = DataManager()
        df = dm.read_excel(CONFIG["data"])
        df = dm.filter_by_date_range(df, CONFIG["start_date"], CONFIG["end_date"])
        mtm = MTMManager(Path(CONFIG["mtm"]))
        df = mtm.map_dataframe(df)
        df = dm.filter_unmapped_mtm(df)
        
        results = run_top_issue_analysis(df, CONFIG["output"] + "/top_issue", 10)
        print(f"✅ Top Issue 测试成功: Top {results['top_n']} 已分析")
        return True
    except Exception as e:
        print(f"❌ Top Issue 测试失败: {e}")
        return False

def test_top_model():
    print("\n" + "="*70)
    print("测试 Top Model Service")
    print("="*70)
    try:
        dm = DataManager()
        df = dm.read_excel(CONFIG["data"])
        df = dm.filter_by_date_range(df, CONFIG["start_date"], CONFIG["end_date"])
        mtm = MTMManager(Path(CONFIG["mtm"]))
        df = mtm.map_dataframe(df)
        df = dm.filter_unmapped_mtm(df)
        
        results = run_top_model_analysis(df, CONFIG["output"] + "/top_model", 15)
        print(f"✅ Top Model 测试成功: Top {results['top_n']} 机型已分析")
        return True
    except Exception as e:
        print(f"❌ Top Model 测试失败: {e}")
        return False

if __name__ == "__main__":
    print("="*70)
    print("QCR v4.0 服务测试")
    print("="*70)
    
    results = {
        "Weekly": test_weekly(),
        "Top Issue": test_top_issue(),
        "Top Model": test_top_model()
    }
    
    print("\n" + "="*70)
    print("测试总结")
    print("="*70)
    for name, success in results.items():
        print(f"  {name:15s}: {'✅' if success else '❌'}")
    print(f"\n通过率: {sum(results.values())}/{len(results)}")

