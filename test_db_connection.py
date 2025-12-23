#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试数据库连接功能
"""

import pandas as pd
import pymysql
from sqlalchemy import create_engine

# 数据库配置
DB_CONFIG = {
    'host': 'localhost',
    'port': 3306,
    'user': 'root',
    'password': '0929',
    'database': 'local_qcr'
}

def test_database_connection():
    """测试数据库连接"""
    try:
        print("开始测试数据库连接...")
        
        # 创建数据库连接
        connection_string = (
            f"mysql+pymysql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@"
            f"{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
        )
        engine = create_engine(connection_string)
        
        # 测试连接
        test_query = "SELECT 1 as test"
        result = pd.read_sql(test_query, engine)
        print("✅ 数据库连接成功！")
        
        # 检查表是否存在
        try:
            table_check = "SHOW TABLES LIKE 'qcr_data'"
            table_result = pd.read_sql(table_check, engine)
            if len(table_result) > 0:
                print("✅ qcr_data 表存在")
                
                # 检查表结构
                describe_query = "DESCRIBE qcr_data"
                describe_result = pd.read_sql(describe_query, engine)
                print("表结构:")
                print(describe_result)
                
                # 检查现有数据量
                count_query = "SELECT COUNT(*) as count FROM qcr_data"
                count_result = pd.read_sql(count_query, engine)
                print(f"当前数据库中的记录数: {count_result['count'].iloc[0]}")
                
            else:
                print("⚠️  qcr_data 表不存在，需要创建")
                
        except Exception as e:
            print(f"检查表时出错: {e}")
            
    except Exception as e:
        print(f"❌ 数据库连接失败: {e}")
        print("请检查:")
        print("1. MySQL服务是否正在运行")
        print("2. 数据库配置是否正确")
        print("3. 数据库和用户是否存在")

if __name__ == "__main__":
    test_database_connection()
