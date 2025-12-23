# -*- coding: utf-8 -*-
"""
=============================================================================
数据库操作模块
=============================================================================
负责与MySQL数据库的交互，包括：
- 连接管理
- 数据去重
- 数据导入
- 字段映射和清洗
=============================================================================
"""

import pandas as pd
from sqlalchemy import create_engine, text
from pathlib import Path
from typing import Optional

import sys
sys.path.append(str(Path(__file__).parent.parent))
from config import (
    DB_CONFIG,
    DB_COLUMN_MAPPING,
    DB_NUMERIC_COLUMNS,
    DB_STRING_COLUMNS,
    DB_STRING_MAX_LENGTHS,
    DB_REQUIRED_COLUMNS
)


class DatabaseManager:
    """数据库管理器"""
    
    def __init__(self, config: Optional[dict] = None):
        """
        初始化数据库管理器
        
        Args:
            config: 数据库配置字典，如果为None则使用默认配置
        """
        self.config = config if config else DB_CONFIG
        self.engine = None
        self.connected = False
    
    def connect(self) -> bool:
        """
        建立数据库连接
        
        Returns:
            连接是否成功
        """
        try:
            connection_string = (
                f"mysql+pymysql://{self.config['user']}:{self.config['password']}@"
                f"{self.config['host']}:{self.config['port']}/{self.config['database']}"
            )
            self.engine = create_engine(connection_string)
            
            # 测试连接
            with self.engine.connect() as conn:
                pass
            
            self.connected = True
            print("✓ 数据库连接成功")
            return True
        except Exception as e:
            print(f"✗ 数据库连接失败: {e}")
            self.connected = False
            return False
    
    def check_table_exists(self, table_name: Optional[str] = None) -> bool:
        """
        检查表是否存在
        
        Args:
            table_name: 表名，默认使用配置中的表名
            
        Returns:
            表是否存在
        """
        if not self.connected:
            return False
        
        table_name = table_name or self.config.get('table_name', 'QCR_data')
        
        try:
            result = pd.read_sql(
                "SELECT COUNT(*) as count FROM information_schema.tables "
                "WHERE table_schema = %s AND table_name = %s",
                self.engine,
                params=(self.config['database'], table_name)
            )
            return result['count'].iloc[0] > 0
        except Exception as e:
            print(f"检查表是否存在失败: {e}")
            return False
    
    def get_existing_service_orders(self, table_name: Optional[str] = None) -> list:
        """
        获取数据库中已存在的服务单号列表
        
        Args:
            table_name: 表名，默认使用配置中的表名
            
        Returns:
            服务单号列表
        """
        if not self.connected:
            return []
        
        table_name = table_name or self.config.get('table_name', 'QCR_data')
        
        if not self.check_table_exists(table_name):
            print(f"表 {table_name} 不存在，将创建新表")
            return []
        
        try:
            existing_orders = pd.read_sql(
                f"SELECT service_order_id FROM {table_name}",
                self.engine
            )['service_order_id'].astype(str).tolist()
            print(f"✓ 数据库中已存在 {len(existing_orders)} 个服务单号")
            return existing_orders
        except Exception as e:
            print(f"查询服务单号失败: {e}")
            return []
    
    def filter_new_records(self, df: pd.DataFrame, service_order_column: str) -> pd.DataFrame:
        """
        筛选数据库中不存在的新记录
        
        Args:
            df: 原始数据DataFrame
            service_order_column: 服务单号列名
            
        Returns:
            新记录的DataFrame
        """
        if service_order_column not in df.columns:
            print(f"警告：未找到'{service_order_column}'列，跳过数据库去重")
            return df
        
        # 获取当前数据中的服务单号
        current_orders = df[service_order_column].dropna().astype(str).tolist()
        print(f"当前数据包含 {len(current_orders)} 个服务单号")
        
        # 获取数据库中已存在的服务单号
        existing_orders = self.get_existing_service_orders()
        
        # 筛选新服务单号
        new_orders = [order for order in current_orders if order not in existing_orders]
        print(f"新服务单号数量: {len(new_orders)}")
        
        # 筛选新数据
        df_new = df[df[service_order_column].astype(str).isin(new_orders)].copy()
        
        if len(df_new) == 0:
            print("没有新数据需要导入和分析")
        else:
            print(f"✓ 筛选出 {len(df_new)} 条新记录")
        
        return df_new
    
    def prepare_for_import(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        准备数据以导入数据库
        包括：列重命名、数据类型转换、数据清洗
        
        Args:
            df: 原始数据DataFrame
            
        Returns:
            准备好的DataFrame
        """
        df_import = df.copy()
        
        # 1. 列重命名 - 优先完全匹配，然后才是包含匹配
        column_mapping = {}
        for col in df_import.columns:
            col_str = str(col).strip()
            # 先尝试完全匹配
            if col_str in DB_COLUMN_MAPPING:
                column_mapping[col] = DB_COLUMN_MAPPING[col_str]
            else:
                # 再尝试包含匹配（向后兼容）
                for key, value in DB_COLUMN_MAPPING.items():
                    if key in col_str:
                        column_mapping[col] = value
                        break
        
        df_import = df_import.rename(columns=column_mapping)
        
        print(f"  列映射: {len(column_mapping)} 个列被映射")
        
        # 2. 确保必需的列存在
        for col in DB_REQUIRED_COLUMNS:
            if col not in df_import.columns:
                df_import[col] = ''
        
        # 3. 处理日期列
        if 'date' in df_import.columns:
            df_import['date'] = pd.to_datetime(df_import['date'], errors='coerce')
            # 填充空日期为当前日期
            df_import['date'] = df_import['date'].fillna(pd.Timestamp.now())
            df_import['date'] = df_import['date'].dt.strftime('%Y-%m-%d')
        
        # 4. 处理数值列（NOT NULL约束）
        for col in DB_NUMERIC_COLUMNS:
            if col in df_import.columns:
                df_import[col] = pd.to_numeric(df_import[col], errors='coerce')
                # 对于NOT NULL的数值字段，使用0填充空值
                df_import[col] = df_import[col].fillna(0)
                df_import[col] = df_import[col].astype('int64')  # 使用int64而不是Int64，避免可空类型
        
        # 5. 处理字符串列（NOT NULL约束）
        for col in DB_STRING_COLUMNS:
            if col in df_import.columns:
                # 先转换为字符串，然后填充空值
                df_import[col] = df_import[col].astype(str)
                # 将'nan', 'None', 'NaN'等替换为空字符串
                df_import[col] = df_import[col].replace(['nan', 'None', 'NaN', '<NA>'], '')
                df_import[col] = df_import[col].fillna('')
                df_import[col] = df_import[col].str.strip()
                # 字符串长度限制
                if col in DB_STRING_MAX_LENGTHS:
                    df_import[col] = df_import[col].str[:DB_STRING_MAX_LENGTHS[col]]
                # 对于特定的NOT NULL字段，如果为空则填充默认值
                if col in ['product_name', 'sn_code', 'customer_account', 'audit_reason', 'mtm']:
                    df_import[col] = df_import[col].replace('', '未知')
                if col == 'issue_description':
                    df_import[col] = df_import[col].replace('', '无描述')
                if col in ['issue_category', 'category']:
                    df_import[col] = df_import[col].replace('', '未分类')
        
        # 6. 删除关键字段无效的行
        if 'service_order_id' in df_import.columns:
            before_drop = len(df_import)
            # 删除服务单号为0或空的记录
            df_import = df_import[df_import['service_order_id'] > 0]
            after_drop = len(df_import)
            if before_drop > after_drop:
                print(f"  删除了 {before_drop - after_drop} 条服务单号无效的记录")
        
        # 7. 只选择数据库需要的列（如果列存在）
        available_columns = [col for col in DB_REQUIRED_COLUMNS if col in df_import.columns]
        missing_columns = [col for col in DB_REQUIRED_COLUMNS if col not in df_import.columns]
        
        if missing_columns:
            print(f"  警告: 以下必需列缺失: {missing_columns}")
            print(f"  可用的列: {df_import.columns.tolist()}")
            # 为缺失的列填充默认值
            for col in missing_columns:
                if col in DB_NUMERIC_COLUMNS:
                    df_import[col] = 0
                else:
                    df_import[col] = ''
        
        df_import = df_import[DB_REQUIRED_COLUMNS]
        
        return df_import
    
    def import_data(self, df: pd.DataFrame, table_name: Optional[str] = None) -> bool:
        """
        导入数据到数据库
        
        Args:
            df: 要导入的DataFrame
            table_name: 表名，默认使用配置中的表名
            
        Returns:
            导入是否成功
        """
        if not self.connected:
            print("错误：数据库未连接")
            return False
        
        if len(df) == 0:
            print("没有数据需要导入")
            return True
        
        table_name = table_name or self.config.get('table_name', 'QCR_data')
        
        try:
            # 显示准备导入的数据信息
            print(f"  准备导入 {len(df)} 条记录到表 {table_name}")
            print(f"  列: {df.columns.tolist()}")
            
            # 检查数据中是否有NULL值（针对NOT NULL字段）
            null_counts = df.isnull().sum()
            if null_counts.sum() > 0:
                print("  警告：发现以下列包含空值：")
                for col, count in null_counts[null_counts > 0].items():
                    print(f"    {col}: {count} 个空值")
            
            # 使用单条插入模式，更容易定位问题
            # 如果数据量小于100条，使用单条插入；否则使用批量插入
            if len(df) < 100:
                df.to_sql(
                    table_name,
                    self.engine,
                    if_exists='append',
                    index=False,
                    method=None  # 使用默认方法，逐条插入
                )
            else:
                # 批量插入，提高效率
                df.to_sql(
                    table_name,
                    self.engine,
                    if_exists='append',
                    index=False,
                    method='multi',
                    chunksize=100  # 每次插入100条
                )
            
            print(f"✓ 成功导入 {len(df)} 条记录到数据库表 {table_name}")
            return True
        except Exception as e:
            print(f"✗ 导入数据失败: {e}")
            print(f"  错误类型: {type(e).__name__}")
            
            # 尝试找出有问题的记录
            print("\n尝试诊断问题...")
            try:
                # 显示前几行数据的信息
                print("\n前3条数据样例：")
                for idx in range(min(3, len(df))):
                    print(f"\n记录 {idx + 1}:")
                    for col in df.columns:
                        val = df.iloc[idx][col]
                        val_type = type(val).__name__
                        print(f"  {col}: {val} (类型: {val_type})")
            except Exception as diag_e:
                print(f"诊断失败: {diag_e}")
            
            import traceback
            print("\n详细错误信息:")
            traceback.print_exc()
            return False
    
    def check_and_import_new_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        完整流程：检查、筛选新数据、导入数据库
        
        Args:
            df: 原始数据DataFrame
            
        Returns:
            新数据的DataFrame
        """
        try:
            # 1. 连接数据库
            if not self.connected:
                if not self.connect():
                    print("数据库连接失败，跳过数据库操作")
                    return df
            
            # 2. 查找服务单号列
            service_order_column = None
            for col in df.columns:
                if '服务单号' in str(col):
                    service_order_column = col
                    break
            
            if service_order_column is None:
                print("警告：未找到'服务单号'列，跳过数据库去重")
                return df
            
            # 3. 筛选新数据
            df_new = self.filter_new_records(df, service_order_column)
            
            if len(df_new) == 0:
                return df_new
            
            # 4. 准备数据
            df_import = self.prepare_for_import(df_new)
            
            # 5. 导入数据库
            self.import_data(df_import)
            
            # 返回新数据用于后续分析
            return df_new
            
        except Exception as e:
            print(f"数据库操作失败: {e}")
            print("将继续分析原始数据，跳过数据库检查和导入")
            return df
    
    def query_by_date_range(self, start_date: str, end_date: str, 
                           table_name: Optional[str] = None) -> pd.DataFrame:
        """
        从数据库按日期范围查询数据
        
        Args:
            start_date: 开始日期 (YYYY-MM-DD)
            end_date: 结束日期 (YYYY-MM-DD)
            table_name: 表名，默认使用配置中的表名
            
        Returns:
            查询结果的DataFrame
        """
        if not self.connected:
            if not self.connect():
                print("数据库连接失败")
                return pd.DataFrame()
        
        table_name = table_name or self.config.get('table_name', 'QCR_data')
        
        if not self.check_table_exists(table_name):
            print(f"表 {table_name} 不存在")
            return pd.DataFrame()
        
        try:
            query = f"""
                SELECT * FROM {table_name}
                WHERE date >= %s AND date <= %s
                ORDER BY date DESC
            """
            df = pd.read_sql(query, self.engine, params=(start_date, end_date))
            print(f"✓ 从数据库查询到 {len(df)} 条记录 ({start_date} ~ {end_date})")
            
            # 将数据库列名映射回Excel列名（反向映射）
            reverse_mapping = {v: k for k, v in DB_COLUMN_MAPPING.items()}
            df = df.rename(columns=reverse_mapping)
            
            return df
        except Exception as e:
            print(f"✗ 查询数据失败: {e}")
            import traceback
            traceback.print_exc()
            return pd.DataFrame()
    
    def update_mtm_mappings(self, mtm_file: str) -> bool:
        """
        从MTM表格更新数据库中的product_name
        
        Args:
            mtm_file: MTM映射表Excel文件路径
            
        Returns:
            更新是否成功
        """
        if not self.connected:
            if not self.connect():
                print("数据库连接失败")
                return False
        
        try:
            # 读取MTM表格
            print(f"📖 读取MTM映射表: {mtm_file}")
            mtm_df = pd.read_excel(mtm_file)
            
            # 查找MTM和产品名称列
            mtm_col = None
            product_col = None
            
            for col in mtm_df.columns:
                col_lower = str(col).lower().strip()
                if 'mtm' in col_lower and mtm_col is None:
                    mtm_col = col
                if ('product' in col_lower or '产品' in col_lower or '机型' in col_lower) and product_col is None:
                    product_col = col
            
            if mtm_col is None or product_col is None:
                print(f"✗ MTM表格格式不正确，需要包含MTM列和产品名称列")
                print(f"   找到的列: {mtm_df.columns.tolist()}")
                return False
            
            print(f"✓ 识别列映射: MTM列='{mtm_col}', 产品名称列='{product_col}'")
            
            # 清理数据
            mtm_df = mtm_df[[mtm_col, product_col]].dropna()
            mtm_df[mtm_col] = mtm_df[mtm_col].astype(str).str.strip()
            mtm_df[product_col] = mtm_df[product_col].astype(str).str.strip()
            
            print(f"✓ 读取到 {len(mtm_df)} 条MTM映射记录")
            
            # 批量更新数据库
            table_name = self.config.get('table_name', 'QCR_data')
            updated_count = 0
            
            print(f"\n🔄 开始更新数据库表 {table_name} 中的product_name...")
            
            with self.engine.begin() as conn:
                for idx, row in mtm_df.iterrows():
                    mtm_code = row[mtm_col]
                    product_name = row[product_col]
                    
                    # 执行UPDATE（使用命名参数）
                    result = conn.execute(
                        text(f"""
                        UPDATE {table_name}
                        SET product_name = :product_name
                        WHERE mtm = :mtm_code
                        """),
                        {"product_name": product_name, "mtm_code": mtm_code}
                    )
                    
                    if result.rowcount > 0:
                        updated_count += result.rowcount
                        if (idx + 1) % 10 == 0:
                            print(f"  已处理 {idx + 1}/{len(mtm_df)} 条映射记录...")
            
            print(f"\n✅ MTM映射更新完成！")
            print(f"   处理了 {len(mtm_df)} 条映射记录")
            print(f"   更新了 {updated_count} 条数据库记录")
            
            return True
            
        except Exception as e:
            print(f"✗ 更新MTM映射失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def import_excel_to_db(self, excel_file: str) -> bool:
        """
        独立功能：将Excel数据导入数据库（带去重）
        
        Args:
            excel_file: Excel文件路径
            
        Returns:
            导入是否成功
        """
        if not self.connected:
            if not self.connect():
                print("数据库连接失败")
                return False
        
        try:
            # 读取Excel
            print(f"\n📖 读取Excel文件: {excel_file}")
            df = pd.read_excel(excel_file, sheet_name=0)
            print(f"✓ 读取到 {len(df)} 条记录")
            
            # 处理日期列（假设第一列是日期）
            date_column = df.columns[0]
            df[date_column] = pd.to_datetime(df[date_column]).dt.date
            
            # 查找服务单号列
            service_order_column = None
            for col in df.columns:
                if '服务单号' in str(col):
                    service_order_column = col
                    break
            
            if service_order_column is None:
                print("✗ 未找到'服务单号'列，无法进行去重")
                return False
            
            # 去重检测
            print(f"\n🔍 开始数据库去重检测...")
            df_new = self.filter_new_records(df, service_order_column)
            
            if len(df_new) == 0:
                print("✓ 没有新数据需要导入")
                return True
            
            # 准备数据
            print(f"\n⚙️  准备数据...")
            df_import = self.prepare_for_import(df_new)
            
            # 导入数据库
            print(f"\n📥 导入数据到数据库...")
            success = self.import_data(df_import)
            
            if success:
                print(f"\n✅ Excel导入完成！")
                print(f"   原始记录: {len(df)} 条")
                print(f"   新记录: {len(df_new)} 条")
                print(f"   已导入数据库")
            
            return success
            
        except Exception as e:
            print(f"✗ 导入Excel失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def close(self):
        """关闭数据库连接"""
        if self.engine:
            self.engine.dispose()
            self.connected = False
            print("✓ 数据库连接已关闭")

