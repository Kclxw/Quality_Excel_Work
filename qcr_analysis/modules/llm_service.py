# -*- coding: utf-8 -*-
"""
=============================================================================
LLM服务模块
=============================================================================
负责与Kimi API交互，生成智能分析摘要
=============================================================================
"""

import json
import requests
import pandas as pd
from typing import Dict, List
from pathlib import Path

import sys
sys.path.append(str(Path(__file__).parent.parent))
from config import KIMI_API_KEY, KIMI_API_URL, KIMI_MODEL


class LLMGenerationError(Exception):
    """LLM生成错误异常"""
    pass


class LLMService:
    """LLM服务类"""
    
    def __init__(self, api_key: str = None, api_url: str = None, model: str = None):
        """
        初始化LLM服务
        
        Args:
            api_key: API密钥
            api_url: API地址
            model: 模型名称
        """
        self.api_key = api_key or KIMI_API_KEY
        self.api_url = api_url or KIMI_API_URL
        self.model = model or KIMI_MODEL
    
    def test_connection(self) -> bool:
        """
        测试Kimi API连通性
        
        Returns:
            连接成功返回True，失败返回False
        """
        print("\n" + "="*60)
        print("🔍 Kimi API 连通性测试")
        print("="*60)
        
        # 检查API Key配置
        print(f"\n1. 检查API Key配置...")
        if not self.api_key or self.api_key == "":
            print("   ❌ 错误: 未配置KIMI_API_KEY")
            print("   请在config.py中配置 DEFAULT_KIMI_API_KEY 或设置环境变量 KIMI_API_KEY")
            return False
        
        # 显示配置信息（隐藏部分密钥）
        masked_key = self.api_key[:10] + "..." + self.api_key[-8:] if len(self.api_key) > 18 else "***"
        print(f"   ✓ API Key: {masked_key}")
        print(f"   ✓ API URL: {self.api_url}")
        print(f"   ✓ 模型: {self.model}")
        
        # 发送测试请求
        print(f"\n2. 发送测试请求...")
        test_message = {
            "role": "user",
            "content": "你好，请简单回复'连接成功'即可。"
        }
        
        try:
            response = self.call_api([test_message], timeout=30)
            print(f"   ✓ 请求成功!")
            print(f"\n3. Kimi 响应:")
            print(f"   {response}")
            
            print("\n" + "="*60)
            print("✅ Kimi API 连接测试成功！")
            print("="*60 + "\n")
            return True
            
        except LLMGenerationError as exc:
            print(f"   ❌ 请求失败: {exc}")
            print("\n" + "="*60)
            print("❌ Kimi API 连接测试失败")
            print("="*60)
            print("\n可能的原因:")
            print("1. API Key 不正确或已过期")
            print("2. 网络连接问题")
            print("3. API 服务暂时不可用")
            print("4. API URL 或模型名称配置错误")
            print("\n请检查配置后重试。\n")
            return False
    
    def call_api(self, messages: List[Dict[str, str]], timeout: int = 60) -> str:
        """
        调用Kimi API
        
        Args:
            messages: 消息列表
            timeout: 超时时间（秒）
            
        Returns:
            API返回的文本内容
            
        Raises:
            LLMGenerationError: API调用失败时抛出
        """
        if not self.api_key or self.api_key == "":
            raise LLMGenerationError("未配置KIMI_API_KEY，请在config.py中配置API Key")
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        payload = {
            "model": self.model,
            "messages": messages,
            "temperature": 0.2
        }
        
        try:
            response = requests.post(
                self.api_url,
                headers=headers,
                data=json.dumps(payload),
                timeout=timeout
            )
        except requests.RequestException as exc:
            raise LLMGenerationError(f"Kimi API 请求异常: {exc}")
        
        if response.status_code != 200:
            raise LLMGenerationError(f"Kimi API 请求失败: {response.status_code} - {response.text}")
        
        data = response.json()
        try:
            return data["choices"][0]["message"]["content"].strip()
        except (KeyError, IndexError) as exc:
            raise LLMGenerationError(f"Kimi API 响应解析失败: {exc}")
    
    def dataframe_to_category_rows(self, df: pd.DataFrame) -> List[Dict[str, str]]:
        """
        将DataFrame转换为分类行格式
        
        Args:
            df: 包含分类数据的DataFrame
            
        Returns:
            分类行列表
        """
        rows = []
        for _, row in df.iterrows():
            rows.append({
                "Category": str(row.get("分类", "")),
                "Count": str(row.get("次数", "")),
                "Share": str(row.get("占比", ""))
            })
        return rows
    
    def build_prompt(self, category_rows: List[Dict[str, str]], 
                    top_n: int, coverage_threshold: float, 
                    focus_threshold: float) -> Dict[str, str]:
        """
        构建Prompt
        
        Args:
            category_rows: 分类数据行
            top_n: TopN参数
            coverage_threshold: 覆盖度阈值
            focus_threshold: 重点拦截阈值
            
        Returns:
            Prompt消息字典
        """
        table_lines = ["分类\t频次\t占比"]
        for row in category_rows:
            table_lines.append(f"{row['Category']}\t{row['Count']}\t{row['Share']}")
        table_text = "\n".join(table_lines)
        
        prompt = f"""# 角色
你是一名PC电脑制造业的质量管理专家与用户反馈分析专家。你的任务是在严格依赖输入表格（包含列：分类、频次、占比）的前提下，不引入外部信息、不自行计算/重算占比，输出高度凝练的核心观点与可执行建议，用于问题拦截与后续复现/根因分析。

## 输入数据表格
{table_text}

## 技能
### 技能 1: 生成核心观点
1. 输入包含"分类""频次""占比"的表格数据。
2. 开篇交代样本处理与总量（如"去除无效后共 N 项"），并直接点名 Top-N：
    - 采用紧凑体例："分类名*频次（占比）"（示例：无法开机*20（27.0%））。
3. 输出剩余分类情况："其余问题分布较为分散，无明显集中性"。
4. Top-N需要满足：
    - 必须输出Top1。
    - 对于Top2和Top3分别需要大于等于15%才可以被输出。
    - 对于Top4可以不输出。

### 技能 2: 生成可执行建议
1. 根据生成的核心观点中的Top-N分类。
2. 明确重点拦截清单，并给出下一步：
    - 给出拦截建议，一般拦截Top-N的机型，如果分类问题的频次较少可以不拦截，话术参考："建议对死机，无法开机等机器进行退机拦截处理，做进一步分析。"
    - 若分类名已明确指向方向（如"适配器-无法充电"），可给出极简方向性线索（从质量管理的角度，给出质量问题的探索方向，要求精简专业），避免越界推断。
    - 对于无理由退机的分类无需给出建议，直接忽略即可

## 限制:
- 不得计算/重算占比：不得基于频次推导占比，不得改写任何单项占比。
- 零幻觉：不添加输入表格之外的类别、原因或数据。
- 保留原词：引用分类名时保持与输入一致（除去多余空格）。
- 风格与数值：中文为主；百分比以输入为准，展示到2位小数（如输入非2位小数，原样输出或四舍五入但需注明）。
- 输出必须按照规定的格式和要求进行组织，不能偏离框架要求。
"""
        
        return {
            "role": "user",
            "content": prompt
        }
    
    def generate_summary(self, category_df: pd.DataFrame, timeout: int = 60,
                        top_n: int = 3, coverage_threshold: float = 80.0,
                        focus_threshold: float = 10.0) -> str:
        """
        生成LLM摘要
        
        Args:
            category_df: 分类数据DataFrame
            timeout: 超时时间
            top_n: TopN参数
            coverage_threshold: 覆盖度阈值
            focus_threshold: 重点拦截阈值
            
        Returns:
            生成的摘要文本
            
        Raises:
            LLMGenerationError: 生成失败时抛出
        """
        category_rows = self.dataframe_to_category_rows(category_df)
        if not category_rows:
            raise LLMGenerationError("分类数据为空，无法生成LLM摘要")
        
        message = self.build_prompt(category_rows, top_n, coverage_threshold, focus_threshold)
        return self.call_api([message], timeout)
    
    def analyze_top_issue(self, issue_name: str, count: int, percentage: float,
                         model_dist: pd.DataFrame, timeout: int = 60) -> str:
        """
        分析Top Issue，从质量管理角度生成分析报告
        
        Args:
            issue_name: Issue名称
            count: 数量
            percentage: 占比
            model_dist: 机型分布DataFrame
            timeout: 超时时间
            
        Returns:
            LLM生成的分析文本
            
        Raises:
            LLMGenerationError: 生成失败时抛出
        """
        # 构建Prompt
        prompt = self._build_top_issue_prompt(issue_name, count, percentage, model_dist)
        
        message = {
            "role": "user",
            "content": prompt
        }
        
        return self.call_api([message], timeout)
    
    def _build_top_issue_prompt(self, issue_name: str, count: int, 
                                percentage: float, model_dist: pd.DataFrame) -> str:
        """
        构建Top Issue分析的Prompt
        
        Args:
            issue_name: Issue名称
            count: 数量
            percentage: 占比
            model_dist: 机型分布DataFrame
            
        Returns:
            Prompt文本
        """
        # 准备机型分布数据
        model_info = []
        for idx, row in model_dist.head(10).iterrows():
            model_info.append(
                f"  - {row['机型名称']}: {row['数量']}台 ({row['占比(%)']:.1f}%)"
            )
        model_str = "\n".join(model_info)
        
        prompt = f"""你是一位专业的质量管理专家。请基于以下数据，从质量管理角度对该问题进行深入分析。

问题信息：
- 问题分类：{issue_name}
- 影响数量：{count}台设备
- 占比：{percentage:.2f}%

机型分布情况（Top 10）：
{model_str}

请按照以下结构进行分析，每个部分不超过3-4句话：

【问题特征】
描述该问题的主要表现形式和特点。

【影响范围】
评估问题的严重程度和影响面。

【机型关联分析】
分析问题在不同机型上的分布特征，是否存在机型集中度，这说明了什么。

【可能原因】
从质量管理角度，列出3个可能导致该问题的原因（产品设计、供应链、制造工艺、使用环境等角度）。

【改进建议】
给出3条具体的质量改进建议。

要求：
1. 语言专业、简洁、直接
2. 基于数据进行客观分析
3. 避免模糊表述，给出可执行的建议
4. 总字数控制在300字以内"""

        return prompt
    
    @staticmethod
    def get_fallback_text(clean_model: str, suffix: str, total_records: int) -> str:
        """
        获取降级方案的默认文本
        
        Args:
            clean_model: 机型名称
            suffix: 类型后缀
            total_records: 记录总数
            
        Returns:
            默认文本
        """
        return (
            "核心观点（Human-Readable Core Insights）\n"
            f"- 样本：{clean_model}{suffix}共 {total_records} 条，暂未能生成自动化摘要。\n"
            "- 暂未获取模型结论，建议人工复核分类表。"
        )

