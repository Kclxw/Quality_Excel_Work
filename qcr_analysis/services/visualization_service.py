# -*- coding: utf-8 -*-
"""可视化服务"""
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from pathlib import Path
from typing import Tuple

import sys
sys.path.append(str(Path(__file__).parent.parent))
from config import MATPLOTLIB_FONTS, CHART_STYLE

matplotlib.rcParams['font.family'] = MATPLOTLIB_FONTS
matplotlib.rcParams['axes.unicode_minus'] = False

class VisualizationService:
    """可视化服务：生成各类图表"""
    
    def __init__(self, output_dir: str or Path):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def generate_pie_chart(self, data, value_column, label_column, title, filename, figsize=(8,8)):
        """生成饼图"""
        plt.figure(figsize=figsize)
        plt.pie(data[value_column], labels=data[label_column], autopct="%1.1f%%")
        plt.title(title)
        plt.tight_layout()
        chart_path = self.output_dir / filename
        plt.savefig(chart_path, dpi=150)
        plt.close()
        return chart_path
    
    def generate_bar_chart(self, data, x_column, y_column, title, filename, figsize=(12,6), orientation='vertical'):
        """生成柱状图"""
        fig, ax = plt.subplots(figsize=figsize)
        if orientation == 'vertical':
            ax.bar(range(len(data)), data[y_column])
            ax.set_xticks(range(len(data)))
            ax.set_xticklabels(data[x_column], rotation=45, ha='right')
        else:
            ax.barh(range(len(data)), data[y_column])
            ax.set_yticks(range(len(data)))
            ax.set_yticklabels(data[x_column])
        ax.set_title(title)
        plt.tight_layout()
        chart_path = self.output_dir / filename
        plt.savefig(chart_path, dpi=150)
        plt.close()
        return chart_path

def create_visualization_service(output_dir):
    return VisualizationService(output_dir)

