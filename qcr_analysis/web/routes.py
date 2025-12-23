# -*- coding: utf-8 -*-
"""路由定义"""
from flask import render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import date
import sys

sys.path.append(str(Path(__file__).parent.parent))
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

def register_routes(app):
    @app.route('/')
    def index():
        return render_template('index.html')
    
    @app.route('/weekly')
    def weekly_form():
        return render_template('weekly_form.html')
    
    @app.route('/top_issue')
    def top_issue_form():
        return render_template('top_issue_form.html')
    
    @app.route('/top_model')
    def top_model_form():
        return render_template('top_model_form.html')
    
    @app.route('/api/analyze/weekly', methods=['POST'])
    def analyze_weekly():
        try:
            data_file = request.files.get('data_file')
            mtm_file = request.files.get('mtm_file')
            
            if not data_file or not mtm_file:
                return jsonify({'error': '请上传数据文件和MTM文件'}), 400
            
            data_path = _save_file(data_file, app)
            mtm_path = _save_file(mtm_file, app)
            
            start_date = _parse_date(request.form.get('start_date'))
            end_date = _parse_date(request.form.get('end_date'))
            
            # 输出目录：用户指定或默认
            custom_output = request.form.get('output_dir')
            if custom_output and custom_output.strip():
                output_dir = Path(custom_output.strip())
            else:
                output_dir = app.config['UPLOAD_FOLDER'] / 'results' / 'weekly'
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # LLM配置
            use_llm = request.form.get('use_llm') == 'true'
            llm_config = None
            if use_llm:
                from config import KIMI_API_KEY, KIMI_API_URL, KIMI_MODEL
                llm_config = {
                    'api_key': KIMI_API_KEY,
                    'api_url': KIMI_API_URL,
                    'model': KIMI_MODEL,
                    'timeout': int(request.form.get('llm_timeout', 60)),
                    'top_n': int(request.form.get('llm_top_n', 3)),
                    'coverage': float(request.form.get('llm_coverage', 80)),
                    'focus': float(request.form.get('llm_focus', 10))
                }
            
            results = run_weekly_analysis(
                data_source=str(data_path),
                mtm_file=str(mtm_path),
                output_dir=str(output_dir),
                start_date=start_date,
                end_date=end_date,
                filter_unmapped=request.form.get('filter_unmapped') == 'true',
                use_llm=use_llm,
                llm_config=llm_config
            )
            if not results:
                return jsonify({'error': '数据为空或缺少必需列，无法生成周报'}), 400
            
            # 生成PPT（如果需要）
            ppt_path = None
            ppt_download = None
            if request.form.get('generate_ppt') == 'true':
                from services.weekly_analysis import WeeklyAnalysisService
                service = WeeklyAnalysisService(output_dir)
                service.results = results  # 注入结果
                payload = service.get_ppt_payload()
                
                batch_name = request.form.get('batch_name', '2024-2025')
                template_path = request.form.get('ppt_template')
                
                ppt_path = generate_weekly_report(
                    payload=payload,
                    output_dir=str(output_dir),
                    batch_name=batch_name,
                    template_path=template_path,
                    use_llm=use_llm,
                    llm_config=llm_config
                )
                # 如果是默认目录，生成下载URL；否则用户需通过"打开输出目录"访问
                if not custom_output or not custom_output.strip():
                    try:
                        rel_path = ppt_path.relative_to(app.config['UPLOAD_FOLDER'])
                        ppt_download = f"/download/{rel_path.as_posix()}"
                    except Exception:
                        ppt_download = None
                else:
                    # 自定义目录，提示用户打开输出目录查看
                    ppt_download = None
            
            return jsonify({
                'success': True,
                'total_records': len(results['total_df']),
                'records_7d': len(results['df_7d']),
                'records_non_7d': len(results['df_non_7d']),
                'output_dir': str(output_dir),
                'ppt_path': str(ppt_path) if ppt_path else None,
                'ppt_download_url': ppt_download
            })
        except Exception as e:
            import traceback
            traceback.print_exc()
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/analyze/top_issue', methods=['POST'])
    def analyze_top_issue():
        try:
            data_file = request.files.get('data_file')
            mtm_file = request.files.get('mtm_file')
            
            if not data_file or not mtm_file:
                return jsonify({'error': '请上传数据文件和MTM文件'}), 400
            
            data_path = _save_file(data_file, app)
            mtm_path = _save_file(mtm_file, app)
            
            data_manager = DataManager()
            df = data_manager.read_excel(str(data_path))
            
            start_date = _parse_date(request.form.get('start_date'))
            end_date = _parse_date(request.form.get('end_date'))
            if start_date or end_date:
                df = data_manager.filter_by_date_range(df, start_date, end_date)
            
            mtm_manager = MTMManager(Path(mtm_path))
            df = mtm_manager.map_dataframe(df)
            
            if request.form.get('filter_unmapped') == 'true':
                df = data_manager.filter_unmapped_mtm(df)
            
            # 输出目录：用户指定或默认
            custom_output = request.form.get('output_dir')
            if custom_output and custom_output.strip():
                output_dir = Path(custom_output.strip())
            else:
                output_dir = app.config['UPLOAD_FOLDER'] / 'results' / 'top_issue'
            output_dir.mkdir(parents=True, exist_ok=True)
            
            top_n = int(request.form.get('top_n', 10))
            use_llm = request.form.get('use_llm') == 'true'
            llm_config = None
            if use_llm:
                from config import KIMI_API_KEY, KIMI_API_URL, KIMI_MODEL
                llm_config = {
                    'api_key': KIMI_API_KEY,
                    'api_url': KIMI_API_URL,
                    'model': KIMI_MODEL,
                    'timeout': int(request.form.get('llm_timeout', 60)),
                    'top_n': int(request.form.get('llm_top_n', 3)),
                    'coverage': float(request.form.get('llm_coverage', 80)),
                    'focus': float(request.form.get('llm_focus', 10))
                }
            results = run_top_issue_analysis(df, str(output_dir), top_n, use_llm=use_llm, llm_config=llm_config)
            if not results:
                return jsonify({'error': '数据为空或缺少必需列，无法生成Top Issue分析'}), 400
            
            # 生成PPT（如果需要）
            ppt_path = None
            ppt_download = None
            if request.form.get('generate_ppt') == 'true':
                from services.top_issue_analysis import TopIssueAnalysisService
                service = TopIssueAnalysisService(output_dir)
                service.results = results  # 注入结果
                payload = service.get_ppt_payload()
                
                batch_name = request.form.get('batch_name', '2024-2025')
                template_path = request.form.get('ppt_template')
                
                ppt_path = generate_top_issue_report(
                    payload=payload,
                    output_dir=str(output_dir),
                    batch_name=batch_name,
                    template_path=template_path,
                    use_llm=use_llm,
                    llm_config=llm_config
                )
                # 如果是默认目录，生成下载URL；否则用户需通过"打开输出目录"访问
                if not custom_output or not custom_output.strip():
                    try:
                        rel_path = ppt_path.relative_to(app.config['UPLOAD_FOLDER'])
                        ppt_download = f"/download/{rel_path.as_posix()}"
                    except Exception:
                        ppt_download = None
                else:
                    # 自定义目录，提示用户打开输出目录查看
                    ppt_download = None
            
            return jsonify({
                'success': True,
                'total_records': results['total_records'],
                'top_n': results['top_n'],
                'output_dir': str(output_dir),
                'ppt_path': str(ppt_path) if ppt_path else None,
                'ppt_download_url': ppt_download
            })
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    
    @app.route('/api/analyze/top_model', methods=['POST'])
    def analyze_top_model():
        try:
            data_file = request.files.get('data_file')
            mtm_file = request.files.get('mtm_file')
            
            if not data_file or not mtm_file:
                return jsonify({'error': '请上传数据文件和MTM文件'}), 400
            
            data_path = _save_file(data_file, app)
            mtm_path = _save_file(mtm_file, app)
            
            data_manager = DataManager()
            df = data_manager.read_excel(str(data_path))
            
            start_date = _parse_date(request.form.get('start_date'))
            end_date = _parse_date(request.form.get('end_date'))
            if start_date or end_date:
                df = data_manager.filter_by_date_range(df, start_date, end_date)
            
            mtm_manager = MTMManager(Path(mtm_path))
            df = mtm_manager.map_dataframe(df)
            
            if request.form.get('filter_unmapped') == 'true':
                df = data_manager.filter_unmapped_mtm(df)
            
            # 输出目录：用户指定或默认
            custom_output = request.form.get('output_dir')
            if custom_output and custom_output.strip():
                output_dir = Path(custom_output.strip())
            else:
                output_dir = app.config['UPLOAD_FOLDER'] / 'results' / 'top_model'
            output_dir.mkdir(parents=True, exist_ok=True)
            
            top_n = int(request.form.get('top_n', 15))
            use_llm = request.form.get('use_llm') == 'true'
            llm_config = None
            if use_llm:
                from config import KIMI_API_KEY, KIMI_API_URL, KIMI_MODEL
                llm_config = {
                    'api_key': KIMI_API_KEY,
                    'api_url': KIMI_API_URL,
                    'model': KIMI_MODEL,
                    'timeout': int(request.form.get('llm_timeout', 60)),
                    'top_n': int(request.form.get('llm_top_n', 3)),
                    'coverage': float(request.form.get('llm_coverage', 80)),
                    'focus': float(request.form.get('llm_focus', 10))
                }
            results = run_top_model_analysis(df, str(output_dir), top_n, use_llm=use_llm, llm_config=llm_config)
            if not results:
                return jsonify({'error': '数据为空或缺少必需列，无法生成Top Model分析'}), 400
            
            # 生成PPT（如果需要）
            ppt_path = None
            ppt_download = None
            if request.form.get('generate_ppt') == 'true':
                from services.top_model_analysis import TopModelAnalysisService
                service = TopModelAnalysisService(output_dir)
                service.results = results  # 注入结果
                payload = service.get_ppt_payload()
                
                batch_name = request.form.get('batch_name', '2024-2025')
                template_path = request.form.get('ppt_template')
                
                ppt_path = generate_top_model_report(
                    payload=payload,
                    output_dir=str(output_dir),
                    batch_name=batch_name,
                    template_path=template_path,
                    use_llm=use_llm,
                    llm_config=llm_config
                )
                # 如果是默认目录，生成下载URL；否则用户需通过"打开输出目录"访问
                if not custom_output or not custom_output.strip():
                    try:
                        rel_path = ppt_path.relative_to(app.config['UPLOAD_FOLDER'])
                        ppt_download = f"/download/{rel_path.as_posix()}"
                    except Exception:
                        ppt_download = None
                else:
                    # 自定义目录，提示用户打开输出目录查看
                    ppt_download = None
            
            return jsonify({
                'success': True,
                'total_records': results['total_records'],
                'total_models': results['total_models'],
                'top_n': results['top_n'],
                'output_dir': str(output_dir),
                'ppt_path': str(ppt_path) if ppt_path else None,
                'ppt_download_url': ppt_download
            })
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    @app.route('/download/<path:filepath>')
    def download_file(filepath):
        """下载文件"""
        try:
            # 构建完整文件路径
            full_path = app.config['UPLOAD_FOLDER'] / 'results' / filepath
            print(f"尝试下载文件: {full_path}")
            
            if full_path.exists():
                return send_file(str(full_path), as_attachment=True)
            else:
                print(f"文件不存在: {full_path}")
                return jsonify({'error': f'文件不存在: {filepath}'}), 404
        except Exception as e:
            print(f"下载文件失败: {e}")
            return jsonify({'error': str(e)}), 500

def _save_file(file, app):
    filename = secure_filename(file.filename)
    filepath = app.config['UPLOAD_FOLDER'] / filename
    file.save(str(filepath))
    return filepath

def _parse_date(date_str):
    if not date_str:
        return None
    try:
        parts = date_str.split('-')
        return date(int(parts[0]), int(parts[1]), int(parts[2]))
    except:
        return None

