# -*- coding: utf-8 -*-
"""Flask应用"""
from flask import Flask
from pathlib import Path
import sys

sys.path.append(str(Path(__file__).parent.parent))

def create_app():
    """创建Flask应用"""
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'qcr-analysis-v4-secret-key'
    app.config['UPLOAD_FOLDER'] = Path(__file__).parent / 'static' / 'uploads'
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
    app.config['UPLOAD_FOLDER'].mkdir(parents=True, exist_ok=True)
    
    from .routes import register_routes
    register_routes(app)
    
    return app

