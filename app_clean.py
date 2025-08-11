from flask import Flask, request, jsonify, send_file, render_template_string
import os
import pandas as pd
import numpy as np
from datetime import datetime
import holidays
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import uuid
import traceback
import re

app = Flask(__name__)

# 添加CORS支持
@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
    return response

# 配置
UPLOAD_FOLDER = './uploads'
PROCESSED_FOLDER = './processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

@app.route('/')
def index():
    return "打卡数据处理系统 - 正在修复中..."

@app.route('/api/status')
def status():
    return jsonify({'status': 'running'})

if __name__ == '__main__':
    print("🚀 启动打卡数据处理系统...")
    print("📱 访问地址: http://localhost:8080")
    app.run(host='0.0.0.0', port=8080, debug=True) 