from flask import Blueprint, request, jsonify, send_file
import os
import uuid
from processors.timecard_processor import TimecardProcessor

api = Blueprint('api', __name__)

def get_processor():
    """获取处理器实例"""
    from app import app
    return TimecardProcessor(app.config['UPLOAD_FOLDER'], app.config['PROCESSED_FOLDER'])

@api.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有选择文件'}), 400

    file = request.files['file']
    if file.filename == '' or not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请选择Excel文件'}), 400

    filename = str(uuid.uuid4()) + '_' + file.filename
    file_path = os.path.join(get_processor().upload_folder, filename)
    file.save(file_path)

    return jsonify({
        'success': True,
        'filename': filename,
        'original_name': file.filename
    })

@api.route('/upload/error', methods=['POST'])
def upload_error_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有选择文件'}), 400

    file = request.files['file']
    if file.filename == '' or not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请选择Excel文件'}), 400

    filename = str(uuid.uuid4()) + '_error_' + file.filename
    file_path = os.path.join(get_processor().processed_folder, filename)
    file.save(file_path)

    return jsonify({
        'success': True,
        'filename': filename,
        'original_name': file.filename
    })

@api.route('/process/step1', methods=['POST'])
def process_step1():
    data = request.json
    filename = data.get('filename')
    if not filename:
        return jsonify({'error': '缺少文件名'}), 400

    file_path = os.path.join(get_processor().upload_folder, filename)
    if not os.path.exists(file_path):
        return jsonify({'error': '文件不存在'}), 404

    processor = get_processor()
    result = processor.process_step1(file_path)
    return jsonify(result)

@api.route('/process/step2', methods=['POST'])
def process_step2():
    data = request.json
    error_filename = data.get('error_filename')
    time_range = data.get('time_range')

    if not error_filename or not time_range:
        return jsonify({'error': '缺少必要参数'}), 400

    error_file_path = os.path.join(get_processor().processed_folder, error_filename)
    if not os.path.exists(error_file_path):
        return jsonify({'error': '中间文件不存在'}), 404

    processor = get_processor()
    result = processor.process_step2(error_file_path, time_range)
    return jsonify(result)

@api.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(get_processor().processed_folder, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name=filename)
    return jsonify({'error': '文件不存在'}), 404

@api.route('/status')
def status():
    processor = get_processor()
    return jsonify({
        'status': 'running',
        'upload_folder': processor.upload_folder,
        'processed_folder': processor.processed_folder
    }) 