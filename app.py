from flask import Flask, render_template, send_file
from config import Config
from routes.api import api

app = Flask(__name__)
app.config.from_object(Config)

# 注册蓝图
app.register_blueprint(api, url_prefix='/api')

# 添加CORS支持
@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
    return response

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/favicon.ico')
def favicon():
    svg_icon = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
        <circle cx="50" cy="50" r="45" fill="#667eea" stroke="#5a6fd8" stroke-width="2"/>
        <text x="50" y="65" text-anchor="middle" font-size="40" fill="white">🕐</text>
    </svg>'''
    return svg_icon, 200, {'Content-Type': 'image/svg+xml'}

if __name__ == '__main__':
    print("🚀 启动模块化打卡数据处理系统...")
    print("📱 访问地址: http://localhost:8080")
    print("✨ 包含完整的错误检测、高亮标记和详细报告功能")
    print("🔄 新增：支持上传修改后的错误表格重新处理")
    print("🏗️ 架构：模块化设计，易于维护和扩展")
    app.run(host='0.0.0.0', port=8080, debug=True)