import os

# 基础配置
UPLOAD_FOLDER = './uploads'
PROCESSED_FOLDER = './processed'
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB

# 创建必要的目录
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# 应用程序配置
class Config:
    UPLOAD_FOLDER = UPLOAD_FOLDER
    PROCESSED_FOLDER = PROCESSED_FOLDER
    MAX_CONTENT_LENGTH = MAX_CONTENT_LENGTH
    SECRET_KEY = 'your-secret-key-here' 