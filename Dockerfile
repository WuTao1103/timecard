
# Dockerfile - 单容器版本
FROM python:3.9-slim

# 设置工作目录
WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    curl \
    && rm -rf /var/lib/apt/lists/*

# 复制requirements文件
COPY requirements.txt .

# 安装Python依赖
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码和所有必要文件
COPY app.py .
COPY config.py .
COPY processors/ ./processors/
COPY utils/ ./utils/
COPY routes/ ./routes/
COPY templates/ ./templates/

# 创建必要的目录
RUN mkdir -p /app/uploads /app/processed /app/logs

# 暴露端口
EXPOSE 811

# 健康检查
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:811/api/status || exit 1

# 启动命令
CMD ["python", "app.py"]
