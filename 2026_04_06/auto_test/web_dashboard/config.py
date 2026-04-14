import os

# 基础目录（项目根目录）
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 日志目录（与现有框架一致）
LOG_DIR = os.path.join(BASE_DIR, 'log')

# 如果 LOG_DIR 不存在则创建
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# 其他配置
DEBUG = True
SECRET_KEY = 'your-secret-key'