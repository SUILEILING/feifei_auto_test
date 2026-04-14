import os
import json
from flask import Flask, render_template, request, jsonify
from datetime import datetime
from config import LOG_DIR, DEBUG

app = Flask(__name__)
app.config.from_object('config')

def get_all_executions():
    """遍历 log 目录，返回所有执行记录的信息"""
    executions = []
    if not os.path.exists(LOG_DIR):
        return executions

    for dir_name in os.listdir(LOG_DIR):
        dir_path = os.path.join(LOG_DIR, dir_name)
        if not os.path.isdir(dir_path):
            continue

        # 查找 JSON 结果文件
        json_files = [f for f in os.listdir(dir_path) if f.endswith('.json')]
        if not json_files:
            continue

        # 取最新的 JSON 文件（如果有多个）
        json_file = max(json_files, key=lambda f: os.path.getmtime(os.path.join(dir_path, f)))
        json_path = os.path.join(dir_path, json_file)

        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # 如果是列表，取第一个（通常只有一个）
            if isinstance(data, list) and data:
                data = data[0]

            # 提取基本信息
            exec_info = {
                'dir': dir_name,
                'json_file': json_file,
                'timestamp': data.get('start_time', data.get('timestamp', '')),
                'script': data.get('script_name', data.get('file', '未知脚本')),
                'device': data.get('device', {}).get('address', 'N/A'),
                'parameters': data.get('parameters', {}),
                'status': '完成' if data.get('overall_success') else '失败',
                'duration': data.get('execution_time', 0),
                'loop_count': data.get('loop_count', 1)
            }
            executions.append(exec_info)
        except Exception as e:
            print(f"读取 {json_path} 失败: {e}")

    # 按时间倒序排列
    executions.sort(key=lambda x: x['timestamp'], reverse=True)
    return executions

def load_execution_data(dir_name, json_file):
    """加载指定执行的详细数据"""
    json_path = os.path.join(LOG_DIR, dir_name, json_file)
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if isinstance(data, list) and data:
        data = data[0]
    return data

@app.route('/')
def index():
    executions = get_all_executions()
    return render_template('index.html', executions=executions)

@app.route('/execution/<dir_name>/<json_file>')
def detail(dir_name, json_file):
    data = load_execution_data(dir_name, json_file)
    return render_template('detail.html', data=data, dir_name=dir_name, json_file=json_file)

@app.route('/api/execution/<dir_name>/<json_file>')
def api_detail(dir_name, json_file):
    data = load_execution_data(dir_name, json_file)
    return jsonify(data)

if __name__ == '__main__':
    app.run(debug=DEBUG, host='0.0.0.0', port=5000)