import re
from pathlib import Path
from datetime import datetime
import os
import sys
import tempfile
import atexit

# ==============================================
# SCV 转换核心函数
# ==============================================

def parse_python_file(file_path):
    """
    解析 Python 文件中的 TEST_STEPS 列表
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 查找 TEST_STEPS 列表的开始位置
    test_steps_pattern = r'TEST_STEPS\s*:\s*List\[TestStep\]\s*=\s*\['  # 修改了这里
    start_match = re.search(test_steps_pattern, content)
    
    if not start_match:
        # 尝试另一种格式
        test_steps_pattern = r'TEST_STEPS\s*=\s*\['  # 修改了这里
        start_match = re.search(test_steps_pattern, content)
    
    if not start_match:
        raise ValueError("在文件中找不到 TEST_STEPS 定义")
    
    start_pos = start_match.end()
    
    # 找到对应的结束括号
    bracket_count = 1
    end_pos = start_pos
    
    for i in range(start_pos, len(content)):
        if content[i] == '[':
            bracket_count += 1
        elif content[i] == ']':
            bracket_count -= 1
            if bracket_count == 0:
                end_pos = i
                break
    
    if bracket_count != 0:
        raise ValueError("TEST_STEPS 列表未正确闭合")
    
    # 提取列表内容
    list_content = content[start_pos:end_pos].strip()
    
    # 分割各个 TestStep 构造函数
    test_steps = []
    step_pattern = r'TestStep\s*\('
    
    # 找到所有 TestStep( 的位置
    matches = list(re.finditer(step_pattern, list_content))
    
    for idx, match in enumerate(matches):
        start_idx = match.start()
        
        # 找到当前 TestStep 的结束位置
        paren_count = 0
        in_string = False
        string_char = None
        end_idx = start_idx
        
        for i in range(start_idx, len(list_content)):
            char = list_content[i]
            
            # 处理字符串
            if char in ('"', "'") and (i == 0 or list_content[i-1] != '\\'):
                if not in_string:
                    in_string = True
                    string_char = char
                elif string_char == char:
                    in_string = False
                    string_char = None
            
            # 如果不是在字符串中，统计括号
            if not in_string:
                if char == '(':
                    paren_count += 1
                elif char == ')':
                    paren_count -= 1
                    if paren_count == 0:
                        end_idx = i + 1
                        break
        
        # 提取 TestStep 的字符串表示
        step_str = list_content[start_idx:end_idx]
        
        # 解析 TestStep 参数
        try:
            test_step = parse_teststep_string(step_str)
            test_steps.append(test_step)
        except Exception as e:
            print(f"警告: 解析第 {idx+1} 个 TestStep 失败: {e}")
            continue
    
    return test_steps

def parse_teststep_string(step_str):
    """
    从 TestStep(...) 字符串解析出 TestStep 对象
    """
    # 提取括号内的内容
    match = re.search(r'TestStep\s*\((.*)\)', step_str, re.DOTALL)
    if not match:
        raise ValueError("无效的 TestStep 格式")
    
    args_str = match.group(1).strip()
    
    # 初始化默认值
    command = ""
    expected_result = ""
    description = ""
    step_type = "NORMAL"
    run_status = "Pass"
    timeout_ms = 10000
    loop_enable = False
    loop_times = 1
    loop_sleep_ms = 0
    loop_expected = ""
    loop_stop_on_fail = False
    
    # 解析参数
    lines = args_str.split('\n')
    for line in lines:
        line = line.strip().rstrip(',')
        if not line:
            continue
        
        # 解析键值对
        if '=' in line:
            key, value = line.split('=', 1)
            key = key.strip()
            value = value.strip()
            
            # 处理字符串值（去掉引号）
            if value.startswith(('"', "'")) and value.endswith(('"', "'")):
                value = value[1:-1]
            
            if key == 'command':
                command = value
            elif key == 'expected_result':
                expected_result = value
            elif key == 'description':
                description = value
            elif key == 'run_status':
                run_status = value
            elif key == 'step_type':
                # 处理 StepType.CONFIGURE 格式
                if 'StepType.' in value:
                    step_type = value.split('.')[1]
                else:
                    step_type = value.strip('"\' ')
            elif key == 'timeout_ms':
                timeout_ms = int(value)
            elif key == 'loop_config':
                # 解析 LoopConfig
                loop_data = parse_loopconfig(value)
                loop_enable = loop_data.get('enable', False)
                loop_times = loop_data.get('times', 1)
                loop_sleep_ms = loop_data.get('sleep_ms', 0)
                loop_expected = loop_data.get('expected_result', '')
                loop_stop_on_fail = loop_data.get('stop_on_fail', False)
    
    # 如果 command 不为空，根据命令确定步骤类型
    if command:
        cmd_upper = command.upper()
        if cmd_upper.startswith("CALL:"):
            step_type = "CALL"
        elif cmd_upper.startswith("SLEEP"):
            step_type = "SLEEP"
        elif "?" in cmd_upper:
            step_type = "QUERY"
        elif cmd_upper.startswith("CONFIGURE:"):
            step_type = "CONFIGURE"
    
    return {
        'command': command,
        'expected_result': expected_result,
        'description': description,
        'step_type': step_type,
        'run_status': run_status,
        'timeout_ms': timeout_ms,
        'loop_enable': loop_enable,
        'loop_times': loop_times,
        'loop_sleep_ms': loop_sleep_ms,
        'loop_expected': loop_expected,
        'loop_stop_on_fail': loop_stop_on_fail
    }

def parse_loopconfig(loop_str):
    """
    解析 LoopConfig(...) 字符串
    """
    loop_data = {
        'enable': False,
        'times': 1,
        'sleep_ms': 0,
        'expected_result': '',
        'stop_on_fail': False
    }
    
    # 提取括号内的内容
    match = re.search(r'LoopConfig\s*\((.*)\)', loop_str, re.DOTALL)
    if not match:
        return loop_data
    
    args_str = match.group(1).strip()
    
    # 解析参数
    lines = args_str.split('\n')
    for line in lines:
        line = line.strip().rstrip(',')
        if not line:
            continue
        
        if '=' in line:
            key, value = line.split('=', 1)
            key = key.strip()
            value = value.strip()
            
            # 处理字符串值
            if value.startswith(('"', "'")) and value.endswith(('"', "'")):
                value = value[1:-1]
            
            if key == 'enable':
                loop_data['enable'] = value.lower() == 'true'
            elif key == 'times':
                loop_data['times'] = int(value)
            elif key == 'sleep_ms':
                loop_data['sleep_ms'] = int(value)
            elif key == 'expected_result':
                loop_data['expected_result'] = value
            elif key == 'stop_on_fail':
                loop_data['stop_on_fail'] = value.lower() == 'true'
    
    return loop_data

def convert_to_scv(test_steps, filename="test_steps.scv"):
    """
    将测试步骤转换为 SCV 文件（标准格式）- 简化版
    格式：命令（仅命令，不要注释）
    """
    scv_content = []
    
    # ⭐⭐⭐ 只添加命令内容，不要注释 ⭐⭐⭐
    # 标准SCV格式应该是每行一个命令，不要包含任何注释
    
    for i, step in enumerate(test_steps, 1):
        # ⭐⭐⭐ 构建简单格式：仅命令 ⭐⭐⭐
        command = step['command'].strip()
        
        # ⭐⭐⭐ 修复：移除命令中的注释部分 ⭐⭐⭐
        # 有些命令可能已经包含注释，需要清除
        if '#' in command:
            command = command.split('#')[0].strip()
        
        # 只添加非空命令
        if command:
            scv_content.append(command)
    
    # 写入文件
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write('\n'.join(scv_content))
        
        print(f"✅ 成功转换 {len(test_steps)} 个步骤到 {filename}")
        
        # 显示预览
        print(f"\n📋 SCV 文件预览（前10行内容）:")
        print("-" * 100)
        with open(filename, 'r', encoding='utf-8') as f:
            for i, line in enumerate(f):
                if i < 10:  # 显示前10行
                    print(f"{i+1:3d}: {line.rstrip()}")
                else:
                    break
        print("-" * 100)
        
        return True
        
    except Exception as e:
        print(f"❌ 转换失败: {e}")
        return False

def convert_to_scv_table(test_steps, filename="test_steps_table.scv"):
    """
    将测试步骤转换为表格格式的 SCV 文件
    """
    scv_content = []
    
    # 文件头
    scv_content.append("=" * 100)
    scv_content.append("测试步骤配置表")
    scv_content.append("=" * 100)
    scv_content.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    scv_content.append(f"总步骤数: {len(test_steps)}")
    scv_content.append("")
    
    # 表格格式
    for i, step in enumerate(test_steps, 1):
        scv_content.append(f"步骤 {i:03d}")
        scv_content.append(f"命令:     {step['command']}")
        
        if step['description']:
            scv_content.append(f"描述:     {step['description']}")
        
        scv_content.append(f"类型:     {step['step_type']}")
        scv_content.append(f"状态:     {step['run_status']}")
        scv_content.append(f"超时:     {step['timeout_ms']}ms")
        
        if step['expected_result']:
            scv_content.append(f"预期结果: {step['expected_result']}")
        
        if step['loop_enable']:
            scv_content.append(f"循环配置: 启用, {step['loop_times']}次")
            if step['loop_sleep_ms'] > 0:
                scv_content.append(f"循环间隔: {step['loop_sleep_ms']}ms")
            if step['loop_expected']:
                scv_content.append(f"循环预期: {step['loop_expected']}")
            if step['loop_stop_on_fail']:
                scv_content.append(f"失败停止: 是")
        
        scv_content.append("-" * 50)
    
    # 写入文件
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write('\n'.join(scv_content))
        
        print(f"✅ 成功创建表格格式 SCV: {filename}")
        
        # 显示预览
        print(f"\n📋 表格格式 SCV 文件预览（前10行）:")
        print("-" * 80)
        with open(filename, 'r', encoding='utf-8') as f:
            content_lines = f.readlines()
            for i in range(min(20, len(content_lines))):  # 显示前20行
                print(content_lines[i].rstrip())
        print("-" * 80)
        
        return True
        
    except Exception as e:
        print(f"❌ 创建表格格式 SCV 失败: {e}")
        return False

# ==============================================
# 动态SCV生成函数（自动化系统调用）
# ==============================================

# 全局变量，用于跟踪临时文件
_temp_scv_files = []

def generate_dynamic_scv_from_python(python_file_path, output_format="standard", 
                                    output_filename=None, auto_cleanup=False):
    """
    从Python脚本生成动态SCV文件
    
    参数:
        python_file_path: Python脚本路径
        output_format: 输出格式 "standard" 或 "table"
        output_filename: 输出文件名（None则自动生成）
        auto_cleanup: 是否在程序结束时自动清理SCV文件
    """
    try:
        print(f"🔄 开始从Python脚本生成动态SCV文件...")
        print(f"📁 源文件: {python_file_path}")
        print(f"📄 输出格式: {output_format}")
        
        # ⭐⭐⭐ 强制使用标准格式 ⭐⭐⭐
        if output_format != "standard":
            print(f"⚠️  注意: 自动将格式从 '{output_format}' 改为 'standard'")
            print(f"    标准格式更适合自动化执行")
            output_format = "standard"
        
        # 解析Python文件
        test_steps = parse_python_file(python_file_path)
        
        if not test_steps:
            print("❌ 未找到任何测试步骤")
            return None
        
        print(f"✅ 成功解析 {len(test_steps)} 个测试步骤")
        
        # 确定输出文件名
        if not output_filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if output_format == "table":
                output_filename = f"dynamic_test_steps_table_{timestamp}.scv"
            else:
                output_filename = f"dynamic_test_steps_{timestamp}.scv"
        
        # 如果是自动化模式且需要自动清理，使用临时文件
        if auto_cleanup:
            # 创建临时目录中的文件
            temp_dir = tempfile.gettempdir()
            output_filename = os.path.join(temp_dir, output_filename)
            print(f"📂 使用临时文件: {output_filename}")
        
        # 生成SCV文件
        if output_format == "table":
            success = convert_to_scv_table(test_steps, output_filename)
        else:
            success = convert_to_scv(test_steps, output_filename)
        
        if success:
            print(f"✅ 动态SCV文件生成成功: {output_filename}")
            
            # 显示生成的SCV文件内容预览
            print(f"\n🔍 生成的SCV文件内容预览:")
            print("-" * 80)
            try:
                with open(output_filename, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    for i, line in enumerate(lines[:10]):  # 显示前10行
                        print(f"{i+1:3d}: {line.rstrip()}")
                    if len(lines) > 10:
                        print(f"... 还有 {len(lines)-10} 行")
            except Exception as e:
                print(f"⚠️  无法读取文件内容: {e}")
            print("-" * 80)
            
            # 如果是自动化模式，注册清理函数
            if auto_cleanup and os.path.exists(output_filename):
                _temp_scv_files.append(output_filename)
                atexit.register(_cleanup_temp_files)
                print(f"🗑️  文件将在程序结束时自动删除")
            
            return output_filename
        else:
            print(f"❌ 动态SCV文件生成失败")
            return None
            
    except Exception as e:
        print(f"❌ 动态生成SCV文件时出错: {e}")
        import traceback
        traceback.print_exc()
        return None

def _cleanup_temp_files():
    """清理临时SCV文件"""
    for file_path in _temp_scv_files:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"🗑️  已清理临时文件: {os.path.basename(file_path)}")
        except Exception as e:
            print(f"⚠️  清理文件失败 {file_path}: {e}")

# ==============================================
# 单独执行时的处理
# ==============================================

def _run_as_standalone(python_file_path, output_format="standard"):
    """
    单独执行时的处理
    
    参数:
        python_file_path: Python文件路径
        output_format: 输出格式 "standard" 或 "table"
    """
    print("📊 SCV 文件转换工具 - 独立执行")
    print("=" * 60)
    
    if not os.path.exists(python_file_path):
        print(f"❌ 错误: 文件不存在: {python_file_path}")
        print(f"当前工作目录: {os.getcwd()}")
        return False
    
    print(f"📁 解析文件: {python_file_path}")
    print(f"📄 输出格式: {output_format}")
    print("=" * 60)
    
    # 确保输出到 save_scv 目录
    save_scv_dir = "save_scv"
    if not os.path.exists(save_scv_dir):
        os.makedirs(save_scv_dir)
        print(f"📁 创建目录: {save_scv_dir}")
    
    # 生成输出文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if output_format == "table":
        output_filename = os.path.join(save_scv_dir, f"test_steps_table_{timestamp}.scv")
    else:
        output_filename = os.path.join(save_scv_dir, f"test_steps_{timestamp}.scv")
    
    # 生成SCV文件（不使用自动清理）
    output_file = generate_dynamic_scv_from_python(
        python_file_path, 
        output_format, 
        output_filename,
        auto_cleanup=False
    )
    
    if output_file:
        print(f"\n🎉 转换完成！")
        print(f"📁 生成的SCV文件: {output_file}")
        print(f"📁 文件大小: {os.path.getsize(output_file)} 字节")
        print(f"📁 绝对路径: {os.path.abspath(output_file)}")
        print("=" * 60)
        return True
    else:
        print(f"\n❌ 转换失败")
        return False

# ==============================================
# 模块化导入检查
# ==============================================
if __name__ == "__main__":
    # ⭐⭐⭐ 在这里直接指定文件路径 ⭐⭐⭐
    
    # 定义要转换的Python文件路径
    PYTHON_FILE = r"D:\feifei\2026_01_15\auto_test\save_scv\night_cycle_update_feifei_steps.py"
    
    # 定义输出格式（可选：standard 或 table）
    OUTPUT_FORMAT = "standard"  # ⭐⭐⭐ 强制使用标准格式 ⭐⭐⭐
    
    print("=" * 60)
    print("🚀 SCV 文件转换工具启动")
    print("=" * 60)
    
    # 直接执行转换
    success = _run_as_standalone(PYTHON_FILE, OUTPUT_FORMAT)
    
    if success:
        print("✅ 转换成功完成！")
    else:
        print("❌ 转换失败")
        sys.exit(1)
    
else:
    # 作为模块导入：供自动化系统调用
    # 提供模块化接口
    __all__ = [
        'parse_python_file',
        'parse_teststep_string',
        'parse_loopconfig',
        'convert_to_scv',
        'convert_to_scv_table',
        'generate_dynamic_scv_from_python'
    ]