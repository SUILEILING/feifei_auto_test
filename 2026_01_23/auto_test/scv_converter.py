from lib.var import *

# ==============================================
# 数据类型定义（从原文件迁移过来）
# ==============================================

from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
from enum import Enum

class StepType(Enum):
    """步骤类型"""
    NORMAL = "Normal"
    CALL = "Call"
    SLEEP = "Sleep"
    QUERY = "Query"
    CONFIGURE = "Configure"


@dataclass
class LoopConfig:
    """循环配置"""
    enable: bool = False
    times: int = 1
    sleep_ms: int = 0
    expected_result: str = ""
    stop_on_fail: bool = False
    
    def to_dict(self) -> Dict:
        """转换为字典"""
        return {
            "enable": self.enable,
            "times": self.times,
            "sleep_ms": self.sleep_ms,
            "expected_result": self.expected_result,
            "stop_on_fail": self.stop_on_fail
        }


@dataclass
class TestStep:
    """测试步骤"""
    command: str  # 命令内容
    expected_result: str = ""  # 预期结果
    loop_config: LoopConfig = field(default_factory=LoopConfig)  # 循环配置
    step_type: StepType = StepType.NORMAL  # 步骤类型
    run_status: str = "Pass"  # 运行状态
    description: str = ""  # 步骤描述
    timeout_ms: int = 10000  # 超时时间(毫秒)
    
    def __post_init__(self):
        """初始化后自动确定步骤类型"""
        cmd = self.command.upper()
        if cmd.startswith("CALL:"):
            self.step_type = StepType.CALL
        elif cmd.startswith("SLEEP"):
            self.step_type = StepType.SLEEP
        elif "?" in cmd:
            self.step_type = StepType.QUERY
        elif cmd.startswith("CONFIGURE:"):
            self.step_type = StepType.CONFIGURE
    
    def to_dict(self) -> Dict:
        """转换为字典"""
        return {
            "command": self.command,
            "expected_result": self.expected_result,
            "loop_config": self.loop_config.to_dict(),
            "step_type": self.step_type.value,
            "run_status": self.run_status,
            "description": self.description,
            "timeout_ms": self.timeout_ms
        }
    
    def __str__(self) -> str:
        """字符串表示"""
        result = f"命令: {self.command}"
        if self.expected_result:
            result += f"\n  预期: {self.expected_result}"
        if self.loop_config.enable:
            result += f"\n  循环: {self.loop_config.times}次"
            if self.loop_config.sleep_ms > 0:
                result += f" (间隔{self.loop_config.sleep_ms}ms)"
            if self.loop_config.expected_result:
                result += f" -> {self.loop_config.expected_result}"
        if self.description:
            result += f"\n  描述: {self.description}"
        return result


# ==============================================
# 统计和分析函数
# ==============================================

def get_statistics(test_steps: List[TestStep]) -> Dict:
    """获取步骤统计信息"""
    stats = {
        "total_steps": len(test_steps),
        "step_types": {},
        "commands_with_loop": 0,
        "commands_with_expected": 0,
        "estimated_duration_ms": 0
    }
    
    # 统计步骤类型
    for step in test_steps:
        step_type = step.step_type.value
        stats["step_types"][step_type] = stats["step_types"].get(step_type, 0) + 1
        
        # 统计循环命令
        if step.loop_config.enable:
            stats["commands_with_loop"] += 1
        
        # 统计有预期结果的命令
        if step.expected_result or step.loop_config.expected_result:
            stats["commands_with_expected"] += 1
        
        # 估算总时长（仅统计SLEEP命令）
        if step.step_type == StepType.SLEEP:
            sleep_match = re.search(r'SLEEP\s+(\d+)', step.command.upper())
            if sleep_match:
                stats["estimated_duration_ms"] += int(sleep_match.group(1))
        
        # 加上循环等待时间
        if step.loop_config.enable:
            stats["estimated_duration_ms"] += step.loop_config.times * step.loop_config.sleep_ms
    
    stats["estimated_duration_sec"] = stats["estimated_duration_ms"] / 1000
    return stats


def print_statistics(test_steps: List[TestStep]) -> None:
    """打印统计信息"""
    stats = get_statistics(test_steps)
    
    print("=" * 60)
    print("📊 测试步骤统计")
    print("=" * 60)
    print(f"总步骤数: {stats['total_steps']}")
    print(f"预计总时长: {stats['estimated_duration_sec']:.1f}秒")
    print()
    print("步骤类型分布:")
    for step_type, count in stats["step_types"].items():
        print(f"  {step_type}: {count}")
    print()
    print(f"循环步骤: {stats['commands_with_loop']}")
    print(f"有预期结果的步骤: {stats['commands_with_expected']}")
    print("=" * 60)


def print_all_steps(test_steps: List[TestStep], show_details: bool = False) -> None:
    """打印所有步骤"""
    print("=" * 60)
    print("📋 所有测试步骤")
    print("=" * 60)

    for i, step in enumerate(test_steps, 1):
        print(f"\n[{i}] {step.command}")
        if show_details:
            if step.description:
                print(f"   描述: {step.description}")
            if step.expected_result:
                print(f"   预期: {step.expected_result}")
            if step.loop_config.enable:
                loop_info = f"   循环: {step.loop_config.times}次"
                if step.loop_config.sleep_ms > 0:
                    loop_info += f" (间隔{step.loop_config.sleep_ms}ms)"
                if step.loop_config.expected_result:
                    loop_info += f" -> {step.loop_config.expected_result}"
                print(loop_info)
            print(f"   类型: {step.step_type.value}")
            print(f"   状态: {step.run_status}")
    
    print("=" * 60)


# ==============================================
# SCV 转换核心函数
# ==============================================

def parse_python_file(file_path):
    """
    解析 Python 文件中的 TEST_STEPS 列表
    同时解析 SKIP_IN_NEXT_CYCLES
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 查找 TEST_STEPS 列表的开始位置
    test_steps_pattern = r'TEST_STEPS\s*:\s*List\[Dict\]\s*=\s*\['  
    start_match = re.search(test_steps_pattern, content)
    
    if not start_match:
        # 尝试另一种格式
        test_steps_pattern = r'TEST_STEPS\s*=\s*\[' 
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
    
    # 解析字典列表
    test_steps = parse_dict_list(list_content)
    
    return test_steps

def parse_dict_list(content: str) -> List[Dict]:
    """解析Python字典列表"""
    steps = []
    
    # 查找所有大括号
    brace_count = 0
    current_dict_start = -1
    in_string = False
    string_char = None
    
    for i, char in enumerate(content):
        # 处理字符串
        if char in ('"', "'") and (i == 0 or content[i-1] != '\\'):
            if not in_string:
                in_string = True
                string_char = char
            elif string_char == char:
                in_string = False
                string_char = None
        
        # 如果不是在字符串中，统计大括号
        if not in_string:
            if char == '{':
                if brace_count == 0:
                    current_dict_start = i
                brace_count += 1
            elif char == '}':
                brace_count -= 1
                if brace_count == 0 and current_dict_start != -1:
                    # 提取字典内容
                    dict_str = content[current_dict_start:i+1]
                    try:
                        # 使用eval解析字典（注意安全性）
                        dict_obj = eval(dict_str)
                        if isinstance(dict_obj, dict):
                            steps.append(dict_obj)
                    except:
                        # 尝试手动解析
                        dict_obj = parse_dict_string(dict_str)
                        if dict_obj:
                            steps.append(dict_obj)
                    current_dict_start = -1
    
    return steps

def parse_dict_string(dict_str: str) -> Dict:
    """解析字典字符串"""
    result = {}
    
    # 移除大括号
    dict_str = dict_str.strip()[1:-1].strip()
    
    # 分割键值对
    lines = dict_str.split('\n')
    for line in lines:
        line = line.strip().rstrip(',')
        if not line:
            continue
        
        if ':' in line:
            # 找到第一个冒号的位置
            colon_pos = line.find(':')
            key = line[:colon_pos].strip()
            value = line[colon_pos+1:].strip()
            
            # 清理键名
            if key.startswith('"') and key.endswith('"'):
                key = key[1:-1]
            elif key.startswith("'") and key.endswith("'"):
                key = key[1:-1]
            
            # 清理值
            if value.startswith('"') and value.endswith('"'):
                value = value[1:-1]
            elif value.startswith("'") and value.endswith("'"):
                value = value[1:-1]
            elif value.lower() == 'true':
                value = True
            elif value.lower() == 'false':
                value = False
            elif value.isdigit():
                value = int(value)
            
            result[key] = value
    
    return result

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
    
    for i, step in enumerate(test_steps, 1):
        command = step['command'].strip()
        
        if '#' in command:
            command = command.split('#')[0].strip()
        
        if command:
            scv_content.append(command)
    
    # 写入文件
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write('\n'.join(scv_content))
        
        print(f"✅ 成功转换 {len(test_steps)} 个步骤到 {filename}")
        
        return True
        
    except Exception as e:
        print(f"❌ 转换失败: {e}")
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
        output_filename: 输出文件名(None则自动生成)
        auto_cleanup: 是否在程序结束时自动清理SCV文件
    """
    try:
        print(f"🔄 开始从Python脚本生成动态SCV文件...")
        print(f"📁 源文件: {python_file_path}")
        print(f"📄 输出格式: {output_format}")
        
        # 强制使用标准格式
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
        
        success = convert_to_scv(test_steps, output_filename)
        
        if success:
            print(f"✅ 动态SCV文件生成成功: {output_filename}")
            
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
    # 在这里直接指定文件路径
    
    # 定义要转换的Python文件路径
    PYTHON_FILE = r"D:\feifei\2026_01_19\auto_test\save_scv\night_cycle_update_feifei_steps.py"
    
    # 定义输出格式（可选：standard 或 table）
    OUTPUT_FORMAT = "standard"  # 强制使用标准格式
    
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
        'StepType',
        'LoopConfig', 
        'TestStep',
        'get_statistics',
        'print_statistics',
        'print_all_steps',
        'parse_python_file',
        'parse_teststep_string',
        'parse_loopconfig',
        'convert_to_scv',
        'generate_dynamic_scv_from_python'
    ]