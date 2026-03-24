import sys
import time
import json
import logging
import os
import re
import pyvisa
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

# 修改ADB导入方式
ADB_AVAILABLE = False
try:
    # 尝试导入ADB控制器
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import adb_integration
    ADB_AVAILABLE = True
    print("✅ ADB模块导入成功")
except ImportError as e:
    print(f"⚠️  未找到adb_integration模块，手机模式功能将不可用: {e}")
    ADB_AVAILABLE = False
except Exception as e:
    print(f"⚠️  导入ADB模块时出错: {e}")
    ADB_AVAILABLE = False

# ==============================================
# ↓↓↓ CALL命令处理器（添加 AT 控制功能和手机模式控制）
# ==============================================
class CallCommandProcessor:
    """CALL命令处理器 - 添加 AT 控制和手机模式控制功能"""
    
    @staticmethod
    def process_call_command(call_command: str, instrument_controller) -> Tuple[bool, str]:
        """处理CALL命令 - 添加 Board AT 控制和手机飞行模式控制"""
        if not call_command:
            return False, "空的CALL命令"
        
        # 清理命令
        original_command = call_command.strip()
        
        # ⭐⭐⭐ 修复：去掉空的命令提示 ⭐⭐⭐
        if not original_command:
            return False, "空的命令"
        
        print(f"📤 准备发送命令: '{original_command}'")
        
        # ⭐⭐⭐ 检查设备类型 ⭐⭐⭐
        from souren_config import DEVICE_TYPE, CELL_COMMANDS_NEED_AT, AT_SEQUENCE
        from souren_config import ADB_WAIT_TIME, CELL_COMMANDS_NEED_ADB
        
        # 1. 先发送命令到仪器
        print(f"📡 先发送命令到仪器: '{original_command}'")
        result = instrument_controller.execute_scpi_command(original_command)
        
        # 2. 根据设备类型执行后续控制
        if DEVICE_TYPE.lower() == 'board':
            # BOARD模式：执行AT序列控制
            need_at_control = False
            if DEVICE_TYPE.lower() == 'board':
                command_upper = original_command.upper()
                need_at_control = any(cmd.upper() == command_upper for cmd in CELL_COMMANDS_NEED_AT)
                
                if need_at_control:
                    print(f"📱 BOARD模式: 检测到 CELL 控制命令")
            
            if need_at_control and result[0]:  # result[0] 是 success
                print(f"🔧 仪器命令发送成功，开始执行 AT 序列: {AT_SEQUENCE}")
                
                try:
                    # 直接发送AT序列
                    from board_at_controller import send_at_sequence_directly
                    at_success, at_summary = send_at_sequence_directly()
                    
                    if at_success:
                        print(f"✅ AT序列执行成功")
                    else:
                        print(f"⚠️  AT序列执行失败: {at_summary}")
                        
                except Exception as e:
                    print(f"⚠️  执行AT序列时出错: {e}")
        
        elif DEVICE_TYPE.lower() == 'phone':
            # PHONE模式：执行ADB飞行模式控制
            need_flight_mode_control = False
            command_upper = original_command.upper()
            need_flight_mode_control = any(cmd.upper() == command_upper for cmd in CELL_COMMANDS_NEED_ADB)
            
            if need_flight_mode_control and result[0]:
                print(f"📱 PHONE模式: 检测到 CELL 控制命令，启动ADB飞行模式控制")
                print(f"   设备类型: {DEVICE_TYPE}")
                print(f"   等待时间: {ADB_WAIT_TIME}秒")
                
                # 检查ADB是否可用
                if not ADB_AVAILABLE:
                    print(f"❌ ADB模块不可用，跳过飞行模式控制")
                    print(f"   请检查adb_integration.py文件是否存在")
                    return result
                
                try:
                    # 创建ADB控制器
                    adb_controller = adb_integration.ADBFlightModeController()
                    
                    if adb_controller.adb_path and adb_controller.device_id:
                        print(f"✅ ADB连接成功，设备: {adb_controller.device_id}")
                        print(f"   开始执行定时飞行模式控制...")
                        
                        # 执行定时飞行模式控制
                        flight_success = adb_controller.timed_flight_mode_control(
                            wait_time=ADB_WAIT_TIME
                        )
                        
                        if flight_success:
                            print(f"✅ ADB飞行模式控制执行成功")
                        else:
                            print(f"⚠️  ADB飞行模式控制执行失败")
                    else:
                        print(f"❌ ADB连接失败，跳过飞行模式控制")
                        print(f"   请检查：")
                        print(f"   1. 手机USB调试是否开启")
                        print(f"   2. USB线是否连接正常")
                        print(f"   3. 电脑是否安装了ADB驱动")
                        
                except Exception as e:
                    print(f"⚠️  执行飞行模式控制时出错: {e}")
                    import traceback
                    traceback.print_exc()
        
        return result

# ==============================================
# ↓↓↓ 仪器通信接口
# ==============================================
class InstrumentController:
    """仪器控制器"""
    
    def __init__(self, device_address: str = None):
        from souren_config import INSTRUMENT_ADDRESS
        
        if device_address is None:
            device_address = INSTRUMENT_ADDRESS
        
        self.device_address = device_address
        self.rm = None
        self.instrument = None
        self.connected = False
        self.timeout = 10000
    
    def connect(self) -> Tuple[bool, str]:
        """连接仪器"""
        try:
            self.rm = pyvisa.ResourceManager()
            print(f"🔌 尝试连接仪器: {self.device_address}")
            
            self.instrument = self.rm.open_resource(self.device_address)
            self.instrument.timeout = self.timeout
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            
            # 测试连接
            try:
                idn = self.instrument.query('*IDN?').strip()
                print(f"✅ 仪器连接成功: {idn}")
                print(f"  📝 确认: 仪器接受CALL:前缀的命令")
                self.connected = True
                return True, f"连接成功: {idn}"
            except Exception as e:
                print(f"❌ 仪器响应测试失败: {e}")
                return False, f"连接成功但仪器无响应: {e}"
                
        except Exception as e:
            error_msg = f"连接失败: {str(e)}"
            print(f"❌ {error_msg}")
            return False, error_msg
    
    def disconnect(self):
        """断开连接"""
        if self.instrument:
            try:
                self.instrument.close()
                print("📴 仪器连接已关闭")
            except:
                pass
            finally:
                self.instrument = None
        self.connected = False
    
    def execute_scpi_command(self, command: str) -> Tuple[bool, str]:
        """执行SCPI命令"""
        if not self.connected or not self.instrument:
            return False, "仪器未连接"
        
        command = command.strip()
        
        # ⭐⭐⭐ 修复：检查是否为空命令 ⭐⭐⭐
        if not command:
            print(f"⚠️  警告：尝试发送空命令，跳过")
            return False, "空的命令"
        
        try:
            from souren_config import SHOW_COMMAND_SENDING
            if SHOW_COMMAND_SENDING:
                print(f"  发送: {command}")
            
            if '?' in command:  # 查询命令
                # ⭐⭐⭐ 特殊处理：CELL1?可能不响应 ⭐⭐⭐
                if command.upper() == "CELL1?":
                    print(f"  ⚠️  CELL1?查询可能不响应，尝试发送...")
                
                result = self.instrument.query(command).strip()
                if SHOW_COMMAND_SENDING:
                    print(f"    结果: {result}")
                return True, result
            else:  # 设置命令
                self.instrument.write(command)
                time.sleep(0.5)
                return True, "命令执行成功"
                
        except pyvisa.errors.VisaIOError as e:
            error_msg = f"仪器通信错误: {str(e)}"
            print(f"  ❌ {error_msg}")
            
            # 检查是否是特殊命令（通信错误也视为通过）
            from souren_config import COMMUNICATION_ERROR_PASS_COMMANDS
            if command in COMMUNICATION_ERROR_PASS_COMMANDS:
                print(f"  ⭐ 是特殊命令，通信错误视为通过")
                return True, f"特殊命令通信错误: {command}"
            
            return False, error_msg
        except Exception as e:
            error_msg = f"命令执行失败: {str(e)}"
            print(f"  ❌ {error_msg}")
            return False, error_msg
    
    def execute_call_command(self, command: str) -> Tuple[bool, str]:
        """执行CALL命令"""
        return CallCommandProcessor.process_call_command(command, self)

# ==============================================
# ↓↓↓ 直接命令执行器
# ==============================================
class DirectCommandExecutor:
    """直接命令执行器"""
    
    instrument_controller = None
    
    @staticmethod
    def initialize() -> bool:
        """初始化仪器连接"""
        print("🔄 初始化仪器连接...")
        
        try:
            DirectCommandExecutor.instrument_controller = InstrumentController()
            success, message = DirectCommandExecutor.instrument_controller.connect()
            
            if success:
                print("✅ 仪器连接成功")
                return True
            else:
                print(f"❌ 仪器连接失败: {message}")
                return False
                
        except ImportError:
            print("❌ 请安装pyvisa库: pip install pyvisa pyvisa-py")
            return False
        except Exception as e:
            print(f"❌ 初始化失败: {e}")
            return False
    
    @staticmethod
    def execute_command(command: str) -> Tuple[bool, str]:
        """执行命令"""
        if not DirectCommandExecutor.instrument_controller:
            return False, "仪器未连接，请先初始化"
        
        command = command.strip()
        
        # ⭐⭐⭐ 修复：检查是否为空命令 ⭐⭐⭐
        if not command:
            print(f"⚠️  警告：尝试发送空命令，跳过")
            return False, "空的命令"
        
        # 处理睡眠命令
        if command.upper().startswith("SLEEP"):
            import re
            sleep_match = re.search(r'SLEEP\s+(\d+)', command.upper())
            if sleep_match:
                sleep_ms = int(sleep_match.group(1))
                from souren_config import SHOW_COMMAND_SENDING
                if SHOW_COMMAND_SENDING:
                    print(f"😴 睡眠 {sleep_ms} 毫秒...")
                time.sleep(sleep_ms / 1000)
                return True, f"睡眠完成 ({sleep_ms}毫秒)"
        
        return DirectCommandExecutor.instrument_controller.execute_call_command(command)
    
    @staticmethod
    def cleanup():
        """清理资源"""
        if DirectCommandExecutor.instrument_controller:
            DirectCommandExecutor.instrument_controller.disconnect()
        print("🧹 资源已清理")

# ==============================================
# ↓↓↓ 日志系统
# ==============================================
class SourenLogger:
    """Souren.ToolSet 日志系统"""
    
    def __init__(self):
        from souren_config import LOG_ENABLED, LOG_LEVEL, LOG_FILE
        
        self.enabled = LOG_ENABLED
        self.log_file = LOG_FILE
        
        if self.enabled:
            log_level = getattr(logging, LOG_LEVEL, logging.INFO)
            
            log_dir = os.path.dirname(self.log_file)
            if log_dir and not os.path.exists(log_dir):
                os.makedirs(log_dir)
            
            logging.getLogger('pyvisa').setLevel(logging.ERROR)
            
            logging.basicConfig(
                level=log_level,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(self.log_file, encoding='utf-8'),
                ]
            )
            
            self.logger = logging.getLogger('SourenToolSet')
        else:
            self.logger = None
    
    def log(self, level: str, message: str, **kwargs):
        if not self.enabled or not self.logger:
            return
        
        log_method = getattr(self.logger, level.lower(), self.logger.warning)
        if kwargs:
            message = f"{message} | {kwargs}"
        log_method(message)
    
    def info(self, message: str, **kwargs):
        self.log('INFO', message, **kwargs)
    
    def error(self, message: str, **kwargs):
        self.log('ERROR', message, **kwargs)
    
    def warning(self, message: str, **kwargs):
        self.log('WARNING', message, **kwargs)
    
    def debug(self, message: str, **kwargs):
        self.log('DEBUG', message, **kwargs)

class SourenResultSaver:
    """Souren.ToolSet 结果保存系统"""
    
    def __init__(self):
        from souren_config import RESULT_FILE
        self.result_file = RESULT_FILE
        
        result_dir = os.path.dirname(self.result_file)
        if result_dir and not os.path.exists(result_dir):
            os.makedirs(result_dir)
        
        self.results = []
    
    def save_result(self, result_data: Dict):
        try:
            if os.path.exists(self.result_file):
                try:
                    with open(self.result_file, 'r', encoding='utf-8') as f:
                        existing_results = json.load(f)
                        if isinstance(existing_results, list):
                            self.results = existing_results
                except:
                    self.results = []
            
            if 'timestamp' not in result_data:
                result_data['timestamp'] = datetime.now().isoformat()
            
            if 'timestamp_readable' not in result_data:
                result_data['timestamp_readable'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            self.results.append(result_data)
            
            with open(self.result_file, 'w', encoding='utf-8') as f:
                json.dump(self.results, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 结果已保存到: {self.result_file}")
            return True
        except Exception as e:
            print(f"❌ 保存结果失败: {e}")
            return False

class SCVParser:
    """SCV文件解析器 - 支持多种格式"""
    
    def __init__(self):
        self.logger = SourenLogger()
    
    def parse_file(self, file_path: str) -> Dict:
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)
        
        self.logger.info(f"解析SCV文件: {file_path}")
        
        try:
            if not os.path.exists(file_path):
                return {
                    "success": False,
                    "error": f"文件不存在: {file_path}"
                }
            
            encodings = ['utf-8', 'gbk', 'latin-1', 'utf-8-sig']
            content = None
            
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        content = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                return {
                    "success": False,
                    "error": "无法解码文件内容"
                }
            
            # ⭐⭐⭐ 首先检查是否为表格格式 ⭐⭐⭐
            lines = content.strip().split('\n')
            
            # 检查是否为表格格式（包含"步骤 "和"命令:"）
            if len(lines) > 5 and any("步骤 " in line for line in lines[:10]) and any("命令:" in line for line in lines[:10]):
                print(f"📋 检测到表格格式SCV文件")
                result = self._parse_table_format(content, file_path)
                if result.get("success"):
                    return result
            
            # 检查是否为标准格式（包含"[001]"和"|"）
            if len(lines) > 0 and any(re.search(r'^\[\d+\]', line) for line in lines[:10]) and "|" in content:
                print(f"📋 检测到标准格式SCV文件")
                result = self._parse_standard_format(content, file_path)
                if result.get("success"):
                    return result
            
            # 检查是否为XML格式
            if content.strip().startswith('<?xml') or '<CmdSet' in content:
                print(f"📋 检测到XML格式SCV文件")
                result = self._parse_xml_format(content, file_path)
                if result.get("success"):
                    return result
            
            # 尝试解析CSV/Tab格式
            if len(lines) > 0:
                first_line = lines[0].strip()
                if '\t' in first_line and ('Content' in first_line or 'Type' in first_line):
                    print(f"📋 检测到CSV/Tab格式SCV文件")
                    headers = first_line.split('\t')
                    steps = []
                    
                    for i in range(1, len(lines)):
                        line = lines[i].strip()
                        if not line:
                            continue
                        
                        columns = line.split('\t')
                        if len(columns) >= 2:
                            step_info = self._create_step_info(i, columns)
                            steps.append(step_info)
                    
                    if steps:
                        analysis = self._analyze_steps(steps)
                        return {
                            "success": True,
                            "file": file_path,
                            "total_steps": len(steps),
                            "steps": steps,
                            "analysis": analysis,
                            "format": "csv_tab"
                        }
            
            # 最后尝试简单格式
            steps = self._parse_simple_format(lines)
            if steps:
                analysis = self._analyze_steps(steps)
                return {
                    "success": True,
                    "file": file_path,
                    "total_steps": len(steps),
                    "steps": steps,
                    "analysis": analysis,
                    "format": "simple"
                }
            
            return {
                "success": False,
                "error": "文件中没有找到有效的测试步骤或不支持的格式"
            }
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.logger.error(f"解析SCV文件失败: {str(e)}")
            return {
                "success": False,
                "error": f"解析SCV文件失败: {str(e)}",
                "details": error_details
            }
    
    def _parse_table_format(self, content: str, file_path: str) -> Dict:
        """解析表格格式的SCV文件"""
        steps = []
        lines = content.strip().split('\n')
        
        step_data = {}
        current_step_num = 0
        in_step_block = False
        step_counter = 0
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # 检测步骤开始
            if line.startswith("步骤 "):
                # 如果有上一个步骤的数据，保存它
                if step_data and "command" in step_data:
                    step_info = self._create_step_from_table_data(current_step_num, step_data)
                    steps.append(step_info)
                    step_counter += 1
                
                # 解析步骤编号
                try:
                    step_num_str = line.replace("步骤", "").strip()
                    current_step_num = int(step_num_str)
                except:
                    # 如果解析失败，递增步骤编号
                    current_step_num = step_counter + 1
                
                step_data = {}
                in_step_block = True
                continue
            
            # 只在步骤块内解析其他字段
            if in_step_block:
                # 解析命令 - 只提取命令内容，去掉标签
                if line.startswith("命令:"):
                    command = line.replace("命令:", "").strip()
                    step_data["command"] = command
                    # ⭐⭐⭐ 关键：这是实际要执行的命令 ⭐⭐⭐
                
                # 解析描述
                elif line.startswith("描述:"):
                    description = line.replace("描述:", "").strip()
                    step_data["description"] = description
                
                # 解析类型
                elif line.startswith("类型:"):
                    step_type = line.replace("类型:", "").strip()
                    step_data["type"] = step_type
                
                # 解析状态
                elif line.startswith("状态:"):
                    run_status = line.replace("状态:", "").strip()
                    step_data["run_status"] = run_status
                
                # 解析超时
                elif line.startswith("超时:"):
                    timeout_str = line.replace("超时:", "").replace("ms", "").strip()
                    try:
                        step_data["timeout_ms"] = int(timeout_str)
                    except:
                        step_data["timeout_ms"] = 10000
                
                # 解析循环配置
                elif line.startswith("循环配置:"):
                    loop_config = line.replace("循环配置:", "").strip()
                    step_data["has_loop"] = True
                
                # 预期结果
                elif line.startswith("预期结果:"):
                    expected_result = line.replace("预期结果:", "").strip()
                    if expected_result and expected_result != "无":
                        step_data["read_string"] = expected_result
                
                # 分隔线表示步骤结束
                elif line.startswith("---") or line.startswith("====") or line.startswith("____") or "---" in line:
                    if step_data and "command" in step_data:
                        step_info = self._create_step_from_table_data(current_step_num, step_data)
                        steps.append(step_info)
                        step_counter += 1
                        step_data = {}
                    in_step_block = False
        
        # 添加最后一个步骤
        if step_data and "command" in step_data:
            step_info = self._create_step_from_table_data(current_step_num, step_data)
            steps.append(step_info)
        
        if not steps:
            return {
                "success": False,
                "error": "表格格式中没有找到有效的测试步骤"
            }
        
        analysis = self._analyze_steps(steps)
        
        print(f"✅ 成功解析表格格式SCV，共 {len(steps)} 个步骤")
        
        return {
            "success": True,
            "file": file_path,
            "total_steps": len(steps),
            "steps": steps,
            "analysis": analysis,
            "format": "table"
        }
    
    def _create_step_from_table_data(self, step_num: int, step_data: Dict) -> Dict:
        """从表格数据创建步骤信息"""
        # ⭐⭐⭐ 关键：只使用命令内容，其他都是描述信息 ⭐⭐⭐
        command = step_data.get("command", "")
        
        step_info = {
            "step": step_num,
            "content": command,  # ⭐⭐⭐ 这是实际要执行的命令 ⭐⭐⭐
            "type": step_data.get("type", "Normal"),
            "run_status": step_data.get("run_status", "Pass"),
            "read_string": step_data.get("read_string", ""),
            "enable": "true",
            "unloop_pass_result_stop": "false",
            "loop_times": "",
            "sleep_ms": "",
            "loop_pass_result": "",
            "cmd_para_table": "",
            "status": "未执行",
            "executed": False,
            "result": None,
            "start_time": None,
            "end_time": None,
            "has_loop": step_data.get("has_loop", False)
        }
        
        # 只提取必要的字段
        if "description" in step_data:
            step_info["description"] = step_data["description"]
        
        if "timeout_ms" in step_data:
            step_info["timeout_ms"] = step_data["timeout_ms"]
        
        return step_info
    
    def _parse_standard_format(self, content: str, file_path: str) -> Dict:
        """解析标准格式的SCV文件（[001] 命令 | 类型 | 描述...）"""
        steps = []
        lines = content.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line or line.startswith('#') or line.startswith('='):
                continue
            
            # 匹配标准格式：[001] 命令 | 类型:XXX | 描述:XXX
            match = re.match(r'^\[(\d+)\]\s+(.+?)(?:\s+\|.*)?$', line)
            if match:
                step_num = int(match.group(1))
                command = match.group(2).strip()
                
                # 提取其他信息
                step_type = "Normal"
                description = ""
                run_status = "Pass"
                
                if "|" in line:
                    parts = line.split('|')
                    if len(parts) > 1:
                        command = parts[0].split(']')[1].strip()
                        
                        for part in parts[1:]:
                            part = part.strip()
                            if part.startswith("类型:"):
                                step_type = part.replace("类型:", "").strip()
                            elif part.startswith("描述:"):
                                description = part.replace("描述:", "").replace("无", "").strip()
                            elif part.startswith("期望:"):
                                expected = part.replace("期望:", "").replace("无", "").strip()
                                if expected:
                                    step_info["read_string"] = expected
                            elif part.startswith("状态:"):
                                run_status = part.replace("状态:", "").strip()
                
                step_info = {
                    "step": step_num,
                    "content": command,
                    "type": step_type,
                    "run_status": run_status,
                    "read_string": "",
                    "enable": "true",
                    "unloop_pass_result_stop": "false",
                    "loop_times": "",
                    "sleep_ms": "",
                    "loop_pass_result": "",
                    "cmd_para_table": "",
                    "status": "未执行",
                    "executed": False,
                    "result": None,
                    "start_time": None,
                    "end_time": None,
                    "has_loop": False
                }
                
                if description:
                    step_info["description"] = description
                
                steps.append(step_info)
        
        if not steps:
            return {"success": False, "error": "标准格式中没有找到有效的测试步骤"}
        
        analysis = self._analyze_steps(steps)
        
        print(f"✅ 成功解析标准格式SCV，共 {len(steps)} 个步骤")
        
        return {
            "success": True,
            "file": file_path,
            "total_steps": len(steps),
            "steps": steps,
            "analysis": analysis,
            "format": "standard"
        }
    
    def _parse_xml_format(self, content: str, file_path: str) -> Dict:
        """解析XML格式的SCV文件"""
        steps = []
        step_counter = 0
        
        # 使用正则表达式提取CmdSet
        cmdset_pattern = r'<CmdSet\s+([^>]+)>(.*?)</CmdSet>'
        cmdset_matches = re.findall(cmdset_pattern, content, re.DOTALL)
        
        for cmdset_attrs, cmdset_content in cmdset_matches:
            step_counter += 1
            
            # 解析属性
            attrs = {}
            attr_pattern = r'(\w+)="([^"]*)"'
            attr_matches = re.findall(attr_pattern, cmdset_attrs)
            for attr_name, attr_value in attr_matches:
                attrs[attr_name.lower()] = attr_value
            
            # 解析ReadString
            read_string = ""
            read_string_match = re.search(r'<ReadString>(.*?)</ReadString>', cmdset_content, re.DOTALL)
            if read_string_match:
                read_string = read_string_match.group(1).strip()
            
            # 解析Loop配置
            loop_config = {}
            loop_match = re.search(r'<Loop\s+([^/]+)/>', cmdset_content)
            if loop_match:
                loop_attrs = loop_match.group(1)
                loop_attr_matches = re.findall(r'(\w+)="([^"]*)"', loop_attrs)
                for attr_name, attr_value in loop_attr_matches:
                    loop_config[attr_name.lower()] = attr_value
            
            # 创建步骤信息
            step_info = self._create_step_info(
                step_counter,
                [
                    attrs.get('content', ''),
                    attrs.get('type', 'Normal'),
                    attrs.get('runstatus', 'Pass'),
                    read_string,
                    "true",
                    "false",
                    loop_config.get('looptimes', '1'),
                    loop_config.get('sleepms', '0'),
                    loop_config.get('looppassresult', ''),
                    ""
                ]
            )
            
            # 添加额外的XML特定信息
            step_info["run_status"] = attrs.get('runstatus', 'Pass')
            step_info["xml_format"] = True
            
            if loop_config:
                step_info.update({
                    "loop_enable": loop_config.get('enable', 'false').lower() == 'true',
                    "loop_times": int(loop_config.get('looptimes', 1)),
                    "loop_sleep_ms": int(loop_config.get('sleepms', 0)),
                    "loop_pass_result": loop_config.get('looppassresult', ''),
                    "loop_unloop_pass_result_stop": loop_config.get('unlooppassresultstop', 'false').lower() == 'true',
                    "has_loop": True,
                    "loop_config": loop_config
                })
            
            steps.append(step_info)
        
        if not steps:
            return {"success": False, "error": "没有找到CmdSet格式的测试步骤"}
        
        analysis = self._analyze_steps(steps)
        
        return {
            "success": True,
            "file": file_path,
            "total_steps": len(steps),
            "steps": steps,
            "analysis": analysis,
            "format": "xml"
        }
    
    def _create_step_info(self, step_num: int, columns: list) -> Dict:
        """创建步骤信息"""
        step_info = {
            "step": step_num,
            "content": columns[0] if len(columns) > 0 else "",
            "type": columns[1] if len(columns) > 1 else "Normal",
            "run_status": columns[2] if len(columns) > 2 else "Pass",
            "read_string": columns[3] if len(columns) > 3 else "",
            "enable": columns[4] if len(columns) > 4 else "true",
            "unloop_pass_result_stop": columns[5] if len(columns) > 5 else "false",
            "loop_times": columns[6] if len(columns) > 6 else "",
            "sleep_ms": columns[7] if len(columns) > 7 else "",
            "loop_pass_result": columns[8] if len(columns) > 8 else "",
            "cmd_para_table": columns[9] if len(columns) > 9 else "",
            "status": "未执行",
            "executed": False,
            "result": None,
            "start_time": None,
            "end_time": None,
            "has_loop": False
        }
        
        # 检查是否有循环配置
        if step_info["loop_times"] and step_info["loop_times"].isdigit() and int(step_info["loop_times"]) > 1:
            step_info["has_loop"] = True
            step_info["loop_config"] = {
                "enable": step_info["enable"],
                "looptimes": step_info["loop_times"],
                "sleepms": step_info["sleep_ms"],
                "looppassresult": step_info["loop_pass_result"],
                "unlooppassresultstop": step_info["unloop_pass_result_stop"]
            }
        
        return step_info
    
    def _parse_simple_format(self, lines):
        """解析简单格式"""
        steps = []
        step_counter = 0
        
        for i, line in enumerate(lines):
            line = line.strip()
            if not line or line.startswith('<?') or line.startswith('<!--') or line.startswith('#'):
                continue
            
            # ⭐⭐⭐ 修复：移除行内注释 ⭐⭐⭐
            if '#' in line:
                # 分割字符串，只取注释前的部分
                line = line.split('#')[0].strip()
            
            step_counter += 1
            step_info = self._create_step_info(step_counter, [line])
            steps.append(step_info)
        
        return steps
    
    def _analyze_steps(self, steps):
        """分析步骤"""
        analysis = {
            "total_steps": len(steps),
            "call_steps": 0,
            "sleep_steps": 0,
            "normal_steps": 0,
            "other_steps": 0,
            "estimated_duration_ms": 0,
            "step_types": {}
        }
        
        for step in steps:
            step_type = step.get("type", "Normal")
            content = step.get("content", "").upper()
            
            analysis["step_types"][step_type] = analysis["step_types"].get(step_type, 0) + 1
            
            if step_type == "Normal":
                analysis["normal_steps"] += 1
                
                if content.startswith("CALL:"):
                    analysis["call_steps"] += 1
                elif "SLEEP" in content:
                    analysis["sleep_steps"] += 1
                    
                    sleep_match = re.search(r'SLEEP\s+(\d+)', content)
                    if sleep_match:
                        sleep_ms = int(sleep_match.group(1))
                        analysis["estimated_duration_ms"] += sleep_ms
            else:
                analysis["other_steps"] += 1
        
        analysis["estimated_duration_sec"] = analysis["estimated_duration_ms"] / 1000
        
        return analysis