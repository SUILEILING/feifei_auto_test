from lib.var import *

# 修改ADB导入方式
ADB_AVAILABLE = False
try:
    # 尝试导入ADB控制器
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import adb_integration
    ADB_AVAILABLE = True
    print("✅ ADB模块导入成功")
except ImportError as e:
    print(f"⚠️  未找到adb_integration模块,手机模式功能将不可用: {e}")
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
        
        if not original_command:
            return False, "空的命令"
        
        print(f"📤 准备发送命令: '{original_command}'")
        
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
                print(f"📱 PHONE模式: 检测到 CELL 控制命令,启动ADB飞行模式控制")
                print(f"   设备类型: {DEVICE_TYPE}")
                print(f"   等待时间: {ADB_WAIT_TIME}秒")
                
                # 检查ADB是否可用
                if not ADB_AVAILABLE:
                    print(f"❌ ADB模块不可用,跳过飞行模式控制")
                    print(f"   请检查adb_integration.py文件是否存在")
                    return result
                
                try:
                    # 创建ADB控制器
                    adb_controller = adb_integration.ADBFlightModeController()
                    
                    if adb_controller.adb_path and adb_controller.device_id:
                        print(f"✅ ADB连接成功,设备: {adb_controller.device_id}")
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
                        print(f"❌ ADB连接失败,跳过飞行模式控制")
                        print(f"   请检查:")
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
        
        if not command:
            print(f"⚠️  警告：尝试发送空命令，跳过")
            return False, "空的命令"
        
        try:
            from souren_config import SHOW_COMMAND_SENDING
            if SHOW_COMMAND_SENDING:
                print(f"  发送: {command}")
            
            if '?' in command: 
                if command.upper() == "CELL1?":
                    print(f"  ⚠️  CELL1?查询可能不响应，尝试发送...")
                
                result = self.instrument.query(command).strip()
                if SHOW_COMMAND_SENDING:
                    print(f"    结果: {result}")
                return True, result
            else:  
                self.instrument.write(command)
                time.sleep(0.5)
                return True, "命令执行成功"
                
        except pyvisa.errors.VisaIOError as e:
            error_msg = f"仪器通信错误: {str(e)}"
            print(f"  ❌ {error_msg}")
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
        from souren_config import LOG_ENABLED, LOG_LEVEL
        
        self.enabled = LOG_ENABLED
        
        # 直接调用函数获取日志文件路径
        try:
            from souren_config import _get_log_file
            self.log_file = _get_log_file()
        except:
            # 如果无法获取，使用默认路径
            base_dir = os.path.dirname(os.path.abspath(__file__))
            log_dir = os.path.join(base_dir, "log")
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.log_file = os.path.join(log_dir, f"souren_execution_{timestamp}.log")
        
        if self.enabled and self.log_file:
            log_level = getattr(logging, LOG_LEVEL, logging.INFO)
            
            # 确保日志目录存在
            if self.log_file:
                log_dir = os.path.dirname(self.log_file)
                if log_dir and not os.path.exists(log_dir):
                    os.makedirs(log_dir, exist_ok=True)
            
            logging.getLogger('pyvisa').setLevel(logging.ERROR)
            
            logging.basicConfig(
                level=log_level,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(self.log_file, encoding='utf-8'),
                ]
            )
            
            self.logger = logging.getLogger('SourenToolSet')
            print(f"📁 日志文件: {self.log_file}")
        else:
            self.logger = None
            print("📁 日志功能已禁用")
    
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
    """Souren.ToolSet 结果保存系统 - 修改为支持自定义目录"""
    
    def __init__(self, result_dir=None, script_name=None):
        self.script_name = script_name
        
        if result_dir:
            # 使用指定的目录
            self.result_dir = result_dir
            # 确保目录存在
            if not os.path.exists(self.result_dir):
                os.makedirs(self.result_dir, exist_ok=True)
                print(f"📁 创建结果目录: {self.result_dir}")
            
            # 生成结果文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if script_name:
                script_base = os.path.splitext(os.path.basename(script_name))[0]
                result_filename = f"{script_base}_results_{timestamp}.json"
            else:
                result_filename = f"souren_results_{timestamp}.json"
            
            self.result_file = os.path.join(self.result_dir, result_filename)
        else:
            # 使用默认配置
            from souren_config import RESULT_FILE
            
            if hasattr(RESULT_FILE, '__call__'):
                self.result_file = RESULT_FILE()
            elif hasattr(RESULT_FILE, 'fget'):
                self.result_file = RESULT_FILE.fget()
            else:
                self.result_file = RESULT_FILE
            
            self.result_dir = os.path.dirname(self.result_file) if self.result_file else None
        
        # 确保目录存在
        if self.result_dir and not os.path.exists(self.result_dir):
            os.makedirs(self.result_dir, exist_ok=True)
        
        self.results = []
    
    def get_result_file(self):
        """获取结果文件路径"""
        return self.result_file
    
    def get_result_dir(self):
        """获取结果目录"""
        return self.result_dir
    
    def save_result(self, result_data: Dict):
        try:
            # 再次确保目录存在（防止多线程或异步操作）
            if self.result_dir and not os.path.exists(self.result_dir):
                os.makedirs(self.result_dir, exist_ok=True)
                print(f"📁 重新创建结果目录: {self.result_dir}")
            
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
            
            # 添加脚本信息
            if self.script_name:
                result_data['script_name'] = self.script_name
            
            self.results.append(result_data)
            
            with open(self.result_file, 'w', encoding='utf-8') as f:
                json.dump(self.results, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 结果已保存到: {self.result_file}")
            return True
        except Exception as e:
            print(f"❌ 保存结果失败: {e}")
            import traceback
            traceback.print_exc()
            return False

# ==============================================
# ↓↓↓ Python脚本解析器
# ==============================================
class PythonScriptParser:
    """Python脚本解析器 - 从Python文件中提取测试步骤"""
    
    def __init__(self):
        self.logger = SourenLogger()
    
    def parse_file(self, file_path: str) -> Dict:
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)
        
        self.logger.info(f"解析Python脚本: {file_path}")
        
        try:
            if not os.path.exists(file_path):
                return {
                    "success": False,
                    "error": f"文件不存在: {file_path}"
                }
            
            # 读取文件内容
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print(f"📖 读取Python文件成功: {os.path.basename(file_path)}")
            
            steps = self._parse_test_steps(content, file_path)
            skip_commands = self._parse_skip_commands(content, file_path)
            
            print(f"✅ 成功解析Python脚本,共 {len(steps)} 个测试步骤")
            print(f"📊 步骤分析:")

            
            # 检查是否有循环步骤
            loop_steps = [step for step in steps if step.get("has_loop", False)]
            if loop_steps:
                print(f"   循环步骤: {len(loop_steps)}")
                for step in loop_steps:
                    print(f"     - 步骤{step['step']}: {step['content'][:50]}...")
            
            # 显示跳过的命令信息
            if skip_commands:
                print(f"   跳过的命令: {len(skip_commands)}")
                for i, cmd in enumerate(sorted(list(skip_commands)), 1):
                    print(f"     {i:2d}. {cmd}")
            
            return {
                "success": True,
                "file": file_path,
                "total_steps": len(steps),
                "steps": steps,
                "skip_commands": skip_commands,
                "format": "python_script"
            }
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.logger.error(f"解析Python脚本失败: {str(e)}")
            print(f"❌ 解析Python脚本时出错: {e}")
            return {
                "success": False,
                "error": f"解析Python脚本失败: {str(e)}",
                "details": traceback.format_exc()
            }
    
    def _parse_test_steps(self, content: str, file_path: str) -> List[Dict]:
        """从Python脚本中解析TEST_STEPS"""
        steps = []
        
        print(f"🔍 开始解析TEST_STEPS...")
        
        try:
            import re
            import ast
            
            # 首先尝试更宽松的正则表达式
            # 匹配 TEST_STEPS = [...] 或 TEST_STEPS: List[Dict] = [...]
            test_steps_patterns = [
                r'TEST_STEPS\s*[:=]\s*\[(.*?)\]',  # 标准模式
                r'TEST_STEPS\s*=\s*\[(.*?)\]',      # 简单模式
                r'TEST_STEPS\s*:\s*List\[Dict\]\s*=\s*\[(.*?)\]',  # 带类型注解
            ]
            
            test_steps_content = None
            matched_pattern = None
            
            for pattern in test_steps_patterns:
                test_steps_match = re.search(pattern, content, re.DOTALL | re.IGNORECASE)
                if test_steps_match:
                    test_steps_content = test_steps_match.group(1)
                    matched_pattern = pattern
                    print(f"✅ 使用模式找到TEST_STEPS定义: {pattern}")
                    break
            
            if not test_steps_content:
                print(f"❌ 未找到TEST_STEPS定义")
                # 尝试查找文件中所有以TEST_STEPS开头的行
                lines = content.split('\n')
                for i, line in enumerate(lines):
                    if line.strip().startswith('TEST_STEPS'):
                        print(f"  第{i+1}行: {line}")
                return steps
            
            # 打印提取的内容用于调试
            print(f"📝 TEST_STEPS内容预览:")
            print(f"  {test_steps_content[:500]}...")
            
            # 方法1：尝试使用ast.literal_eval解析整个列表
            print(f"🔄 尝试使用ast解析TEST_STEPS...")
            try:
                # 尝试将字符串解析为Python列表
                test_steps_str = f"[{test_steps_content}]"
                parsed_steps = ast.literal_eval(test_steps_str)
                
                if isinstance(parsed_steps, list):
                    print(f"✅ 使用ast解析成功，找到 {len(parsed_steps)} 个步骤")
                    
                    for i, step_dict in enumerate(parsed_steps):
                        if isinstance(step_dict, dict):
                            command = step_dict.get("command", "").strip()
                            if not command:
                                continue
                            
                            expected_result = step_dict.get("expected_result", "")
                            loop_config = step_dict.get("loop_config", {})
                            
                            has_loop = bool(loop_config and loop_config.get("enable", False))
                            
                            # 创建步骤信息
                            step_info = {
                                "step": i + 1,
                                "content": command,
                                "type": "Normal",
                                "run_status": "Pass",
                                "read_string": expected_result,
                                "enable": "true",
                                "unloop_pass_result_stop": "false",
                                "loop_times": str(loop_config.get("times", 1)) if loop_config else "1",
                                "sleep_ms": str(loop_config.get("sleep_ms", 0)) if loop_config else "0",
                                "loop_pass_result": loop_config.get("expected_result", "") if loop_config else "",
                                "cmd_para_table": "",
                                "status": "未执行",
                                "executed": False,
                                "result": None,
                                "start_time": None,
                                "end_time": None,
                                "has_loop": has_loop,
                                "expected_result": expected_result,
                                "raw_loop_config": loop_config  # 保存原始loop_config
                            }
                            
                            if has_loop:
                                step_info["loop_config"] = {
                                    "enable": "true",
                                    "looptimes": str(loop_config.get("times", 1)),
                                    "sleepms": str(loop_config.get("sleep_ms", 0)),
                                    "looppassresult": loop_config.get("expected_result", "")
                                }
                                print(f"  ✅ 步骤 {i+1} 已启用循环: 次数={loop_config.get('times', 1)}, 期望结果='{loop_config.get('expected_result', '')}'")
                            
                            steps.append(step_info)
                            print(f"  ✅ 步骤 {i+1}: {command[:80]}...")
                    
                    return steps
            except Exception as e:
                print(f"⚠️  ast解析失败，使用正则表达式解析: {e}")
            
            # 方法2：如果ast解析失败，使用正则表达式
            print(f"🔄 使用正则表达式解析TEST_STEPS...")
            # 尝试多种方法解析字典 - 改进的正则表达式
            # 使用非贪婪匹配来匹配每个字典
            dict_pattern = r'\{(?:[^{}]|(?:\{[^{}]*\}))*?\}'
            dict_matches = re.findall(dict_pattern, test_steps_content, re.DOTALL)
            
            print(f"🔍 使用改进的正则表达式找到 {len(dict_matches)} 个字典")
            
            if not dict_matches:
                # 回退到简单模式
                dict_pattern = r'\{[^{}]*\}'
                dict_matches = re.findall(dict_pattern, test_steps_content)
                print(f"🔍 使用简单模式找到 {len(dict_matches)} 个字典")
            
            for i, dict_content in enumerate(dict_matches):
                # 清理字典内容
                dict_content = dict_content.strip()
                if not dict_content or dict_content == '{}':
                    continue
                
                # 提取command - 使用多种引号格式
                command_patterns = [
                    r'"command"\s*:\s*"([^"]+)"',
                    r"'command'\s*:\s*'([^']+)'",
                    r'"command"\s*:\s*\'([^\']+)\'',  # 混合引号
                    r"'command'\s*:\s*\"([^\"]+)\"",  # 混合引号
                ]
                
                command = None
                for pattern in command_patterns:
                    command_match = re.search(pattern, dict_content)
                    if command_match:
                        command = command_match.group(1).strip()
                        break
                
                if not command:
                    # 尝试查找不带引号的command
                    command_match = re.search(r'command\s*:\s*([^,\n\}]+)', dict_content)
                    if command_match:
                        command = command_match.group(1).strip().strip('"').strip("'")
                
                if not command:
                    print(f"⚠️  第{i+1}个步骤没有command字段，跳过")
                    print(f"   字典内容: {dict_content}")
                    continue
                
                # 提取expected_result
                expected_result = ""
                expected_patterns = [
                    r'"expected_result"\s*:\s*"([^"]+)"',
                    r"'expected_result'\s*:\s*'([^']+)'",
                ]
                
                for pattern in expected_patterns:
                    expected_match = re.search(pattern, dict_content)
                    if expected_match:
                        expected_result = expected_match.group(1).strip()
                        break
                
                # 提取loop_config - 改进的正则表达式
                loop_config = {}
                # 查找loop_config字典，支持多行
                loop_config_match = re.search(r'loop_config\s*:\s*\{([^}]+(?:\{[^}]*\}[^}]*)*)\}', dict_content, re.DOTALL)
                
                if loop_config_match:
                    loop_config_content = loop_config_match.group(1)
                    
                    # 提取loop_config中的各个字段
                    config_patterns = {
                        "enable": r'enable\s*:\s*(True|False|true|false)',
                        "times": r'times\s*:\s*(\d+)',
                        "sleep_ms": r'sleep_ms\s*:\s*(\d+)',
                        "expected_result": r'expected_result\s*:\s*["\']([^"\']+)["\']',
                    }
                    
                    for key, pattern in config_patterns.items():
                        match = re.search(pattern, loop_config_content, re.IGNORECASE)
                        if match:
                            if key == "enable":
                                loop_config[key] = match.group(1).lower() == "true"
                            elif key == "times" or key == "sleep_ms":
                                loop_config[key] = int(match.group(1))
                            else:
                                loop_config[key] = match.group(1)
                    
                    print(f"  🔄 步骤 {i+1} 找到loop_config: {loop_config}")
                else:
                    # 尝试查找简单的loop_config
                    if 'loop_config' in dict_content.lower():
                        print(f"  ⚠️  步骤 {i+1} 包含loop_config但解析失败")
                        print(f"     字典内容: {dict_content[:200]}...")
                
                # 检查是否有loop_config
                has_loop = bool(loop_config and loop_config.get("enable", False))
                
                # 创建步骤信息
                step_info = {
                    "step": i + 1,
                    "content": command,
                    "type": "Normal",
                    "run_status": "Pass",
                    "read_string": expected_result,
                    "enable": "true",
                    "unloop_pass_result_stop": "false",
                    "loop_times": str(loop_config.get("times", 1)) if loop_config else "1",
                    "sleep_ms": str(loop_config.get("sleep_ms", 0)) if loop_config else "0",
                    "loop_pass_result": loop_config.get("expected_result", "") if loop_config else "",
                    "cmd_para_table": "",
                    "status": "未执行",
                    "executed": False,
                    "result": None,
                    "start_time": None,
                    "end_time": None,
                    "has_loop": has_loop,
                    "expected_result": expected_result,
                    "raw_loop_config": loop_config  # 保存原始loop_config
                }
                
                if has_loop:
                    step_info["loop_config"] = {
                        "enable": "true",
                        "looptimes": str(loop_config.get("times", 1)),
                        "sleepms": str(loop_config.get("sleep_ms", 0)),
                        "looppassresult": loop_config.get("expected_result", "")
                    }
                    print(f"  ✅ 步骤 {i+1} 已启用循环: 次数={loop_config.get('times', 1)}, 期望结果='{loop_config.get('expected_result', '')}'")
                
                steps.append(step_info)
                print(f"  ✅ 步骤 {i+1}: {command[:80]}...")
            
            # 验证步骤顺序
            print(f"📋 验证步骤顺序:")
            for i, step in enumerate(steps):
                print(f"  步骤 {i+1}: {step['content'][:50]}...")
                if step.get('has_loop'):
                    print(f"    包含循环配置: 次数={step.get('loop_times')}, 期望={step.get('loop_pass_result')}")
            
            return steps
            
        except Exception as e:
            print(f"❌ 解析TEST_STEPS时出错: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def _parse_skip_commands(self, content: str, file_path: str) -> Set[str]:
        """从Python脚本中解析SKIP_IN_NEXT_CYCLES - 修复版本"""
        skip_commands = set()
        
        print(f"🔍 开始解析SKIP_IN_NEXT_CYCLES...")
        
        try:
            import re
            import ast
            
            # 尝试使用ast解析整个文件
            try:
                tree = ast.parse(content)
                
                for node in ast.walk(tree):
                    # 寻找赋值节点，且变量名为SKIP_IN_NEXT_CYCLES
                    if isinstance(node, ast.Assign):
                        for target in node.targets:
                            if isinstance(target, ast.Name) and target.id == 'SKIP_IN_NEXT_CYCLES':
                                # 获取赋值的值
                                value = node.value
                                if isinstance(value, ast.List):
                                    for element in value.elts:
                                        if isinstance(element, ast.Dict):
                                            # 在字典中查找键为'command'的值
                                            keys = element.keys
                                            values = element.values
                                            for k, v in zip(keys, values):
                                                if isinstance(k, ast.Str) and k.s == 'command':
                                                    if isinstance(v, ast.Str):
                                                        skip_commands.add(v.s.strip())
                                                    # 如果是常量，也可以处理其他类型的节点，这里只处理字符串
                                                    elif isinstance(v, ast.Constant):
                                                        skip_commands.add(str(v.value).strip())
                                
                                print(f"✅ 使用ast解析SKIP_IN_NEXT_CYCLES成功，提取到 {len(skip_commands)} 个需要跳过的命令")
                                return skip_commands
            except Exception as ast_e:
                print(f"⚠️  ast解析SKIP_IN_NEXT_CYCLES失败，尝试正则表达式: {ast_e}")
            
            # 如果ast解析失败，使用正则表达式
            # 改进的正则表达式，匹配更多格式
            skip_patterns = [
                r'SKIP_IN_NEXT_CYCLES\s*[:=]\s*\[(.*?)\]',  # 原始模式
                r'SKIP_IN_NEXT_CYCLES\s*=\s*\[(.*?)\]',      # 简单赋值
                r'SKIP_IN_NEXT_CYCLES\s*:\s*List\[.*?\]\s*=\s*\[(.*?)\]',  # 带类型注解
            ]
            
            skip_match = None
            skip_content = ""
            for pattern in skip_patterns:
                skip_match = re.search(pattern, content, re.DOTALL | re.IGNORECASE)
                if skip_match:
                    skip_content = skip_match.group(1)
                    print(f"✅ 使用模式找到SKIP_IN_NEXT_CYCLES: {pattern}")
                    break
            
            if not skip_match:
                print(f"⚠️  未找到SKIP_IN_NEXT_CYCLES定义,使用空跳过列表")
                return skip_commands
        
            try:
                skip_list_str = f"[{skip_content}]"
                parsed_skip_list = ast.literal_eval(skip_list_str)
                
                if isinstance(parsed_skip_list, list):
                    for item in parsed_skip_list:
                        if isinstance(item, dict) and "command" in item:
                            skip_commands.add(item["command"].strip())
                    

                    
                    return skip_commands
            except Exception as e:
                print(f"⚠️  ast解析SKIP_IN_NEXT_CYCLES失败,使用正则表达式: {e}")
            
            # 使用正则表达式作为备选
            # 首先匹配所有字典
            dict_pattern = r'\{(?:[^{}]|(?:\{[^{}]*\}))*?\}'
            dict_matches = re.findall(dict_pattern, skip_content, re.DOTALL)
            
            print(f"🔍 正则表达式找到 {len(dict_matches)} 个字典")
            
            for dict_content in dict_matches:
                # 清理字典内容
                dict_content = dict_content.strip()
                if not dict_content or dict_content == '{}':
                    continue
                
                # 提取command - 使用多种引号格式
                command_patterns = [
                    r'"command"\s*:\s*"([^"]+)"',
                    r"'command'\s*:\s*'([^']+)'",
                    r'"command"\s*:\s*\'([^\']+)\'',  # 混合引号
                    r"'command'\s*:\s*\"([^\"]+)\"",  # 混合引号
                ]
                
                command = None
                for pattern in command_patterns:
                    command_match = re.search(pattern, dict_content)
                    if command_match:
                        command = command_match.group(1).strip()
                        break
                
                if not command:
                    # 尝试查找不带引号的command
                    command_match = re.search(r'command\s*:\s*([^,\n\}]+)', dict_content)
                    if command_match:
                        command = command_match.group(1).strip().strip('"').strip("'")
                
                if command:
                    skip_commands.add(command)
        except Exception as e:
            print(f"❌ 解析SKIP_IN_NEXT_CYCLES时出错: {e}")
            import traceback
            traceback.print_exc()
        
        return skip_commands