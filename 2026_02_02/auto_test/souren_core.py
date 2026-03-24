from lib.var import *
from souren_config import *

ADB_AVAILABLE = False
try:
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import adb_integration
    ADB_AVAILABLE = True
    print("✅ ADB模块导入成功")
except ImportError:
    print("⚠️  未找到adb_integration模块,手机模式功能将不可用")
    ADB_AVAILABLE = False

class CallCommandProcessor:
    """CALL命令处理器 - 添加 AT 控制和手机模式控制功能"""
    
    @staticmethod
    def process_call_command(call_command: str, instrument_controller) -> Tuple[bool, str]:
        """处理CALL命令 - 添加 Board AT 控制和手机飞行模式控制"""
        if not call_command:
            return False, "空的CALL命令"
        
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
            command_upper = original_command.upper()
            need_at_control = any(cmd.upper() == command_upper for cmd in CELL_COMMANDS_NEED_AT)
            
            if need_at_control and result[0]:
                print(f"🔧 仪器命令发送成功，开始执行 AT 序列: {AT_SEQUENCE}")
                
                try:
                    from board_at_controller import send_at_sequence_directly
                    at_success, at_summary = send_at_sequence_directly()
                    
                    if at_success:
                        print(f"✅ AT序列执行成功")
                    else:
                        print(f"⚠️  AT序列执行失败: {at_summary}")
                        
                except Exception as e:
                    print(f"⚠️  执行AT序列时出错: {e}")
        
        elif DEVICE_TYPE.lower() == 'phone':
            command_upper = original_command.upper()
            need_flight_mode_control = any(cmd.upper() == command_upper for cmd in CELL_COMMANDS_NEED_ADB)
            
            if need_flight_mode_control and result[0]:
                print(f"📱 PHONE模式: 检测到 CELL 控制命令,启动ADB飞行模式控制")
                print(f"   设备类型: {DEVICE_TYPE}")
                print(f"   等待时间: {ADB_WAIT_TIME}秒")
                
                if not ADB_AVAILABLE:
                    print(f"❌ ADB模块不可用,跳过飞行模式控制")
                    return result
                
                try:
                    adb_controller = adb_integration.ADBFlightModeController()
                    
                    if adb_controller.adb_path and adb_controller.device_id:
                        print(f"✅ ADB连接成功,设备: {adb_controller.device_id}")
                        print(f"   开始执行定时飞行模式控制...")
                        
                        flight_success = adb_controller.timed_flight_mode_control(
                            wait_time=ADB_WAIT_TIME
                        )
                        
                        if flight_success:
                            print(f"✅ ADB飞行模式控制执行成功")
                        else:
                            print(f"⚠️  ADB飞行模式控制执行失败")
                    else:
                        print(f"❌ ADB连接失败,跳过飞行模式控制")
                        
                except Exception as e:
                    print(f"⚠️  执行飞行模式控制时出错: {e}")
        
        return result

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
            
            try:
                idn = self.instrument.query('*IDN?').strip()
                print(f"✅ 仪器连接成功: {idn}")
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
        
        if not command:
            print(f"⚠️  警告：尝试发送空命令，跳过")
            return False, "空的命令"
        
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

class SourenLogger:
    """Souren.ToolSet 日志系统"""
    
    def __init__(self):
        from souren_config import LOG_ENABLED, LOG_LEVEL
        
        self.enabled = LOG_ENABLED
        
        try:
            from souren_config import _get_log_file
            self.log_file = _get_log_file()
        except:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            log_dir = os.path.join(base_dir, "log")
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.log_file = os.path.join(log_dir, f"souren_execution_{timestamp}.log")
        
        if self.enabled and self.log_file:
            log_level = getattr(logging, LOG_LEVEL, logging.INFO)
            
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
            self.result_dir = result_dir
            if not os.path.exists(self.result_dir):
                os.makedirs(self.result_dir, exist_ok=True)
                print(f"📁 创建结果目录: {self.result_dir}")
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if script_name:
                script_base = os.path.splitext(os.path.basename(script_name))[0]
                result_filename = f"{script_base}_results_{timestamp}.json"
            else:
                result_filename = f"souren_results_{timestamp}.json"
            
            self.result_file = os.path.join(self.result_dir, result_filename)
        else:
            from souren_config import RESULT_FILE
            
            if hasattr(RESULT_FILE, '__call__'):
                self.result_file = RESULT_FILE()
            elif hasattr(RESULT_FILE, 'fget'):
                self.result_file = RESULT_FILE.fget()
            else:
                self.result_file = RESULT_FILE
            
            self.result_dir = os.path.dirname(self.result_file) if self.result_file else None
        
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
            
            if self.script_name:
                result_data['script_name'] = self.script_name
            
            self.results.append(result_data)
            
            with open(self.result_file, 'w', encoding='utf-8') as f:
                json.dump(self.results, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 结果已保存到: {self.result_file}")
            return True
        except Exception as e:
            print(f"❌ 保存结果失败: {e}")
            return False

class PythonScriptParser:
    """Python脚本解析器 - 从Python文件中提取测试步骤，支持if-else条件分支"""
    
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
            
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print(f"📖 读取Python文件成功: {os.path.basename(file_path)}")
            
            steps = self._parse_test_steps(content, file_path)
            skip_commands = self._parse_skip_commands(content, file_path)
            
            print(f"✅ 成功解析Python脚本,共 {len(steps)} 个测试步骤")
            
            # 统计不同类型的步骤
            loop_steps = [step for step in steps if step.get("has_loop", False)]
            if_steps = [step for step in steps if step.get("type") == "IfCondition"]
            else_steps = [step for step in steps if step.get("type") == "ElseCondition"]
            for_steps = [step for step in steps if step.get("type") == "ForLoop"]
            
            if loop_steps:
                print(f"   循环步骤: {len(loop_steps)}")
            if if_steps:
                print(f"   条件判断步骤: {len(if_steps)}")
            if else_steps:
                print(f"   条件分支步骤: {len(else_steps)}")
            if for_steps:
                print(f"   循环控制步骤: {len(for_steps)}")
            
            if skip_commands:
                print(f"   跳过的命令: {len(skip_commands)}")
            
            return {
                "success": True,
                "file": file_path,
                "total_steps": len(steps),
                "steps": steps,
                "skip_commands": skip_commands,
                "format": "python_script"
            }
            
        except Exception as e:
            print(f"❌ 解析Python脚本时出错: {e}")
            return {
                "success": False,
                "error": f"解析Python脚本失败: {str(e)}",
            }
    
    def _parse_test_steps(self, content: str, file_path: str) -> List[Dict]:
        """从Python脚本中解析TEST_STEPS，支持if-else条件分支"""
        steps = []
        
        print(f"🔍 开始解析TEST_STEPS...")
        
        try:
            import re
            import ast
            
            # 尝试使用ast解析
            try:
                tree = ast.parse(content)
                
                for node in ast.walk(tree):
                    if isinstance(node, ast.Assign):
                        for target in node.targets:
                            if isinstance(target, ast.Name) and target.id == 'TEST_STEPS':
                                if isinstance(node.value, ast.List):
                                    steps = self._parse_ast_list(node.value, file_path)
                                    print(f"✅ 使用ast解析成功，找到 {len(steps)} 个步骤")
                                    return steps
            except Exception as ast_e:
                print(f"⚠️  ast解析失败，使用正则表达式解析: {ast_e}")
            
            # 使用正则表达式解析
            test_steps_patterns = [
                r'TEST_STEPS\s*[:=]\s*\[(.*?)\]',
                r'TEST_STEPS\s*=\s*\[(.*?)\]',
                r'TEST_STEPS\s*:\s*List\[Dict\]\s*=\s*\[(.*?)\]',
            ]
            
            test_steps_content = None
            for pattern in test_steps_patterns:
                test_steps_match = re.search(pattern, content, re.DOTALL | re.IGNORECASE)
                if test_steps_match:
                    test_steps_content = test_steps_match.group(1)
                    print(f"✅ 使用模式找到TEST_STEPS定义: {pattern}")
                    break
            
            if not test_steps_content:
                print(f"❌ 未找到TEST_STEPS定义")
                return steps
            
            return self._parse_with_regex(test_steps_content, file_path)
            
        except Exception as e:
            print(f"❌ 解析TEST_STEPS时出错: {e}")
            return []
    
    def _parse_ast_list(self, list_node: ast.List, file_path: str) -> List[Dict]:
        """解析ast列表节点"""
        steps = []
        
        for i, element in enumerate(list_node.elts):
            step_info = self._parse_ast_element(element, i + 1, file_path)
            if step_info:
                steps.append(step_info)
        
        return steps
    
    def _parse_ast_element(self, element, index: int, file_path: str) -> Dict:
        """解析ast元素"""
        if isinstance(element, ast.Dict):
            return self._parse_ast_dict(element, index, file_path)
        return None
    
    def _parse_ast_dict(self, dict_node: ast.Dict, index: int, file_path: str) -> Dict:
        """解析ast字典节点"""
        step_dict = {}
        
        # 提取字典的键值对
        for key, value in zip(dict_node.keys, dict_node.values):
            if isinstance(key, ast.Str):
                key_str = key.s
                
                # 处理不同类型的值
                if isinstance(value, ast.Str):
                    step_dict[key_str] = value.s
                elif isinstance(value, ast.Num):
                    step_dict[key_str] = value.n
                elif isinstance(value, ast.NameConstant):  # True, False, None
                    step_dict[key_str] = value.value
                elif isinstance(value, ast.Dict):
                    # 处理嵌套字典，如loop_config
                    nested_dict = {}
                    for n_key, n_value in zip(value.keys, value.values):
                        if isinstance(n_key, ast.Str):
                            if isinstance(n_value, ast.Str):
                                nested_dict[n_key.s] = n_value.s
                            elif isinstance(n_value, ast.Num):
                                nested_dict[n_key.s] = n_value.n
                            elif isinstance(n_value, ast.NameConstant):
                                nested_dict[n_key.s] = n_value.value
                    step_dict[key_str] = nested_dict
                elif isinstance(value, ast.List):
                    # 处理列表
                    list_values = []
                    for item in value.elts:
                        if isinstance(item, ast.Str):
                            list_values.append(item.s)
                        elif isinstance(item, ast.Num):
                            list_values.append(item.n)
                        elif isinstance(item, ast.NameConstant):
                            list_values.append(item.value)
                    step_dict[key_str] = list_values
        
        # 转换为标准步骤格式
        return self._create_step_from_dict(step_dict, index, file_path)
    
    def _parse_with_regex(self, content: str, file_path: str) -> List[Dict]:
        """使用正则表达式解析步骤"""
        steps = []
        
        # 匹配字典
        dict_pattern = r'\{(?:[^{}]|(?:\{[^{}]*\}))*?\}'
        dict_matches = re.findall(dict_pattern, content, re.DOTALL)
        
        print(f"🔍 找到 {len(dict_matches)} 个字典")
        
        for i, dict_content in enumerate(dict_matches):
            dict_content = dict_content.strip()
            if not dict_content or dict_content == '{}':
                continue
            
            # 解析字典内容
            step_dict = self._parse_dict_content(dict_content)
            step_info = self._create_step_from_dict(step_dict, i + 1, file_path)
            
            if step_info:
                steps.append(step_info)
                # 打印步骤信息
                step_type = step_info.get("type", "Normal")
                if step_type == "IfCondition":
                    print(f"  🔍 步骤 {i+1}: 条件判断 - {step_info.get('condition', '')}")
                elif step_type == "ElseCondition":
                    print(f"  🔄 步骤 {i+1}: 条件分支 - Else")
                elif step_type == "ForLoop":
                    for_info = step_info.get("for", "")
                    times_info = f", times={step_info.get('times', 1)}" if step_info.get("times") else ""
                    print(f"  🔄 步骤 {i+1}: 循环标记 - {for_info}{times_info}")
                else:
                    command = step_info.get("content", "")[:80]
                    print(f"  ✅ 步骤 {i+1}: {command}...")
        
        return steps
    
    def _parse_dict_content(self, dict_content: str) -> Dict:
        """解析字典字符串内容"""
        result = {}
        
        # 处理特殊情况：嵌套字典
        dict_content = dict_content.strip('{}').strip()
        
        # 分割键值对
        pattern = r'"([^"]+)"\s*:\s*(?:("(?:[^"\\]|\\.)*")|(\'(?:[^\'\\]|\\.)*\')|([^,}]+))'
        matches = re.findall(pattern, dict_content, re.DOTALL)
        
        for match in matches:
            key = match[0]
            # 获取值（可能在不同的捕获组中）
            value_str = ""
            for group in match[1:]:
                if group:
                    value_str = group
                    break
            
            # 清理值
            if value_str.startswith('"') and value_str.endswith('"'):
                value_str = value_str[1:-1]
            elif value_str.startswith("'") and value_str.endswith("'"):
                value_str = value_str[1:-1]
            
            # 处理特殊键
            if key in ["command", "expected_result", "condition", "expected_value", "data_type", "for", "if", "else"]:
                result[key] = value_str
            elif key in ["check_index", "threshold", "times", "sleep_ms"]:
                try:
                    result[key] = float(value_str) if '.' in value_str else int(value_str)
                except:
                    result[key] = value_str
            elif key == "enable":
                result[key] = value_str.lower() in ["true", "yes", "1", "on"]
            elif key == "loop_config":
                # 解析嵌套的loop_config
                loop_config_match = re.search(r'loop_config\s*:\s*\{([^}]+(?:\{[^}]*\}[^}]*)*)\}', dict_content, re.DOTALL)
                if loop_config_match:
                    loop_config_content = loop_config_match.group(1)
                    loop_config = self._parse_loop_config(loop_config_content)
                    result[key] = loop_config
        
        return result
    
    def _parse_loop_config(self, content: str) -> Dict:
        """解析loop_config"""
        config = {}
        
        # 提取键值对
        patterns = [
            (r'enable\s*:\s*(True|False|true|false)', 'enable', lambda x: x.lower() == "true"),
            (r'times\s*:\s*(\d+)', 'times', int),
            (r'sleep_ms\s*:\s*(\d+)', 'sleep_ms', int),
            (r'expected_result\s*:\s*"([^"]+)"', 'expected_result', str),
            (r'expected_result\s*:\s*\'([^\']+)\'', 'expected_result', str),
        ]
        
        for pattern, key, converter in patterns:
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                try:
                    config[key] = converter(match.group(1))
                except:
                    pass
        
        return config
    
    def _create_step_from_dict(self, step_dict: Dict, index: int, file_path: str) -> Dict:
        """从字典创建标准步骤"""
        step_type = "Normal"
        
        # 检查步骤类型
        if "for" in step_dict:
            step_type = "ForLoop"
        elif "if" in step_dict:
            step_type = "IfCondition"
        elif "else" in step_dict:
            step_type = "ElseCondition"
        
        # 构建基础步骤信息
        step_info = {
            "step": index,
            "type": step_type,
            "run_status": "Pass",
            "enable": "true",
            "unloop_pass_result_stop": "false",
            "cmd_para_table": "",
            "status": "未执行",
            "executed": False,
            "result": None,
            "start_time": None,
            "end_time": None,
            "has_loop": False,
        }
        
        # 根据类型添加特定字段
        if step_type == "ForLoop":
            step_info["content"] = f"for:{step_dict.get('for', '')}"
            step_info["read_string"] = ""
            step_info["loop_times"] = "1"
            step_info["sleep_ms"] = "0"
            step_info["loop_pass_result"] = ""
            step_info["for"] = step_dict.get("for", "")
            step_info["times"] = step_dict.get("times", 1)  # 添加times参数
        
        elif step_type == "IfCondition":
            step_info["content"] = f"if:{step_dict.get('if', '')}"
            step_info["read_string"] = ""
            step_info["loop_times"] = "1"
            step_info["sleep_ms"] = "0"
            step_info["loop_pass_result"] = ""
            step_info["condition"] = step_dict.get("if", "")
            step_info["condition_type"] = "python"  # 支持Python表达式
        
        elif step_type == "ElseCondition":
            step_info["content"] = "else:"
            step_info["read_string"] = ""
            step_info["loop_times"] = "1"
            step_info["sleep_ms"] = "0"
            step_info["loop_pass_result"] = ""
        
        else:  # Normal步骤
            command = step_dict.get("command", "").strip()
            if not command:
                return None
            
            step_info["content"] = command
            step_info["read_string"] = step_dict.get("expected_result", "")
            step_info["expected_result"] = step_dict.get("expected_result", "")
            
            # 检查是否有循环配置
            loop_config = step_dict.get("loop_config", {})
            has_loop = loop_config.get("enable", False)
            
            if has_loop:
                step_info["has_loop"] = True
                step_info["loop_times"] = str(loop_config.get("times", 1))
                step_info["sleep_ms"] = str(loop_config.get("sleep_ms", 0))
                step_info["loop_pass_result"] = loop_config.get("expected_result", "")
                step_info["loop_config"] = {
                    "enable": "true",
                    "looptimes": str(loop_config.get("times", 1)),
                    "sleepms": str(loop_config.get("sleep_ms", 0)),
                    "looppassresult": loop_config.get("expected_result", "")
                }
            
            else:
                step_info["loop_times"] = "1"
                step_info["sleep_ms"] = "0"
                step_info["loop_pass_result"] = ""
            
            # 添加停止条件字段
            stop_condition_fields = ["check_index", "condition", "threshold", "data_type", "expected_value"]
            for field in stop_condition_fields:
                if field in step_dict:
                    step_info[field] = step_dict[field]
        
        return step_info
    
    def _parse_skip_commands(self, content: str, file_path: str) -> Set[str]:
        """从Python脚本中解析SKIP_IN_NEXT_CYCLES - 修复版本"""
        skip_commands = set()
        
        print(f"🔍 开始解析SKIP_IN_NEXT_CYCLES...")
        
        try:
            import re
            import ast
            
            try:
                tree = ast.parse(content)
                
                for node in ast.walk(tree):
                    if isinstance(node, ast.Assign):
                        for target in node.targets:
                            if isinstance(target, ast.Name) and target.id == 'SKIP_IN_NEXT_CYCLES':
                                value = node.value
                                if isinstance(value, ast.List):
                                    for element in value.elts:
                                        if isinstance(element, ast.Dict):
                                            keys = element.keys
                                            values = element.values
                                            for k, v in zip(keys, values):
                                                if isinstance(k, ast.Str) and k.s == 'command':
                                                    if isinstance(v, ast.Str):
                                                        skip_commands.add(v.s.strip())
                                                    elif isinstance(v, ast.Constant):
                                                        skip_commands.add(str(v.value).strip())
                                
                                print(f"✅ 使用ast解析SKIP_IN_NEXT_CYCLES成功，提取到 {len(skip_commands)} 个需要跳过的命令")
                                return skip_commands
            except Exception as ast_e:
                print(f"⚠️  ast解析SKIP_IN_NEXT_CYCLES失败，尝试正则表达式: {ast_e}")
            
            skip_patterns = [
                r'SKIP_IN_NEXT_CYCLES\s*[:=]\s*\[(.*?)\]',
                r'SKIP_IN_NEXT_CYCLES\s*=\s*\[(.*?)\]',
                r'SKIP_IN_NEXT_CYCLES\s*:\s*List\[.*?\]\s*=\s*\[(.*?)\]',
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
            
            dict_pattern = r'\{(?:[^{}]|(?:\{[^{}]*\}))*?\}'
            dict_matches = re.findall(dict_pattern, skip_content, re.DOTALL)
            
            print(f"🔍 正则表达式找到 {len(dict_matches)} 个字典")
            
            for dict_content in dict_matches:
                dict_content = dict_content.strip()
                if not dict_content or dict_content == '{}':
                    continue
                
                command_patterns = [
                    r'"command"\s*:\s*"([^"]+)"',
                    r"'command'\s*:\s*'([^']+)'",
                    r'"command"\s*:\s*\'([^\']+)\'',
                    r"'command'\s*:\s*\"([^\"]+)\"",
                ]
                
                command = None
                for pattern in command_patterns:
                    command_match = re.search(pattern, dict_content)
                    if command_match:
                        command = command_match.group(1).strip()
                        break
                
                if not command:
                    command_match = re.search(r'command\s*:\s*([^,\n\}]+)', dict_content)
                    if command_match:
                        command = command_match.group(1).strip().strip('"').strip("'")
                
                if command:
                    skip_commands.add(command)
        except Exception as e:
            print(f"❌ 解析SKIP_IN_NEXT_CYCLES时出错: {e}")
        
        return skip_commands