from lib.var import *
from souren_config import (
    INSTRUMENT_ADDRESS,
    LOG_ENABLED,
    LOG_LEVEL,
    SHOW_COMMAND_SENDING,
    _get_log_file,
    RESULT_FILE
)

try:
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import common
    print("✅ common模块导入成功")
except ImportError as e:
    print(f"⚠️  导入common模块失败: {e}")
    common = None

class CallCommandProcessor:
    @staticmethod
    def process_call_command(call_command: str, instrument_controller) -> Tuple[bool, str]:
        if not call_command:
            return False, "空的CALL命令"
        original_command = call_command.strip()
        if not original_command:
            return False, "空的命令"
        print(f"📡 发送命令到仪器: '{original_command}'")
        return instrument_controller.execute_scpi_command(original_command)

class InstrumentController:
    def __init__(self, device_address: str = None):
        self.device_address = device_address or INSTRUMENT_ADDRESS
        self.rm = None
        self.instrument = None
        self.connected = False
        self.timeout = 10000

    def connect(self) -> Tuple[bool, str]:
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
                return False, f"连接成功但仪器无响应: {e}"
        except Exception as e:
            error_msg = f"连接失败: {str(e)}"
            print(f"❌ {error_msg}")
            return False, error_msg

    def disconnect(self):
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
        if not self.connected or not self.instrument:
            return False, "仪器未连接"
        command = command.strip()
        if not command:
            return False, "空的命令"
        try:
            if '?' in command:
                result = self.instrument.query(command).strip()
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
        return CallCommandProcessor.process_call_command(command, self)

class DirectCommandExecutor:
    instrument_controller = None

    @staticmethod
    def initialize() -> bool:
        print("🔄 初始化仪器连接...")
        try:
            DirectCommandExecutor.instrument_controller = InstrumentController()
            success, message = DirectCommandExecutor.instrument_controller.connect()
            if success:
                print("✅ 仪器连接成功")
                if common:
                    common.setup_instrument_controller(DirectCommandExecutor.instrument_controller)
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
        if not DirectCommandExecutor.instrument_controller:
            return False, "仪器未连接，请先初始化"
        command = command.strip()
        if not command:
            return False, "空的命令"
        if command.upper().startswith("SLEEP"):
            import re
            sleep_match = re.search(r'SLEEP\s+(\d+)', command.upper())
            if sleep_match:
                sleep_ms = int(sleep_match.group(1))
                if SHOW_COMMAND_SENDING:
                    print(f"😴 睡眠 {sleep_ms} 毫秒...")
                time.sleep(sleep_ms / 1000)
                return True, f"睡眠完成 ({sleep_ms}毫秒)"
        return DirectCommandExecutor.instrument_controller.execute_call_command(command)

    @staticmethod
    def cleanup():
        if DirectCommandExecutor.instrument_controller:
            DirectCommandExecutor.instrument_controller.disconnect()
        print("🧹 资源已清理")

class SourenLogger:
    def __init__(self):
        self.enabled = LOG_ENABLED
        try:
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
                handlers=[logging.FileHandler(self.log_file, encoding='utf-8')]
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

    def info(self, message: str, **kwargs): self.log('INFO', message, **kwargs)
    def error(self, message: str, **kwargs): self.log('ERROR', message, **kwargs)
    def warning(self, message: str, **kwargs): self.log('WARNING', message, **kwargs)
    def debug(self, message: str, **kwargs): self.log('DEBUG', message, **kwargs)

class SourenResultSaver:
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

    def get_result_file(self): return self.result_file
    def get_result_dir(self): return self.result_dir

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

class PythonScriptExecutor:
    def __init__(self):
        self.logger = SourenLogger()
        self.current_loop_iteration = 1
        self.total_loop_count = 1
        self.extracted_data = []
        self.execution_details = []
        self.step_counter = 0
        self._pending_check = None
        self._current_command_is_query = False
        self.query_expected_map = {}

    def reset(self):
        self.step_counter = 0
        self.execution_details = []
        self.extracted_data = []
        self._pending_check = None
        self._current_command_is_query = False
        self.query_expected_map.clear()

    def set_loop_info(self, loop_iteration: int, total_loop_count: int):
        self.current_loop_iteration = loop_iteration
        self.total_loop_count = total_loop_count

    def execute_script(self, file_path: str, parameters: Dict = None,
                       loop_iteration: int = 1, total_loop_count: int = 1) -> Tuple[bool, Dict]:
        if not os.path.exists(file_path):
            return False, {"error": f"文件不存在: {file_path}"}
        self.reset()
        self.current_loop_iteration = loop_iteration
        self.total_loop_count = total_loop_count
        print(f"🚀 开始执行Python脚本: {os.path.basename(file_path)} (循环 {loop_iteration}/{total_loop_count})")
        with open(file_path, 'r', encoding='utf-8') as f:
            script_content = f.read()
        self._build_query_expected_map(script_content, file_path)
        local_env = {
            'os': os, 'sys': sys, 'time': time, 'datetime': datetime,
            '__file__': file_path, 'external_params': parameters or {}, 'self': self
        }
        if common:
            local_env['check_phone_at'] = common.check_phone_at
        class APWrapper:
            def __init__(self, executor): 
                self.executor = executor

            def send(self, command, extract_index=None, should_extract=False, chart_title=None, x_label=None):
                self.executor._current_command_is_query = False
                return self.executor._execute_ap_command(command, extract_index, should_extract, chart_title, x_label)

            def query(self, command, extract_index=None, should_extract=False, chart_title=None, x_label=None):
                self.executor._current_command_is_query = True
                return self.executor._execute_ap_command(command, extract_index, should_extract, chart_title, x_label)

            def sleep(self, ms):
                self.executor._current_command_is_query = False
                return self.executor._execute_sleep(ms, self.executor.step_counter+1)
        ap_wrapper = APWrapper(self)
        local_env['ap'] = ap_wrapper
        def my_sleep_wrapper(seconds):
            ms = int(seconds * 1000)
            command = f"SLEEP {ms}"
            self._current_command_is_query = False
            return self._execute_ap_command(command, None, False)
        local_env['my_sleep'] = my_sleep_wrapper
        if common:
            try:
                common.ap = ap_wrapper
                common.my_sleep = my_sleep_wrapper
                print("✅ 已强制替换 common.ap 和 common.my_sleep 为我们的包装器")
            except Exception as e:
                print(f"⚠️ 替换 common 对象失败: {e}")
        try:
            code = compile(script_content, file_path, 'exec')
            exec(code, local_env)
            if 'update_parameters' in local_env:
                print("\n🔄 调用update_parameters更新参数...")
                local_env['update_parameters'](parameters or {})
            elif 'parameter' in local_env and parameters:
                print("\n🔄 更新脚本参数...")
                for key, value in parameters.items():
                    if key in local_env['parameter']:
                        print(f"   {key}: {local_env['parameter'][key]} -> {value}")
                        local_env['parameter'][key] = value
                    else:
                        print(f"   {key}: {value} (新参数)")
            if 'case_start' in local_env:
                print("\n🔧 执行 case_start()...")
                local_env['case_start']()
            if 'case_body' in local_env:
                print("\n🔧 执行 case_body()...")
                local_env['case_body']()
            if 'case_clear' in local_env:
                print("\n🧹 执行 case_clear()...")
                local_env['case_clear']()
            self._finalize_pending_check(forced=True)
            return True, {
                "success": True,
                "execution_details": self.execution_details,
                "extracted_data": self.extracted_data,
                "step_count": self.step_counter,
                "parameters": parameters,
                "loop_iteration": self.current_loop_iteration,
                "loop_count": self.total_loop_count
            }
        except Exception as e:
            error_msg = f"执行脚本失败: {str(e)}"
            print(f"❌ {error_msg}")
            import traceback
            traceback.print_exc()
            return False, {"error": error_msg}

    def _build_query_expected_map(self, script_content: str, file_path: str):
        try:
            tree = ast.parse(script_content)
            query_calls = []  
            for node in ast.walk(tree):
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                    if node.func.attr == 'query' and isinstance(node.func.value, ast.Name) and node.func.value.id == 'ap':
                        command = self._get_command_from_call(node, script_content)
                        if command:
                            query_calls.append((node.lineno, command))

            if_nodes = [] 
            for node in ast.walk(tree):
                if isinstance(node, ast.If):
                    expected = self._extract_expected_from_condition(node.test)
                    if expected:
                        if_nodes.append((node.lineno, expected, node))

            query_calls.sort(key=lambda x: x[0])
            if_nodes.sort(key=lambda x: x[0])

            used_if = set()
            for q_lineno, q_cmd in query_calls:
                for i, (if_lineno, expected, if_node) in enumerate(if_nodes):
                    if if_lineno > q_lineno and i not in used_if:
                        self.query_expected_map[q_cmd] = expected
                        used_if.add(i)
                        print(f"📌 动态提取预期: {q_cmd} -> {expected} (行 {q_lineno})")
                        break
        except Exception as e:
            print(f"⚠️  AST解析失败,动态预期提取将不可用: {e}")

    def _extract_expected_from_condition(self, node) -> Optional[str]:
        if isinstance(node, ast.Compare):
            left, comparators, ops = node.left, node.comparators, node.ops
            for op, right in zip(ops, comparators):
                if isinstance(op, (ast.Eq, ast.In)):
                    for expr in (left, right):
                        val = self._get_constant_str(expr)
                        if val is not None:
                            return val
        elif isinstance(node, (ast.And, ast.Or)):
            for val in node.values:
                expected = self._extract_expected_from_condition(val)
                if expected:
                    return expected
        return None

    def _get_constant_str(self, node) -> Optional[str]:
        raw_val = None
        if hasattr(ast, 'Str') and isinstance(node, ast.Str):
            raw_val = node.s
        elif isinstance(node, ast.Constant) and isinstance(node.value, str):
            raw_val = node.value
        if raw_val is not None:
            return raw_val.strip().strip('"').strip("'")
        return None

    def _get_command_from_call(self, call_node: ast.Call, script_content: str) -> Optional[str]:
        try:
            if len(call_node.args) > 0:
                arg = call_node.args[0]
                val = self._get_constant_str(arg)
                if val:
                    return val
                if hasattr(call_node, 'lineno') and hasattr(call_node, 'col_offset'):
                    lines = script_content.splitlines()
                    if call_node.lineno <= len(lines):
                        line = lines[call_node.lineno - 1]
                        import re
                        match = re.search(r'ap\.query\([\'\"]([^\'\"]+)[\'\"]', line)
                        if match:
                            return match.group(1)
        except:
            pass
        return None

    def _execute_ap_command(self, command: str, extract_index=None, should_extract=False,chart_title=None, x_label=None) -> str:
        step_start_time = time.time()
        if isinstance(command, str) and command.upper().startswith("SLEEP"):
            import re
            sleep_match = re.search(r'SLEEP\s+(\d+)', command.upper())
            if sleep_match:
                sleep_ms = int(sleep_match.group(1))
                if self._pending_check and self._current_command_is_query is False:
                    self._pending_check['total_sleep'] += sleep_ms
                    self._pending_check['duration'] += sleep_ms / 1000
                else:
                    self._execute_sleep(sleep_ms, self.step_counter + 1)
                time.sleep(sleep_ms / 1000)
                return f"睡眠完成 ({sleep_ms}毫秒)"
            else:
                success, result = DirectCommandExecutor.execute_command(str(command))
                return result
        self.step_counter += 1
        step_num = self.step_counter
        print(f"\n📌【步骤 {step_num}】命令: {command} (来源: {'ap.query' if self._current_command_is_query else 'ap.send'})")
        success, result = DirectCommandExecutor.execute_command(str(command))
        clean_result = None
        if isinstance(result, str):
            clean_result = result.strip().strip('"').strip("'")
        extracted = None
        if should_extract and extract_index is not None:
            extracted = self._extract_data_from_result(result, extract_index)
            if extracted is not None:
                self.extracted_data.append({
                    "step": step_num, "command": command,
                    "extracted_data": extracted,
                    "loop_iteration": self.current_loop_iteration,
                    "chart_title": chart_title,
                    "x_label": x_label,   
                })
                print(f"  📊 提取数据: {extracted} (标题: {chart_title}, 横坐标: {x_label})")
        if self._current_command_is_query:
            expected = self.query_expected_map.get(command, None)
            is_success = False
            if expected and clean_result and expected in clean_result:
                is_success = True
            if self._pending_check is None or self._pending_check['command'] != command:
                self._finalize_pending_check()
                self._pending_check = {
                    'step': step_num, 'command': command, 'attempts': 1,
                    'success': is_success, 'expected': expected,
                    'first_result': result, 'last_result': result,
                    'start_time': step_start_time, 'duration': 0.0,
                    'total_sleep': 0, 'extract_index': extract_index,
                    'extracted_data': extracted
                }
            else:
                self._pending_check['attempts'] += 1
                self._pending_check['last_result'] = result
                self._pending_check['duration'] = time.time() - self._pending_check['start_time']
                if is_success:
                    self._pending_check['success'] = True
                if should_extract and extract_index is not None and extracted is not None:
                    self._pending_check['extracted_data'] = extracted
                    for item in self.extracted_data:
                        if item['step'] == self._pending_check['step']:
                            item['extracted_data'] = extracted
                            break
            if is_success:
                self._finalize_pending_check()
        else:
            self._finalize_pending_check()
            self._record_command_step(step_num, command, success, result, step_start_time, extract_index, extracted)
        return result

    def _finalize_pending_check(self, forced=False):
        if not self._pending_check:
            return
        p = self._pending_check
        step_num, command = p['step'], p['command']
        attempts, success, expected = p['attempts'], p['success'], p['expected']
        last_result, duration = p['last_result'], time.time() - p['start_time']
        if success:
            status_msg = f"第{attempts}次查询达到预期"
            if expected:
                status_msg += f' "{expected}"'
            status = "success"
        else:
            status_msg = f"第{attempts}次查询未达到预期"
            if expected:
                status_msg += f' "{expected}"'
            status = "failed"
        detail = {
            "step": step_num, "type": "Check", "function": "unknown",
            "content": command, "status": status, "duration": duration,
            "result": f"{status_msg} | 末次结果: {last_result[:200]}" if last_result else status_msg,
            "start_time": p['start_time'], "end_time": time.time(),
            "loop_iteration": self.current_loop_iteration,
            "loop_count": self.total_loop_count,
            "attempts": attempts, "expected": expected,
            "extracted_data": p.get('extracted_data')
        }
        self.execution_details.append(detail)
        print(f"  ✅ Check合并完成 - 尝试{attempts}次, 状态: {status}")
        self._pending_check = None

    def _record_command_step(self, step_num, command, success, result, start_time, extract_index, extracted):
        detail = {
            "step": step_num, "type": "Command", "function": "unknown",
            "content": command, "status": "success" if success else "failed",
            "duration": time.time() - start_time,
            "result": result if isinstance(result, str) else str(result),
            "start_time": start_time, "end_time": time.time(),
            "loop_iteration": self.current_loop_iteration,
            "loop_count": self.total_loop_count,
            "extracted_data": extracted
        }
        self.execution_details.append(detail)
        print(f"  ✅ 命令执行完成 - 耗时: {detail['duration']:.2f}秒")

    def _execute_sleep(self, sleep_ms: int, step_num: int) -> str:
        print(f"😴 独立睡眠 {sleep_ms} 毫秒...")
        time.sleep(sleep_ms / 1000)
        detail = {
            "step": step_num, "type": "Sleep", "function": "unknown",
            "content": f"SLEEP {sleep_ms}", "status": "success",
            "duration": sleep_ms / 1000,
            "result": f"睡眠完成 ({sleep_ms}毫秒)",
            "start_time": time.time() - sleep_ms / 1000, "end_time": time.time(),
            "loop_iteration": self.current_loop_iteration,
            "loop_count": self.total_loop_count,
            "extracted_data": None
        }
        self.execution_details.append(detail)
        return f"睡眠完成 ({sleep_ms}毫秒)"

    def _extract_data_from_result(self, result: str, extract_index: int) -> Optional[float]:
        try:
            if not result:
                return None
            result_str = str(result).strip()
            error_keywords = ["仪器通信错误", "VI_ERROR_TMO", "Timeout", "通信失败", "错误", "ERROR", "失败"]
            if any(keyword in result_str.upper() for keyword in [k.upper() for k in error_keywords]):
                print(f"⚠️  检测到错误信息: {result_str[:100]}")
                return None

            if ',' in result_str:
                parts = [p.strip() for p in result_str.split(',')]
                if 0 <= extract_index < len(parts):
                    try:
                        return float(parts[extract_index])
                    except ValueError:
                        import re
                        num_match = re.search(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?', parts[extract_index])
                        if num_match:
                            return float(num_match.group())
            else:
                try:
                    return float(result_str)
                except ValueError:
                    import re
                    num_match = re.search(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?', result_str)
                    if num_match:
                        return float(num_match.group())
            return None
        except Exception as e:
            print(f"❌ 提取数据失败: {e}")
            return None