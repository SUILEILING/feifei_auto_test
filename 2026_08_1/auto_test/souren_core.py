from lib.var import *
from souren_config import (
    INSTRUMENT_ADDRESS,
    LOG_ENABLED,
    LOG_LEVEL,
    SHOW_COMMAND_SENDING,
    _get_log_file,
    RESULT_FILE,
    LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS
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


class VisaInstrumentController:
    scpi_comm_log = []
    SCPI_LOG_MAX_LINES = 200000  # 防止超长压测撑爆内存

    @classmethod
    def _log_scpi(cls, direction, text):
        if len(cls.scpi_comm_log) < cls.SCPI_LOG_MAX_LINES:
            ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S:%f')[:-3]
            cls.scpi_comm_log.append(f"{ts}   {direction} {text}")

    def __init__(self, device_address: str = None):
        self.device_address = device_address or INSTRUMENT_ADDRESS
        self.rm = None
        self.instrument = None
        self.connected = False
        self.timeout = 30000         
        self.max_retries = 3          
        self.reconnect_attempts = 0
        self.max_reconnect_attempts = 2

    def _create_resource_manager(self):
        try:
            self.rm = pyvisa.ResourceManager()
            return True
        except Exception as e:
            print(f"❌ 创建 ResourceManager 失败: {e}")
            return False

    def connect(self) -> Tuple[bool, str]:
        if not self._create_resource_manager():
            return False, "无法创建 ResourceManager"
        try:
            print(f"🔌 尝试连接仪器: {self.device_address}")
            self.instrument = self.rm.open_resource(self.device_address)
            self.instrument.timeout = self.timeout
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            # 验证连接
            idn = self.instrument.query('*IDN?').strip()
            print(f"✅ 仪器连接成功: {idn}")
            self.connected = True
            self.reconnect_attempts = 0
            return True, f"连接成功: {idn}"
        except Exception as e:
            error_msg = f"连接失败: {str(e)}"
            print(f"❌ {error_msg}")
            self.connected = False
            self.instrument = None
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

    def reconnect(self) -> bool:
        print("🔄 尝试重新连接仪器 (重建资源管理器)...")
        self.disconnect()
        self.rm = None
        time.sleep(1)
        success, msg = self.connect()
        if success:
            print("✅ 重新连接成功")
        else:
            print(f"❌ 重新连接失败: {msg}")
        return success

    def execute_scpi_command(self, command: str) -> Tuple[bool, str]:
        if not self.connected or not self.instrument:
            if not self.reconnect():
                return False, "仪器未连接，重连失败"

        command = command.strip()
        if not command:
            return False, "空的命令"

        for attempt in range(self.max_retries):
            try:
                if '?' in command:
                    VisaInstrumentController._log_scpi("Send->", command)
                    result = self.instrument.query(command).strip()
                    VisaInstrumentController._log_scpi("Recv<-", result)
                    return True, result
                else:
                    VisaInstrumentController._log_scpi("Send->", command)
                    self.instrument.write(command)
                    time.sleep(0.1)
                    return True, "命令执行成功"
            except pyvisa.errors.VisaIOError as e:
                error_msg = str(e)
                if ("10053" in error_msg or "connection" in error_msg.lower() or
                    "VI_ERROR_RSRC_NFOUND" in error_msg or "resource not present" in error_msg.lower()):
                    print(f"⚠️ 检测到连接/资源错误 (尝试 {attempt+1}/{self.max_retries})...")
                    if attempt < self.max_retries - 1:
                        if self.reconnect():
                            continue
                        else:
                            return False, f"重连失败: {error_msg}"
                    else:
                        return False, f"仪器通信错误，重试失败: {error_msg}"
                else:
                    # 其他 VISA 错误（如超时），不重连
                    return False, f"仪器通信错误: {error_msg}"
            except Exception as e:
                error_msg = str(e)
                if ("10053" in error_msg or "connection" in error_msg.lower() or
                    "VI_ERROR_RSRC_NFOUND" in error_msg):
                    print(f"⚠️ 检测到连接/资源错误 (尝试 {attempt+1}/{self.max_retries})...")
                    if attempt < self.max_retries - 1:
                        if self.reconnect():
                            continue
                        else:
                            return False, f"重连失败: {error_msg}"
                    else:
                        return False, f"命令执行失败，重试失败: {error_msg}"
                else:
                    return False, f"命令执行失败: {error_msg}"
        return False, "达到最大重试次数"

    def execute_call_command(self, command: str) -> Tuple[bool, str]:
        try:
            return CallCommandProcessor.process_call_command(command, self)
        except Exception as e:
            return False, f"处理 CALL 命令异常: {str(e)}"


class DirectCommandExecutor:
    instrument_controller = None

    @staticmethod
    def initialize() -> bool:
        print("🔄 初始化仪器连接...")
        try:
            DirectCommandExecutor.instrument_controller = VisaInstrumentController()
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
            if not DirectCommandExecutor.initialize():
                return False, "仪器未连接，初始化失败"

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

        max_outer_retries = 2
        for outer in range(max_outer_retries):
            try:
                success, result = DirectCommandExecutor.instrument_controller.execute_call_command(command)
                if success:
                    return success, result
                if "RSRC_NFOUND" in str(result) or "resource not present" in str(result).lower():
                    print(f"⚠️ 检测到资源未找到，尝试完全重建控制器 (尝试 {outer+1}/{max_outer_retries})")
                    DirectCommandExecutor.cleanup()
                    time.sleep(2)
                    if DirectCommandExecutor.initialize():
                        continue
                    else:
                        return False, f"重建控制器失败: {result}"
                else:
                    return False, result
            except Exception as e:
                error_msg = str(e)
                if "not enough values to unpack" in error_msg:
                    return False, f"内部解包错误，可能控制器返回异常: {error_msg}"
                return False, f"执行命令异常: {error_msg}"
        return False, "达到最大外部重试次数"

    @staticmethod
    def cleanup():
        if DirectCommandExecutor.instrument_controller:
            DirectCommandExecutor.instrument_controller.disconnect()
        DirectCommandExecutor.instrument_controller = None
        print("🧹 资源已清理")


class _PyvisaCloseNoiseFilter(logging.Filter):
    def filter(self, record):
        try:
            return 'Error closing VISA link' not in record.getMessage()
        except Exception:
            return True

def _install_pyvisa_close_noise_filter():
    pv = logging.getLogger('pyvisa')
    if not any(isinstance(flt, _PyvisaCloseNoiseFilter) for flt in pv.filters):
        pv.addFilter(_PyvisaCloseNoiseFilter())


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
            _install_pyvisa_close_noise_filter()
            self._file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
            self._file_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            )
            logging.basicConfig(
                level=log_level,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                handlers=[self._file_handler]
            )
            self.logger = logging.getLogger('SourenToolSet')
            print(f"📁 日志文件: {self.log_file}")
        else:
            self._file_handler = None
            self.logger = None
            print("📁 日志功能已禁用")

    def close(self):
        handler = getattr(self, '_file_handler', None)
        if handler is not None:
            try:
                logging.getLogger().removeHandler(handler)
                handler.close()
            except Exception:
                pass
            self._file_handler = None

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
        self._query_expected_patterns = []
        self._extracted_index_map = {}

    def reset(self):
        self.step_counter = 0
        self.execution_details = []
        self.extracted_data = []
        self._pending_check = None
        self._current_command_is_query = False
        self.query_expected_map.clear()
        self._query_expected_patterns.clear()
        self._extracted_index_map.clear()

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
            print("✅ 添加check_phone_at函数到脚本环境")

        class APWrapper:
            def __init__(self, executor): self.executor = executor
            def send(self, command, extract_index=None, should_extract=False, chart_title=None, x_label=None,
                     separate_loop_chart=False, keep_duplicate_in_loop=False, status=None, record_step=True):
                self.executor._current_command_is_query = False
                return self.executor._execute_ap_command(command, extract_index, should_extract, chart_title,
                                                         x_label, separate_loop_chart, keep_duplicate_in_loop, status,
                                                         record_step)
            def query(self, command, extract_index=None, should_extract=False, chart_title=None, x_label=None,
                      separate_loop_chart=False, keep_duplicate_in_loop=False, status=None):
                self.executor._current_command_is_query = True
                return self.executor._execute_ap_command(command, extract_index, should_extract, chart_title,
                                                         x_label, separate_loop_chart, keep_duplicate_in_loop, status)
            def sleep(self, ms):
                self.executor._current_command_is_query = False
                return self.executor._execute_sleep(ms, self.executor.step_counter+1)
            def tag_last(self, status, count=1):
                return self.executor._tag_last_extracted(status, count)
            def check(self, content, passed, detail=None, attempts=1):
                return self.executor._record_check(content, passed, detail, attempts)

        ap_wrapper = APWrapper(self)
        local_env['ap'] = ap_wrapper

        def my_sleep_wrapper(seconds):
            ms = int(seconds * 1000)
            return self._execute_ap_command(f"SLEEP {ms}", None, False)
        local_env['my_sleep'] = my_sleep_wrapper

        if common:
            try:
                common.ap = ap_wrapper
                common.my_sleep = my_sleep_wrapper
                common._active_executor = self  # 登记当前执行器,供 common 占位 ap 的 check/tag_last 转调
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
            script_name = os.path.splitext(os.path.basename(file_path))[0]
            if script_name in LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS:
                print(f"[跳过] {script_name} 在 LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS 中，跳过 LTE_MeasurementReport 检查")
            elif parameters and parameters.get('lte_band'):
                self._append_lte_measurement_report_check()
            return True, {
                "success": True,
                "execution_details": self.execution_details,
                "extracted_data": self.extracted_data,
                "step_count": self.step_counter,
                "parameters": parameters,
                "loop_iteration": self.current_loop_iteration,
                "loop_count": self.total_loop_count
            }
        except KeyboardInterrupt:
            self._finalize_pending_check(forced=True)
            print("\n⏹️ 脚本执行被用户中断，已保存部分执行数据")
            return False, {
                "success": False,
                "error": "用户中断",
                "interrupted": True,
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
            self.query_expected_map.clear()
            nodes = []
            for node in ast.walk(tree):
                nodes.append(node)
            nodes.sort(key=lambda n: getattr(n, 'lineno', 0))
            last_query_cmd = None
            last_query_lineno = 0
            for node in nodes:
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                    if node.func.attr == 'query' and isinstance(node.func.value, ast.Name) and node.func.value.id == 'ap':
                        cmd = self._get_command_from_call(node, script_content)
                        if cmd:
                            last_query_cmd = cmd
                            last_query_lineno = getattr(node, 'lineno', 0)
                elif isinstance(node, ast.If):
                    expected = self._extract_expected_from_condition(node.test)
                    if expected and last_query_cmd is not None:
                        if_lineno = getattr(node, 'lineno', 0)
                        if if_lineno > last_query_lineno and if_lineno - last_query_lineno <= 50:
                            if last_query_cmd not in self.query_expected_map:
                                self.query_expected_map[last_query_cmd] = expected
                                print(f"📌 动态提取预期: {last_query_cmd} -> {expected} (行 {last_query_lineno} -> {if_lineno})")
                                if '{' in last_query_cmd and '}' in last_query_cmd:
                                    try:
                                        parts = re.split(r'\{[^}]*\}', last_query_cmd)
                                        pattern = '^' + '.*?'.join(re.escape(p) for p in parts) + '$'
                                        self._query_expected_patterns.append(
                                            (re.compile(pattern), expected))
                                        print(f"📌 (f-string)预期正则: {pattern} -> {expected}")
                                    except Exception as e:
                                        print(f"⚠️  构建 f-string 预期正则失败: {e}")
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
        elif isinstance(node, ast.And):
            for val in node.values:
                expected = self._extract_expected_from_condition(val)
                if expected:
                    return expected
        elif isinstance(node, ast.Or):
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
                if isinstance(arg, ast.JoinedStr):
                    template = ''
                    for v in arg.values:
                        if isinstance(v, ast.Constant) and isinstance(v.value, str):
                            template += v.value
                        elif hasattr(ast, 'Str') and isinstance(v, ast.Str):
                            template += v.s
                        else:
                            template += '{}'
                    if template:
                        return template
                if hasattr(call_node, 'lineno') and hasattr(call_node, 'col_offset'):
                    lines = script_content.splitlines()
                    if call_node.lineno <= len(lines):
                        line = lines[call_node.lineno - 1]
                        import re
                        match = re.search(r'ap\.query\(\s*[a-zA-Z]*[\'\"]([^\'\"]+)[\'\"]', line)
                        if match:
                            return match.group(1)
        except:
            pass
        return None

    def _tag_last_extracted(self, status, count=1):
        tagged = 0
        for item in reversed(self.extracted_data):
            if tagged >= count:
                break
            item['status'] = status
            tagged += 1
        return tagged

    def _execute_ap_command(self, command: str, extract_index=None, should_extract=False,
                            chart_title=None, x_label=None, separate_loop_chart=False,
                            keep_duplicate_in_loop=False, status=None, record_step=True) -> str:
        step_start_time = time.time()
        if isinstance(command, str) and command.upper().startswith("SLEEP"):
            import re
            sleep_match = re.search(r'SLEEP\s+(\d+)', command.upper())
            if sleep_match:
                sleep_ms = int(sleep_match.group(1))
                if self._pending_check:
                    self._pending_check['duration'] += sleep_ms / 1000
                    self._pending_check['total_sleep'] = self._pending_check.get('total_sleep', 0) + sleep_ms
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
                loop = self.current_loop_iteration
                if keep_duplicate_in_loop:
                    extracted_item = {
                        "step": step_num,
                        "command": command,
                        "extracted_data": extracted,
                        "loop_iteration": loop,
                        "chart_title": chart_title,
                        "x_label": x_label,
                        "separate_loop_chart": separate_loop_chart,
                        "status": status
                    }
                    self.extracted_data.append(extracted_item)
                    print(f"  📊 提取数据(保留重复): {extracted} (标题: {chart_title}, 横坐标: {x_label}, 单独图表: {separate_loop_chart})")
                else:
                    key = (loop, command, chart_title, x_label)
                    if key in self._extracted_index_map:
                        idx = self._extracted_index_map[key]
                        self.extracted_data[idx] = {
                            "step": step_num,
                            "command": command,
                            "extracted_data": extracted,
                            "loop_iteration": loop,
                            "chart_title": chart_title,
                            "x_label": x_label,
                            "separate_loop_chart": separate_loop_chart,
                            "status": status
                        }
                        print(f"  📊 更新已有提取数据: {extracted} (标题: {chart_title}, 横坐标: {x_label})")
                    else:
                        self._extracted_index_map[key] = len(self.extracted_data)
                        self.extracted_data.append({
                            "step": step_num,
                            "command": command,
                            "extracted_data": extracted,
                            "loop_iteration": loop,
                            "chart_title": chart_title,
                            "x_label": x_label,
                            "separate_loop_chart": separate_loop_chart,
                            "status": status
                        })
                        print(f"  📊 提取数据: {extracted} (标题: {chart_title}, 横坐标: {x_label})")

        if self._current_command_is_query:
            expected = self.query_expected_map.get(command, None)
            if expected is None and self._query_expected_patterns:
                for pat, exp in self._query_expected_patterns:
                    if pat.match(command):
                        expected = exp
                        break
            expected_matched = False
            if expected and clean_result:
                expected_clean = expected.strip().strip('"').strip("'")
                result_clean = clean_result.strip().strip('"').strip("'")
                if expected_clean == result_clean:
                    expected_matched = True

            if self._pending_check is None or self._pending_check['command'] != command:
                self._finalize_pending_check()
                self._pending_check = {
                    'step': step_num, 'command': command, 'attempts': 1,
                    'cmd_success': success,
                    'expected': expected,
                    'expected_matched': expected_matched,
                    'first_result': result, 'last_result': result,
                    'start_time': step_start_time, 'duration': 0.0,
                    'total_sleep': 0, 'extract_index': extract_index,
                    'extracted_data': extracted
                }
            else:
                self._pending_check['attempts'] += 1
                self._pending_check['last_result'] = result
                self._pending_check['duration'] = time.time() - self._pending_check['start_time']
                if expected_matched:
                    self._pending_check['expected_matched'] = True
                if should_extract and extract_index is not None and extracted is not None:
                    self._pending_check['extracted_data'] = extracted

            if expected is not None and expected_matched:
                self._finalize_pending_check()
        else:
            self._finalize_pending_check()
            # record_step=False: 只抽取数据(供图表/tag_last),不在详细执行记录里留 Command。
            # 用于"取值+ap.check 手动判定"的场景,避免同一命令既有 Command 又有 Check 的重复行。
            if record_step:
                self._record_command_step(step_num, command, success, result, step_start_time, extract_index, extracted)

        return result

    def _finalize_pending_check(self, forced=False):
        if not self._pending_check:
            return
        p = self._pending_check
        step_num, command = p['step'], p['command']
        attempts, cmd_success = p['attempts'], p['cmd_success']
        expected = p.get('expected')
        expected_matched = p.get('expected_matched', False)
        last_result, duration = p['last_result'], time.time() - p['start_time']
        if not cmd_success:
            status = "failed"
            status_msg = f"第{attempts}次查询命令执行失败"
        else:
            if expected is not None and not expected_matched:
                status = "failed"
                status_msg = f"第{attempts}次查询命令执行失败，预期值“{expected}”"
            else:
                status = "success"
                status_msg = f"第{attempts}次查询命令执行成功"
        detail = {
            "step": step_num, "type": "Check", "function": "unknown",
            "content": command, "status": status, "duration": duration,
            "result": f"{status_msg} | 末次结果: {last_result[:200]}" if last_result else status_msg,
            "start_time": p['start_time'], "end_time": time.time(),
            "loop_iteration": self.current_loop_iteration,
            "loop_count": self.total_loop_count,
            "attempts": attempts,
            "expected": expected,
            "expected_matched": expected_matched,
            "extracted_data": p.get('extracted_data')
        }
        self.execution_details.append(detail)
        print(f"  ✅ Check合并完成 - 尝试{attempts}次, 状态: {status}")
        self._pending_check = None

    def _append_lte_measurement_report_check(self):
        import glob
        # Ubuntu 服务端不可达时,基站没起、signal 日志也没抓到,不做 LTE_MeasurementReport 检查。
        try:
            if common and not common.remote_client.RemoteClient.ping(common.host, common.port):
                print("[跳过] Ubuntu 服务端不可达，跳过 LTE_MeasurementReport 检查")
                return
        except Exception as e:
            print(f"[跳过] 无法确认 Ubuntu 服务端状态({e})，跳过 LTE_MeasurementReport 检查")
            return
        start = time.time()
        self.step_counter += 1
        step_num = self.step_counter
        try:
            candidates = glob.glob(os.path.join(os.getcwd(), '*_signal_message_log.txt'))
        except Exception:
            candidates = []
        codes = []
        differ = False
        if candidates:
            signal_file = max(candidates, key=os.path.getmtime)
            try:
                with open(signal_file, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                codes = re.findall(r'LTE_MeasurementReport\s*\(code:\s*([0-9A-Fa-f]+)\)', content)
                differ = len({len(c) for c in codes}) > 1
            except Exception as e:
                print(f"[WARN] 读取/解析 signal 日志失败: {e}")
        else:
            print("[WARN] 未找到 signal_message_log, 跳过 LTE_MeasurementReport 检查")

        if differ:
            status, status_msg, last_result = "success", "第1次查询命令执行成功", "differ"
        else:
            status, status_msg, last_result = "failed", "第1次查询命令执行失败", "same"

        detail = {
            "step": step_num, "type": "Check", "function": "unknown",
            "content": "LTE_MeasurementReport", "status": status,
            "duration": time.time() - start,
            "result": f"{status_msg} | 末次结果: {last_result}",
            "start_time": start, "end_time": time.time(),
            "loop_iteration": self.current_loop_iteration,
            "loop_count": self.total_loop_count,
            "attempts": 1,
            "expected": None,
            "expected_matched": differ,
            "extracted_data": None
        }
        self.execution_details.append(detail)
        print(f"  ✅ LTE_MeasurementReport 检查完成 - 共{len(codes)}条, "
              f"长度{'不一致(differ)' if differ else '一致(same)'}, 状态: {status}")

    def _record_check(self, content, passed, detail=None, attempts=1):
        """供测试脚本通过 ap.check() 手动追加一条 Check 记录到详细执行记录,
        用于框架自动 == 判定覆盖不到的场景(如数值容差比较)。
        passed=True -> success, False -> failed。
        attempts: 实际查询次数,体现在"第N次查询命令执行..."。"""
        self._finalize_pending_check()  # 先刷掉可能存在的待合并 query check,保证顺序
        start = time.time()
        self.step_counter += 1
        step_num = self.step_counter
        if passed:
            status, status_msg = "success", f"第{attempts}次查询命令执行成功"
        else:
            status, status_msg = "failed", f"第{attempts}次查询命令执行失败"
        record = {
            "step": step_num, "type": "Check", "function": "unknown",
            "content": content, "status": status, "duration": time.time() - start,
            "result": f"{status_msg} | 末次结果: {detail}" if detail is not None else status_msg,
            "start_time": start, "end_time": time.time(),
            "loop_iteration": self.current_loop_iteration,
            "loop_count": self.total_loop_count,
            "attempts": attempts, "expected": None, "expected_matched": passed,
            "extracted_data": None
        }
        self.execution_details.append(record)
        print(f"  ✅ 手动Check记录 - {content}: {status} ({detail})")
        return passed

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