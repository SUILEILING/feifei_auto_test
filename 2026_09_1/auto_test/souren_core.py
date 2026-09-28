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
import souren_config

try:
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    import common
    print("✅ common模块导入成功")
except ImportError as e:
    print(f"⚠️  导入common模块失败: {e}")
    common = None

STOP_EVENT = threading.Event()


def request_stop():
    STOP_EVENT.set()


def clear_stop():
    STOP_EVENT.clear()


def is_stop_requested() -> bool:
    return STOP_EVENT.is_set()


class StopRequested(KeyboardInterrupt):
    pass


class GnbProcessDead(BaseException):
    pass


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
    # 内存里只保留最近若干行(deque 自动丢最旧的)用于兜底/显示;
    # 完整 SCPI 记录改为“流式落盘”到每个 case 的文件,长挂测每条都写、绝不丢,内存也不涨。
    SCPI_LOG_MAX_LINES = 200000  # 内存 tail 上限(仅当没有落盘文件时才作为兜底来源)
    scpi_comm_log = collections.deque(maxlen=SCPI_LOG_MAX_LINES)
    _scpi_fh = None          # 当前 case 的 SCPI 落盘文件句柄
    _scpi_path = None        # 当前 case 的 SCPI 落盘文件路径
    # 熔断截止时间(类级别: 控制器实例被重建也不丢失)：自愈彻底失败后快速失败到此时刻
    _down_until = 0.0

    @classmethod
    def open_scpi_log(cls, path):
        cls.close_scpi_log()
        try:
            cls._scpi_fh = open(path, 'a', encoding='utf-8')
            cls._scpi_path = path
        except Exception as e:
            print(f"⚠️ 打开 SCPI 落盘文件失败(退回内存记录): {e}")
            cls._scpi_fh = None
            cls._scpi_path = None
        return cls._scpi_path

    @classmethod
    def close_scpi_log(cls):
        fh = cls._scpi_fh
        cls._scpi_fh = None
        path = cls._scpi_path
        cls._scpi_path = None
        if fh is not None:
            try:
                fh.flush()
                fh.close()
            except Exception:
                pass
        return path

    @classmethod
    def _log_scpi(cls, direction, text):
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S:%f')[:-3]
        line = f"{ts}   {direction} {text}"
        fh = cls._scpi_fh
        if fh is not None:
            try:
                fh.write(line + "\n")
                fh.flush()
            except Exception:
                pass
        cls.scpi_comm_log.append(line)

    def __init__(self, device_address: str = None):
        if device_address is None:
            import souren_config as _sc
            device_address = getattr(_sc, "INSTRUMENT_ADDRESS", None) or INSTRUMENT_ADDRESS
        self.device_address = device_address
        self.rm = None
        self.instrument = None
        self.connected = False
        self.timeout = 30000
        self.max_retries = 3
        self.reconnect_attempts = 0
        self.max_reconnect_attempts = 2
        self._last_io_time = 0.0
        self.idle_heartbeat_sec = 30
        self._last_conn_error = ""

    def _create_resource_manager(self):
        errors = []
        for backend in ("", "@py"):
            try:
                self.rm = pyvisa.ResourceManager(backend) if backend else pyvisa.ResourceManager()
                if backend:
                    print(f"✅ 使用 pyvisa 后端: {backend}")
                return True
            except Exception as e:
                label = backend or "默认(NI-VISA)"
                errors.append(f"[{label}] {type(e).__name__}: {e}")
                print(f"❌ 创建 ResourceManager 失败 {label}: {e}")
        self._rm_error = "；".join(errors)
        return False

    def connect(self) -> Tuple[bool, str]:
        if not self._create_resource_manager():
            detail = getattr(self, "_rm_error", "")
            return False, (
                "无法创建 VISA ResourceManager（缺少 VISA 后端）。\n"
                "请安装其一：\n"
                "  · NI-VISA 运行库（推荐，官网下载安装）\n"
                "  · 或 Python 环境执行: pip install pyvisa-py zeroconf\n"
                f"原始错误: {detail}"
            )
        try:
            print(f"🔌 尝试连接仪器: {self.device_address}")
            self.instrument = self.rm.open_resource(self.device_address)
            self.instrument.timeout = self.timeout
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            self._enable_tcp_keepalive()
            idn = self.instrument.query('*IDN?').strip()
            print(f"✅ 仪器连接成功: {idn}")
            self.connected = True
            self.reconnect_attempts = 0
            self._last_io_time = time.time()
            return True, f"连接成功: {idn}"
        except Exception as e:
            addr = self.device_address or ""
            ip = ""
            try:
                parts = addr.split("::")
                if len(parts) >= 2:
                    ip = parts[1]
            except Exception:
                pass
            hint = (
                f"仪器连接失败\n"
                f"  当前地址: {addr}\n"
                f"  目标IP: {ip or '未知'}\n"
                f"  请检查:\n"
                f"    1) 仪器/基站是否已开机\n"
                f"    2) 网线是否插好、与电脑是否同一网段\n"
                f"    3) 界面“全局配置”里的 IP 是否与仪器实际 IP 一致\n"
                f"    4) 能否 ping 通该 IP (cmd: ping {ip or '<IP>'})\n"
                f"    5) 是否已安装 NI-VISA 运行库\n"
                f"  原始错误: {str(e)}"
            )
            error_msg = hint
            print(f"❌ {error_msg}")
            self.connected = False
            self.instrument = None
            return False, error_msg

    def _enable_tcp_keepalive(self):
        try:
            try:
                from pyvisa import constants as _c
                self.instrument.set_visa_attribute(
                    _c.VI_ATTR_TCPIP_KEEPALIVE, True
                )
            except Exception as e:
                print(f"⚠️ 设置 VISA keepalive 属性失败(忽略): {e}")

            sock = self._get_underlying_socket()
            if sock is not None:
                import socket as _s
                sock.setsockopt(_s.SOL_SOCKET, _s.SO_KEEPALIVE, 1)
                if hasattr(_s, "SIO_KEEPALIVE_VALS"):
                    sock.ioctl(_s.SIO_KEEPALIVE_VALS, (1, 20000, 3000))
                else:
                    for opt, val in (("TCP_KEEPIDLE", 20), ("TCP_KEEPINTVL", 3),
                                     ("TCP_KEEPCNT", 5)):
                        if hasattr(_s, opt):
                            sock.setsockopt(_s.IPPROTO_TCP, getattr(_s, opt), val)
                print("✅ 已开启 TCP keepalive(空闲20s探测)")
        except Exception as e:
            print(f"⚠️ 开启 TCP keepalive 失败(忽略): {e}")

    def _get_underlying_socket(self):
        try:
            visalib = getattr(self.instrument, "visalib", None)
            sessions = getattr(visalib, "sessions", None)
            sess = None
            if sessions:
                sess = sessions.get(self.instrument.session)
            if sess is None:
                return None
            iface = getattr(sess, "interface", None)
            sock = getattr(iface, "sock", None)
            if sock is None and hasattr(iface, "setsockopt"):
                sock = iface
            return sock
        except Exception:
            return None

    def disconnect(self):
        if self.instrument:
            try:
                self.instrument.close()
                print("📴 仪器连接已关闭")
            except:
                pass
            finally:
                self.instrument = None
        if self.rm is not None:
            try:
                self.rm.close()
            except Exception:
                pass
            finally:
                self.rm = None
        self.connected = False
    def reconnect(self, retries: int = None, backoff: float = 2.0) -> bool:
        if retries is None:
            try:
                import souren_config as _sc
                retries = getattr(_sc, "RECONNECT_RETRIES", 15)
                backoff_max = getattr(_sc, "RECONNECT_BACKOFF_MAX", 30.0)
            except Exception:
                retries, backoff_max = 15, 30.0
        else:
            backoff_max = 15.0
        self._last_conn_error = ""
        self._last_recovery_info = None
        _t0 = time.time()
        for i in range(1, retries + 1):
            print(f"🔄 第 {i}/{retries} 次尝试重新连接仪器 (重建资源管理器)...")
            self.disconnect()
            self.rm = None
            time.sleep(1)
            success, msg = self.connect()
            if success:
                print(f"✅ 重新连接成功 (第 {i} 次)")
                VisaInstrumentController._down_until = 0.0
                self._last_recovery_info = {
                    'attempts': i, 'remote_restart': False,
                    'elapsed': time.time() - _t0,
                }
                self._log_recovery_summary()
                return True
            self._last_conn_error = msg
            print(f"❌ 第 {i}/{retries} 次重连失败: {msg}")
            if i < retries:
                wait = min(backoff * i, backoff_max)
                print(f"⏳ 等待 {wait:.0f}s 后再试...")
                time.sleep(wait)

        if self._try_remote_instrument_restart():
            for i in range(1, 4):
                print(f"🚑 仪器软件重启后第 {i}/3 次尝试连接...")
                self.disconnect()
                self.rm = None
                time.sleep(2)
                success, msg = self.connect()
                if success:
                    print("✅ 远程重启仪器后连接恢复")
                    VisaInstrumentController._down_until = 0.0
                    self._last_recovery_info = {
                        'attempts': retries, 'remote_restart': True,
                        'restart_attempts': i, 'elapsed': time.time() - _t0,
                    }
                    self._log_recovery_summary()
                    return True
                self._last_conn_error = msg
                time.sleep(10)

        try:
            import souren_config as _sc
            cooldown = getattr(_sc, "RECONNECT_DOWN_COOLDOWN", 120)
        except Exception:
            cooldown = 120
        VisaInstrumentController._down_until = time.time() + cooldown
        print(f"⛔ 自愈失败，进入熔断：{cooldown}s 内命令将快速失败，之后自动再试")
        return False

    def _log_recovery_summary(self):
        """
        把这次断连自愈的过程醒目地记录出来(控制台 + SCPI 落盘日志 + SourenToolSet 日志),
        说明是"怎么救回来的": 重连了几次、是否动用了远程重启仪器软件、总共花了多久。
        长挂测里 OOM 杀掉 nr-softmodem 等场景, 事后翻日志一眼能看懂发生过什么。
        """
        info = getattr(self, '_last_recovery_info', None)
        if not info:
            return
        try:
            if info.get('remote_restart'):
                how = (f"本地重连 {info.get('attempts', '?')} 次均失败 → "
                       f"远程重启仪器软件(DIAG:restart) → 重启后第 {info.get('restart_attempts', '?')}/3 次连接成功")
            else:
                how = f"本地重连第 {info.get('attempts', '?')} 次成功(重建VISA会话)"
            line = (f"🚑 [断连自愈] 连接已恢复 | 方式: {how} | "
                    f"总耗时 {info.get('elapsed', 0):.0f}s | 触发原因: {getattr(self, '_last_conn_error', '') or '连接中断'}")
            print(line)
            # 写进 SCPI 落盘日志(与命令时间线放一起, 方便对照哪段命令间隔是自愈耗掉的)
            VisaInstrumentController._log_scpi("Info--", line)
            try:
                logging.getLogger('SourenToolSet').warning(line)
            except Exception:
                pass
        except Exception:
            pass

    def _try_remote_instrument_restart(self) -> bool:
        try:
            import souren_config as _sc
            if not getattr(_sc, "RECONNECT_REMOTE_RESTART", True):
                return False
            from lib import remote_client as _rc
            host = _sc.DEFAULT_IP
            port = _sc.REMOTE_SERVER_PORT
            if not _rc.RemoteClient.ping(host, port):
                print("[跳过] Ubuntu 服务端不可达，无法远程重启仪器软件")
                return False
            print("🚑 重连彻底失败，尝试通过远程服务重启仪器软件...")
            client = _rc.RemoteClient(host, port)
            ok = client._send_command_and_check_ok("DIAG:restart", "[OK]")
            if ok:
                print("🚑 远程重启命令已发出，等待 30s 仪器软件恢复...")
                time.sleep(30)
                return True
            print("⚠️ 远程重启命令未确认成功")
            return False
        except Exception as e:
            print(f"⚠️ 远程重启仪器软件异常: {e}")
            return False

    @staticmethod
    def _is_recoverable_conn_error(error_msg: str) -> bool:
        if not error_msg:
            return False
        msg = str(error_msg).lower()

# 码	系统名	         中文含义	               什么时候出现	                                         谁的锅
# 10053	WSAECONNABORTED	软件导致连接中止	      本地协议栈主动放弃连接(超时/资源问题触发)	               本地侧
# 10054	WSAECONNRESET	远程主机强迫关闭连接	  对端直接发 TCP RST 把连接掐了                           对端(仪器)
# 10060	WSAETIMEDOUT	连接超时	             发了包对端没响应,等到超时	                             网络/对端无响应
# 10061	WSAECONNREFUSED	连接被拒绝	             对端端口没在监听(服务没起/端口错)	                    对端没开门

        keywords = [
            '10053', '10054', '10060', '10061',
            'connection', 'reset', 'broken pipe', 'aborted',
            '远程主机', '强迫关闭', '连接', '中止',
            'vi_error_rsrc_nfound', 'resource not present',
            'vi_error_conn_lost', 'vi_error_inv_object',
            # 长挂测最常见的“连接被静默回收/仪器瞬时无响应”：
            # VI_ERROR_IO(-1073807298) I/O 错误、VI_ERROR_TMO(-1073807339) 超时、
            # 以及会话被内部关闭后再用引发的 NoneType。这些都应触发重连而不是直接判死。
            'vi_error_io', '-1073807298', 'i/o error',
            'could not perform operation',
            'vi_error_tmo', '-1073807339', 'timeout', 'timed out',
            "'nonetype'", 'nonetype',
        ]
        return any(k in msg for k in keywords)

    def _idle_heartbeat_if_needed(self):
        if not self.idle_heartbeat_sec or not self.connected or not self.instrument:
            return
        idle = time.time() - (self._last_io_time or 0)
        if idle < self.idle_heartbeat_sec:
            return
        try:
            print(f"💓 空闲 {idle:.0f}s，发送心跳探活 (*IDN?)...")
            self.instrument.query('*IDN?')
            self._last_io_time = time.time()
            print("💓 心跳正常，连接有效")
        except Exception as e:
            print(f"💤 空闲期间连接疑似断开，具体错误: {e}")
            print("🔄 自动重连中...")
            if self.reconnect():
                print("✅ 重连成功，继续执行")
            else:
                reason = getattr(self, "_last_conn_error", "")
                print(f"⚠️ 重连失败，触发原因[{e}] 重连原因[{reason}]")

    def execute_scpi_command(self, command: str) -> Tuple[bool, str]:
        if time.time() < self._down_until:
            remain = int(self._down_until - time.time())
            return False, f"仪器不可达(熔断冷却中,{remain}s后自动重试)"
        if not self.connected or not self.instrument:
            if not self.reconnect():
                reason = getattr(self, "_last_conn_error", "") or "未知原因"
                return False, f"仪器未连接，重连失败 | 具体原因: {reason}"

        self._idle_heartbeat_if_needed()

        command = command.strip()
        if not command:
            return False, "空的命令"

        for attempt in range(self.max_retries):
            try:
                if '?' in command:
                    VisaInstrumentController._log_scpi("Send->", command)
                    result = self.instrument.query(command).strip()
                    VisaInstrumentController._log_scpi("Recv<-", result)
                    self._last_io_time = time.time()
                    return True, result
                else:
                    VisaInstrumentController._log_scpi("Send->", command)
                    self.instrument.write(command)
                    time.sleep(0.1)
                    self._last_io_time = time.time()
                    return True, "命令执行成功"
            except pyvisa.errors.VisaIOError as e:
                error_msg = str(e)
                self.connected = False
                if self._is_recoverable_conn_error(error_msg):
                    print(f"⚠️ 检测到连接/资源错误 (尝试 {attempt+1}/{self.max_retries}): {error_msg}")
                    if attempt < self.max_retries - 1:
                        if self.reconnect():
                            continue
                        else:
                            reason = getattr(self, "_last_conn_error", "")
                            return False, f"断连[{error_msg}] 重连失败[{reason}]"
                    else:
                        if self.reconnect():
                            try:
                                if '?' in command:
                                    VisaInstrumentController._log_scpi("Send->", command)
                                    result = self.instrument.query(command).strip()
                                    VisaInstrumentController._log_scpi("Recv<-", result)
                                    self._last_io_time = time.time()
                                    return True, result
                                else:
                                    VisaInstrumentController._log_scpi("Send->", command)
                                    self.instrument.write(command)
                                    time.sleep(0.1)
                                    self._last_io_time = time.time()
                                    return True, "命令执行成功"
                            except Exception as e2:
                                return False, f"仪器通信错误，重连后仍失败: {e2}"
                        return False, f"仪器通信错误，重试失败: {error_msg}"
                else:
                    if self.reconnect():
                        continue
                    return False, f"仪器通信错误: {error_msg}"
            except Exception as e:
                error_msg = str(e)
                self.connected = False
                if self._is_recoverable_conn_error(error_msg):
                    print(f"⚠️ 检测到连接/资源错误 (尝试 {attempt+1}/{self.max_retries}): {error_msg}")
                    if attempt < self.max_retries - 1:
                        if self.reconnect():
                            continue
                        else:
                            reason = getattr(self, "_last_conn_error", "")
                            return False, f"断连[{error_msg}] 重连失败[{reason}]"
                    else:
                        if self.reconnect():
                            continue
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
            old = DirectCommandExecutor.instrument_controller
            if old is not None:
                try:
                    old.disconnect()
                except Exception:
                    pass
            ctrl = VisaInstrumentController()
            DirectCommandExecutor.instrument_controller = ctrl
            success, message = ctrl.connect()
            if not success:
                print(f"⚠️ 首次连接失败，启动重试: {message}")
                success = ctrl.reconnect(retries=3, backoff=1.0)
                if success:
                    message = "重连成功"
                else:
                    message = getattr(ctrl, "_last_conn_error", "") or message
            if success:
                print("✅ 仪器连接成功")
                if common:
                    common.setup_instrument_controller(ctrl)
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

        max_outer_retries = 3
        for outer in range(max_outer_retries):
            try:
                success, result = DirectCommandExecutor.instrument_controller.execute_call_command(command)
                if success:
                    return success, result
                if "熔断" in str(result):
                    return False, result
                res_l = str(result).lower()
                need_rebuild = (
                    "rsrc_nfound" in res_l or "resource not present" in res_l or
                    VisaInstrumentController._is_recoverable_conn_error(result) or
                    "重连失败" in str(result) or "断连" in str(result) or
                    "未连接" in str(result)
                )
                if need_rebuild:
                    print(f"⚠️ 检测到连接异常，尝试完全重建控制器 (尝试 {outer+1}/{max_outer_retries})")
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
            result_filename = f"case_results_{timestamp}.json"
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
            tmp_file = self.result_file + ".tmp"
            with open(tmp_file, 'w', encoding='utf-8') as f:
                json.dump(self.results, f, ensure_ascii=False, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_file, self.result_file)
            print(f"✅ 结果已保存到: {self.result_file}")
            return True
        except Exception as e:
            print(f"❌ 保存结果失败: {e}")
            try:
                _tmp = self.result_file + ".tmp"
                if os.path.exists(_tmp):
                    os.remove(_tmp)
            except Exception:
                pass
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
        self._parameters = None
        self._ue_disconnect_count = 0

    def reset(self):
        self.step_counter = 0
        self.execution_details = []
        self.extracted_data = []
        self._pending_check = None
        self._current_command_is_query = False
        self.query_expected_map.clear()
        self._query_expected_patterns.clear()
        self._extracted_index_map.clear()
        self._ue_disconnect_count = 0

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
        self._parameters = parameters or {}
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
            def check(self, content, passed, detail=None, attempts=1, status_msg=None):
                return self.executor._record_check(content, passed, detail, attempts, status_msg)
            def record_value(self, title, value, x_label=None):
                return self.executor._record_extracted_value(title, value, x_label)

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
                common._active_executor = self
                common._current_case = os.path.splitext(os.path.basename(file_path))[0]
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
            gnb_dead_msg = None
            try:
                if 'case_start' in local_env:
                    print("\n🔧 执行 case_start()...")
                    local_env['case_start']()
                if 'case_body' in local_env:
                    print("\n🔧 执行 case_body()...")
                    local_env['case_body']()
            except GnbProcessDead as e:
                gnb_dead_msg = str(e)
                print(f"\n⛔ 基站进程异常,跳过 case_body 剩余步骤,强制执行 case_clear: {gnb_dead_msg}")
            finally:
                if 'case_clear' in local_env:
                    print("\n🧹 执行 case_clear()...")
                    try:
                        local_env['case_clear']()
                    except GnbProcessDead as e2:
                        print(f"[WARN] case_clear 期间基站进程仍异常(忽略,已尽力收尾): {e2}")
            self._finalize_pending_check(forced=True)
            script_name = os.path.splitext(os.path.basename(file_path))[0]
            if gnb_dead_msg:
                print(f"[跳过] 基站进程异常提前结束,跳过 LTE_MeasurementReport 检查")
            elif script_name in LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS:
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
                "loop_count": self.total_loop_count,
                "gnb_process_dead": bool(gnb_dead_msg),
                "gnb_process_dead_msg": gnb_dead_msg,
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
                    if expected:
                        inline_call = self._find_query_call_in_node(node.test)
                        if inline_call is not None:
                            cmd = self._get_command_from_call(inline_call, script_content)
                            cmd_lineno = getattr(inline_call, 'lineno', getattr(node, 'lineno', 0))
                        else:
                            cmd = last_query_cmd
                            cmd_lineno = last_query_lineno
                        if cmd is not None:
                            if_lineno = getattr(node, 'lineno', 0)
                            if if_lineno >= cmd_lineno and if_lineno - cmd_lineno <= 50:
                                if cmd not in self.query_expected_map:
                                    self.query_expected_map[cmd] = expected
                                    print(f"📌 动态提取预期: {cmd} -> {expected} (行 {cmd_lineno} -> {if_lineno})")
                                    if '{' in cmd and '}' in cmd:
                                        try:
                                            parts = re.split(r'\{[^}]*\}', cmd)
                                            pattern = '^' + '.*?'.join(re.escape(p) for p in parts) + '$'
                                            self._query_expected_patterns.append(
                                                (re.compile(pattern), expected))
                                            print(f"📌 (f-string)预期正则: {pattern} -> {expected}")
                                        except Exception as e:
                                            print(f"⚠️  构建 f-string 预期正则失败: {e}")
        except Exception as e:
            print(f"⚠️  AST解析失败,动态预期提取将不可用: {e}")

    def _find_query_call_in_node(self, node):
        for sub in ast.walk(node):
            if (isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute)
                    and sub.func.attr == 'query'
                    and isinstance(sub.func.value, ast.Name) and sub.func.value.id == 'ap'):
                return sub
        return None

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

    def _maybe_check_gnb_on_ue_state(self, command, clean_result):
        try:
            if not isinstance(command, str) or not command.strip().upper().endswith(":UE:STATE?"):
                return
        except Exception:
            return

        is_connected = bool(clean_result) and "CONNECTED" in clean_result.upper()
        if is_connected:
            self._ue_disconnect_count = 0
            return

        self._ue_disconnect_count += 1
        every = getattr(souren_config, 'GNB_PROC_CHECK_EVERY', 30)
        if not every or every <= 0:
            return
        if self._ue_disconnect_count % every != 0:
            return

        if not common or not hasattr(common, 'check_gnb_processes_alive'):
            return
        try:
            all_alive, dead_keys, detail = common.check_gnb_processes_alive(self._parameters or {})
        except Exception as e:
            print(f"[WARN] 基站进程存活检查异常(忽略,继续等待): {e}")
            return

        if all_alive:
            print(f"ℹ️ UE 已连续 {self._ue_disconnect_count} 次未连接,但基站进程均存活 "
                  f"({detail}),继续等待")
            return

        # 进程已死: 先把当前挂起的未连接 Check 落账(记为 failed),再补一条明确的进程存活 Check,
        # 然后抛异常中断 case_body,由框架强制转入 case_clear。
        self._finalize_pending_check(forced=True)
        msg = (f"UE 已连续 {self._ue_disconnect_count} 次未连接,且检测到基站进程已退出: "
               f"{dead_keys} ({detail})")
        try:
            self._record_check("基站关键进程存活检查", False, detail=msg,
                               status_msg="进程已退出,提前结束当前case")
        except Exception:
            pass
        print(f"🛑 {msg} → 立即结束当前 case,进入 case_clear")
        raise GnbProcessDead(msg)

    def _execute_ap_command(self, command: str, extract_index=None, should_extract=False,
                            chart_title=None, x_label=None, separate_loop_chart=False,
                            keep_duplicate_in_loop=False, status=None, record_step=True) -> str:
        if STOP_EVENT.is_set():
            raise StopRequested("用户请求停止")
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
                self._interruptible_sleep(sleep_ms / 1000)
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

            # 框架级 UE 连接自愈: 在 pending_check 更新完毕后再检测,保证尝试次数已同步。
            # case 轮询 `...:UE:STATe?` 期望 Connected 时,每累计 N 次仍未连上就查一次基站进程,
            # 进程死了立即抛 GnbProcessDead 中断本 case 转 case_clear。
            self._maybe_check_gnb_on_ue_state(command, clean_result)
        else:
            self._finalize_pending_check()
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

    def _record_check(self, content, passed, detail=None, attempts=1, status_msg=None):
        self._finalize_pending_check()  
        start = time.time()
        self.step_counter += 1
        step_num = self.step_counter
        if passed:
            status = "success"
            default_msg = f"第{attempts}次查询命令执行成功"
        else:
            status = "failed"
            default_msg = f"第{attempts}次查询命令执行失败"
        msg = status_msg if status_msg is not None else default_msg
        record = {
            "step": step_num, "type": "Check", "function": "unknown",
            "content": content, "status": status, "duration": time.time() - start,
            "result": f"{msg} | 末次结果: {detail}" if detail is not None else msg,
            "start_time": start, "end_time": time.time(),
            "loop_iteration": self.current_loop_iteration,
            "loop_count": self.total_loop_count,
            "attempts": attempts, "expected": None, "expected_matched": passed,
            "extracted_data": None
        }
        self.execution_details.append(record)
        print(f"  ✅ 手动Check记录 - {content}: {status} ({detail})")
        return passed

    def _record_extracted_value(self, title, value, x_label=None):
        item = {
            "step": self.step_counter,
            "command": f"RECORD:{title}",
            "extracted_data": value,
            "loop_iteration": self.current_loop_iteration,
            "chart_title": title,
            "x_label": x_label,
            "separate_loop_chart": False,
            "status": None,
            "summary_only": True
        }
        self.extracted_data.append(item)
        print(f"  📊 记录汇总数据: {title} = {value}")
        return value

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

    def _interruptible_sleep(self, seconds: float):
        end = time.time() + max(0.0, seconds)
        while True:
            if STOP_EVENT.is_set():
                raise StopRequested("用户请求停止")
            remaining = end - time.time()
            if remaining <= 0:
                return
            time.sleep(min(0.1, remaining))

    def _execute_sleep(self, sleep_ms: int, step_num: int) -> str:
        print(f"😴 独立睡眠 {sleep_ms} 毫秒...")
        self._interruptible_sleep(sleep_ms / 1000)
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