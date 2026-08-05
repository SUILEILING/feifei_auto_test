from lib.var import *
from lib import remote_client
import souren_config
host = souren_config.DEFAULT_IP
port = souren_config.REMOTE_SERVER_PORT 

try:
    import adb_integration
    ADB_INTEGRATION_AVAILABLE = True
    print("✅ adb_integration 模块导入成功")
except ImportError as e:
    ADB_INTEGRATION_AVAILABLE = False
    print(f"⚠️  导入 adb_integration 模块失败: {e}")

try:
    from board_at_controller import find_fibocom_at_port, send_at_sequence
    AT_CONTROLLER_AVAILABLE = True
except ImportError as e:
    AT_CONTROLLER_AVAILABLE = False
    print(f"⚠️  导入 board_at_controller 模块失败: {e}")


_active_executor = None

class ScriptInstrumentController:
    def __init__(self, instrument_controller=None):
        self.instrument = instrument_controller
        self.logger = logging.getLogger('SourenCommon')
        self.last_result = None
        
    def send(self, command: str, extract_index: Optional[int] = None, should_extract: bool = False, record_step: bool = True) -> Union[str, float, None]:
        try:
            if not command:
                self.logger.warning("尝试发送空命令")
                return None
            
            if isinstance(command, str) and command.upper().startswith("SLEEP"):
                sleep_match = re.search(r'SLEEP\s+(\d+)', command.upper())
                if sleep_match:
                    sleep_ms = int(sleep_match.group(1))
                    return self.sleep_ms(sleep_ms)
            
            self.logger.info(f"发送命令: {command}")
            
            if '?' in command:
                success, result = self.instrument.execute_call_command(command)
                self.last_result = result if success else None
                
                if success:
                    self.logger.info(f"命令响应: {result}")
                    if should_extract and extract_index is not None:
                        extracted_value = self._extract_data(result, extract_index)
                        if extracted_value is not None:
                            self.logger.info(f"提取索引 {extract_index} 的数据: {extracted_value}")
                            return extracted_value
                        else:
                            self.logger.warning(f"无法从响应中提取索引 {extract_index} 的数据")
                            return None
                    else:
                        return result
                else:
                    self.logger.error(f"命令执行失败: {result}")
                    return None
            else:
                success, result = self.instrument.execute_call_command(command)
                self.last_result = result if success else None
                
                if success:
                    self.logger.info("命令执行成功")
                    return "命令执行成功"
                else:
                    self.logger.error(f"命令执行失败: {result}")
                    return None
                    
        except Exception as e:
            self.logger.error(f"发送命令时发生错误: {e}")
            return None
    
    def query(self, command: str) -> str:
        return self.send(command)

    def check(self, content, passed, detail=None, attempts=1):
        if _active_executor is not None:
            return _active_executor._record_check(content, passed, detail, attempts)
        status = "success" if passed else "failed"
        self.logger.info(f"[check] {content}: {status} (第{attempts}次) - {detail}")
        return passed

    def tag_last(self, status, count=1):
        if _active_executor is not None:
            return _active_executor._tag_last_extracted(status, count)
        self.logger.info(f"[tag_last] status={status}, count={count} (占位:无执行上下文,未标记)")
        return 0

    def sleep(self, seconds: float):
        self.logger.info(f"睡眠 {seconds} 秒")
        time.sleep(seconds)
    
    def sleep_ms(self, milliseconds: float):
        seconds = milliseconds / 1000
        self.logger.info(f"睡眠 {milliseconds} 毫秒 ({seconds:.2f} 秒)")
        time.sleep(seconds)
        return f"睡眠完成 ({milliseconds}毫秒)"
    
    def _extract_data(self, result_str: str, index: int) -> Optional[float]:
        try:
            result_str = str(result_str).strip()
            error_keywords = ["仪器通信错误", "VI_ERROR_TMO", "Timeout", "通信失败", "错误", "ERROR", "失败"]
            if any(keyword in result_str.upper() for keyword in [k.upper() for k in error_keywords]):
                self.logger.warning(f"检测到错误信息: {result_str[:100]}")
                return None
            
            if ',' in result_str:
                parts = [part.strip() for part in result_str.split(',')]
                if 0 <= index < len(parts):
                    try:
                        return float(parts[index])
                    except ValueError:
                        num_match = re.search(r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?', parts[index])
                        if num_match:
                            try:
                                return float(num_match.group())
                            except:
                                pass
            else:
                try:
                    return float(result_str)
                except ValueError:
                    pass
            
            pattern = r'[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?'
            matches = re.findall(pattern, result_str)
            if 0 <= index < len(matches):
                try:
                    return float(matches[index])
                except ValueError:
                    pass
            return None
        except Exception as e:
            self.logger.error(f"提取数据时发生错误: {e}")
            return None


class ADBController:
    def __init__(self):
        self.adb_controller = None
        self.device_id = None
        self._init_adb_controller()
    
    def _init_adb_controller(self):
        if ADB_INTEGRATION_AVAILABLE:
            try:
                self.adb_controller = adb_integration.ADBFlightModeController()
                if self.adb_controller and self.adb_controller.device_id:
                    self.device_id = self.adb_controller.device_id
                    print(f"✅ ADB控制器初始化成功,设备ID: {self.device_id}")
                else:
                    print("⚠️  ADB控制器初始化,但未找到设备")
            except Exception as e:
                print(f"❌ 初始化ADB控制器失败: {e}")
                self.adb_controller = None
        else:
            print("❌ adb_integration 模块不可用")
            self.adb_controller = None
    
    def check_connection(self) -> bool:
        if not self.adb_controller:
            print("❌ ADB控制器未初始化")
            return False
        try:
            if self.adb_controller.device_id:
                print(f"✅ 检测到设备: {self.adb_controller.device_id}")
                return True
            else:
                print("❌ 未检测到设备")
                return False
        except Exception as e:
            print(f"❌ 检查设备连接失败: {e}")
            return False
    
    def timed_flight_mode_control(self, wait_time: int = 5) -> bool:
        if not self.adb_controller:
            print("❌ ADB控制器未初始化")
            return False
        if not self.adb_controller.device_id:
            print("❌ 未检测到设备")
            return False
        try:
            print(f"📱 执行定时飞行模式控制，等待时间: {wait_time}秒")
            success = self.adb_controller.timed_flight_mode_control(wait_time)
            if success:
                print(f"✅ 定时飞行模式控制成功")
            else:
                print(f"❌ 定时飞行模式控制失败")
            return success
        except Exception as e:
            print(f"❌ 定时飞行模式控制失败: {e}")
            return False

class ATController:
    def __init__(self):
        self.port = None
        self.baudrate = 115200
        self.timeout = 3
    
    def execute_at_sequence(self) -> bool:
        print("\n📡 开始执行AT序列 (自动检测端口)...")
        port = find_fibocom_at_port()
        if not port:
            print("❌ 未检测到Fibocom AT端口,跳过AT序列")
            return False
        
        success, _ = send_at_sequence(port)
        return success

_adb_controller = None
_at_controller = None

def get_adb_controller() -> ADBController:
    global _adb_controller
    if _adb_controller is None:
        _adb_controller = ADBController()
    return _adb_controller

def get_at_controller() -> ATController:
    global _at_controller
    if _at_controller is None:
        _at_controller = ATController()
    return _at_controller


def check_phone_at(wait_time: int = 5) -> bool:
    print("\n" + "="*50)
    print("📱 开始设备类型检测和控制...")
    print("="*50)
    
    if ADB_INTEGRATION_AVAILABLE:
        adb = get_adb_controller()
        if adb.adb_controller and adb.check_connection():
            print("✅ 检测到手机设备，使用手机模式")
            success = adb.timed_flight_mode_control(wait_time)
            if success:
                print("✅ 手机飞行模式控制完成")
                return True
            else:
                print("❌ 手机飞行模式控制失败,切换到AT板模式")
        else:
            print("❌ 未检测到手机设备,使用AT板模式")
    else:
        print("❌ adb_integration 模块不可用,使用AT板模式")
    
    if AT_CONTROLLER_AVAILABLE:
        at = get_at_controller()
        success = at.execute_at_sequence()
        if success:
            print("✅ AT序列执行完成")
        else:
            print("❌ AT序列执行失败")
    else:
        print("❌ board_at_controller 模块不可用,无法执行AT序列")
    
    return False

def my_sleep(seconds: float):
    print(f"😴 睡眠 {seconds} 秒...")
    time.sleep(seconds)
    print(f"✅ 睡眠完成")


ap = ScriptInstrumentController()

def setup_instrument_controller(instrument_controller):
    global ap
    ap.instrument = instrument_controller
    print("✅ 仪器控制器已设置")


def config_line_loss(parameter):
    if parameter['lte_band']:
        ap.send("CONFigure:NSASa:SWITch NSA")
    else:
        ap.send("CONFigure:NSASa:SWITch SA")

    ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,{parameter['lineLoss1']},6000000000,{parameter['lineLoss1']}")
    ap.send("CONFigure:BASE:FDCorrection:SAVE")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX")

    if parameter['lineLoss3'] is not None:
        ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_3,100000000,{parameter['lineLoss3']},6000000000,{parameter['lineLoss3']}")
        ap.send("CONFigure:BASE:FDCorrection:SAVE")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,IO,RXTX")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,OUT,TX")

    ap.send("CONFigure:RFINdex:CLear:ALL")
    if parameter['lte_band'] is not None:
        ap.send("CONFigure:RFINdex:DL LTE,3")
        ap.send("CONFigure:RFINdex:UL LTE,3")
    ap.send("CONFigure:RFINdex:DL NR,1")
    ap.send("CONFigure:RFINdex:UL NR,1")
    ap.send("CONFigure:RFINdex1:CONNector IO")
    ap.send("CONFigure:RFINdex3:CONNector IO")
    ap.send("CONFigure:RFINdex:apply")

    my_sleep(12) if parameter['lte_band'] else my_sleep(6)


def config_cell_band(parameter):
    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {parameter['nr_band']}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{parameter['nr_bw']}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{parameter['scs']}")
    ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {parameter['range']}")
    ap.send("CONFigure:CELL1:NR:UE:MReport ON")

    if parameter['lte_band'] is not None:
        ap.send(f"CONFigure:CELL1:LTE:SIGN:BAND:DL OB{parameter['lte_band']}")
    if parameter['lte_bw'] is not None:
        ap.send(f"CONFigure:CELL1:LTE:SIGN:BWidth BW_{parameter['lte_bw']}")


DEFAULT_NR_SLOTS = {
    "DL": [
        {3:  {"TIND": 5, "MCS1": 4}},
        {4:  {"TIND": 4, "MCS1": 4}},
        {5:  {"TIND": 3, "MCS1": 4}},
        {6:  {"TIND": 2, "MCS1": 4}},
        {10: {"TIND": 8, "MCS1": 4}},
        {11: {"TIND": 7, "MCS1": 4}},
        {12: {"TIND": 6, "MCS1": 4}},
        {13: {"TIND": 5, "MCS1": 4}},
        {14: {"TIND": 4, "MCS1": 4}},
        {15: {"TIND": 3, "MCS1": 4}},
        {16: {"TIND": 2, "MCS1": 4}},
    ],
    "UL": [
        {8:  {"MCS1": 2}},
        {9:  {"MCS1": 2}},
        {18: {"MCS1": 2}},
        {19: {"MCS1": 2}},
    ],
}

DEFAULT_UL_RB_SPECIAL = {
    (100, 30): "67,135",
    (20, 15):  "25,50",
}

_SLOT_CTYPE = {"DL": "PDSCh", "UL": "PUSCh"}


def config_nr_slots(parameter, slots=None):
    slots = slots or DEFAULT_NR_SLOTS
    for direction in ("DL", "UL"):
        ctype = _SLOT_CTYPE[direction]
        for slot_item in slots.get(direction, []):
            for slot_num, fields in slot_item.items():
                ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT{slot_num}:CTYPe {ctype}")
                for field, value in fields.items():
                    ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT{slot_num}:{direction}:{field} {value}")

    rb_mode = parameter.get('rb_mode', 'Inner_Full')
    ul_slots = [n for item in slots.get("UL", []) for n in item.keys()]
    if rb_mode == 'Inner_Full':
        key = (parameter['nr_bw'], parameter['scs'])
        rb_value = DEFAULT_UL_RB_SPECIAL.get(key)
        if rb_value is not None:
            for slot in ul_slots:
                ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT{slot}:UL:RB {rb_value}")
        else:
            print(f"[WARN] Inner_Full 未找到 (nr_bw={parameter['nr_bw']}, scs={parameter['scs']}) 的 RB, 回退 SLOT:UPDate")
            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
    elif rb_mode == 'Outer_Full':
        ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
    else:
        print(f"[WARN] 未知 rb_mode='{rb_mode}', 按 Outer_Full 处理")
        ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")

    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
    ap.send('CONFigure:CELL1:NR:SIGN:DDETection:SWITch ON,1')
    my_sleep(2)


def config_lte_subframes(parameter):
    if parameter['lte_band'] is None:
        return
    ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame3:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame4:CTYPe PDSCh")
    resource_allocation_type = parameter.get('resource_allocation_type')
    if resource_allocation_type is not None:
        ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame3:DL:RATYpe TYPE{resource_allocation_type}")
        ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame4:DL:RATYpe TYPE{resource_allocation_type}")
    ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame8:CTYPe PUSCh")
    ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame9:CTYPe PUSCh")
    ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame:APPLy")
    my_sleep(2)
    ap.send("CONFigure:LTE:TXP:MSUBframe 8")


def remote_diag_start(result_dir=None):
    try:
        from souren_core import VisaInstrumentController
        VisaInstrumentController.scpi_comm_log.clear()
    except Exception:
        pass
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，跳过远程日志抓取")
        return False
    client = remote_client.RemoteClient(host, port)
    if result_dir is None:
        result_dir = os.getcwd()
    print(f"[*] 启动远程日志抓取，保存目录: {result_dir}")
    return client.start_log(result_dir)


def _save_client_scpi_log():
    try:
        from souren_core import VisaInstrumentController
        lines = VisaInstrumentController.scpi_comm_log
        if not lines:
            return None
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_path = os.path.join(os.getcwd(), f"{ts}_scpi_message_log.txt")
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
        VisaInstrumentController.scpi_comm_log.clear()
        return file_path
    except Exception as e:
        print(f"[WARN] 保存客户端 SCPI 日志失败: {e}")
        return None


def remote_diag_stop():
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达")
        scpi_file = _save_client_scpi_log()
        if scpi_file:
            print(f"[OK] SCPI 日志(客户端记录)已保存: {scpi_file}")
        return None, scpi_file
    client = remote_client.RemoteClient(host, port)
    print("[*] 停止远程日志抓取并获取 SCPI 日志...")
    signal_file, scpi_file = client.stop_log()
    if signal_file:
        print(f"[OK] Signal 日志已保存: {signal_file}")
    else:
        print("[WARN] 未收到 Signal 日志")
    if not scpi_file:
        scpi_file = _save_client_scpi_log()
        if scpi_file:
            print(f"[OK] SCPI 日志(客户端记录)已保存: {scpi_file}")
        else:
            print("[WARN] 未收到 SCPI 日志(服务器与客户端均无记录)")
    else:
        print(f"[OK] SCPI 日志已保存: {scpi_file}")
    return signal_file, scpi_file


def remote_restart(restart_time=20, rst_time=5):
    if remote_client.RemoteClient.ping(host, port):
        if restart_time != 0:
            print("[*] 尝试通过远程服务重启网页...")
        client = remote_client.RemoteClient(host, port)
        success = client._send_command_and_check_ok("DIAG:restart", "[OK]")
        if success:
            if restart_time != 0:
                print(f"[OK] 远程重启成功，等待 {restart_time} 秒...")
                time.sleep(restart_time)
            return True
        else:
            print("[WARN] 远程重启失败，将回退到本地 *rst")
    else:
        print("[跳过] Ubuntu 服务端不可达，使用本地重启")

    try:
        print("[*] 执行本地 SCPI 命令 *rst")
        ap.send("*rst")
        print(f"[OK] 本地重启完成，等待 {rst_time} 秒...")
        time.sleep(rst_time)
        return True
    except Exception as e:
        print(f"[FAIL] 本地重启失败: {e}")
        return False

def remote_pvt_screenshot(save_dir=None):
    host = souren_config.DEFAULT_IP
    port = souren_config.REMOTE_SERVER_PORT
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，无法远程截取 PVT 截图")
        return None
    client = remote_client.RemoteClient(host, port)
    if save_dir is None:
        save_dir = os.getcwd()
    print(f"[*] 请求 PVT 截图，保存目录: {save_dir}")
    return client.pvt_screenshot(save_dir)


def remote_gnb_start():
    host = souren_config.DEFAULT_IP
    port = souren_config.REMOTE_SERVER_PORT
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，无法配置基站")
        return False

    from souren_config import LOG_LEVEL_PARAMS
    log_level = str(LOG_LEVEL_PARAMS.get('log_level', 'info')).strip().lower()
    checkbox_items = [f"{k},{int(v)}" for k, v in LOG_LEVEL_PARAMS.items() if k != 'log_level']
    sub_parts = [log_level] + checkbox_items
    scpi_body = ";".join(sub_parts)
    scpi_cmd = f'CONFigure:VERSion:LOG:STATe "{scpi_body}"'

    try:
        ap.send(scpi_cmd)
        print(f"[*] 已下发 SCPI 日志配置命令")
    except Exception as e:
        print(f"[WARN] 下发日志配置 SCPI 失败: {e}")
    try:
        ap.send("*rst")
        print("[*] 已下发 *rst,等待 5 秒...")
        time.sleep(5)
    except Exception as e:
        print(f"[WARN] 下发 *rst 失败: {e}")

    print("[*] 通过远程服务重启...")
    success = remote_restart(restart_time=0)
    if success:
        print("[OK] 远程重启成功，等待 20 秒基站稳定...")
        time.sleep(20)
        return True
    else:
        print("[WARN] 远程重启失败")
        return False

def remote_gnb_stop(wait_time=10):
    host = souren_config.DEFAULT_IP
    port = souren_config.REMOTE_SERVER_PORT
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达")
        return False
    try:
        # sync 必须在 Ubuntu 主机(基站写日志的机器)上执行；本地 Windows 客户端没有 sync 命令，
        # 且本地 sync 也刷不到远端文件系统。改为通过远程服务在 Ubuntu 上执行。
        client_sync = remote_client.RemoteClient(host, port)
        if client_sync._send_command_and_check_ok("sync && echo SYNC_DONE", "SYNC_DONE"):
            print("[*] 已在 Ubuntu 主机执行 sync,强制刷新文件系统缓存")
        else:
            print("[WARN] 远程 sync 未确认完成(不影响后续取日志)")
    except Exception as e:
        print(f"[WARN] 远程 sync 执行失败: {e}")

    print(f"[*] CALL:CELL1 OFF 后等待 {wait_time} 秒，等基站写完 current 日志...")
    time.sleep(wait_time)

    client = remote_client.RemoteClient(host, port)
    return client.collect_current_logs(os.getcwd()) is not None

def remote_cleanup_on_interrupt(save_dir=None):
    host = souren_config.DEFAULT_IP
    port = souren_config.REMOTE_SERVER_PORT
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，无法远程清理/取日志")
        return False

    import signal as _signal
    old_handler = None
    try:
        old_handler = _signal.signal(_signal.SIGINT, _signal.SIG_IGN)
    except Exception:
        old_handler = None 

    original_cwd = None
    try:
        if save_dir:
            try:
                os.makedirs(save_dir, exist_ok=True)
                original_cwd = os.getcwd()
                os.chdir(save_dir)
            except Exception as e:
                print(f"[WARN] 切换到日志目录失败({save_dir}): {e}")
        try:
            remote_diag_stop()
        except Exception as e:
            print(f"[WARN] 中断时获取 signal/scpi 日志失败: {e}")

        print("[*] 中断清理：回传基站 current 日志(不杀进程)...")
        try:
            client = remote_client.RemoteClient(host, port)
            client.collect_current_logs(save_dir or os.getcwd())
            print("[OK] 中断清理完成：日志已回传")
        except Exception as e:
            print(f"[WARN] 回传 gNB current 日志异常: {e}")
        return True
    finally:
        if original_cwd:
            try:
                os.chdir(original_cwd)
            except Exception:
                pass
        if old_handler is not None:
            try:
                _signal.signal(_signal.SIGINT, old_handler)
            except Exception:
                pass


def remote_collect_core_logs(save_dir=None):
    host = souren_config.DEFAULT_IP
    port = souren_config.REMOTE_SERVER_PORT
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，跳过核心网日志收集")
        return None
    client = remote_client.RemoteClient(host, port)
    return client.collect_core_logs(save_dir)