from lib.var import *
from lib import remote_client
import souren_config
host = souren_config.DEFAULT_IP
port = souren_config.REMOTE_SERVER_PORT 

first_rst_flag = True
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
_current_case = None  


def _current_case_name():
    if _current_case is not None:
        return _current_case
    ex = _active_executor
    if ex is not None and getattr(ex, 'script_name', None):
        return os.path.splitext(os.path.basename(ex.script_name))[0]
    return None


# 在此 list 里的用例: config_line_loss / config_cell_band 走“按当前 SA/NSA 状态判断再切换”
# 的优化流程(不必要不切换/不重发 RFINdex); 其它用例保持原有流程不变。
NSA_SA_AUTO_SWITCH_SCRIPTS = [
    "test_nsa_channel_switch",
]


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

    def check(self, content, passed, detail=None, attempts=1, status_msg=None):
        if _active_executor is not None:
            return _active_executor._record_check(content, passed, detail, attempts, status_msg)
        status = "success" if passed else "failed"
        self.logger.info(f"[check] {content}: {status} (第{attempts}次) - {detail}")
        return passed

    def tag_last(self, status, count=1):
        if _active_executor is not None:
            return _active_executor._tag_last_extracted(status, count)
        self.logger.info(f"[tag_last] status={status}, count={count} (占位:无执行上下文,未标记)")
        return 0

    def record_value(self, title, value, x_label=None):
        if _active_executor is not None:
            return _active_executor._record_extracted_value(title, value, x_label)
        self.logger.info(f"[record_value] {title}={value} (占位:无执行上下文,未记录)")
        return value

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

    # AT 口优先：Fibocom 模组自带 ADB 接口，先探 ADB 会被误判成手机。
    # 先扫 Fibocom AT 串口，找到就发 AT 序列；没找到才用 adb 找手机。
    if AT_CONTROLLER_AVAILABLE:
        try:
            at_port = find_fibocom_at_port()
        except Exception as e:
            print(f"⚠️ 扫描 Fibocom AT 端口异常: {e}")
            at_port = None
        if at_port:
            print(f"✅ 检测到 Fibocom AT 端口: {at_port}，使用 AT 板模式")
            success, _ = send_at_sequence(at_port)
            if success:
                print("✅ AT序列执行完成")
                return True
            print("❌ AT序列执行失败,尝试 ADB 手机模式")
        else:
            print("❌ 未检测到 Fibocom AT 端口,尝试 ADB 手机模式")
    else:
        print("❌ board_at_controller 模块不可用,尝试 ADB 手机模式")

    if ADB_INTEGRATION_AVAILABLE:
        adb = get_adb_controller()
        if adb.adb_controller and adb.check_connection():
            print("✅ 检测到手机设备，使用手机模式")
            success = adb.timed_flight_mode_control(wait_time)
            if success:
                print("✅ 手机飞行模式控制完成")
                return True
            print("❌ 手机飞行模式控制失败")
        else:
            print("❌ 未检测到手机设备")
    else:
        print("❌ adb_integration 模块不可用")

    print("❌ AT 板和手机均不可用，未能触发注册")
    return False

def check_gnb_processes_alive(parameter=None, check_lte=None, timeout=15):
    if check_lte is None:
        check_lte = bool(parameter and parameter.get('lineLoss3') is not None)

    keywords = getattr(souren_config, 'GNB_PROCESS_KEYWORDS', {
        'nr': 'nr-softmodem', 'lte': 'lte-softmodem', 'ssm': 'ssmainctrl',
    })
    keys = ['nr', 'ssm'] + (['lte'] if check_lte else [])
    targets = {k: keywords[k] for k in keys if k in keywords}

    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，无法检查基站进程存活，按存活处理")
        return True, [], "服务端不可达,跳过检查"

    def _anti_self(n):
        return f'[{n[0]}]{n[1:]}' if n else n

    probe = "; ".join(
        f'pgrep -f "{_anti_self(name)}" >/dev/null 2>&1 && echo "PROC_CHECK:{key}:ALIVE" || echo "PROC_CHECK:{key}:DEAD"'
        for key, name in targets.items()
    )
    client = remote_client.RemoteClient(host, port)
    try:
        _, lines = client._execute_command_and_receive_output(probe, ready_string='', timeout=timeout)
    except Exception as e:
        print(f"[WARN] 检查基站进程存活失败: {e}, 按存活处理")
        return True, [], f"检测异常: {e}"

    status = {}
    for line in lines:
        m = re.search(r'PROC_CHECK:(\w+):(ALIVE|DEAD)', line)
        if m:
            status[m.group(1)] = (m.group(2) == 'ALIVE')

    missing = [k for k in targets if k not in status]
    if missing:
        print(f"[WARN] 进程检查未取到结果: {missing} (对应 {[targets[k] for k in missing]}), 按存活处理")

    dead_keys = [k for k in targets if status.get(k) is False]
    all_alive = not dead_keys
    detail = ", ".join(f"{k}({targets[k]})={'DEAD' if k in dead_keys else 'OK'}" for k in targets)
    if all_alive:
        print(f"✅ 基站关键进程检查正常: {detail}")
    else:
        print(f"❌ 检测到基站关键进程已退出: {detail}")
    return all_alive, dead_keys, detail


def my_sleep(seconds: float):
    print(f"😴 睡眠 {seconds} 秒...")
    time.sleep(seconds)
    print(f"✅ 睡眠完成")


def check_txp(label, txp_raw, index=1, tech=""):
    prefix = f"{tech} " if tech else ""
    try:
        parts = str(txp_raw).split(',')
        txp_val = float(parts[index].strip())
        passed = txp_val >= 0
        msg = f"check {prefix}TXP>0 pass" if passed else f"check {prefix}TXP<0 fail"
        ap.check(label, passed, detail=f"{prefix}TXP={txp_val} dBm", status_msg=msg)
    except (TypeError, ValueError, IndexError, AttributeError):
        ap.check(label, False, detail=f"{prefix}TXP 读取失败: {txp_raw}", status_msg=f"check {prefix}TXP fail")


ap = ScriptInstrumentController()

def setup_instrument_controller(instrument_controller):
    global ap
    ap.instrument = instrument_controller
    print("✅ 仪器控制器已设置")

def modify_config_line_loss(parameter):   
    
    ap.send("CONFigure:CELL1:NR:UE:MReport ON")
    UE_RSRP_val_str = ap.query(f"CONFigure:CELL1:NR:UEReport:RSRP?")
    TX_RSRP_val_str = ap.query(f"CONFigure:CELL1:NR:SIGN:RSRP?")
    UE_RSRP_val = float(UE_RSRP_val_str.split(',')[1])
    TX_RSRP_val = float(TX_RSRP_val_str.split(',')[0])
    line_loss = parameter['lineLoss1'] + (TX_RSRP_val - UE_RSRP_val) - 3
    print(f"line_loss {line_loss:.2f} parameter['lineLoss1'] {parameter['lineLoss1']} UE_RSRP_val {UE_RSRP_val} TX_RSRP_val{TX_RSRP_val}")
    ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,{line_loss:.2f},6000000000,{line_loss:.2f}")
    ap.send("CONFigure:BASE:FDCorrection:SAVE")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX")
    
    if parameter.get('lineLoss3') is not None:
        ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_3,100000000,{line_loss:.2f},6000000000,{line_loss:.2f}")
        ap.send("CONFigure:BASE:FDCorrection:SAVE")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,IO,RXTX")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,OUT,TX")
    

def config_line_loss(parameter):
    global _ofdm_last_wf, _baseline_line_loss, _skip_next_calibrate
    _ofdm_last_wf = None
    _skip_next_calibrate = False
    _baseline_line_loss = {
        'lineLoss1': parameter.get('lineLoss1'),
        'lineLoss3': parameter.get('lineLoss3'),
    }

    if _current_case_name() in NSA_SA_AUTO_SWITCH_SCRIPTS:
        _config_line_loss_switch(parameter)
        return

    if parameter.get('lte_band'):
        ap.send("CONFigure:NSASa:SWITch NSA")
    else:
        ap.send("CONFigure:NSASa:SWITch SA")

    ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,{parameter['lineLoss1']},6000000000,{parameter['lineLoss1']}")
    ap.send("CONFigure:BASE:FDCorrection:SAVE")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX")

    if parameter.get('lineLoss3') is not None:
        ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_3,100000000,{parameter['lineLoss3']},6000000000,{parameter['lineLoss3']}")
        ap.send("CONFigure:BASE:FDCorrection:SAVE")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,IO,RXTX")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,OUT,TX")

    ap.send("CONFigure:RFINdex:CLear:ALL")
    if parameter.get('lte_band') is not None:
        ap.send("CONFigure:RFINdex:DL LTE,3")
        ap.send("CONFigure:RFINdex:UL LTE,3")
    ap.send("CONFigure:RFINdex:DL NR,1")
    ap.send("CONFigure:RFINdex:UL NR,1")
    ap.send("CONFigure:RFINdex1:CONNector IO")
    ap.send("CONFigure:RFINdex3:CONNector IO")
    ap.send("CONFigure:RFINdex:apply")

    my_sleep(12) if parameter.get('lte_band') else my_sleep(6)


def _config_line_loss_switch(parameter):
    RFINdex_apply_flag = False
    Standalone_type = ap.query("CONFigure:NSASa:SWITch?")
    if parameter.get('lte_band') is not None and Standalone_type == "SA":
        ap.send("CALL:CELL1 OFF")
        ap.send("CONFigure:NSASa:SWITch NSA")
        RFINdex_apply_flag = True
    elif parameter.get('lte_band') is None and Standalone_type == "NSA":
        ap.send("CALL:CELL1 OFF")
        ap.send("CONFigure:NSASa:SWITch SA")
        RFINdex_apply_flag = True

    ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,{parameter['lineLoss1']},6000000000,{parameter['lineLoss1']}")
    ap.send("CONFigure:BASE:FDCorrection:SAVE")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX")

    if parameter.get('lineLoss3') is not None:
        ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_3,100000000,{parameter['lineLoss3']},6000000000,{parameter['lineLoss3']}")
        ap.send("CONFigure:BASE:FDCorrection:SAVE")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,IO,RXTX")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,OUT,TX")

    if RFINdex_apply_flag:
        ap.send("CONFigure:RFINdex:CLear:ALL")
        if parameter.get('lte_band') is not None:
            ap.send("CONFigure:RFINdex:DL LTE,3")
            ap.send("CONFigure:RFINdex:UL LTE,3")
        ap.send("CONFigure:RFINdex:DL NR,1")
        ap.send("CONFigure:RFINdex:UL NR,1")
        ap.send("CONFigure:RFINdex1:CONNector IO")
        ap.send("CONFigure:RFINdex3:CONNector IO")
        ap.send("CONFigure:RFINdex:apply")
        print("CONFigure:RFINdex:apply")
        my_sleep(12) if parameter.get('lte_band') else my_sleep(6)


def _query_rsrp_value(cmd, index=None):
    raw = ap.query(cmd)
    if raw is None:
        return None
    try:
        parts = str(raw).split(',')
        s = parts[index] if index is not None and index < len(parts) else parts[-1]
        return float(s.strip())
    except (TypeError, ValueError, IndexError):
        print(f"[WARN] RSRP 解析失败: {cmd} -> {raw!r}")
        return None


def _query_ue_rsrp(ue_cmd, max_tries=3, wait=5.0):
    rsrp = None
    last_raw = ""
    tries = 0
    for i in range(max_tries):
        tries = i + 1
        raw = ap.query(ue_cmd)
        last_raw = "" if raw is None else str(raw).strip()
        if raw is None:
            return None, False, last_raw, tries
        parts = last_raw.split(',')
        try:
            rel = int(float(parts[0].strip()))
        except (ValueError, IndexError):
            rel = None
        try:
            rsrp = float(parts[1].strip()) if len(parts) > 1 else None
        except (ValueError, IndexError):
            rsrp = None
        if rel != 7:
            return rsrp, True, last_raw, tries
        print(f"[线损校准] UEReport:RSRP reliability=7(测报未就绪), 第{tries}/{max_tries}次, 等待{wait}s后重读")
        if i < max_tries - 1:
            my_sleep(wait)
    print(f"[线损校准] UEReport:RSRP reliability 持续=7, 重试{max_tries}次仍无效")
    return rsrp, False, last_raw, tries


def _send_line_loss_table(table_name, rf_index, value):
    ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate {table_name},100000000,{value},6000000000,{value}")
    ap.send("CONFigure:BASE:FDCorrection:SAVE")
    ap.send(f"CONFigure:FDCorrection:ACTivate {table_name},{rf_index},IO,RXTX")
    ap.send(f"CONFigure:FDCorrection:ACTivate {table_name},{rf_index},OUT,TX")


def _calibrate_one_tech(parameter, tech, ll_key, table_name, rf_index, tolerance):
    ue_cmd = f"CONFigure:CELL1:{tech}:UEReport:RSRP?"
    sign_cmd = f"CONFigure:CELL1:{tech}:SIGN:RSRP?"
    tag = f"{tech}线损校准"

    ue_rsrp, ue_valid, ue_raw, ue_tries = _query_ue_rsrp(ue_cmd)
    if not ue_valid:
        cur_ll = parameter.get(ll_key)
        print(f"[WARN] {tag}: UE测报未使能(reliability=7), 跳过线损校准, 保持线损={cur_ll}")
        ap.check(tag, False,
                 detail=f"{ue_raw} 测报使能",
                 attempts=ue_tries,
                 status_msg=f"第{ue_tries}次查询命令执行成功")
        return False, cur_ll

    sign_rsrp = _query_rsrp_value(sign_cmd)
    if ue_rsrp is None or sign_rsrp is None:
        print(f"[WARN] {tag}: RSRP 读取失败, 跳过本次校准")
        ap.check(tag, False,
                 detail=f"UEReport={ue_rsrp}, SIGN={sign_rsrp}",
                 status_msg="RSRP读取失败,跳过线损校准 fail")
        return False, parameter.get(ll_key)

    diff = sign_rsrp - ue_rsrp   # SIGN比UE大->diff>0->加; SIGN比UE小->diff<0->减
    old_ll = float(parameter.get(ll_key) or 0)
    print(f"[{tag}] UEReport:RSRP={ue_rsrp}, SIGN:RSRP={sign_rsrp}, 差值={diff:+.2f}dB "
          f"(容差±{tolerance}dB), 当前 {ll_key}={old_ll}")

    if abs(diff) <= tolerance:
        ap.check(tag, True, detail=f"UE={ue_rsrp}, SIGN={sign_rsrp}, 差值={diff:+.2f}dB ≤ ±{tolerance}dB",
                 status_msg=f"误差{diff:+.2f}dB在容差内,无需重配线损 pass")
        print(f"[{tag}] 误差在容差内, 无需重新配置")
        return False, old_ll

    new_ll = round(old_ll + diff, 2)
    print(f"[{tag}] 误差超限, 重配线损: {old_ll} {'+' if diff >= 0 else '-'} {abs(diff):.2f} = {new_ll}")
    parameter[ll_key] = new_ll
    _send_line_loss_table(table_name, rf_index, new_ll)
    ap.check(tag, True,
             detail=f"UE={ue_rsrp}, SIGN={sign_rsrp}, 差值={diff:+.2f}dB, {ll_key} {old_ll}→{new_ll}",
             status_msg=f"误差{diff:+.2f}dB超差,线损已重配({old_ll}→{new_ll}) pass")
    return True, new_ll


def calibrate_line_loss(parameter, tolerance=2.0, skip_next=False):
    global _skip_next_calibrate
    if _skip_next_calibrate:
        _skip_next_calibrate = False
        print("[线损校准] 沿用初始校准结果, 跳过本次校准")
        return False, {'nr': parameter.get('lineLoss1'), 'lte': parameter.get('lineLoss3')}

    nr_changed, nr_ll = _calibrate_one_tech(
        parameter, 'NR', 'lineLoss1', 'LineLossTable_1', 1, tolerance)

    lte_changed, lte_ll = False, None
    if parameter.get('lineLoss3') is not None:
        lte_changed, lte_ll = _calibrate_one_tech(
            parameter, 'LTE', 'lineLoss3', 'LineLossTable_3', 3, tolerance)

    if skip_next:
        _skip_next_calibrate = True

    return (nr_changed or lte_changed), {'nr': nr_ll, 'lte': lte_ll}


_baseline_line_loss = {}
_skip_next_calibrate = False


def restore_line_loss(parameter):
    base_ll1 = _baseline_line_loss.get('lineLoss1')
    base_ll3 = _baseline_line_loss.get('lineLoss3')

    restored = []
    if base_ll1 is not None and parameter.get('lineLoss1') != base_ll1:
        cur = parameter.get('lineLoss1')
        parameter['lineLoss1'] = base_ll1
        ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,{base_ll1},6000000000,{base_ll1}")
        ap.send("CONFigure:BASE:FDCorrection:SAVE")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX")
        restored.append(f"lineLoss1 {cur}→{base_ll1}")

    if base_ll3 is not None and parameter.get('lineLoss3') != base_ll3:
        cur = parameter.get('lineLoss3')
        parameter['lineLoss3'] = base_ll3
        ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_3,100000000,{base_ll3},6000000000,{base_ll3}")
        ap.send("CONFigure:BASE:FDCorrection:SAVE")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,IO,RXTX")
        ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_3,3,OUT,TX")
        restored.append(f"lineLoss3 {cur}→{base_ll3}")

    if restored:
        print(f"[线损还原] 校准后线损已还原为基准值: {', '.join(restored)}")
        return True
    return False


def set_channel_switch_flag(parameter):
    global  channel_switch_flag
    channel_switch_flag = parameter
    print(f"set channel_switch_flag {channel_switch_flag} ")

def get_channel_switch_flag():
    print(f"get channel_switch_flag {channel_switch_flag} ")
    return channel_switch_flag

def config_cell_band(parameter):
    if _current_case_name() in NSA_SA_AUTO_SWITCH_SCRIPTS:
        _config_cell_band_switch(parameter)
        return
    channel_switch_flag = False
    nr_band = parameter.get('nr_band')
    if nr_band is None:
        band_list = parameter.get('nr_band_list') or []
        nr_band = band_list[0] if band_list else None
    if nr_band is None:
        print("[WARN] config_cell_band: 无 nr_band/nr_band_list,跳过 NR band 配置")
    else:
        nr_bw, scs = parameter.get('nr_bw'), parameter.get('scs')
        # nr_bw/scs 只接受标量; channel switch 用例的 scs 是 dict({'tdd':[30],'fdd':[15]}),
        # 先归一成 None, 再按首 band 的 duplex 取标量(有 bw/scs dict 用 duplex_bw_scs, 否则按 band 推)
        if not isinstance(nr_bw, (int, float)):
            nr_bw = None
        if not isinstance(scs, (int, float)):
            scs = None
        if nr_bw is None or scs is None:
            if isinstance(parameter.get('bw'), dict) or isinstance(parameter.get('scs'), dict):
                d_bw, d_scs = duplex_bw_scs(parameter, nr_band, 0)
            else:
                d_bw, d_scs = _bw_scs_for_band(nr_band)
            nr_bw = nr_bw if nr_bw is not None else d_bw
            scs = scs if scs is not None else d_scs
        ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {nr_band}")
        ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{nr_bw}")
        ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{scs}")
        if parameter.get('dl_arfcn') is not None:
            ap.send(f"CONFigure:CELL1:NR:CONFig:PMODe MANUAL")
            ap.send(f"CONFigure:CELL1:NR:SIGN:CFSCommand {parameter['dl_arfcn']},AUTO,AUTO")
        elif parameter.get('range') is not None:
            ap.send(f"CONFigure:CELL1:NR:CONFig:PMODe AUTO")
            ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {parameter['range']}")
    ap.send("CONFigure:CELL1:NR:UE:MReport ON")
    ap.send('CONFigure:CELL1:NR:SIGN:DDETection:SWITch ON,1')

    if parameter.get('lte_band') is not None:
        ap.send(f"CONFigure:CELL1:LTE:SIGN:BAND:DL OB{parameter['lte_band']}")
    if parameter.get('lte_bw') is not None:
        ap.send(f"CONFigure:CELL1:LTE:SIGN:BWidth BW_{parameter['lte_bw']}")


def _config_cell_band_switch(parameter):
    channel_switch_flag = False

    nr_band = parameter.get('nr_band')
    if nr_band is None:
        band_list = parameter.get('nr_band_list') or []
        nr_band = band_list[0] if band_list else None
    if nr_band is None:
        print("[WARN] config_cell_band: 无 nr_band/nr_band_list,跳过 NR band 配置")
    else:
        nr_bw, scs = parameter.get('nr_bw'), parameter.get('scs')
        if not isinstance(nr_bw, (int, float)):
            nr_bw = None
        if not isinstance(scs, (int, float)):
            scs = None
        if nr_bw is None or scs is None:
            if isinstance(parameter.get('bw'), dict) or isinstance(parameter.get('scs'), dict):
                d_bw, d_scs = duplex_bw_scs(parameter, nr_band, 0)
            else:
                d_bw, d_scs = _bw_scs_for_band(nr_band)
            nr_bw = nr_bw if nr_bw is not None else d_bw
            scs = scs if scs is not None else d_scs

        current_nr_band = ap.query("CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator?")
        if current_nr_band != f"{nr_band}":
            ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {nr_band}")
            channel_switch_flag = True
        current_nr_bw = ap.query("CONFigure:CELL1:NR:SIGN:BWidth:DL?")
        if current_nr_bw != f"BW{nr_bw}":
            ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{nr_bw}")
            channel_switch_flag = True
        current_nr_scs = ap.query("CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing?")
        if current_nr_scs != f"kHz{scs}":
            ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{scs}")
            channel_switch_flag = True

        if parameter.get('dl_arfcn') is not None:
            current_nr_dl_arfcn_str = ap.query(f"CONFigure:CELL1:NR:SIGN:ARFCn?")
            current_nr_dl_arfcn = (current_nr_dl_arfcn_str.split(',')[0])
            if current_nr_dl_arfcn != f"{parameter['dl_arfcn']}":
                ap.send(f"CONFigure:CELL1:NR:CONFig:PMODe MANUAL")
                ap.send(f"CONFigure:CELL1:NR:SIGN:CFSCommand {parameter['dl_arfcn']},AUTO,AUTO")
                channel_switch_flag = True
        elif parameter.get('range') is not None:
            current_nr_range = ap.query("CONFigure:CELL1:NR:CONFig:RANGe?")
            if current_nr_range != f"{parameter['range']}":
                ap.send(f"CONFigure:CELL1:NR:CONFig:PMODe AUTO")
                ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {parameter['range']}")
    ap.send("CONFigure:CELL1:NR:UE:MReport ON")

    if parameter.get('lte_band') is not None:
        current_Lte_band = ap.query("CONFigure:CELL1:LTE:SIGN:BAND:DL?")
        if current_Lte_band != f"OB{parameter['lte_band']}":
            ap.send(f"CONFigure:CELL1:LTE:SIGN:BAND:DL OB{parameter['lte_band']}")
            channel_switch_flag = True
    if parameter.get('lte_bw') is not None:
        current_Lte_bw = ap.query("CONFigure:CELL1:LTE:SIGN:BWidth?")
        lte_bw_val = int(parameter['lte_bw']) * 10
        if current_Lte_bw != f"BW_{lte_bw_val}":
            ap.send(f"CONFigure:CELL1:LTE:SIGN:BWidth BW_{lte_bw_val}")
            channel_switch_flag = True
    if parameter.get('lte_dl_arfcn') is not None:
        current_lte_dl_arfcn = ap.query("CONFigure:CELL1:LTE:SIGN:ARFCn:DL?")
        if current_lte_dl_arfcn != f"{parameter['lte_dl_arfcn']}":
            ap.send(f"CONFigure:CELL1:LTE:SIGN:ARFCn:DL {parameter['lte_dl_arfcn']}")
            channel_switch_flag = True
    if parameter.get('lte_ul_arfcn') is not None:
        current_lte_ul_arfcn = ap.query("CONFigure:CELL1:LTE:SIGN:ARFCn:UL?")
        if current_lte_ul_arfcn != f"{parameter['lte_ul_arfcn']}":
            ap.send(f"CONFigure:CELL1:LTE:SIGN:ARFCn:UL {parameter['lte_ul_arfcn']}")
            channel_switch_flag = True
    if parameter.get('nr_start_power') is not None:
        ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {parameter['nr_start_power']}")
    if parameter.get('nsa_start_power') is not None:
        ap.send(f"CONFigure:CELL1:LTE:SIGN:POWer {parameter['nsa_start_power']}")

    UE_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
    if channel_switch_flag and UE_status == '"Connected"':
        ap.send("CONFigure:CELL1:NR:SIGN:CHANnel:SWITCH")
    elif UE_status != '"Connected"':
        ap.send("CALL:CELL1 OFF")
        my_sleep(1)



from ul_rb_table import UL_RB_TABLE, get_ul_rb


_SLOT_CTYPE = {"DL": "PDSCh", "UL": "PUSCh"}

_ofdm_last_wf = None


def _merge_slots(base, override):
    base = base or {}
    override = override or {}
    if not override:
        return base
    merged = {}
    for d in ("DL", "UL"):
        if d in override:
            v = override[d]
            merged[d] = v if v else list(base.get(d, []))
    return merged


def _bw_scs_for_band(band):
    if band is None:
        return None, None
    return (100, 30) if is_tdd_band(band) else (20, 15)



_NR_TDD_BANDS = {34, 38, 39, 40, 41, 42, 43, 46, 47, 48, 50, 51, 53,
                 77, 78, 79, 90, 96, 102, 104}



def nr_duplex_by_band(band):
    try:
        b = int(band)
    except (TypeError, ValueError):
        return "FDD"
    if b in _NR_TDD_BANDS:
        return "TDD"
    return "FDD"


def is_tdd_band(band):
    return nr_duplex_by_band(band) == "TDD"


def num_rounds(parameter):
    n = 1
    for cfg in (parameter.get('bw'), parameter.get('scs')):
        if isinstance(cfg, dict):
            for v in cfg.values():
                if isinstance(v, (list, tuple)):
                    n = max(n, len(v))
    return n


def duplex_bw_scs(parameter, band, round_idx=0):
    duplex = 'tdd' if is_tdd_band(band) else 'fdd'
    fb_bw, fb_scs = (100, 30) if duplex == 'tdd' else (20, 15)

    def pick(cfg, fallback):
        v = cfg.get(duplex) if isinstance(cfg, dict) else None
        if isinstance(v, (list, tuple)):
            return v[min(round_idx, len(v) - 1)] if v else fallback
        return v if v is not None else fallback

    return pick(parameter.get('bw'), fb_bw), pick(parameter.get('scs'), fb_scs)


def slots_for_band(band, nr_slots):
    if is_tdd_band(band):
        return nr_slots
    out = {}
    for d in ("DL", "UL"):
        out[d] = [it for it in (nr_slots or {}).get(d, [])
                  if all(int(n) <= 9 for n in it.keys())]
    return out


def config_nr_slots(parameter, slots=None, nr_bw=None, nr_scs=None, band=None, clear=True):
    global _ofdm_last_wf
    if clear:
        ap.send("CONFigure:CELL1:NR:SIGN:SLOT:CLEar")

    wf = str(parameter.get('waveform', 'DFTS')).upper()
    ofdm = "DFT" if wf.startswith("DFT") else "CP"
    if ofdm != _ofdm_last_wf:
        ap.send(f"CONFigure:CELL1:NR:SIGN:OFDM {ofdm}")
        print(f"[OFDM] 上行波形已设为 {ofdm} (waveform={wf})")
        _ofdm_last_wf = ofdm

    slots = _merge_slots(slots, parameter.get('nr_slots'))
    _filter_band = band
    if _filter_band is None:
        _bl = parameter.get('nr_band_list') or []
        _filter_band = _bl[0] if _bl else parameter.get('nr_band')
    if _filter_band is not None:
        slots = slots_for_band(_filter_band, slots)
    if not slots:
        print("[WARN] config_nr_slots: 未提供 NR slots(应在 case 内定义并传入),跳过 slot 配置")
    ul_explicit_rb = set()
    mcs_override = parameter.get('mcs')   # 软件/配置里给了 mcs 就覆盖 slot 表的 MCS1
    for direction in ("DL", "UL"):
        ctype = _SLOT_CTYPE[direction]
        for slot_item in slots.get(direction, []):
            for slot_num, fields in slot_item.items():
                ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT{slot_num}:CTYPe {ctype}")
                for field, value in fields.items():
                    if mcs_override is not None and str(field).upper() == "MCS1":
                        value = mcs_override
                    ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT{slot_num}:{direction}:{field} {value}")
                    if direction == "UL" and str(field).upper() == "RB":
                        ul_explicit_rb.add(slot_num)

    rb_mode = parameter.get('rb_mode', 'Inner_Full')
    ul_slots = [n for item in slots.get("UL", []) for n in item.keys()]
    if rb_mode in UL_RB_TABLE:
        _bw, _scs = nr_bw, nr_scs
        if _bw is None or _scs is None:
            p_bw, p_scs = parameter.get('nr_bw'), parameter.get('scs')
            if p_bw is None or p_scs is None:
                _seed = band
                if _seed is None:
                    band_list = parameter.get('nr_band_list') or []
                    _seed = band_list[0] if band_list else parameter.get('nr_band')
                b_bw, b_scs = _bw_scs_for_band(_seed)
                p_bw = p_bw if p_bw is not None else b_bw
                p_scs = p_scs if p_scs is not None else b_scs
            _bw = _bw if _bw is not None else p_bw
            _scs = _scs if _scs is not None else p_scs
        nr_bw, scs = _bw, _scs
        waveform = parameter.get('waveform', 'DFTS')
        rb_value = get_ul_rb(nr_bw, scs, rb_mode, waveform)
        if rb_value is not None:
            for slot in ul_slots:
                if slot in ul_explicit_rb:
                    continue
                ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT{slot}:UL:RB {rb_value}")
            print(f"[UL RB] 首配 slot{ul_slots} = {rb_value} (bw={nr_bw}, scs={scs}, {rb_mode}/{waveform})")
        else:
            print(f"[WARN] {rb_mode} 未找到 (nr_bw={nr_bw}, scs={scs}, waveform={waveform}) 的 RB, 回退 SLOT:UPDate")
            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
    elif rb_mode in ('Outer_Full', '无'):
        ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
    else:
        print(f"[WARN] 未知 rb_mode='{rb_mode}', 按 Outer_Full 处理")
        ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")

    global _ul_rb_last_key
    _seed_band = band
    if _seed_band is None:
        _bl = parameter.get('nr_band_list') or []
        _seed_band = _bl[0] if _bl else parameter.get('nr_band')
    _seed_bw = nr_bw if nr_bw is not None else parameter.get('nr_bw')
    _seed_scs = nr_scs if nr_scs is not None else parameter.get('scs')
    if _seed_bw is None or _seed_scs is None:
        _b_bw, _b_scs = _bw_scs_for_band(_seed_band)
        _seed_bw = _seed_bw if _seed_bw is not None else _b_bw
        _seed_scs = _seed_scs if _seed_scs is not None else _b_scs
    _ul_rb_last_key = (
        nr_duplex_by_band(_seed_band) if _seed_band is not None else None,
        _seed_bw, _seed_scs, rb_mode, parameter.get('waveform', 'DFTS'),
    )

    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
    

_ul_rb_last_key = None


def reset_ul_rb_cache():
    global _ul_rb_last_key
    _ul_rb_last_key = None


def switch_ul_slot_rb(ul_slots, nr_bw, scs, rb_mode='Inner_Full', waveform='DFTS',
                      ul_explicit_rb=None, sleep_after=2, band=None):
    global _ul_rb_last_key
    ul_explicit_rb = ul_explicit_rb or set()

    duplex = nr_duplex_by_band(band) if band is not None else None
    key = (duplex, nr_bw, scs, rb_mode, waveform)

    if key == _ul_rb_last_key:
        print(f"[UL RB] {duplex or ''} bw={nr_bw} scs={scs} {rb_mode}/{waveform} 与上次相同, 跳过 UPDate/UL RB, 直接 APPLy")
    else:
        ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
        rb_value = get_ul_rb(nr_bw, scs, rb_mode, waveform) if rb_mode in UL_RB_TABLE else None
        if rb_value is not None:
            for slot in ul_slots:
                if slot in ul_explicit_rb:
                    continue
                ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT{slot}:UL:RB {rb_value}")
            print(f"[UL RB] 配置 slot{list(ul_slots)} = {rb_value} ({duplex or ''} bw={nr_bw}, scs={scs}, {rb_mode}/{waveform})")
        elif rb_mode in UL_RB_TABLE:
            print(f"[WARN] {rb_mode} 未找到 (nr_bw={nr_bw}, scs={scs}, waveform={waveform}) 的 RB, 仅 UPDate")
        _ul_rb_last_key = key

    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
    if sleep_after:
        my_sleep(sleep_after)



LTE_BW_RB_TABLE = {
    1.4: (6, 6), 3: (15, 8), 5: (25, 13),
    10: (50, 17), 15: (75, 19), 20: (100, 25),
}


def _lte_bw_rb(lte_bw):
    try:
        return LTE_BW_RB_TABLE.get(float(lte_bw), (None, None))
    except (TypeError, ValueError):
        return None, None


def config_lte_subframes(parameter, slots=None, lte_bw=None):
    if parameter.get('lte_band') is None:
        return
    slots = _merge_slots(slots, parameter.get('lte_slots'))
    if not slots:
        print("[WARN] config_lte_subframes: 未提供 LTE slots(应在 case 内定义并传入),跳过子帧配置")
    ra_type = parameter.get('resource_allocation_type')
    if ra_type is None:
        ra_type = 0

    bw = lte_bw if lte_bw is not None else parameter.get('lte_bw')
    dl_rb, bitmap_bits = _lte_bw_rb(bw)
    ul_rb_cnt = dl_rb // 2 if dl_rb else None
    if dl_rb:
        print(f"[LTE RB] lte_bw={bw}MHz -> DL {dl_rb}RB/bitmap {bitmap_bits}bits, "
              f"UL RB 0,{ul_rb_cnt} (RA type{ra_type})")

    for direction in ("DL", "UL"):
        ctype = _SLOT_CTYPE[direction]
        for slot_item in slots.get(direction, []):
            for sf_num, fields in slot_item.items():
                fk = {str(k).upper() for k in fields}
                ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame{sf_num}:CTYPe {ctype}")

                if direction == "DL":
                    # 满配下行: RBGBitmap 全1(TYPE0) + DL:RB 0,dl_rb + RATYpe TYPE0, 随 lte_bw 动态
                    if bitmap_bits and ra_type == 0 and "RBGBITMAP" not in fk:
                        ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame{sf_num}:DL:RBGBitmap {'1' * bitmap_bits}")
                    if dl_rb and "RB" not in fk:
                        ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame{sf_num}:DL:RB 0,{dl_rb}")
                    for field, value in fields.items():
                        ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame{sf_num}:DL:{field} {value}")
                    ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame{sf_num}:DL:RATYpe TYPE{ra_type}")
                else:
                    for field, value in fields.items():
                        ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame{sf_num}:UL:{field} {value}")
                    if ul_rb_cnt and "RB" not in fk:
                        ap.send(f"CONFigure:CELL1:LTE:SIGN:SUBFrame{sf_num}:UL:RB 0,{ul_rb_cnt}")

    ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame:APPLy")
    my_sleep(1)

    ul_sfs = [n for item in slots.get("UL", []) for n in item.keys()]
    msubframe = ul_sfs[0] if ul_sfs else 8
    ap.send(f"CONFigure:LTE:TXP:MSUBframe {msubframe}")


def remote_diag_start(result_dir=None):
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，跳过远程日志抓取")
        return False
    client = remote_client.RemoteClient(host, port)
    if result_dir is None:
        result_dir = os.getcwd()
    print(f"[*] 启动远程日志抓取，保存目录: {result_dir}")
    return client.start_log(result_dir)


def remote_diag_stop():
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，无法获取远程日志")
        return None, None
    client = remote_client.RemoteClient(host, port)
    print("[*] 停止远程日志抓取并获取 SCPI 日志...")
    signal_file, scpi_file = client.stop_log()
    if signal_file:
        print(f"[OK] Signal 日志已保存: {signal_file}")
    else:
        print("[WARN] 未收到 Signal 日志")
    if scpi_file:
        print(f"[OK] SCPI 日志已保存: {scpi_file}")
    else:
        print("[WARN] 未收到 SCPI 日志")
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


def remote_osc_screenshot(save_dir=None, tag=None):
    host = souren_config.DEFAULT_IP
    port = souren_config.REMOTE_SERVER_PORT
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，无法远程截取 OSC 截图")
        return None
    client = remote_client.RemoteClient(host, port)
    if save_dir is None:
        save_dir = os.getcwd()
    print(f"[*] 请求 OSC 截图，保存目录: {save_dir}" + (f", tag={tag}" if tag else ""))
    return client.osc_screenshot(save_dir, tag)


def remote_set_duty_cycle(value, save_dir=None):
    host = souren_config.DEFAULT_IP
    port = souren_config.REMOTE_SERVER_PORT
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，无法设置 TDD Uplink Duty Cycle")
        return False
    client = remote_client.RemoteClient(host, port)
    if save_dir is None:
        save_dir = os.getcwd()
    print(f"[*] 请求设置 TDD Uplink Duty Cycle = {value}")
    return client.set_duty_cycle(value, save_dir)
def set_rst_flag(parameter): 
    global  first_rst_flag 
    first_rst_flag = parameter
    print(f"set first_rst_flag {first_rst_flag} ")
def get_rst_flag():
    print(f"get first_rst_flag {first_rst_flag} ")
    return  first_rst_flag 


_debug_log_configured = False


def reset_debug_log_flag():
    global _debug_log_configured
    _debug_log_configured = False


def remote_gnb_start():
    global _debug_log_configured
    host = souren_config.DEFAULT_IP
    port = souren_config.REMOTE_SERVER_PORT
    if not remote_client.RemoteClient.ping(host, port):
        print("[跳过] Ubuntu 服务端不可达，无法配置基站")
        return False

    if not getattr(souren_config, 'DEBUG_LOG_ENABLED', True):
        print("[*] 未开启抓 log(DEBUG_LOG_ENABLED=False),跳过 调试日志级别 与 *rst 下发")
    elif _debug_log_configured:
        print("[*] 调试日志级别本轮批跑已下发过(首个case已配置),本case跳过重复下发")
    else:
        from souren_config import LOG_LEVEL_PARAMS
        log_level = str(LOG_LEVEL_PARAMS.get('log_level', 'info')).strip().lower()
        checkbox_items = [f"{k},{int(v)}" for k, v in LOG_LEVEL_PARAMS.items() if k != 'log_level']
        sub_parts = [log_level] + checkbox_items
        scpi_body = ";".join(sub_parts)
        scpi_cmd = f'CONFigure:VERSion:LOG:STATe "{scpi_body}"'

        sent_ok = False
        try:
            ap.send(scpi_cmd)
            print(f"[*] 已下发 SCPI 日志配置命令")
            sent_ok = True
        except Exception as e:
            print(f"[WARN] 下发日志配置 SCPI 失败: {e}")
        try:
            ap.send("*rst")
            print("[*] 已下发 *rst,等待 5 秒...")
            time.sleep(5)
        except Exception as e:
            print(f"[WARN] 下发 *rst 失败: {e}")
        if sent_ok:
            _debug_log_configured = True

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