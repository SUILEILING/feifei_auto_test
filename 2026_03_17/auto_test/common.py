from lib.var import *

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

class InstrumentController:
    
    def __init__(self, instrument_controller=None):
        self.instrument = instrument_controller
        self.logger = logging.getLogger('SourenCommon')
        self.last_result = None
        
    def send(self, command: str, extract_index: Optional[int] = None, should_extract: bool = False) -> Union[str, float, None]:
        try:
            if not command:
                self.logger.warning("尝试发送空命令")
                return None
            
            if isinstance(command, str) and command.upper().startswith("SLEEP"):
                import re
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


ap = InstrumentController()

def setup_instrument_controller(instrument_controller):
    global ap
    ap.instrument = instrument_controller
    print("✅ 仪器控制器已设置")