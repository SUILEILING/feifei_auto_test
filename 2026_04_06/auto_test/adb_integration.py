from lib.var import *

class ADBFlightModeController:
    def __init__(self, adb_path=None):
        self.adb_path = adb_path or self.find_adb()
        self.device_id = None

        if not self.adb_path:
            print("❌ 未找到ADB,请安装ADB工具")
            self.show_install_guide()
            return

        print(f"📱 使用ADB: {self.adb_path}")

        if not self.check_adb_working():
            print("❌ ADB无法正常工作")
            return

        self.connect_device_once()

    def run_adb_command(self, command: str, timeout: int = 10) -> Tuple[str, str, int]:
        try:
            if self.adb_path != "adb":
                cmd = f'"{self.adb_path}" {command}'
            else:
                cmd = f"adb {command}"

            result = subprocess.run(cmd, shell=True, capture_output=True, text=True,
                                    timeout=timeout, encoding='utf-8', errors='ignore')
            return result.stdout.strip(), result.stderr.strip(), result.returncode
        except subprocess.TimeoutExpired:
            return "", "命令超时", -1
        except Exception as e:
            return "", str(e), -1

    def check_adb_working(self) -> bool:
        try:
            version_cmd = f'"{self.adb_path}" version' if self.adb_path != "adb" else "adb version"
            result = subprocess.run(version_cmd, shell=True, capture_output=True, text=True,
                                    timeout=5, encoding='utf-8', errors='ignore')
            if result.returncode != 0:
                print(f"❌ ADB无法运行: {result.stderr}")
                return False
            version_line = result.stdout.split('\n')[0] if result.stdout else "未知版本"
            print(f"✅ ADB版本: {version_line}")
            return True
        except Exception as e:
            print(f"❌ ADB检查失败: {e}")
            return False

    def find_adb(self) -> Optional[str]:
        d_drive_paths = [
            r"D:\adb\platform-tools\adb.exe",
            r"D:\adb\adb.exe",
            r"D:\platform-tools\adb.exe",
        ]
        print("🔍 正在检查D盘ADB...")
        for path in d_drive_paths:
            if os.path.exists(path):
                print(f"✅ 在D盘找到ADB: {path}")
                return path

        print("🔍 检查PATH环境变量中的ADB...")
        try:
            result = subprocess.run("adb version", shell=True, capture_output=True, text=True,
                                    timeout=5, encoding='utf-8', errors='ignore')
            if result.returncode == 0:
                print("✅ 在PATH中找到ADB")
                return "adb"
        except:
            pass

        print("🔍 检查其他常见路径...")
        common_paths = [
            r"C:\adb\adb.exe",
            r"C:\Program Files (x86)\Android\android-sdk\platform-tools\adb.exe",
        ]
        for path in common_paths:
            if os.path.exists(path):
                print(f"✅ 在其他路径找到ADB: {path}")
                return path

        print("❌ 未在任何路径找到ADB")
        return None

    def show_install_guide(self):
        print("="*60)
        print("ADB未找到,请按以下步骤安装：")
        print("="*60)

    def connect_device_once(self) -> bool:
        try:
            print("\n🔄 检查设备连接...")
            devices_cmd = f'"{self.adb_path}" devices' if self.adb_path != "adb" else "adb devices"
            result = subprocess.run(devices_cmd, shell=True, capture_output=True, text=True,
                                    timeout=10, encoding='utf-8', errors='ignore')
            if result.returncode != 0:
                print("❌ 检查设备失败")
                return False

            lines = result.stdout.strip().split('\n')
            devices = []
            for line in lines[1:]:
                line = line.strip()
                if line and "device" in line and "offline" not in line and "unauthorized" not in line:
                    devices.append(line.split()[0])

            if devices:
                self.device_id = devices[0]
                print(f"✅ 已连接设备: {self.device_id}")
                return True
            else:
                print("❌ 未检测到已连接的设备")
                return False

        except subprocess.TimeoutExpired:
            print("❌ ADB命令超时")
            return False
        except Exception as e:
            print(f"❌ 连接设备失败: {e}")
            return False

    def get_flight_mode_status(self) -> Optional[bool]:
        if not self.device_id:
            return None
        out, _, code = self.run_adb_command(f"-s {self.device_id} shell settings get global airplane_mode_on")
        if code == 0:
            return out.strip() == "1"
        return None

    def _force_radio_off(self):
        print("📡 强制关闭无线电硬件...")

        print("  1. 关闭移动数据...")
        self.run_adb_command(f"-s {self.device_id} shell svc data disable")
        time.sleep(0.3)
        self.run_adb_command(f"-s {self.device_id} shell settings put global mobile_data 0")
        time.sleep(0.3)

        print("  2. 关闭WiFi和蓝牙...")
        self.run_adb_command(f"-s {self.device_id} shell svc wifi disable")
        time.sleep(0.2)
        self.run_adb_command(f"-s {self.device_id} shell svc bluetooth disable")
        time.sleep(0.2)
        print("  3. 设置飞行模式标志...")
        self.run_adb_command(f"-s {self.device_id} shell settings put global airplane_mode_on 1")
        time.sleep(0.5)

        self.run_adb_command(
            f"-s {self.device_id} shell content insert --uri content://settings/global "
            f"--bind name:s:airplane_mode_on --bind value:i:1"
        )
        time.sleep(0.3)

        self.run_adb_command(f"-s {self.device_id} shell cmd connectivity airplane-mode enable")
        time.sleep(0.3)

        print("  6. 发送系统广播...")
        self.run_adb_command(
            f"-s {self.device_id} shell am broadcast -a android.intent.action.AIRPLANE_MODE --ez state true"
        )
        time.sleep(0.2)
        self.run_adb_command(
            f"-s {self.device_id} shell am broadcast -a android.intent.action.AIRPLANE_MODE_CHANGED --ez state true"
        )
        time.sleep(0.2)

        print("  7. 设置基带属性...")
        self.run_adb_command(f"-s {self.device_id} shell setprop gsm.radio.disabled 1")
        time.sleep(0.2)
        self.run_adb_command(f"-s {self.device_id} shell setprop persist.radio.airplane_mode_on 1")
        time.sleep(0.2)

        time.sleep(1)
        print("✅ 无线电强制关闭完成")
        
    def _force_radio_on(self):
        print("📡 强制开启无线电硬件...")

        print("  1. 清除飞行模式标志...")
        self.run_adb_command(f"-s {self.device_id} shell settings put global airplane_mode_on 0")
        time.sleep(0.5)
        self.run_adb_command(
            f"-s {self.device_id} shell content insert --uri content://settings/global "
            f"--bind name:s:airplane_mode_on --bind value:i:0"
        )
        time.sleep(0.3)

        print("  2. 清除基带属性...")
        self.run_adb_command(f"-s {self.device_id} shell setprop gsm.radio.disabled 0")
        time.sleep(0.2)
        self.run_adb_command(f"-s {self.device_id} shell setprop persist.radio.airplane_mode_on 0")
        time.sleep(0.2)

        print("  3. 开启网络服务...")
        self.run_adb_command(f"-s {self.device_id} shell svc data enable")    
        time.sleep(0.3)
        self.run_adb_command(f"-s {self.device_id} shell svc wifi enable")     
        time.sleep(0.3)
        self.run_adb_command(f"-s {self.device_id} shell svc bluetooth enable") 
        time.sleep(0.3)

        print("  4. 设置移动数据属性...")
        self.run_adb_command(f"-s {self.device_id} shell settings put global mobile_data 1")
        time.sleep(0.3)
        self.run_adb_command(f"-s {self.device_id} shell cmd connectivity data enable")
        time.sleep(0.3)
        self.run_adb_command(f"-s {self.device_id} shell cmd connectivity airplane-mode disable")
        time.sleep(0.3)

        print("  5. 发送网络状态变更广播...")
        self.run_adb_command(
            f"-s {self.device_id} shell am broadcast -a android.net.conn.CONNECTIVITY_CHANGE"
        )
        time.sleep(0.2)
        self.run_adb_command(
            f"-s {self.device_id} shell am broadcast -a android.net.wifi.STATE_CHANGE"
        )
        time.sleep(0.2)
        self.run_adb_command(
            f"-s {self.device_id} shell am broadcast -a android.intent.action.SERVICE_STATE"
        )
        time.sleep(0.2)

        time.sleep(1)
        print("✅ 无线电强制开启完成")

    def enable_flight_mode(self) -> bool:
        if not self.device_id:
            print("❌ 未连接设备")
            return False
        print("🔄 开启飞行模式...")
        self._force_radio_off()
        return True

    def disable_flight_mode(self) -> bool:
        if not self.device_id:
            print("❌ 未连接设备")
            return False
        print("🔄 关闭飞行模式...")
        self._force_radio_on()
        return True

    def timed_flight_mode_control(self, wait_time: int = 5) -> bool:
        if not self.device_id:
            print("❌ 未连接设备")
            return False

        print(f"📱 开始定时飞行模式控制, 等待时间: {wait_time}秒")

        if not self.enable_flight_mode():
            return False

        print(f"⏳ 等待 {wait_time} 秒...")
        time.sleep(wait_time)

        if not self.disable_flight_mode():
            return False

        print("✅ 定时飞行模式控制完成")
        return True

if __name__ == "__main__":
    print("="*60)
    print("📱 ADB飞行模式控制工具")
    print("="*60)

    controller = ADBFlightModeController()
    if controller.device_id:
        print(f"\n✅ 设备已连接: {controller.device_id}")
        controller.timed_flight_mode_control(3)
    else:
        print("\n❌ 未连接设备，退出")