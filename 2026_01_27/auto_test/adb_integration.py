from lib.var import sys, subprocess, os, time

class ADBFlightModeController:
    """ADB飞行模式控制器"""
    
    def __init__(self, adb_path=None):
        """
        初始化
        :param adb_path: ADB路径,如果为None则自动查找
        """
        # 默认先检查D盘
        self.adb_path = adb_path or self.find_adb()
        self.device_id = None
        
        if not self.adb_path:
            print("❌ 未找到ADB,请安装ADB工具")
            self.show_install_guide()
            return
        
        print(f"📱 使用ADB: {self.adb_path}")
        
        # 检查ADB是否正常工作
        if not self.check_adb_working():
            print("❌ ADB无法正常工作")
            return
            
        self.connect_device()
    
    def run_adb_command(self, command):
        """运行ADB命令"""
        try:
            if self.adb_path != "adb":
                cmd = f'"{self.adb_path}" {command}'
            else:
                cmd = f"adb {command}"
            
            result = subprocess.run(cmd, shell=True, 
                                  capture_output=True, text=True, 
                                  timeout=10, encoding='utf-8', errors='ignore')
            return result.stdout.strip(), result.stderr.strip(), result.returncode
            
        except subprocess.TimeoutExpired:
            return "", "命令超时", -1
        except Exception as e:
            return "", str(e), -1
    
    def check_adb_working(self):
        """检查ADB是否能正常工作"""
        try:
            version_cmd = f'"{self.adb_path}" version' if self.adb_path != "adb" else "adb version"
            result = subprocess.run(version_cmd, shell=True, 
                                  capture_output=True, text=True, 
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
    
    def find_adb(self):
        """查找ADB,优先检查D盘"""
        
        # 1. 优先检查D盘标准路径
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
        
        # 2. 检查PATH环境变量中的adb
        print("🔍 检查PATH环境变量中的ADB...")
        try:
            result = subprocess.run("adb version", shell=True, 
                                  capture_output=True, text=True, 
                                  timeout=5, encoding='utf-8', errors='ignore')
            if result.returncode == 0:
                print("✅ 在PATH中找到ADB")
                return "adb"
        except:
            pass
        
        # 3. 检查其他常见路径
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
        """显示安装指南"""
        print("="*60)
        print("ADB未找到,请按以下步骤安装：")
        print("="*60)
        print("\n📱 安装步骤：")
        print("1. 下载ADB工具包:")
        print("   访问 https://developer.android.com/studio/releases/platform-tools")
        print("   下载 'platform-tools-latest-windows.zip'")
        print("\n2. 解压到 D:\\adb 目录")
        print("\n3. 手机设置：")
        print("   - 进入设置 > 关于手机")
        print("   - 点击'版本号'7次启用开发者选项")
        print("   - 返回设置 > 开发者选项")
        print("   - 开启'USB调试'")
        print("\n4. 重新运行此程序")
        print("="*60)
        
        input("\n按回车键退出...")
    
    def connect_device(self, retry_count=3):
        """连接设备"""
        for attempt in range(retry_count):
            try:
                print(f"\n🔄 连接设备 (第 {attempt + 1} 次)...")
                
                # 检查设备
                devices_cmd = f'"{self.adb_path}" devices' if self.adb_path != "adb" else "adb devices"
                result = subprocess.run(devices_cmd, shell=True, 
                                      capture_output=True, text=True, 
                                      timeout=10, encoding='utf-8', errors='ignore')
                
                if result.returncode != 0:
                    print(f"❌ 检查设备失败")
                    continue
                
                # 解析设备
                lines = result.stdout.strip().split('\n')
                devices = []
                
                for line in lines[1:]:  # 跳过第一行标题
                    line = line.strip()
                    if line and "device" in line and "offline" not in line:
                        device_id = line.split()[0]
                        devices.append(device_id)
                
                if not devices:
                    print("❌ 未找到已连接的设备")
                    
                    if attempt == retry_count - 1:
                        print("\n💡 检查USB连接和USB调试")
                        response = input("\n是否要重启ADB服务?(y/N): ").strip().lower()
                        if response == 'y':
                            self.restart_adb_server()
                            continue
                            
                    time.sleep(2)
                    continue
                
                self.device_id = devices[0]
                print(f"✅ 已连接设备: {self.device_id}")
                return True
                
            except subprocess.TimeoutExpired:
                print("❌ ADB命令超时")
            except Exception as e:
                print(f"❌ 连接设备失败")
            
            if attempt < retry_count - 1:
                time.sleep(2)
        
        return False
    
    def restart_adb_server(self):
        """重启ADB服务"""
        print("\n🔄 重启ADB服务...")
        try:
            kill_cmd = f'"{self.adb_path}" kill-server' if self.adb_path != "adb" else "adb kill-server"
            subprocess.run(kill_cmd, shell=True, capture_output=True, 
                          timeout=5, encoding='utf-8', errors='ignore')
            
            start_cmd = f'"{self.adb_path}" start-server' if self.adb_path != "adb" else "adb start-server"
            subprocess.run(start_cmd, shell=True, capture_output=True, 
                          timeout=5, encoding='utf-8', errors='ignore')
            
            print("✅ ADB服务已重启")
            return True
        except Exception as e:
            print(f"❌ 重启ADB服务失败")
            return False
    
    def get_flight_mode_status(self):
        """获取飞行模式状态"""
        if not self.device_id:
            return None
        
        cmd = f"-s {self.device_id} shell settings get global airplane_mode_on"
        output, error, code = self.run_adb_command(cmd)
        
        if code == 0:
            status = output.strip()
            if status in ["0", "1"]:
                return status == "1"
        
        return None
    
    def _force_radio_off(self):
        """强制关闭无线电硬件"""
        print("📡 强制关闭无线电硬件...")
        
        # 方法1: 先关闭移动数据
        print("  1. 关闭移动数据...")
        data_cmds = [
            f"-s {self.device_id} shell svc data disable",
            f"-s {self.device_id} shell settings put global mobile_data 0",
        ]
        
        for cmd in data_cmds:
            output, error, code = self.run_adb_command(cmd)
            time.sleep(0.3)
        
        # 方法2: 关闭网络相关服务
        print("  2. 关闭网络服务...")
        network_cmds = [
            f"-s {self.device_id} shell svc wifi disable",
            f"-s {self.device_id} shell svc bluetooth disable",
        ]
        
        for cmd in network_cmds:
            self.run_adb_command(cmd)
            time.sleep(0.3)
        
        # 方法3: 使用content命令设置飞行模式
        print("  3. 底层设置飞行模式...")
        content_cmd = f"-s {self.device_id} shell content insert --uri content://settings/global --bind name:s:airplane_mode_on --bind value:i:1"
        self.run_adb_command(content_cmd)
        
        # 方法4: Android 8.0+ 专用命令
        print("  4. 系统级飞行模式命令...")
        cmd_connectivity = f"-s {self.device_id} shell cmd connectivity airplane-mode enable"
        self.run_adb_command(cmd_connectivity)
        
        # 方法5: 设置基带相关属性
        print("  5. 设置基带属性...")
        radio_cmds = [
            f"-s {self.device_id} shell setprop gsm.radio.disabled 1",
            f"-s {self.device_id} shell setprop persist.radio.airplane_mode_on 1",
        ]
        
        for cmd in radio_cmds:
            self.run_adb_command(cmd)
            time.sleep(0.3)
        
        time.sleep(2)
        print("✅ 无线电强制关闭完成")
    
    def _force_radio_on(self):
        """强制开启无线电硬件"""
        print("📡 强制开启无线电硬件...")
        
        # 步骤1: 清除飞行模式标志
        print("  1. 清除飞行模式标志...")
        flight_cmds = [
            f"-s {self.device_id} shell settings put global airplane_mode_on 0",
            f"-s {self.device_id} shell content insert --uri content://settings/global --bind name:s:airplane_mode_on --bind value:i:0",
        ]
        
        for cmd in flight_cmds:
            self.run_adb_command(cmd)
            time.sleep(0.3)
        
        # 步骤2: 清除基带属性
        print("  2. 清除基带属性...")
        radio_cmds = [
            f"-s {self.device_id} shell setprop gsm.radio.disabled 0",
            f"-s {self.device_id} shell setprop persist.radio.airplane_mode_on 0",
        ]
        
        for cmd in radio_cmds:
            self.run_adb_command(cmd)
            time.sleep(0.3)
        
        # 步骤3: 开启所有网络服务
        print("  3. 开启网络服务...")
        network_cmds = [
            f"-s {self.device_id} shell svc data enable",          # 开启移动数据
            f"-s {self.device_id} shell svc wifi enable",          # 开启WiFi
            f"-s {self.device_id} shell svc bluetooth enable",     # 开启蓝牙
        ]
        
        for cmd in network_cmds:
            output, error, code = self.run_adb_command(cmd)
            time.sleep(0.3)
        
        # 步骤4: 设置移动数据属性
        print("  4. 设置移动数据属性...")
        data_cmds = [
            f"-s {self.device_id} shell settings put global mobile_data 1",
            f"-s {self.device_id} shell cmd connectivity data enable",
            f"-s {self.device_id} shell cmd connectivity airplane-mode disable",
        ]
        
        for cmd in data_cmds:
            self.run_adb_command(cmd)
            time.sleep(0.3)
        
        # 步骤5: 发送网络状态变更广播
        print("  5. 发送网络状态变更广播...")
        broadcast_cmds = [
            f"-s {self.device_id} shell am broadcast -a android.net.conn.CONNECTIVITY_CHANGE",
            f"-s {self.device_id} shell am broadcast -a android.net.wifi.STATE_CHANGE",
            f"-s {self.device_id} shell am broadcast -a android.intent.action.SERVICE_STATE",
        ]
        
        for cmd in broadcast_cmds:
            self.run_adb_command(cmd)
            time.sleep(0.3)
        
        time.sleep(2)
        print("✅ 无线电强制开启完成")
    
    def enable_flight_mode(self):
        """开启飞行模式"""
        if not self.device_id:
            print("❌ 未连接设备")
            return False
        
        print("🔄 开启飞行模式...")
        
        # 1. 强制关闭无线电
        self._force_radio_off()
        
        # 2. 设置飞行模式标志
        cmd = f"-s {self.device_id} shell settings put global airplane_mode_on 1"
        output, error, code = self.run_adb_command(cmd)
        
        if code != 0:
            print("❌ 设置飞行模式失败")
            return False
        
        # 3. 发送广播
        print("📢 发送系统广播...")
        broadcast_cmds = [
            f"-s {self.device_id} shell am broadcast -a android.intent.action.AIRPLANE_MODE --ez state true",
            f"-s {self.device_id} shell am broadcast -a android.intent.action.AIRPLANE_MODE_CHANGED --ez state true",
        ]
        
        for cmd in broadcast_cmds:
            self.run_adb_command(cmd)
            time.sleep(0.2)
        
        # 4. 验证
        time.sleep(2)
        status = self.get_flight_mode_status()
        
        if status:
            print("✅ 飞行模式已开启")
            return True
        else:
            print("⚠️ 状态未知，但命令已执行")
            return True
    
    def disable_flight_mode(self):
        """关闭飞行模式"""
        if not self.device_id:
            print("❌ 未连接设备")
            return False
        
        print("🔄 关闭飞行模式...")
        
        # 1. 清除飞行模式标志
        cmd = f"-s {self.device_id} shell settings put global airplane_mode_on 0"
        output, error, code = self.run_adb_command(cmd)
        
        if code != 0:
            print("❌ 关闭飞行模式失败")
            return False
        
        # 2. 发送关闭广播
        print("📢 发送关闭广播...")
        broadcast_cmds = [
            f"-s {self.device_id} shell am broadcast -a android.intent.action.AIRPLANE_MODE --ez state false",
            f"-s {self.device_id} shell am broadcast -a android.intent.action.AIRPLANE_MODE_CHANGED --ez state false",
        ]
        
        for cmd in broadcast_cmds:
            self.run_adb_command(cmd)
            time.sleep(0.2)
        
        # 3. 强制开启无线电硬件（关键修复）
        self._force_radio_on()
        
        # 4. 验证
        time.sleep(2)
        status = self.get_flight_mode_status()
        
        if status is not None and not status:
            print("✅ 飞行模式已关闭，无线电已开启")
            return True
        else:
            print("⚠️ 状态未知，但命令已执行")
            return True
    
    
    def timed_flight_mode_control(self, wait_time=5):

        if not self.device_id:
            print("❌ 未连接设备")
            return False
        
        print(f"📱 开始定时飞行模式控制")
        print(f"设备: {self.device_id}")
        print(f"等待时间: {wait_time}秒")
        
        try:
            print(f"\n1. 开启飞行模式...")
            enable_success = self.enable_flight_mode()
            
            if not enable_success:
                print("❌ 开启飞行模式失败")
                return False
            
            # 2. 等待
            print(f"\n2. 等待 {wait_time} 秒...")
            for i in range(wait_time, 0, -1):
                print(f"\r剩余: {i} 秒", end="", flush=True)
                time.sleep(1)
            print(f"\n✅ 等待完成")
            
            # 3. 关闭飞行模式
            print(f"\n3. 关闭飞行模式...")
            disable_success = self.disable_flight_mode()
            
            if not disable_success:
                print("❌ 关闭飞行模式失败")
                return False
            
            print(f"\n4. 等待网络恢复...")
            time.sleep(5)
            
            print(f"\n✅ 控制完成")
            print(f"总时间: {wait_time + 7}秒")
            
            return enable_success and disable_success
            
        except Exception as e:
            print(f"❌ 定时控制失败: {e}")
            return False

def interactive_menu():
    """交互式菜单"""
    print("="*50)
    print("📱 Android飞行模式控制工具")
    print("="*50)
    
    # 创建控制器
    controller = ADBFlightModeController()
    
    if not controller.adb_path or not controller.device_id:
        print("\n❌ 无法继续，请解决以上问题后重新运行程序")
        input("\n按回车键退出...")
        return
    
    print("\n✅ 准备就绪")
    
    while True:
        print("\n" + "="*30)
        print("📱 飞行模式控制")
        print("="*30)
        
        # 显示当前状态
        status = controller.get_flight_mode_status()
        if status is not None:
            if status:
                print("当前状态: 🔴 飞行模式已开启")
            else:
                print("当前状态: 🟢 飞行模式已关闭")
        else:
            print("当前状态: ❓ 无法获取状态")
        
        print("\n请选择操作:")
        print("1. 开启飞行模式")
        print("2. 关闭飞行模式") 
        print("4. 定时控制")
        print("5. 退出程序")
        
        try:
            choice = input("\n请输入选项 (1-5): ").strip()
        except KeyboardInterrupt:
            print("\n\n👋 程序被中断")
            break
        
        if choice == "1":
            print("\n正在开启飞行模式...")
            if controller.enable_flight_mode():
                print("✅ 操作完成")
            else:
                print("❌ 操作失败")
        elif choice == "2":
            print("\n正在关闭飞行模式...")
            if controller.disable_flight_mode():
                print("✅ 操作完成")
            else:
                print("❌ 操作失败")
        elif choice == "4":
            print("\n定时飞行模式控制...")
            wait_time = 5  # 默认5秒
            try:
                wait_input = input("请输入飞行模式开启后的等待时间(秒,默认5): ").strip()
                if wait_input:
                    wait_time = int(wait_input)
                    if wait_time <= 0:
                        wait_time = 5
            except:
                wait_time = 5
            
            if controller.timed_flight_mode_control(wait_time):
                print("✅ 定时控制完成")
            else:
                print("❌ 定时控制失败")
        elif choice == "5":
            print("\n👋 再见！")
            break
        else:
            print("❌ 无效选项，请重新输入")
        
        time.sleep(1)

def quick_test():
    """快速测试"""
    print("="*50)
    print("📱 快速飞行模式测试")
    print("="*50)
    
    # 获取等待时间
    try:
        wait_input = input("\n请输入开启飞行模式后的等待时间(秒,默认5): ").strip()
        if wait_input:
            wait_time = int(wait_input)
            if wait_time <= 0:
                print("⚠️ 时间必须大于0,设置为5秒")
                wait_time = 5
        else:
            wait_time = 5
    except ValueError:
        print("⚠️ 无效输入,设置为5秒")
        wait_time = 5
    
    print(f"\n⏰ 将执行: 开启飞行模式 → 等待{wait_time}秒 → 关闭飞行模式")
    print("="*50)
    
    # 创建控制器
    controller = ADBFlightModeController()
    
    if not controller.adb_path or not controller.device_id:
        print("\n❌ ADB或设备连接失败")
        return False
    
    print(f"\n✅ 连接成功!")
    print(f"设备ID: {controller.device_id}")
    
    print(f"\n🚀 开始飞行模式测试...")
    return controller.timed_flight_mode_control(wait_time)

if __name__ == "__main__":
    print("="*60)
    print("📱 飞行模式控制工具 v1.4(双向无线电强制控制版）")
    print("="*60)
    
    while True:
        print("\n选择模式:")
        print("1. 快速测试")
        print("2. 交互式菜单")
        print("3. 退出程序")
        
        try:
            mode = input("\n请选择模式 (1-3): ").strip()
        except KeyboardInterrupt:
            print("\n\n程序退出")
            sys.exit(0)
        
        if mode == "1":
            success = quick_test()
            
            if success:
                print("\n✅ 快速测试完成")
            else:
                print("\n❌ 快速测试失败")
            
            cont = input("\n是否继续其他操作?(y/N): ").strip().lower()
            if cont != 'y':
                print("👋 再见！")
                break
                
        elif mode == "2":
            interactive_menu()
            break
        elif mode == "3":
            print("👋 再见！")
            break
        else:
            print("❌ 无效选择，请重新输入")