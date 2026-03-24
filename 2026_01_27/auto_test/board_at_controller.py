from lib.var import serial, Tuple, List, time

class BoardATController:
    """Board AT 指令控制器 - 简化版"""
    
    def __init__(self):
        """
        初始化 AT 控制器 - 所有参数从 souren_config.py 读取
        """
        try:
            from souren_config import USB_AT_PORT, USB_AT_BAUDRATE, USB_AT_TIMEOUT, AT_SEQUENCE
            
            self.port = USB_AT_PORT
            self.baudrate = USB_AT_BAUDRATE
            self.timeout = USB_AT_TIMEOUT
            self.at_sequence = AT_SEQUENCE
            self.serial_conn = None
            
            print(f"📡 Board AT 控制器初始化完成")
            print(f"   端口: {self.port}")
            print(f"   波特率: {self.baudrate}")
            print(f"   AT序列: {self.at_sequence}")
            
        except ImportError as e:
            print(f"❌ 导入配置失败: {e}")
            raise
    
    def connect(self) -> bool:
        """连接串口"""
        try:
            print(f"🔌 连接串口 {self.port}...")
            self.serial_conn = serial.Serial(
                port=self.port,
                baudrate=self.baudrate,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                bytesize=serial.EIGHTBITS,
                timeout=self.timeout
            )
            
            print(f"✅ 串口连接成功")
            return True
            
        except Exception as e:
            print(f"❌ 连接串口失败: {e}")
            return False
    
    def send_at_command(self, command: str) -> Tuple[bool, str]:
        if not self.serial_conn or not self.serial_conn.is_open:
            success = self.connect()
            if not success:
                return False, "串口未连接"
        
        try:
            # 发送 AT 指令
            print(f"📤 发送: {command}")
            self.serial_conn.write(f"{command}\r\n".encode("utf-8"))
            
            # 读取响应
            time.sleep(0.5)  # 等待设备响应
            response = self.serial_conn.read_all().decode("utf-8", errors="ignore")
            
            print(f"📥 响应: {response.strip()}")
            return True, response
            
        except Exception as e:
            print(f"❌ 发送失败: {e}")
            return False, str(e)
    
    def execute_at_sequence(self) -> Tuple[bool, str]:
        """
        执行 AT 指令序列
        
        Returns:
            (success, summary)
        """
        print(f"🔄 执行 AT 序列: {self.at_sequence}")
        
        results = []
        
        for i, (command, wait_time) in enumerate(self.at_sequence, 1):
            print(f"\n{i}/{len(self.at_sequence)}: 执行 {command}")
            
            # 发送命令
            success, response = self.send_at_command(command)
            
            # 检查响应
            if success and "OK" in response.upper():
                results.append(f"{command}: 成功")
                print(f"   ✅ 成功")
            else:
                results.append(f"{command}: 失败")
                print(f"   ❌ 失败")
                return False, f"第{i}步失败: {command}"
            
            # 如果不是最后一步，等待配置的时间
            if i < len(self.at_sequence) and wait_time > 0:
                print(f"   ⏳ 等待 {wait_time} 秒...")
                time.sleep(wait_time)
        
        summary = " | ".join(results)
        print(f"\n✅ AT序列完成: {summary}")
        return True, summary
    
    def disconnect(self):
        """断开串口连接"""
        if self.serial_conn and self.serial_conn.is_open:
            try:
                self.serial_conn.close()
                print(f"🔌 串口已关闭")
            except:
                pass
            finally:
                self.serial_conn = None

# 全局控制器实例
_controller = None

def get_controller():
    """获取全局控制器"""
    global _controller
    return _controller

def init_controller() -> bool:
    """初始化控制器"""
    global _controller
    
    try:
        from souren_config import DEVICE_TYPE
        
        # 只在 BOARD 模式下初始化
        if DEVICE_TYPE.lower() != 'board':
            print(f"📱 Phone模式，跳过AT控制器")
            return True
            
        print(f"🔄 初始化AT控制器...")
        
        _controller = BoardATController()
        return _controller.connect()
        
    except Exception as e:
        print(f"❌ 初始化失败: {e}")
        return False

def cleanup_controller():
    """清理控制器"""
    global _controller
    if _controller:
        _controller.disconnect()
        _controller = None

# 直接调用函数
def send_at_sequence_directly() -> Tuple[bool, str]:
    """直接发送AT序列（不依赖全局控制器）"""
    try:
        from souren_config import DEVICE_TYPE, USB_AT_PORT, USB_AT_BAUDRATE, AT_SEQUENCE
        
        if DEVICE_TYPE.lower() != 'board':
            return True, "Phone模式，跳过AT序列"
        
        print(f"📡 直接发送AT序列到 {USB_AT_PORT}")
        print(f"📋 序列: {AT_SEQUENCE}")
        
        # 创建临时控制器
        controller = BoardATController()
        if not controller.connect():
            return False, "串口连接失败"
        
        # 执行序列
        success, summary = controller.execute_at_sequence()
        controller.disconnect()
        
        return success, summary
        
    except Exception as e:
        print(f"❌ 发送AT序列失败: {e}")
        return False, str(e)

# 测试代码
if __name__ == "__main__":
    print("🧪 测试AT控制器...")
    
    success, summary = send_at_sequence_directly()
    print(f"\n🎯 结果: {'成功' if success else '失败'} - {summary}")