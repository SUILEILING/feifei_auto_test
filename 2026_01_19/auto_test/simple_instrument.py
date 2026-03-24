from lib.var import *

class SimpleInstrumentController:
    """简化的仪器控制器 - 直接发送命令并等待"""
    
    def __init__(self, ip_address='192.168.30.122'):
        self.ip_address = ip_address
        self.address = f"TCPIP0::{ip_address}::inst0::INSTR"
        self.rm = None
        self.instrument = None
        self.connected = False
        
    def connect(self):
        """连接仪器"""
        try:
            self.rm = pyvisa.ResourceManager()
            print(f"🔌 连接仪器: {self.address}")
            
            self.instrument = self.rm.open_resource(self.address)
            self.instrument.timeout = 10000  # 10秒超时
            self.instrument.read_termination = '\n'
            self.instrument.write_termination = '\n'
            
            # 测试连接
            try:
                idn = self.instrument.query('*IDN?').strip()
                print(f"✅ 连接成功: {idn}")
                self.connected = True
                return True, idn
            except Exception as e:
                print(f"⚠️  连接测试失败: {e}")
                return False, f"连接失败: {e}"
                
        except Exception as e:
            print(f"❌ 连接失败: {e}")
            return False, str(e)
    
    def disconnect(self):
        """断开连接"""
        if self.instrument:
            try:
                self.instrument.close()
                print("📴 断开连接")
            except:
                pass
        self.connected = False
    
    def write(self, command, wait_time=0.5):
        """写入命令并等待"""
        if not self.connected:
            print(f"❌ 仪器未连接: {command}")
            return False, "仪器未连接"
        
        try:
            print(f"📤 发送: {command}")
            self.instrument.write(command)
            
            if wait_time > 0:
                time.sleep(wait_time)
            
            return True, "命令已发送"
        except Exception as e:
            print(f"❌ 发送失败: {command} - {e}")
            return False, str(e)
    
    def query(self, command, timeout=None):
        """查询命令"""
        if not self.connected:
            print(f"❌ 仪器未连接: {command}")
            return False, "仪器未连接"
        
        try:
            print(f"📤 查询: {command}")
            
            # 临时设置超时时间
            original_timeout = self.instrument.timeout
            if timeout:
                self.instrument.timeout = timeout
            
            result = self.instrument.query(command).strip()
            
            # 恢复原始超时时间
            if timeout:
                self.instrument.timeout = original_timeout
            
            print(f"📥 响应: {result}")
            return True, result
            
        except Exception as e:
            print(f"❌ 查询失败: {command} - {e}")
            return False, str(e)