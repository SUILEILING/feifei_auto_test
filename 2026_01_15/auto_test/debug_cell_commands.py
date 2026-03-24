import socket
import time
import threading
import pyvisa
import argparse
from datetime import datetime

# 从配置文件中导入网络配置
from souren_config import DEFAULT_IP, DEFAULT_PORT, get_visa_address

class NetworkMonitor:
    """网络流量监控器"""
    
    def __init__(self, host=None, port=None):
        # 使用配置文件中的默认值
        self.host = host or DEFAULT_IP
        self.port = port or DEFAULT_PORT
        self.sock = None
        self.monitoring = False
        self.traffic_log = []
        
    def start_monitoring(self):
        """启动网络监控"""
        print(f"🌐 开始监控网络流量: {self.host}:{self.port}")
        
        try:
            # 创建socket监听
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.sock.settimeout(2.0)
            
            # 尝试绑定到仪器端口（可能需要管理员权限）
            try:
                self.sock.bind((self.host, self.port))
                self.sock.listen(1)
                print(f"✅ 端口监听已启动")
            except OSError as e:
                print(f"⚠️  无法绑定到端口 {self.port}: {e}")
                print(f"💡 尝试使用Wireshark或tcpdump抓包")
                return
            
            self.monitoring = True
            
            # 启动监听线程
            monitor_thread = threading.Thread(target=self._listen_for_traffic)
            monitor_thread.daemon = True
            monitor_thread.start()
            
            print(f"\n📡 监听中... (等待Souren.ToolSet连接)")
            print(f"💡 请执行:")
            print(f"   1. 打开Souren.ToolSet")
            print(f"   2. 连接到 {self.host}")
            print(f"   3. 发送 CALL:CELL1 ON")
            print(f"   4. 观察这里的输出")
            
            # 保持主线程运行
            try:
                while self.monitoring:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\n⏹️ 停止监控")
                self.monitoring = False
                
        except Exception as e:
            print(f"❌ 监控失败: {e}")
    
    def _listen_for_traffic(self):
        """监听网络流量"""
        while self.monitoring:
            try:
                client_socket, client_address = self.sock.accept()
                print(f"🔗 检测到连接来自: {client_address}")
                
                # 接收数据
                client_socket.settimeout(1.0)
                try:
                    while True:
                        data = client_socket.recv(1024)
                        if data:
                            timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
                            decoded = self._decode_data(data)
                            print(f"[{timestamp}] 📥 接收到: {repr(decoded)}")
                            self.traffic_log.append({
                                'time': timestamp,
                                'direction': 'in',
                                'data': decoded,
                                'raw': data.hex()
                            })
                        else:
                            break
                except socket.timeout:
                    pass
                except Exception as e:
                    print(f"接收错误: {e}")
                
                client_socket.close()
                
            except socket.timeout:
                continue
            except Exception as e:
                if self.monitoring:
                    print(f"监听错误: {e}")
                break
    
    def _decode_data(self, data):
        """解码数据"""
        try:
            # 尝试UTF-8
            return data.decode('utf-8', errors='ignore')
        except:
            try:
                # 尝试ASCII
                return data.decode('ascii', errors='ignore')
            except:
                # 返回十六进制
                return data.hex()

class VISACommandLogger:
    """VISA命令记录器（代理模式）"""
    
    def __init__(self, instrument_address=None):
        # 使用配置文件中的默认值
        if instrument_address is None:
            instrument_address = get_visa_address()
        self.instrument_address = instrument_address
        self.rm = pyvisa.ResourceManager()
        self.real_instrument = None
        self.command_log = []
        
    def connect_as_proxy(self):
        """作为代理连接仪器"""
        print(f"\n🔌 作为代理连接到仪器...")
        
        try:
            # 连接仪器
            self.real_instrument = self.rm.open_resource(self.instrument_address)
            self.real_instrument.timeout = 5000
            
            print(f"✅ 仪器连接成功")
            print(f"💡 现在可以通过这个代理发送命令，所有通信都会被记录")
            
            return True
        except Exception as e:
            print(f"❌ 连接失败: {e}")
            return False
    
    def log_command(self, command, response=None, error=None):
        """记录命令和响应"""
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
        
        log_entry = {
            'time': timestamp,
            'command': command,
            'response': response,
            'error': error
        }
        
        self.command_log.append(log_entry)
        
        # 实时显示
        print(f"[{timestamp}] 📤 发送: {repr(command)}")
        if response:
            print(f"[{timestamp}] 📥 响应: {repr(response)}")
        if error:
            print(f"[{timestamp}] ❌ 错误: {error}")

def test_souren_commands(instrument_address=None):
    """测试Souren.ToolSet可能的命令"""
    print("\n" + "="*60)
    print(f"🔍 直接测试可能的命令 - 仪器地址: {instrument_address}")
    print("="*60)
    
    rm = pyvisa.ResourceManager()
    
    # 列出所有资源
    print("🔍 可用VISA资源:")
    resources = rm.list_resources()
    for i, resource in enumerate(resources):
        print(f"  {i+1}. {resource}")
    
    # 连接到仪器
    try:
        if instrument_address is None:
            instrument_address = get_visa_address()
        instrument = rm.open_resource(instrument_address)
        instrument.timeout = 5000
        instrument.read_termination = '\n'
        instrument.write_termination = '\n'
        
        print(f"\n✅ 仪器连接成功")
        
        # 测试可能的命令序列
        # Souren.ToolSet可能在发送CELL命令前需要执行其他命令
        
        test_sequences = [
            # 序列1: 初始化命令
            ["*IDN?", "*CLS", "*RST", "CALL:CELL1 ON"],
            
            # 序列2: 配置命令后发送
            ["CONFigure:NR5G:MEAS:BLER", "INITiate", "CALL:CELL1 ON"],
            
            # 序列3: 使用FETCh命令
            ["FETCh:NR5G:MEAS:BLER:ALL?", "CALL:CELL1 ON"],
        ]
        
        for seq_idx, sequence in enumerate(test_sequences, 1):
            print(f"\n🧪 测试序列 {seq_idx}:")
            print(f"   命令序列: {' -> '.join(sequence)}")
            
            for cmd in sequence:
                print(f"   📤 发送: {cmd}")
                try:
                    if '?' in cmd:
                        response = instrument.query(cmd).strip()
                        print(f"      响应: {response}")
                    else:
                        instrument.write(cmd)
                        print(f"      已发送")
                    time.sleep(0.5)
                except Exception as e:
                    print(f"      ❌ 错误: {e}")
                    break
            
            time.sleep(2)
        
        instrument.close()
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")

def manual_command_test(instrument_address=None):
    """手动命令测试"""
    print("\n" + "="*60)
    print(f"🔧 手动命令测试 - 仪器地址: {instrument_address}")
    print("="*60)
    
    rm = pyvisa.ResourceManager()
    
    try:
        if instrument_address is None:
            instrument_address = get_visa_address()
        instrument = rm.open_resource(instrument_address)
        instrument.timeout = 5000
        
        print("✅ 仪器连接成功")
        print("💡 输入命令进行测试（输入 'quit' 退出）")
        print("   例如: CALL:CELL1 ON")
        print("         CELL1?")
        print("         *IDN?")
        
        while True:
            cmd = input("\n📤 输入命令: ").strip()
            
            if cmd.lower() == 'quit':
                break
            
            if not cmd:
                continue
            
            try:
                print(f"   发送: {repr(cmd)}")
                
                if '?' in cmd:
                    # 查询命令
                    try:
                        response = instrument.query(cmd).strip()
                        print(f"   响应: {repr(response)}")
                    except Exception as e:
                        print(f"   ❌ 查询错误: {e}")
                else:
                    # 设置命令
                    try:
                        instrument.write(cmd)
                        print(f"   已发送")
                    except Exception as e:
                        print(f"   ❌ 发送错误: {e}")
                
            except Exception as e:
                print(f"   ❌ 执行错误: {e}")
        
        instrument.close()
        
    except Exception as e:
        print(f"❌ 连接失败: {e}")

def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='Souren.ToolSet通信分析工具')
    parser.add_argument('--ip', default=DEFAULT_IP, 
                       help=f'仪器IP地址 (默认: {DEFAULT_IP})')
    parser.add_argument('--address', default=None, 
                       help=f'仪器VISA地址 (默认: 自动生成)')
    parser.add_argument('--port', type=int, default=DEFAULT_PORT, 
                       help=f'仪器端口 (默认: {DEFAULT_PORT})')
    
    args = parser.parse_args()
    
    ip_address = args.ip
    port = args.port
    
    # 生成VISA地址
    if args.address:
        visa_address = args.address
    else:
        visa_address = get_visa_address(ip_address)
    
    print("="*60)
    print("Souren.ToolSet通信分析工具")
    print("="*60)
    print(f"📡 仪器IP: {ip_address}")
    print(f"🔌 VISA地址: {visa_address}")
    print(f"🚪 端口: {port}")
    print("="*60)
    
    while True:
        print("\n选择操作:")
        print("1. 监控网络流量（需要管理员权限）")
        print("2. 测试命令序列")
        print("3. 手动命令测试")
        print("4. 退出")
        
        choice = input("> ").strip()
        
        if choice == "1":
            monitor = NetworkMonitor(host=ip_address, port=port)
            monitor.start_monitoring()
        elif choice == "2":
            test_souren_commands(visa_address)
        elif choice == "3":
            manual_command_test(visa_address)
        elif choice == "4":
            print("退出")
            break
        else:
            print("无效选择")

if __name__ == "__main__":
    main()