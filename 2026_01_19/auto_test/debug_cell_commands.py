import pyvisa
from lib.var import socket, time, threading, argparse,datetime

from souren_config import DEFAULT_IP, get_visa_address

class VISACommandLogger:
    
    def __init__(self, instrument_address=None):
        if instrument_address is None:
            instrument_address = get_visa_address()
        self.instrument_address = instrument_address
        self.rm = pyvisa.ResourceManager()
        self.real_instrument = None
        self.command_log = []

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
                    try:
                        response = instrument.query(cmd).strip()
                        print(f"   响应: {repr(response)}")
                    except Exception as e:
                        print(f"   ❌ 查询错误: {e}")
                else:
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
    parser = argparse.ArgumentParser(description='Test scpi cmd sending yc1100 ')
    parser.add_argument('--ip', default=DEFAULT_IP, 
                       help=f'仪器IP地址 (默认: {DEFAULT_IP})')
    parser.add_argument('--address', default=None, 
                       help=f'仪器VISA地址 (默认: 自动生成)')

    args = parser.parse_args()
    
    ip_address = args.ip
    
    # 生成VISA地址
    if args.address:
        visa_address = args.address
    else:
        visa_address = get_visa_address(ip_address)
    
    print("="*60)
    print("Test scpi cmd sending yc1100")
    print("="*60)
    print(f"📡 仪器IP: {ip_address}")
    print(f"🔌 VISA地址: {visa_address}")
    print("="*60)
    
    while True:
        print("\n选择操作:")
        print("1. 手动命令测试")
        print("2. 退出")
        
        choice = input("> ").strip()
        
        if choice == "1":
            manual_command_test(visa_address)
        elif choice == "2":
            print("退出")
            break
        else:
            print("无效选择")

if __name__ == "__main__":
    main()