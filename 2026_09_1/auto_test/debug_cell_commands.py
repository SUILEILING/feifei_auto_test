import pyvisa
from lib.var import socket, time, threading, argparse, datetime

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
    """手动交互式命令测试"""
    print("\n" + "=" * 60)
    print(f"🔧 手动命令测试 - 仪器地址: {instrument_address}")
    print("=" * 60)

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


def fei(instrument_address):
    print("\n🚀 开始执行预定义配置...")
    rm = pyvisa.ResourceManager()
    try:
        instrument = rm.open_resource(instrument_address)
        instrument.timeout = 5000
        print("✅ 仪器连接成功")

        # 定义所有要发送的命令
        commands = [
            'CONFigure:CELL1:NR:Sign:SLOT:Clear',
            'CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT5:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT6:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT3:DL:TIND 5',
            'CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 4',
            'CONFigure:CELL1:NR:SIGN:SLOT5:DL:TIND 3',
            'CONFigure:CELL1:NR:SIGN:SLOT6:DL:TIND 2',
            'CONFigure:CELL1:NR:SIGN:SLOT3:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT4:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT5:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT6:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT8:UL:MCS1 2',
            'CONFigure:CELL1:NR:SIGN:SLOT9:UL:MCS1 2',
            'CONFigure:CELL1:NR:SIGN:SLOT8:UL:RB 67,135',
            'CONFigure:CELL1:NR:SIGN:SLOT9:UL:RB 67,135',
            'CONFigure:CELL1:NR:SIGN:SLOT10:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT11:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT12:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT13:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT14:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT15:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT16:CTYPe PDSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT10:DL:TIND 8',
            'CONFigure:CELL1:NR:SIGN:SLOT11:DL:TIND 7',
            'CONFigure:CELL1:NR:SIGN:SLOT12:DL:TIND 6',
            'CONFigure:CELL1:NR:SIGN:SLOT13:DL:TIND 5',
            'CONFigure:CELL1:NR:SIGN:SLOT14:DL:TIND 4',
            'CONFigure:CELL1:NR:SIGN:SLOT15:DL:TIND 3',
            'CONFigure:CELL1:NR:SIGN:SLOT16:DL:TIND 2',
            'CONFigure:CELL1:NR:SIGN:SLOT10:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT11:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT12:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT13:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT14:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT15:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT16:DL:MCS1 4',
            'CONFigure:CELL1:NR:SIGN:SLOT18:CTYPe PUSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT19:CTYPe PUSCh',
            'CONFigure:CELL1:NR:SIGN:SLOT18:UL:MCS1 2',
            'CONFigure:CELL1:NR:SIGN:SLOT19:UL:MCS1 2',
            'CONFigure:CELL1:NR:SIGN:SLOT18:UL:RB 67,135',
            'CONFigure:CELL1:NR:SIGN:SLOT19:UL:RB 67,135',
            # 'CONFigure:CELL1:NR:SIGN:SLOT:APPLy',
            # 'CONFigure:CELL1:NR:SIGN:SLOT:UPDate'
        ]

        for cmd in commands:
            print(f"   发送: {cmd}")
            instrument.write(cmd)
            time.sleep(0.05)

        print("✅ 所有命令执行完成")
        instrument.close()

    except Exception as e:
        print(f"❌ 执行配置时出错: {e}")




def main():
    parser = argparse.ArgumentParser(description='Test scpi cmd sending yc1100')
    parser.add_argument('--ip', default=DEFAULT_IP,
                        help=f'仪器IP地址 (默认: {DEFAULT_IP})')
    parser.add_argument('--address', default=None,
                        help=f'仪器VISA地址 (默认: 自动生成)')

    args = parser.parse_args()

    ip_address = args.ip

    if args.address:
        visa_address = args.address
    else:
        visa_address = get_visa_address(ip_address)

    print("=" * 60)
    print("Test scpi cmd sending yc1100")
    print("=" * 60)
    print(f"📡 仪器IP: {ip_address}")
    print(f"🔌 VISA地址: {visa_address}")
    print("=" * 60)

    while True:
        print("\n选择操作:")
        print("1. 手动命令测试")
        print("2. 退出")
        print("3. 执行预定义配置 (fei)")


        choice = input("> ").strip()

        if choice == "1":
            manual_command_test(visa_address)
        elif choice == "2":
            print("退出")
            break
        elif choice == "3":
            fei(visa_address)
        else:
            print("无效选择")


if __name__ == "__main__":
    main()