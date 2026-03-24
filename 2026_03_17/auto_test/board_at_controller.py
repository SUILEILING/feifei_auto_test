from lib.var import *

# ==================== 内部默认配置 ====================
DEFAULT_BAUDRATE = 115200
DEFAULT_TIMEOUT = 3
DEFAULT_AT_SEQUENCE = [
    ("AT+CFUN=0", 5),  
    ("AT+CFUN=1", 5),   
]
# ====================================================

def find_fibocom_at_port() -> Optional[str]:
    print("🔍 正在扫描Fibocom USB AT端口...")
    ports = serial.tools.list_ports.comports()
    for port in ports:
        desc = (port.description or "").lower()
        hwid = (port.hwid or "").lower()
        if "fibocom" in desc and "at" in desc:
            print(f"✅ 找到Fibocom AT端口: {port.device} - {port.description}")
            return port.device
        if "fibocom" in hwid and "at" in hwid:
            print(f"✅ 找到Fibocom AT端口 (HWID): {port.device}")
            return port.device
    print("⚠️  未找到Fibocom USB AT端口")
    return None

def send_at_sequence(
    port: Optional[str] = None,
    baudrate: int = DEFAULT_BAUDRATE,
    timeout: int = DEFAULT_TIMEOUT,
    at_sequence: List[Tuple[str, int]] = DEFAULT_AT_SEQUENCE
) -> Tuple[bool, str]:
    if port is None:
        port = find_fibocom_at_port()
        if port is None:
            return False, "❌ 未检测到Fibocom AT端口"

    print(f"\n📡 开始执行AT序列, 端口: {port}, 波特率: {baudrate}")
    print(f"📋 AT序列: {at_sequence}")

    ser = None
    try:
        ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout
        )
        print(f"✅ 串口连接成功: {port} ({baudrate}bps)")

        total_cmds = len(at_sequence)
        success_count = 0
        summary = []

        for idx, (cmd, wait_time) in enumerate(at_sequence, 1):
            print(f"\n{idx}/{total_cmds}: 执行 {cmd}")
            
            cmd_line = f"{cmd}\r\n"
            ser.write(cmd_line.encode())
            
            response = ""
            start_time = time.time()
            while time.time() - start_time < timeout:
                if ser.in_waiting:
                    line = ser.readline().decode(errors='ignore').strip()
                    if line:
                        print(f"  响应: {line}")
                        response += line + "\n"
                        if "OK" in line or "ERROR" in line:
                            break
                else:
                    time.sleep(0.05)
            
            if "OK" in response:
                print(f"  ✅ 成功")
                success_count += 1
                summary.append(f"{cmd}: OK")
            else:
                print(f"  ❌ 失败 (最后响应: {response[-50:] if response else '无响应'})")
                summary.append(f"{cmd}: FAIL")
            
            if wait_time > 0:
                print(f"  ⏳ 等待 {wait_time} 秒...")
                time.sleep(wait_time)

        ser.close()
        
        overall_success = (success_count == total_cmds)
        status = "✅ 全部成功" if overall_success else f"⚠️  部分成功 ({success_count}/{total_cmds})"
        print(f"\n{status}")
        return overall_success, f"{status} | {'; '.join(summary)}"

    except serial.SerialException as e:
        error_msg = f"❌ 串口打开失败: {e}"
        print(error_msg)
        return False, error_msg
    except Exception as e:
        error_msg = f"❌ AT序列执行异常: {e}"
        print(error_msg)
        return False, error_msg
    finally:
        if ser and ser.is_open:
            ser.close()

if __name__ == "__main__":
    print("="*60)
    print("📱 Fibocom AT端口独立调试工具")
    print("="*60)

    port = find_fibocom_at_port()
    if not port:
        print("\n❌ 未检测到Fibocom AT端口,请检查设备连接。")
        print("\n📋 当前所有可用串口:")
        ports = serial.tools.list_ports.comports()
        for p in ports:
            print(f"   {p.device} - {p.description} ({p.hwid})")
        sys.exit(1)

    print(f"\n✅ 自动检测到端口: {port}")
    
    try:
        choice = input("\n是否立即执行默认AT序列 (AT+CFUN=0→等待5秒→AT+CFUN=1) ? (Y/n): ").strip().lower()
        if choice not in ['n', 'no']:
            send_at_sequence(port)
        else:
            print("已取消执行AT序列。")
    except KeyboardInterrupt:
        print("\n\n用户中断")
        sys.exit(0)