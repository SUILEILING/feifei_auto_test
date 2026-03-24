from lib.var import *

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCV_FOLDER = os.path.join(BASE_DIR, "save_scv")
DEFAULT_SCV_NAME = "night_cycle_update_feifei.scv"  

USE_DYNAMIC_SCV = True  # ⭐⭐⭐ 设置为 True 使用py生成SCV 

PYTHON_SCRIPT_NAME = "sa_test.py" # ⭐⭐⭐ 使用函数动态获取脚本路径

EXECUTION_MODE = 'loop_info'  # ⭐⭐⭐ loop_info or run_all  
LOOP_COUNT = 2  # 只有loop_info模式才能使用

DEFAULT_IP = '192.168.30.122'  # ⭐⭐⭐ 网络配置

DEVICE_TYPE = 'phone'  # ⭐⭐⭐ 可选值: 'board' 或 'phone' 

USE_CHART_DATA_EXTRACTION = True  # ⭐⭐⭐  ==True 时使用 # 格式：{命令模式: [提取位置1, 提取位置2, ...]}

# 示例：{"FETCh:NR:BLER:DL:RESult?": [7]} 表示从该命令响应中提取第7个数值（从1开始计数）
CHART_DATA_EXTRACTION_CONFIG = {
    "FETCh:NR:BLER:UL:RESult?": [9], 
    "FETCh:NR:BLER:DL:RESult?": [9], 
    "FETCh:NR:MEValuation:TXP:AVG?": [2],  

}

CHART_STEPS_FILTER = [14, 15, 16, 20]  # 只显示这些步骤的数据  USE_CHART_DATA_EXTRACTION==False 时使用


SCV_CONVERTER_MODULE = "scv_converter"  # SCV转换器模块名 

# =========================（仅 board 模式使用）=====================
# 在 CELL_COMMANDS_NEED_AT 后执行 USB AT 串口配置
CELL_COMMANDS_NEED_AT = [
    "CALL:CELL1 ON",
]
USB_AT_PORT = "COM14"
USB_AT_BAUDRATE = 115200
USB_AT_TIMEOUT = 3  
AT_SEQUENCE = [
    ("AT+CFUN=0", 5),   # 发送 AT+CFUN=0，等待 5秒
    ("AT+CFUN=1", 1)    # 发送 AT+CFUN=1，等待 1秒
]
# =================================================================


# ============================（仅 phone 模式使用）==================
# 在 CELL_COMMANDS_NEED_AT 后执行 ADB飞行模式控制
CELL_COMMANDS_NEED_ADB = [
    "CALL:CELL1 ON",
]
ADB_WAIT_TIME = 3  # 开启飞行模式后的等待时间（秒）后关闭飞行模式
# ================================================================


# ==============================================
# ⭐ Excel导出配置 ⭐

# 默认行高和列宽
EXCEL_DEFAULT_ROW_HEIGHT = 13.5
EXCEL_DEFAULT_COLUMN_WIDTH = 9
EXCEL_HEADER_ROW_HEIGHT = 20

# 执行汇总表 特殊列宽配置
EXCEL_SUMMARY_COLUMN_WIDTHS = {
    "B": 20,
    "C": 15,
    "D": 37,
    "N": 63,
}

# 详细执行记录表 特殊列宽配置
EXCEL_DETAILS_COLUMN_WIDTHS = {
    "B": 24,
    "F": 50,
    "I": 62,
}


# 表头样式配置
EXCEL_HEADER_FILL_COLOR = "366092"
EXCEL_HEADER_FONT_COLOR = "FFFFFF"
EXCEL_QUERY_RESULT_FILL = "E2F0D9"
EXCEL_QUERY_RESULT_FONT = "006400"

# 字体配置
EXCEL_HEADER_FONT_SIZE = 11
EXCEL_DATA_FONT_SIZE = 10
# ==============================================

# =============== 特殊命令配置 ===============================
# 即使仪器通信错误也视为通过的特定命令
COMMUNICATION_ERROR_PASS_COMMANDS = [
    # "CALL:CELL1?",  # 已知查询不响应
    # "CELL1?",       # 已知查询不响应
]
# ==============================================


# 日志配置
LOG_ENABLED = True
LOG_LEVEL = 'INFO'

# 简化的CALL命令映射
# 仪器直接接受CALL:前缀的命令，不需要转换
CALL_COMMAND_MAPPINGS = {
    # 查询命令（如果响应）
    "CALL:*IDN?": "*IDN?",
    "CALL:*OPC?": "*OPC?",
    
    # 设置命令（直接发送）
    "CALL:*RST": "*RST",
    "CALL:*CLS": "*CLS",
}

# 输出配置
SHOW_STEP_TYPE = False
SHOW_STEP_CONTENT = True
SHOW_COMMAND_SENDING = False
SHOW_COMMAND_RESULT = False


def _get_execution_dir():
    """获取执行目录，仅在需要时创建"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    execution_dir = os.path.join(BASE_DIR, "log", f"execution_{timestamp}")
    
    if LOG_ENABLED:
        os.makedirs(execution_dir, exist_ok=True)
    
    return execution_dir

def _get_log_file():
    """获取日志文件路径"""
    if not LOG_ENABLED:
        return None
    
    execution_dir = _get_execution_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(execution_dir, f"souren_execution_{timestamp}.log")

def _get_result_file():
    """获取结果文件路径"""
    execution_dir = _get_execution_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(execution_dir, f"souren_results_{timestamp}.json")


EXECUTION_DIR = _get_execution_dir() if LOG_ENABLED else None
LOG_FILE = _get_log_file()
RESULT_FILE = _get_result_file()


def get_scv_file_path():
    if USE_DYNAMIC_SCV:
        print(f"🔄 动态SCV模式激活,将生成新的SCV文件")
        scv_file = generate_dynamic_scv()
        if scv_file:
            scv_file_name = os.path.basename(scv_file)
            print(f"📄 生成的SCV文件: {scv_file_name}")
            print(f"📁 完整路径: {scv_file}")
            return scv_file
        else:
            print("⚠️  动态SCV生成失败,使用默认SCV文件")
            scv_path = os.path.join(os.path.abspath(SCV_FOLDER), DEFAULT_SCV_NAME)
            return scv_path
    else:
        print(f"📄 使用静态SCV文件: {DEFAULT_SCV_NAME}")
    scv_path = os.path.join(os.path.abspath(SCV_FOLDER), DEFAULT_SCV_NAME)
    return scv_path

def generate_dynamic_scv():
    try:
        print("🔄 开始动态生成SCV文件...")
        
        python_script_path = PYTHON_SCRIPT_PATH
        
        if not os.path.exists(python_script_path):
            print(f"❌ Python脚本不存在: {python_script_path}")
            print(f"🔍 尝试在以下位置查找Python文件:")
            print(f"   1. {SCV_FOLDER}")
            print(f"   2. {BASE_DIR}")
            
            if os.path.exists(SCV_FOLDER):
                py_files = [f for f in os.listdir(SCV_FOLDER) if f.endswith('.py')]
                if py_files:
                    print(f"📄 找到 {len(py_files)} 个Python文件:")
                    for py_file in py_files:
                        print(f"   - {py_file}")
                    python_script_path = os.path.join(SCV_FOLDER, py_files[0])
                    print(f"🔄 使用第一个找到的Python文件: {py_files[0]}")
                else:
                    print("❌ 在save_scv文件夹中未找到Python文件")
                    return None
            else:
                print(f"❌ SCV文件夹不存在: {SCV_FOLDER}")
                return None
        
        print(f"📁 解析Python脚本: {os.path.basename(python_script_path)}")
        print(f"📄 输出格式: {SCV_OUTPUT_FORMAT}")
        
        try:
            from scv_converter import generate_dynamic_scv_from_python
        except ImportError as e:
            print(f"❌ 无法导入SCV转换器: {e}")
            print(f"💡 请确保 scv_converter.py 在同一目录下")
            return None
        
        scv_file = generate_dynamic_scv_from_python(
            python_script_path,
            output_format=SCV_OUTPUT_FORMAT,
            output_filename=DYNAMIC_SCV_NAME,
            auto_cleanup=True  
        )
        
        if scv_file:
            print(f"✅ SCV文件生成成功: {scv_file}")
            print(f"🗑️  注意: 此文件将在程序结束时自动删除")
            
            try:
                with open(scv_file, 'r', encoding='utf-8') as f:
                    first_line = f.readline().strip()
                    print(f"🔍 第一行内容: {first_line}")
                    
                    if "命令:" in first_line or "步骤" in first_line:
                        print(f"⚠️  警告: 检测到表格格式，可能无法正确执行")
                    elif "CALL:" in first_line or "CONFigure" in first_line:
                        print(f"✅ 格式正确: 包含有效命令")
                    else:
                        print(f"❓ 未知格式，请检查")
            except Exception as e:
                print(f"⚠️  无法验证文件格式: {e}")
            
            return scv_file
        else:
            print(f"❌ SCV文件生成失败")
            return None
            
    except Exception as e:
        print(f"❌ 动态生成SCV文件时出错: {e}")
        import traceback
        traceback.print_exc()
        return None

def find_python_script(script_name):
    scv_folder_path = os.path.join(SCV_FOLDER, script_name)
    if os.path.exists(scv_folder_path):
        return scv_folder_path
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    current_dir_path = os.path.join(current_dir, script_name)
    if os.path.exists(current_dir_path):
        return current_dir_path
    
    if os.path.exists(SCV_FOLDER):
        for file in os.listdir(SCV_FOLDER):
            if file.endswith('.py'):
                print(f"📄 找到Python文件: {file}")
                return os.path.join(SCV_FOLDER, file)

    return os.path.join(SCV_FOLDER, script_name)


PYTHON_SCRIPT_PATH = find_python_script(PYTHON_SCRIPT_NAME)
DYNAMIC_SCV_NAME = "dynamic_test_steps.scv"  
SCV_OUTPUT_FORMAT = "standard"  

def get_visa_address(ip_address=None):
    """根据IP地址生成VISA地址"""
    if ip_address is None:
        ip_address = DEFAULT_IP
    return f"TCPIP0::{ip_address}::inst0::INSTR"

INSTRUMENT_ADDRESS = get_visa_address()


def display_config_info():
    if LOG_ENABLED and EXECUTION_DIR:
        print(f"📁 本次执行目录: {os.path.abspath(EXECUTION_DIR)}")
        print(f"📁 日志文件: {os.path.abspath(LOG_FILE)}")
        print(f"📁 结果文件: {os.path.abspath(RESULT_FILE)}")
    else:
        print(f"📁 日志功能已禁用或未启用")
    
    print(f"🔌 仪器地址: {INSTRUMENT_ADDRESS}")
    print(f"📝 确认: 仪器直接接受CALL:前缀的命令")
    print(f"📱 设备类型: {DEVICE_TYPE.upper()}")
    print(f"📋 SCV模式: {'动态生成' if USE_DYNAMIC_SCV else '使用现有SCV文件'}")
    
    if USE_DYNAMIC_SCV:
        print(f"🔄 动态SCV源: {os.path.basename(PYTHON_SCRIPT_PATH)}")
        print(f"📄 生成格式: {SCV_OUTPUT_FORMAT}")
    else:
        print(f"📄 静态SCV文件: {DEFAULT_SCV_NAME}")
    
    if DEVICE_TYPE.lower() == 'board':
        print(f"📡 USB AT端口: {USB_AT_PORT} (波特率: {USB_AT_BAUDRATE})")
        print(f"🔧 AT控制命令: {CELL_COMMANDS_NEED_AT}")
    elif DEVICE_TYPE.lower() == 'phone':
        print(f"📱 手机模式: ADB飞行模式控制")
        print(f"⏰ 等待时间: {ADB_WAIT_TIME}秒")
        print(f"🔧 ADB控制命令: {CELL_COMMANDS_NEED_ADB}")