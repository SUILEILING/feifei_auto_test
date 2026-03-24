from lib.var import *


PYTHON_SCRIPT_NAME = [

    # {"script": "sa_power.py","lineLoss": 25.00, "band": 78, "bw": 100, "scs": 30, "range": "LOW","current_power": "-80", "power_decrement": 1},

    # {"script": "sa_N1.py","lineLoss": 25.00, "band": 1, "bw": 20, "scs": 15, "range": "LOW"},
    # {"script": "sa_N1.py","lineLoss": 25.00, "band": 1, "bw": 20, "scs": 15, "range": "Mid"},
    # {"script": "sa_N1.py","lineLoss": 25.00, "band": 1, "bw": 20, "scs": 15, "range": "High"},

    # {"script": "sa_N41.py","lineLoss": 25.00, "band": 41, "bw": 100, "scs": 30, "range": "LOW"},
    # {"script": "sa_N41.py","lineLoss": 25.00, "band": 41, "bw": 100, "scs": 30, "range": "Mid"},
    # {"script": "sa_N41.py","lineLoss": 25.00, "band": 41, "bw": 100, "scs": 30, "range": "High"},

    {"script": "sa_N77.py","lineLoss": 25.00, "band": 77, "bw": 100, "scs": 30, "range": "LOW"},
    # {"script": "sa_N77.py","lineLoss": 25.00, "band": 77, "bw": 100, "scs": 30, "range": "Mid"},
    # {"script": "sa_N77.py","lineLoss": 25.00, "band": 77, "bw": 100, "scs": 30, "range": "High"},

    {"script": "sa_N78.py", "lineLoss": 25.00, "band": 78, "bw": 100, "scs": 30, "range": "LOW"},
    # {"script": "sa_N78.py", "lineLoss": 25.00, "band": 78, "bw": 100, "scs": 30, "range": "Mid"},
    # {"script": "sa_N78.py", "lineLoss": 25.00, "band": 78, "bw": 100, "scs": 30, "range": "High"},

    # {"script": "sa_N79.py", "lineLoss": 30.00, "band": 79, "bw": 100, "scs": 30, "range": "LOW"},
    # {"script": "sa_N79.py", "lineLoss": 30.00, "band": 79, "bw": 100, "scs": 30, "range": "Mid"},
    # {"script": "sa_N79.py", "lineLoss": 30.00, "band": 79, "bw": 100, "scs": 30, "range": "High"},

]

EXECUTION_MODE = 'loop_info'    
LOOP_COUNT =3  

DEFAULT_IP = '192.168.30.122'  # 网络配置

DEVICE_TYPE = 'phone'  # 可选值: 'board' 或 'phone' 

USE_CHART_DATA_EXTRACTION = True  # ==True 时使用 # 格式：{命令模式: [提取位置1, 提取位置2, ...]}

# 示例：{"FETCh:NR:BLER:DL:RESult?": [7]} 表示从该命令响应中提取第7个数值（从1开始计数）
CHART_DATA_EXTRACTION_CONFIG = {
    "FETCh:NR:BLER:UL:RESult?": [8], 
    "FETCh:NR:BLER:DL:RESult?": [8], 
    "FETCh:NR:MEValuation:TXP:AVG?": [2],  
}

CHART_STEPS_FILTER = [14, 15, 16, 20]  # 只显示这些步骤的数据  USE_CHART_DATA_EXTRACTION==False 时使用


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

# 执行汇总表 特殊列宽配置 - 更新为最新列结构
EXCEL_SUMMARY_COLUMN_WIDTHS = {
    "B": 5,    # 执行时间
    "C": 18,    # SCV文件
    "D": 30,    # 设备
    "E": 37,    # 参数信息
    "F": 10,    # 执行模式
    "G": 10,    # 循环次数
    "H": 10,    # 总步骤数
    "I": 10,    # 已执行步骤
    "J": 10,    # 通过步骤
    "K": 10,    # 失败步骤
    "L": 10,    # 成功率(%)
    "M": 10,    # 总耗时(秒)
    "N": 10,    # 状态
    "O": 10,    # 状态消息
}

# 详细执行记录表 特殊列宽配置
EXCEL_DETAILS_COLUMN_WIDTHS = {
    "B": 24,
    "F": 56,
    "I": 62,
}

# 数据分析图表表 特殊列宽配置
EXCEL_CHART_COLUMN_WIDTHS = {
    "default": 15,
    "A": 30,
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

# 全局执行目录变量
_execution_dir = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 参数化配置 - 添加lineLoss参数
PARAMETER_COMMAND_MAPPINGS = {
    "lineLoss": {
        "command": "CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,{lineLoss},6000000000,{lineLoss}",
        "pattern": r'CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,\d+\.\d+,6000000000,\d+\.\d+'
    },
    "band": {
        "command": "CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {band}",
        "pattern": r'CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator\s+\d+'
    },
    "bw": {
        "command": "CONFigure:CELL1:NR:SIGN:BWidth:DL BW{bw}",
        "pattern": r'CONFigure:CELL1:NR:SIGN:BWidth:DL\s+BW\d+'
    },
    "scs": {
        "command": "CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{scs}",
        "pattern": r'CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing\s+kHz\d+'
    },
    "range": {
        "command": "CONFigure:CELL1:NR:CONFig:RANGe {range}",
        "pattern": r'CONFigure:CELL1:NR:CONFig:RANGe\s+\w+'
    },
    "power": {
        "command": "CONFigure:CELL1:NR:SIGN:POWer {power}",
        "pattern": r'CONFigure:CELL1:NR:SIGN:POWer\s+[-\d]+'
    },
    "power_decrement": {
        "command": None,
        "pattern": None
    }
}



def set_execution_dir(dir_path):
    global _execution_dir
    _execution_dir = dir_path
    print(f"📁 设置执行目录为: {_execution_dir}")

def _get_execution_dir():
    """获取执行目录"""
    global _execution_dir
    
    if _execution_dir:
        if LOG_ENABLED and not os.path.exists(_execution_dir):
            os.makedirs(_execution_dir, exist_ok=True)
        return _execution_dir
    
    if LOG_ENABLED:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        _execution_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "log", f"execution_{timestamp}")
        os.makedirs(_execution_dir, exist_ok=True)
        print(f"⚠️ 自动创建执行目录: {_execution_dir}")
    
    return _execution_dir

def _get_log_file():
    """获取日志文件路径"""
    if not LOG_ENABLED:
        return None
    
    execution_dir = _get_execution_dir()
    return os.path.join(execution_dir, "souren_execution.log")

def _get_result_file():
    """获取结果文件路径"""
    execution_dir = _get_execution_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(execution_dir, f"souren_results_{timestamp}.json")

EXECUTION_DIR = _get_execution_dir() if LOG_ENABLED else None
LOG_FILE = _get_log_file()
RESULT_FILE = _get_result_file()


def get_visa_address(ip_address=None):
    """根据IP地址生成VISA地址"""
    if ip_address is None:
        ip_address = DEFAULT_IP
    return f"TCPIP0::{ip_address}::inst0::INSTR"

INSTRUMENT_ADDRESS = get_visa_address()


def find_script_file(script_name, base_dir=None):
    """查找Python脚本文件"""
    import os
    
    if base_dir is None:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    possible_paths = [
        # 1. save_scv目录（相对于base_dir）
        os.path.join(base_dir, "save_scv", script_name),
        # 2. 直接在当前目录下
        os.path.join(base_dir, script_name),
        # 3. 在当前工作目录的save_scv目录下
        os.path.join(os.getcwd(), "save_scv", script_name),
        # 4. 在当前工作目录下
        os.path.join(os.getcwd(), script_name),
        # 5. 绝对路径
        script_name,
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return os.path.abspath(path)
    
    return None


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

    print(f"📄 Python脚本配置:")
    for i, script_config in enumerate(PYTHON_SCRIPT_NAME, 1):
        if isinstance(script_config, str):
            print(f"  {i}. {script_config} (无参数)")
        elif isinstance(script_config, dict):
            script_name = script_config.get('script', '未知脚本')
            params = {k: v for k, v in script_config.items() if k != 'script'}
            print(f"  {i}. {script_name}")
            for param, value in params.items():
                print(f"     - {param}: {value}")
    
    if DEVICE_TYPE.lower() == 'board':
        print(f"📡 USB AT端口: {USB_AT_PORT} (波特率: {USB_AT_BAUDRATE})")
        print(f"🔧 AT控制命令: {CELL_COMMANDS_NEED_AT}")
    elif DEVICE_TYPE.lower() == 'phone':
        print(f"📱 手机模式: ADB飞行模式控制")
        print(f"⏰ 等待时间: {ADB_WAIT_TIME}秒")
        print(f"🔧 ADB控制命令: {CELL_COMMANDS_NEED_ADB}")