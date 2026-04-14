from lib.var import *

CASE_CONFIG = {
    "yc1100": {
        "directory": "yc1100",
    },
    # "yc2100": {
    #     "directory": "yc2100", 
    # }
}

PYTHON_SCRIPT_NAME = [
    #------------------------------------------------phone ----------------------------------------------------------

    # {"script": "sa_test", "lineLoss": 25.00, "band": 1, "bw": 20, "scs": 15, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 25.00, "band": 1, "bw": 20, "scs": 15, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 25.00, "band": 1, "bw": 20, "scs": 15, "range": "High", "case_dir": "yc1100"},

    # {"script": "sa_test", "lineLoss": 25.00, "band": 41, "bw": 100, "scs": 30, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 25.00, "band": 41, "bw": 100, "scs": 30, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 25.00, "band": 41, "bw": 21000, "scs": 30, "range": "High", "case_dir": "yc1100"},

    # {"script": "sa_test", "lineLoss": 25.00, "band": 77, "bw": 100, "scs": 30, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 25.00, "band": 77, "bw": 100, "scs": 30, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 25.00, "band": 77, "bw": 100, "scs": 30, "range": "High", "case_dir": "yc1100"},

    # {"script": "sa_test", "lineLoss": 25.00, "band": 78, "bw": 100, "scs": 30, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 25.00, "band": 78, "bw": 100, "scs": 30, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 25.00, "band": 78, "bw": 100, "scs": 30, "range": "High", "case_dir": "yc1100"},

    # {"script": "sa_test", "lineLoss": 30.00, "band": 79, "bw": 100, "scs": 30, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 30.00, "band": 79, "bw": 100, "scs": 30, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 30.00, "band": 79, "bw": 100, "scs": 30, "range": "High", "case_dir": "yc1100"},

    #-----------------------------------------------------------------------------------------------------------------


    #------------------------------------------------at --------------------------------------------------------------
    
    # {"script": "sa_test", "lineLoss": 5.00, "band": 1, "bw": 20, "scs": 15, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 1, "bw": 20, "scs": 15, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 1, "bw": 20, "scs": 15, "range": "High", "case_dir": "yc1100"},

    # {"script": "sa_test", "lineLoss": 5.00, "band": 41, "bw": 100, "scs": 30, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 41, "bw": 100, "scs": 30, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 41, "bw": 21000, "scs": 30, "range": "High", "case_dir": "yc1100"},

    # {"script": "sa_test", "lineLoss": 5.00, "band": 77, "bw": 100, "scs": 30, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 77, "bw": 100, "scs": 30, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 77, "bw": 100, "scs": 30, "range": "High", "case_dir": "yc1100"},

    # {"script": "sa_test", "lineLoss": 5.00, "band": 78, "bw": 100, "scs": 30, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 78, "bw": 100, "scs": 30, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 78, "bw": 100, "scs": 30, "range": "High", "case_dir": "yc1100"},

    # {"script": "sa_test", "lineLoss": 5.00, "band": 79, "bw": 100, "scs": 30, "range": "LOW", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 79, "bw": 100, "scs": 30, "range": "Mid", "case_dir": "yc1100"},
    # {"script": "sa_test", "lineLoss": 5.00, "band": 79, "bw": 100, "scs": 30, "range": "High", "case_dir": "yc1100"},

    #-----------------------------------------------------------------------------------------------------------------


    # {"script": "sa_power_test", "lineLoss":7, "band": 79, "bw": 100, "scs": 30, "range": "LOW", "start_power": -40, "end_power": -95, "step": -2, "fallback_delta": 10, "case_dir": "yc1100"},

    {"script": "sa_power_test_ot", "lineLoss": 7, "band": 78, "bw": 100, "scs": 30, "range": "LOW", "start_power": -50, "end_power": -130, "step": -2, "fallback_delta": 10, "case_dir": "yc1100"},

]

EXECUTION_MODE = 'loop_info'    
LOOP_COUNT = 1  

DEFAULT_IP = '192.168.30.122'  



# ==============================================
EXCEL_DEFAULT_ROW_HEIGHT = 13.5
EXCEL_DEFAULT_COLUMN_WIDTH = 9
EXCEL_HEADER_ROW_HEIGHT = 20

# 执行汇总表 特殊列宽配置
EXCEL_SUMMARY_COLUMN_WIDTHS = {
    "B": 17,    # 执行时间
    "C": 18,    # SCV文件
    "D": 32,    # 设备
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
    "O": 12,    # 状态消息
}

# 详细执行记录表 特殊列宽配置
EXCEL_DETAILS_COLUMN_WIDTHS = {
    "B": 24,
    "F": 80,
    "I": 46,
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

LOG_ENABLED = True
LOG_LEVEL = 'INFO'

SHOW_STEP_TYPE = False
SHOW_STEP_CONTENT = True
SHOW_COMMAND_SENDING = False
SHOW_COMMAND_RESULT = False

_execution_dir = None
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_case_directory(case_dir: str) -> str:
    if case_dir in CASE_CONFIG:
        return os.path.join(BASE_DIR, CASE_CONFIG[case_dir]["directory"])
    return os.path.join(BASE_DIR, case_dir)

def set_execution_dir(dir_path):
    global _execution_dir
    _execution_dir = dir_path
    print(f"📁 设置执行目录为: {_execution_dir}")

def _get_execution_dir():
    global _execution_dir
    if _execution_dir:
        if LOG_ENABLED and not os.path.exists(_execution_dir):
            os.makedirs(_execution_dir, exist_ok=True)
        return _execution_dir
    if LOG_ENABLED:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        _execution_dir = os.path.join(BASE_DIR, "log", f"execution_{timestamp}")
        os.makedirs(_execution_dir, exist_ok=True)
    return _execution_dir

def _get_log_file():
    if not LOG_ENABLED:
        return None
    return os.path.join(_get_execution_dir(), "souren_execution.log")

def _get_result_file():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(_get_execution_dir(), f"souren_results_{timestamp}.json")

EXECUTION_DIR = _get_execution_dir() if LOG_ENABLED else None
LOG_FILE = _get_log_file()
RESULT_FILE = _get_result_file()

def get_visa_address(ip_address=None):
    return f"TCPIP0::{ip_address or DEFAULT_IP}::inst0::INSTR"

INSTRUMENT_ADDRESS = get_visa_address()

def display_config_info():
    if LOG_ENABLED and EXECUTION_DIR:
        print(f"📁 本次执行目录: {os.path.abspath(EXECUTION_DIR)}")
        print(f"📁 日志文件: {os.path.abspath(LOG_FILE)}")
        print(f"📁 结果文件: {os.path.abspath(RESULT_FILE)}")
    else:
        print("📁 日志功能已禁用")
    print(f"🔌 仪器地址: {INSTRUMENT_ADDRESS}")
    print("📄 Python脚本配置:")
    for i, sc in enumerate(PYTHON_SCRIPT_NAME, 1):
        if isinstance(sc, str):
            print(f"  {i}. {sc} (无参数)")
        else:
            name = sc.get('script', '未知')
            params = {k:v for k,v in sc.items() if k!='script'}
            print(f"  {i}. {name}")
            for p,v in params.items():
                print(f"     - {p}: {v}")
    print("📁 CASE目录配置:")
    for cn, cc in CASE_CONFIG.items():
        print(f"  {cn}: {cc['directory']}")