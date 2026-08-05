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

    ####----------------------------------------------------------------------phone ------------------------------------------------------------------------------------------------------
    ###-------- sa ---------
    # test fixed power
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "nr_band": 1, "nr_bw": 20, "scs": 15, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "nr_band": 5, "nr_bw": 20, "scs": 15, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "nr_band": 8, "nr_bw": 20, "scs": 15, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "nr_band": 28, "nr_bw": 20, "scs": 15, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "nr_band": 41, "nr_bw": 100, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "nr_band": 77, "nr_bw": 100, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "nr_band": 78, "nr_bw": 100, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},


    # ### power test ot
    # {"script": "power_test_ot", "lineLoss1": 25.00, "nr_band": 1, "nr_bw": 20, "scs": 15, "range": "LOW", "nr_start_power": -54, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "power_test_ot", "lineLoss1": 25.00, "nr_band": 5, "nr_bw": 20, "scs": 15, "range": "LOW", "nr_start_power": -54, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "power_test_ot", "lineLoss1": 25.00, "nr_band": 8, "nr_bw": 20, "scs": 15, "range": "LOW", "nr_start_power": -54, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "power_test_ot", "lineLoss1": 25.00, "nr_band": 28,"nr_bw": 20, "scs": 15, "range": "LOW", "nr_start_power": -54, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},

    # {"script": "power_test_ot", "lineLoss1": 25.00, "nr_band": 41, "nr_bw": 100, "scs": 30, "range": "LOW", "nr_start_power": -50, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "power_test_ot", "lineLoss1": 25.00, "nr_band": 77, "nr_bw": 100, "scs": 30, "range": "LOW", "nr_start_power": -50, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "power_test_ot", "lineLoss1": 25.00, "nr_band": 78, "nr_bw": 100, "scs": 30, "range": "LOW", "nr_start_power": -50, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},


    # ## test sa channel switch
    # {"script": "test_sa_channel_switch", "lineLoss1": 25.00, "nr_band_list": [78,77,79,41,1,5,8,28], "range_list": ["Low","Mid","High","Low"], "rb_mode": "Outer_Full", "case_dir": "yc1100"},


    # ##  nr target pusch power 
    # {"script": "test_nr_target_pusch_power", "lineLoss1": 25.00, "nr_band": 78, "nr_bw": 100, "scs": 30, "range": "LOW", "min_power": -20, "max_power": None, "step": 2, "rb_mode": "Inner_Full", "case_dir": "yc1100"},


    # ##--------- nsa ---------
    # test fixed power
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 41, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 77, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 78, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 79, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},

    # ### check lte type2
    # {"script": "test_fixed_power", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 78, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60,"resource_allocation_type":2, "rb_mode": "Inner_Full", "case_dir": "yc1100"},

    # ## power test_ot
    #  {"script": "power_test_ot", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 41, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "nr_start_power": -50, "nsa_start_power": -54,"end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    #  {"script": "power_test_ot", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 77, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "nr_start_power": -50, "nsa_start_power": -54,"end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    #  {"script": "power_test_ot", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 78, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "nr_start_power": -50, "nsa_start_power": -54,"end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    #  {"script": "power_test_ot", "lineLoss1": 25.00, "lineLoss3": 25.00, "nr_band": 79, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "nr_start_power": -50, "nsa_start_power": -54,"end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},


    # ## test nsa channel switch
    # {"script": "test_nsa_channel_switch", "lineLoss1": 25.00, "lineLoss3": 25.00, "lte_band": 3, "lte_bw": 20, "nr_band_list": [78,77,79,41], "lte_band_list": [{78:1,20:[300]},{77:1,20:[300]}], "nr_range_list": ["Low","Mid","High","Low"], "rb_mode": "Outer_Full", "case_dir": "yc1100"},
 

    ### ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



    # # ####------------------------------------------------------------------AT----------------------------------------------------------------------------------------------------------------
    ####------ sa --------
    ## test fixed power
    {"script": "test_fixed_power", "lineLoss1": 5.00, "nr_band": 1, "nr_bw": 20, "scs": 15, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "test_fixed_power", "lineLoss1": 5.00, "nr_band": 5, "nr_bw": 20, "scs": 15, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "test_fixed_power", "lineLoss1": 5.00, "nr_band": 8, "nr_bw": 20, "scs": 15, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "test_fixed_power", "lineLoss1": 5.00, "nr_band": 28, "nr_bw": 20, "scs": 15, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    
    {"script": "test_fixed_power", "lineLoss1": 5.00, "nr_band": 41, "nr_bw": 100, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "test_fixed_power", "lineLoss1": 5.00, "nr_band": 77, "nr_bw": 100, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "test_fixed_power", "lineLoss1": 5.00, "nr_band": 78, "nr_bw": 100, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},


    ### power test ot
    {"script": "power_test_ot", "lineLoss1": 5.00, "nr_band": 1, "nr_bw": 20, "scs": 15, "range": "LOW", "nr_start_power": -54, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "power_test_ot", "lineLoss1": 5.00, "nr_band": 5, "nr_bw": 20, "scs": 15, "range": "LOW", "nr_start_power": -54, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "power_test_ot", "lineLoss1": 5.00, "nr_band": 8, "nr_bw": 20, "scs": 15, "range": "LOW", "nr_start_power": -54, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "power_test_ot", "lineLoss1": 5.00, "nr_band": 28,"nr_bw": 20, "scs": 15, "range": "LOW", "nr_start_power": -54, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},

    {"script": "power_test_ot", "lineLoss1": 5.00, "nr_band": 41, "nr_bw": 100, "scs": 30, "range": "LOW", "nr_start_power": -50, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "power_test_ot", "lineLoss1": 5.00, "nr_band": 77, "nr_bw": 100, "scs": 30, "range": "LOW", "nr_start_power": -50, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    {"script": "power_test_ot", "lineLoss1": 5.00, "nr_band": 78, "nr_bw": 100, "scs": 30, "range": "LOW", "nr_start_power": -50, "end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},


    ### test sa channel switch 
    {"script": "test_sa_channel_switch", "lineLoss1": 5.00, "nr_band_list": [78,77,79,41,1,5,8,28], "range_list": ["Low","Mid","High","Low"], "rb_mode": "Outer_Full", "case_dir": "yc1100"},


    #  nr target pusch power 
    {"script": "test_nr_target_pusch_power", "lineLoss1": 5.00, "nr_band": 78, "nr_bw": 100, "scs": 30, "range": "LOW", "min_power": -20, "max_power": None, "step": 2, "rb_mode": "Inner_Full", "case_dir": "yc1100"},



    ###----- nsa -------
    ### test fixed power
    # {"script": "test_fixed_power", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 41, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 77, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 78, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    # {"script": "test_fixed_power", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 79, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60, "rb_mode": "Inner_Full", "case_dir": "yc1100"},

    ### check lte type2
    # {"script": "test_fixed_power", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 78, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "TD": 60,"resource_allocation_type":2, "rb_mode": "Inner_Full", "case_dir": "yc1100"},


    ### power test ot
    #  {"script": "power_test_ot", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 41, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "nr_start_power": -50, "nsa_start_power": -54,"end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    #  {"script": "power_test_ot", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 77, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "nr_start_power": -50, "nsa_start_power": -54,"end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    #  {"script": "power_test_ot", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 78, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "nr_start_power": -50, "nsa_start_power": -54,"end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},
    #  {"script": "power_test_ot", "lineLoss1": 5.00, "lineLoss3": 5.00, "nr_band": 79, "lte_band": 3, "nr_bw": 100, "lte_bw": 20, "scs": 30, "range": "LOW", "nr_start_power": -50, "nsa_start_power": -54,"end_power": -150, "step": -1, "fallback_delta": 10, "rb_mode": "Inner_Full", "case_dir": "yc1100"},


    ### test nsa chanel switch
    #  {"script": "test_nsa_channel_switch", "lineLoss1": 5.00, "lineLoss3": 5.00, "lte_band": 3, "lte_bw": 20, "nr_band_list": [78,77,79,41], "lte_band_list": [{78:1,20:[300]},{77:1,20:[300]}], "nr_range_list": ["Low","Mid","High","Low"], "rb_mode": "Outer_Full", "case_dir": "yc1100"},

    ### ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------






    ### ----------------------------------------------------------------------new case------------------------------------------------------------------------------------------------------------------------


    # {"script": "test_pvt", "lineLoss1": 1.00, "nr_band": 78, "nr_bw": 100, "scs": 30, "range": "LOW", "TD": 10, "case_dir": "yc1100"},


    ### ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------





]


DEFAULT_IP = '192.168.30.122'  
REMOTE_SERVER_PORT = 9999  # visitor:8000 yc:9999
REMOTE_SUDO_PASSWORD = 'yc'   


# ========== 基站软件版本 ==========
VERSION = "YC1100.1.00.03.10"

# ========== 核心网日志收集 ==========
COLLECT_CORE_NETWORK_LOGS = True

# ========== 基站调试日志级别 ==========
LOG_LEVEL_PARAMS = {
    "log_level": "info",        # default:info       disable/error/warning/analysis/info/debug/trace
    "global_log_level": 0,      # default:0
    "sctp_log_level": 0,        # default:0
    "hw_log_level": 0,          # default:0
    "phy_log_level": 0,         # default:0
    "mac_log_level": 0,         # default:0
    "rlc_log_level": 0,         # default:0
    "pdcp_log_level": 0,        # default:0
    "rrc_log_level": 0,         # default:0
    "ngap_log_level": 0,        # default:0
    "f1ap_log_level": 0,        # default:0
    "x2ap_log_level": 1,        # default:1
    "s1ap_log_level": 0,        # default:0
    "ULSCH_dump": 0,            # default:0
    "RRC_dump": 0,              # default:0
    "NAS_dump": 0,              # default:0
}

LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS = ["test_nsa_channel_switch"]


# ==============================================

EXECUTION_MODE = 'loop_info'    
LOOP_COUNT = 1  

EXCEL_DEFAULT_ROW_HEIGHT = 13.5
EXCEL_DEFAULT_COLUMN_WIDTH = 9
EXCEL_HEADER_ROW_HEIGHT = 20

# 执行汇总表 特殊列宽配置
EXCEL_SUMMARY_COLUMN_WIDTHS = {
    "B": 18,    # 执行时间
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

def _compute_default_execution_dir():
    global _execution_dir
    if _execution_dir:
        return _execution_dir
    if LOG_ENABLED:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        _execution_dir = os.path.join(BASE_DIR, "log", f"execution_{timestamp}")
    return _execution_dir

def _get_execution_dir():
    d = _compute_default_execution_dir()
    if d and LOG_ENABLED and not os.path.exists(d):
        os.makedirs(d, exist_ok=True)
    return d

def _get_log_file():
    if not LOG_ENABLED:
        return None
    return os.path.join(_get_execution_dir(), "souren_execution.log")

def _get_result_file():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(_get_execution_dir(), f"souren_results_{timestamp}.json")

EXECUTION_DIR = _compute_default_execution_dir() if LOG_ENABLED else None
LOG_FILE = os.path.join(EXECUTION_DIR, "souren_execution.log") if (LOG_ENABLED and EXECUTION_DIR) else None
RESULT_FILE = (os.path.join(EXECUTION_DIR, f"souren_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
               if EXECUTION_DIR else None)

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