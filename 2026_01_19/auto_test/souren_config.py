from lib.var import *

# ==============================================
#  SCV 文件配置选项 
# ==============================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCV_FOLDER = os.path.join(BASE_DIR, "save_scv")
DEFAULT_SCV_NAME = "night_cycle_update_feifei.scv"  # 默认SCV文件 ⭐⭐⭐

# SCV生成选项
# 如果为 True 则使用python脚本动态生成SCV
# 如果为 False 则使用现有的SCV文件
USE_DYNAMIC_SCV = False  # 设置为 True 使用py生成SCV ⭐⭐⭐

# 动态SCV生成配置
SCV_CONVERTER_MODULE = "scv_converter"  # SCV转换器模块名 
PYTHON_SCRIPT_NAME = "night_cycle_update_feifei_steps.py"  # Python脚本名称 ⭐⭐⭐
DYNAMIC_SCV_NAME = "dynamic_test_steps.scv"  # 动态生成的SCV文件名
SCV_OUTPUT_FORMAT = "standard"  

# ==============================================
#  
# ==============执行模式配置================================
EXECUTION_MODE = 'loop_info'  # loop_info or run_all  ⭐⭐⭐
LOOP_COUNT = 2  # 只有loop_info模式才能使用

# ==============================================
# ⭐⭐⭐ 网络配置 ⭐⭐⭐
# ==============================================
DEFAULT_IP = '192.168.30.122' 

# ==============================================
# ⭐⭐⭐ 仪器地址配置 ⭐⭐⭐
# ==============================================
def get_visa_address(ip_address=None):
    """根据IP地址生成VISA地址"""
    if ip_address is None:
        ip_address = DEFAULT_IP
    return f"TCPIP0::{ip_address}::inst0::INSTR"

INSTRUMENT_ADDRESS = get_visa_address()

# ==============================================
# ⭐⭐⭐ 设备类型配置 ⭐⭐⭐
# ==============================================
DEVICE_TYPE = 'phone'  # 可选值: 'board' 或 'phone'

# =========================（仅 board 模式使用）=====================
# 当执行这些命令时，会先发送 AT+CFUN=0 → 等待2秒 → AT+CFUN=1 序列
CELL_COMMANDS_NEED_AT = [
    "CALL:CELL1 ON",
]
# USB AT 串口配置
USB_AT_PORT = "COM14"
USB_AT_BAUDRATE = 115200
USB_AT_TIMEOUT = 3  # 串口超时时间（秒）
# AT 指令序列配置 
AT_SEQUENCE = [
    ("AT+CFUN=0", 5),   # 发送 AT+CFUN=0，等待 5秒
    ("AT+CFUN=1", 1)    # 发送 AT+CFUN=1，等待 1秒
]
# ==============================================

# ============================（仅 phone 模式使用）==================
# 只在 CALL:CELL1 ON 命令后执行ADB飞行模式控制
CELL_COMMANDS_NEED_ADB = [
    "CALL:CELL1 ON",
]
ADB_WAIT_TIME = 3  # 开启飞行模式后的等待时间（秒）后关闭飞行模式
# ==============================================


# ==============================================
# ⭐⭐⭐ Excel导出配置 ⭐⭐⭐
# ==============================================

# ⭐⭐⭐ 图表步骤筛选配置 ⭐⭐⭐
CHART_STEPS_FILTER = [14, 15, 16, 20]  # 只显示这些步骤的数据


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

# 三维折线图 列宽配置
EXCEL_CHART_COLUMN_WIDTHS = {
    "A": 80,
    "default": 15,
}


# 图表颜色生成配置 
def generate_chart_colors(count: int):
    """
    动态生成图表颜色
    """
    import colorsys
    
    colors = []
    
    # 基础颜色集合（较柔和的颜色）
    base_colors = [
        "9933FF",  # 紫罗兰
        "2E75B6",  # 蓝色
        "70AD47",  # 绿色
        "FFC000",  # 橙色
        "5B9BD5",  # 浅蓝
        "ED7D31",  # 深橙
        "A5A5A5",  # 灰色
        "4472C4",  # 深蓝
        "C00000",  # 暗红（比FF0000柔和）
        "00B050",  # 亮绿
        "FF6600",  # 亮橙
        "0066CC",  # 亮蓝
        "FF3399",  # 粉红
        "33CCCC",  # 青色
    ]
    
    # 如果需要的颜色数量少于基础颜色数量，直接使用基础颜色
    if count <= len(base_colors):
        return base_colors[:count]
    
    # 如果需要更多颜色，动态生成
    for i in range(count):
        # 使用HSL色彩空间生成颜色
        hue = (i * 137.508) % 360  # 黄金角度，确保颜色分布均匀
        
        # 调整饱和度和亮度，生成柔和颜色
        saturation = 60 + (i % 20)  # 60-80% 饱和度
        lightness = 50 + ((i // 10) % 10)  # 50-60% 亮度
        
        # 将HSL转换为RGB
        r, g, b = colorsys.hls_to_rgb(hue / 360, lightness / 100, saturation / 100)
        
        # 将RGB转换为十六进制
        hex_color = f"{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"
        colors.append(hex_color)
    
    return colors

# 保留原来的颜色定义用于其他用途（如果其他地方需要）
EXCEL_CHART_COLORS = [
    "9933FF",  # 紫罗兰
    "2E75B6",  # 蓝色
    "70AD47",  # 绿色
    "FFC000",  # 橙色
    "5B9BD5",  # 浅蓝
    "ED7D31",  # 深橙
    "A5A5A5",  # 灰色
]

# 边框颜色
EXCEL_BORDER_COLORS = [
    "000000",
    "FF0000",
    "0000FF",
    "008000",
    "FFA500",
    "800080",
]

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

# ==============================================
#  ⭐⭐⭐ 修复：延迟创建日志目录 ⭐⭐⭐
# ==============================================
def _get_execution_dir():
    """获取执行目录，仅在需要时创建"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    execution_dir = os.path.join(BASE_DIR, "log", f"execution_{timestamp}")
    
    # 只在需要记录日志时创建目录
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

# ==============================================
#  SCV文件处理函数 
# ==============================================
def get_scv_file_path():
    """
    获取SCV文件路径
    根据USE_DYNAMIC_SCV配置决定使用静态SCV还是动态生成的SCV
    """
    if USE_DYNAMIC_SCV:
        # 动态生成SCV文件
        print(f"🔄 动态SCV模式激活,将生成新的SCV文件")
        scv_file = generate_dynamic_scv()
        if scv_file:
            # 获取文件名（不带路径）
            scv_file_name = os.path.basename(scv_file)
            print(f"📄 生成的SCV文件: {scv_file_name}")
            print(f"📁 完整路径: {scv_file}")
            return scv_file
        else:
            print("⚠️  动态SCV生成失败,使用默认SCV文件")
            # 使用静态SCV文件作为备选
            scv_path = os.path.join(os.path.abspath(SCV_FOLDER), DEFAULT_SCV_NAME)
            return scv_path
    else:
        # 使用静态SCV文件
        print(f"📄 使用静态SCV文件: {DEFAULT_SCV_NAME}")
    scv_path = os.path.join(os.path.abspath(SCV_FOLDER), DEFAULT_SCV_NAME)
    return scv_path

def generate_dynamic_scv():
    """
    动态从Python脚本生成SCV文件
    """
    try:
        print("🔄 开始动态生成SCV文件...")
        
        python_script_path = os.path.join(os.path.abspath(SCV_FOLDER), PYTHON_SCRIPT_NAME)
        
        if not os.path.exists(python_script_path):
            print(f"❌ Python脚本不存在: {python_script_path}")
            return None
        
        print(f"📁 解析Python脚本: {python_script_path}")
        print(f"📄 输出格式: {SCV_OUTPUT_FORMAT}")
        
        # 导入SCV转换器（延迟导入，避免循环依赖）
        try:
            from scv_converter import generate_dynamic_scv_from_python
        except ImportError as e:
            print(f"❌ 无法导入SCV转换器: {e}")
            print(f"💡 请确保 scv_converter.py 在同一目录下")
            return None
        
        # 生成SCV文件 - 使用自动清理
        scv_file = generate_dynamic_scv_from_python(
            python_script_path,
            output_format=SCV_OUTPUT_FORMAT,
            output_filename=DYNAMIC_SCV_NAME,
            auto_cleanup=True  # 程序结束后自动删除
        )
        
        if scv_file:
            print(f"✅ SCV文件生成成功: {scv_file}")
            print(f"🗑️  注意: 此文件将在程序结束时自动删除")
            
            try:
                with open(scv_file, 'r', encoding='utf-8') as f:
                    first_line = f.readline().strip()
                    print(f"🔍 第一行内容: {first_line}")
                    
                    # 检查格式是否正确
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

def display_config_info():
    """显示配置信息（需要时手动调用）"""
    # 只有在日志启用时才显示这些信息
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
        print(f"🔄 动态SCV源: {PYTHON_SCRIPT_NAME}")
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


