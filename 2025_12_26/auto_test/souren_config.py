TOOLSET_PATH = r"D:\SourenToolset\Souren.ToolSet.exe"
SCV_FOLDER = r"D:\feifei\2025_12_26\auto_test\save_scv"
COORDINATES_FILE = "souren_complete_coords.py"

WAIT_TIMES = {
    'startup': 10,
    'menu_expand': 1.5,
    'dialog_open': 2.0,
    'file_load': 3.0,
    'test_start': 5.0,
    'hover_time': 1.0,
    'loop_dialog': 2.0,
}

CALIBRATION_STEPS = [
    ('click_toolset_menu', '点击工具集菜单'),
    ('hover_command_menu', '悬停命令菜单'),
    ('click_commander_visa', '点击Commander.Visa'),
    ('click_open_file_menu', '点击打开文件'),
    ('click_open_scv_option', '点击打开SCV'),
    ('click_scv_file', '点击文件对话框中的SCV文件'),
    ('click_ok_button', '点击文件对话框的OK按钮'),
    ('click_device_area', '点击设备区域'),
    ('click_device_list', '点击设备列表'),
    ('click_run_all', '点击全部运行'),
    ('click_loop_info', '点击循环信息'),
    ('input_loop_count', '输入循环次数'),
    ('click_loop_ok', '点击循环对话框的OK按钮'),
]

EXECUTION_SEQUENCE_RUN_ALL = [
    ('click_toolset_menu', '点击工具集菜单', True),
    ('hover_command_menu', '悬停命令菜单', False),
    ('click_commander_visa', '点击Commander.Visa', True),
    ('click_open_file_menu', '点击打开文件', True),
    ('click_open_scv_option', '点击打开SCV', True),
    ('click_scv_file', '点击SCV文件', True),
    ('click_ok_button', '点击OK按钮', True),
    ('click_device_area', '点击设备区域', True),
    ('click_device_list', '点击设备列表', True),
    ('click_run_all', '点击全部运行', True),
]

EXECUTION_SEQUENCE_LOOP_INFO = [
    ('click_toolset_menu', '点击工具集菜单', True),
    ('hover_command_menu', '悬停命令菜单', False),
    ('click_commander_visa', '点击Commander.Visa', True),
    ('click_open_file_menu', '点击打开文件', True),
    ('click_open_scv_option', '点击打开SCV', True),
    ('click_scv_file', '点击SCV文件', True),
    ('click_ok_button', '点击OK按钮', True),
    ('click_device_area', '点击设备区域', True),
    ('click_device_list', '点击设备列表', True),
    ('click_loop_info', '点击循环信息', True),
    ('input_loop_count', '输入循环次数', False),
    ('click_loop_ok', '点击循环对话框的OK按钮', True),
]

DEFAULT_EXECUTION_MODE = 'run_all'
DEFAULT_LOOP_COUNT = 1