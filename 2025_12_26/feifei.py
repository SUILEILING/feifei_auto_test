import time
import subprocess
import os
import pyautogui
import pygetwindow as gw
from typing import Dict, Tuple, Optional
import sys

class SourenCompleteAutomator:
    """Souren.ToolSet 完整版自动化 - 包含所有步骤"""
    
    def __init__(self):
        self.toolset_path = r"D:\SourenToolset\Souren.ToolSet.exe"
        self.scv_path = r"D:\code\night_cycle.scv"
        self.config_file = "souren_complete_coords.py"
        self.coordinates: Dict[str, Tuple[int, int]] = {}
        self.window_offset: Tuple[int, int] = (0, 0)
        self.process = None
        self.window = None
        
        # 需要获取的坐标点（按实际执行顺序）
        self.calibration_steps = [
            ('click_toolset_menu', '点击【工具集】菜单'),
            ('hover_command_menu', '鼠标悬停在【命令】菜单上（不要点击）'),
            ('click_commander_visa', '点击【Commander.Visa】选项'),
            ('click_open_file_menu', '点击【打开文件】按钮或菜单'),
            ('click_open_scv_option', '点击【打开SCV】选项'),
            ('click_open_dialog_button', '在打开对话框中点击【打开】按钮'),
            ('click_device_area', '点击【设备选择】区域'),
            ('click_device_list', '点击【设备列表】'),
            ('click_run_all', '点击【全部运行】按钮'),
        ]
        
        # 执行步骤（获取坐标后使用）
        self.execution_sequence = [
            ('click_toolset_menu', '点击工具集菜单', True),
            ('hover_command_menu', '悬停命令菜单', False),
            ('click_commander_visa', '点击Commander.Visa', True),
            ('click_open_file_menu', '点击打开文件', True),
            ('click_open_scv_option', '点击打开SCV', True),
            ('input_scv_path', '输入SCV文件路径', False),  # 特殊步骤：输入路径
            ('click_open_dialog_button', '点击打开按钮', True),
            ('click_device_area', '点击设备区域', True),
            ('click_device_list', '点击设备列表', True),
            ('click_run_all', '点击全部运行', True),
        ]
    
    def start_and_get_window(self) -> bool:
        """启动程序并获取窗口"""
        print("=" * 70)
        print("🚀 启动 Souren.ToolSet")
        print("=" * 70)
        
        try:
            if not os.path.exists(self.toolset_path):
                print(f"❌ 找不到程序: {self.toolset_path}")
                return False
            
            # 关闭已运行的实例
            self._close_running_instances()
            
            # 启动新实例
            print(f"📁 启动程序...")
            self.process = subprocess.Popen([self.toolset_path])
            
            # 等待窗口出现
            print("⏳ 等待程序启动...")
            self.window = self._wait_for_window("Souren.ToolSet", 30)
            
            if not self.window:
                print("❌ 启动超时，无法找到窗口")
                return False
            
            # 激活窗口
            if self.window.isMinimized:
                self.window.restore()
            self.window.activate()
            time.sleep(3)  # 等待界面完全加载
            
            # 记录窗口位置
            self.window_offset = (self.window.left, self.window.top)
            print(f"✅ 窗口位置: ({self.window.left}, {self.window.top})")
            print(f"✅ 窗口大小: {self.window.width}x{self.window.height}")
            
            return True
            
        except Exception as e:
            print(f"❌ 启动失败: {e}")
            return False
    
    def _wait_for_window(self, title: str, timeout: int = 30) -> Optional[gw.Window]:
        """等待指定标题的窗口出现"""
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                windows = gw.getWindowsWithTitle(title)
                if windows:
                    return windows[0]
            except:
                pass
            
            time.sleep(1)
            
        return None
    
    def _close_running_instances(self):
        """关闭正在运行的实例"""
        try:
            windows = gw.getWindowsWithTitle('Souren.ToolSet')
            for window in windows:
                if window.visible:
                    print("🔒 关闭已存在的实例...")
                    window.close()
                    time.sleep(3)
        except:
            pass
    
    def calibrate_all_positions(self):
        """校准所有位置"""
        print("=" * 70)
        print("🎯 位置校准向导")
        print("=" * 70)
        print("请按照以下顺序校准9个关键位置：")
        
        # 显示校准步骤
        for i, (key, desc) in enumerate(self.calibration_steps, 1):
            print(f"  {i}. {desc}")
        
        print("\n📋 步骤说明：")
        print("  1. 点击工具集菜单 - 展开工具集菜单")
        print("  2. 悬停命令菜单 - 鼠标移到命令菜单上（不要点击）")
        print("  3. 点击Commander.Visa - 选择Commander.Visa工具")
        print("  4. 点击打开文件 - 打开文件菜单")
        print("  5. 点击打开SCV - 选择打开SCV文件选项")
        print("  6. 点击打开按钮 - 在文件对话框中点击打开按钮")
        print("  7. 点击设备区域 - 点击设备选择区域")
        print("  8. 点击设备列表 - 选择设备")
        print("  9. 点击全部运行 - 开始运行测试")
        
        # 启动程序
        if not self.start_and_get_window():
            return
        
        # 开始校准
        print("\n" + "=" * 70)
        print("开始校准...")
        
        for step_key, step_desc in self.calibration_steps:
            self._calibrate_position(step_key, step_desc)
        
        # 保存配置
        self._save_configuration()
        
        # 询问是否测试
        print("\n" + "=" * 70)
        test_now = input("✅ 校准完成！是否立即测试执行？(y/n): ").lower()
        if test_now == 'y':
            self.execute_full_sequence()
    
    def _calibrate_position(self, step_key: str, step_desc: str):
        """校准单个位置"""
        print(f"\n{'='*60}")
        print(f"📍 {step_desc}")
        print(f"{'='*60}")
        
        # 显示操作提示
        if '悬停' in step_desc:
            print("💡 操作：将鼠标移动到目标位置即可，不要点击")
        else:
            print("💡 操作：将鼠标移动到目标位置准备点击")
        
        attempts = 0
        max_attempts = 3
        
        while attempts < max_attempts:
            print(f"\n📝 尝试 {attempts + 1}/{max_attempts}")
            print("请切换到Souren.ToolSet窗口，将鼠标移动到目标位置")
            
            # 显示倒计时
            print("⏱️  3秒后获取坐标...")
            for i in range(3, 0, -1):
                print(f"  {i}...")
                time.sleep(1)
            
            # 获取坐标
            abs_x, abs_y = pyautogui.position()
            
            # 计算相对坐标
            if self.window:
                rel_x = abs_x - self.window.left
                rel_y = abs_y - self.window.top
                
                print(f"\n📊 坐标信息:")
                print(f"  屏幕坐标: ({abs_x}, {abs_y})")
                print(f"  相对坐标: ({rel_x}, {rel_y})")
                print(f"  窗口位置: ({self.window.left}, {self.window.top})")
                
                # 确认
                if '悬停' in step_desc:
                    print(f"\n📌 这将是一个悬停位置（不点击）")
                    confirm = input("记录这个悬停坐标？(y-是/n-重新获取/s-跳过): ").lower()
                else:
                    print(f"\n📌 这将是一个点击位置")
                    confirm = input("记录这个点击坐标？(y-是/n-重新获取/s-跳过): ").lower()
                
                if confirm == 'y':
                    self.coordinates[step_key] = (rel_x, rel_y)
                    
                    # 显示记录结果
                    if '悬停' in step_desc:
                        print(f"✅ 已记录悬停位置: {step_desc}")
                    else:
                        print(f"✅ 已记录点击位置: {step_desc}")
                    print(f"   相对坐标: ({rel_x}, {rel_y})")
                    return
                elif confirm == 'n':
                    attempts += 1
                    continue
                elif confirm == 's':
                    print(f"⏭ 跳过此位置")
                    return
            else:
                print("❌ 无法获取窗口信息")
                attempts += 1
        
        print(f"⚠️ 超过最大尝试次数，跳过此位置")
    
    def _save_configuration(self):
        """保存配置"""
        if not self.coordinates:
            print("❌ 没有校准数据")
            return
        
        print("\n📊 校准结果汇总:")
        print("=" * 70)
        
        desc_map = {
            'click_toolset_menu': '点击工具集菜单',
            'hover_command_menu': '悬停命令菜单',
            'click_commander_visa': '点击Commander.Visa',
            'click_open_file_menu': '点击打开文件',
            'click_open_scv_option': '点击打开SCV',
            'click_open_dialog_button': '点击打开按钮',
            'click_device_area': '点击设备区域',
            'click_device_list': '点击设备列表',
            'click_run_all': '点击全部运行',
        }
        
        for key, (rel_x, rel_y) in self.coordinates.items():
            desc = desc_map.get(key, key)
            if '悬停' in desc:
                print(f"  🖱️  {desc}: 悬停坐标({rel_x}, {rel_y})")
            else:
                print(f"  👆 {desc}: 点击坐标({rel_x}, {rel_y})")
        
        # 生成配置文件
        config_content = '''# Souren.ToolSet 完整版自动化配置（相对坐标）
# 生成时间: {}
# 这些坐标是相对于窗口左上角的

relative_coordinates = {{
'''.format(time.strftime('%Y-%m-%d %H:%M:%S'))
        
        for key, (rel_x, rel_y) in self.coordinates.items():
            config_content += f"    '{key}': ({rel_x}, {rel_y}),\n"
        
        config_content += '''}

# 执行顺序说明:
# 1. click_toolset_menu: 点击工具集菜单
# 2. hover_command_menu: 鼠标悬停在命令菜单上（不点击）
# 3. click_commander_visa: 点击Commander.Visa选项
# 4. click_open_file_menu: 点击打开文件按钮/菜单
# 5. click_open_scv_option: 点击打开SCV选项
# 6. input_scv_path: 输入SCV文件路径（自动输入，无需坐标）
# 7. click_open_dialog_button: 在打开对话框中点击打开按钮
# 8. click_device_area: 点击设备选择区域
# 9. click_device_list: 点击设备列表
# 10. click_run_all: 点击全部运行按钮
'''
        
        with open(self.config_file, "w", encoding="utf-8") as f:
            f.write(config_content)
        
        print(f"\n✅ 配置文件已保存: {self.config_file}")
        print(f"📁 路径: {os.path.abspath(self.config_file)}")
    
    def load_configuration(self) -> bool:
        """加载配置"""
        try:
            if not os.path.exists(self.config_file):
                print(f"❌ 找不到配置文件: {self.config_file}")
                return False
            
            import importlib.util
            spec = importlib.util.spec_from_file_location("complete_coords", self.config_file)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            
            if hasattr(module, 'relative_coordinates'):
                self.coordinates = module.relative_coordinates
                print(f"✅ 加载配置成功: {len(self.coordinates)} 个位置")
                
                # 显示加载的配置
                print("\n📋 加载的位置配置:")
                desc_map = {
                    'click_toolset_menu': '点击工具集菜单',
                    'hover_command_menu': '悬停命令菜单',
                    'click_commander_visa': '点击Commander.Visa',
                    'click_open_file_menu': '点击打开文件',
                    'click_open_scv_option': '点击打开SCV',
                    'click_open_dialog_button': '点击打开按钮',
                    'click_device_area': '点击设备区域',
                    'click_device_list': '点击设备列表',
                    'click_run_all': '点击全部运行',
                }
                
                for key, (rel_x, rel_y) in self.coordinates.items():
                    desc = desc_map.get(key, key)
                    action = "🖱️ 悬停" if 'hover' in key else "👆 点击"
                    print(f"  {action} {desc}: ({rel_x}, {rel_y})")
                
                return True
            else:
                print("❌ 配置文件中没有找到 relative_coordinates")
                return False
                
        except Exception as e:
            print(f"❌ 加载配置失败: {e}")
            return False
    
    def get_or_start_window(self) -> bool:
        """获取或启动窗口"""
        print("\n🔍 检查Souren.ToolSet窗口...")
        
        # 先尝试获取现有窗口
        try:
            windows = gw.getWindowsWithTitle('Souren.ToolSet')
            if windows:
                self.window = windows[0]
                if self.window.isMinimized:
                    self.window.restore()
                self.window.activate()
                time.sleep(2)
                
                self.window_offset = (self.window.left, self.window.top)
                print(f"✅ 找到现有窗口: ({self.window.left}, {self.window.top})")
                return True
        except:
            pass
        
        # 如果没找到，启动新实例
        print("📱 未找到运行中的窗口，启动新实例...")
        return self.start_and_get_window()
    
    def execute_full_sequence(self):
        """执行完整序列"""
        print("=" * 70)
        print("🤖 开始执行完整自动化序列")
        print("=" * 70)
        
        # 1. 检查SCV文件
        if not os.path.exists(self.scv_path):
            print(f"❌ 找不到SCV文件: {self.scv_path}")
            custom_path = input("请输入SCV文件路径: ").strip()
            if custom_path and os.path.exists(custom_path):
                self.scv_path = custom_path
            else:
                print("❌ 文件不存在，退出")
                return
        
        print(f"📄 SCV文件: {self.scv_path}")
        
        # 2. 加载配置
        if not self.load_configuration():
            print("❌ 请先运行位置校准")
            return
        
        # 3. 检查必需位置
        required_keys = ['click_toolset_menu', 'hover_command_menu', 'click_commander_visa',
                        'click_open_file_menu', 'click_open_scv_option', 'click_open_dialog_button',
                        'click_device_area', 'click_device_list', 'click_run_all']
        
        missing_keys = [key for key in required_keys if key not in self.coordinates]
        if missing_keys:
            print(f"❌ 缺少必需位置: {missing_keys}")
            print("请运行位置校准获取所有位置")
            return
        
        # 4. 获取窗口
        if not self.get_or_start_window():
            print("❌ 无法获取窗口")
            return
        
        # 5. 显示执行计划
        print("\n📋 执行计划:")
        print("=" * 60)
        
        step_descriptions = [
            "1. 👆 点击工具集菜单",
            "2. 🖱️  悬停命令菜单（不点击）",
            "3. 👆 点击Commander.Visa",
            "4. 👆 点击打开文件",
            "5. 👆 点击打开SCV",
            "6. 📝 输入SCV文件路径",
            "7. 👆 点击打开按钮",
            "8. 👆 点击设备区域",
            "9. 👆 点击设备列表",
            "10. 👆 点击全部运行"
        ]
        
        for desc in step_descriptions:
            print(f"  {desc}")
        
        # 显示坐标信息
        print("\n📍 坐标信息:")
        for key in required_keys:
            if key in self.coordinates:
                rel_x, rel_y = self.coordinates[key]
                abs_x = self.window_offset[0] + rel_x
                abs_y = self.window_offset[1] + rel_y
                action = "悬停" if 'hover' in key else "点击"
                desc = {
                    'click_toolset_menu': '工具集菜单',
                    'hover_command_menu': '命令菜单',
                    'click_commander_visa': 'Commander.Visa',
                    'click_open_file_menu': '打开文件',
                    'click_open_scv_option': '打开SCV',
                    'click_open_dialog_button': '打开按钮',
                    'click_device_area': '设备区域',
                    'click_device_list': '设备列表',
                    'click_run_all': '全部运行',
                }.get(key, key)
                print(f"  {action} {desc}: ({abs_x}, {abs_y})")
        
        # 6. 确认执行
        print("\n" + "=" * 60)
        print("⚠️ 即将开始自动化执行")
        print("   执行过程中请不要操作鼠标键盘")
        
        confirm = input("是否开始执行？(y-开始/n-取消): ").lower()
        if confirm != 'y':
            print("取消执行")
            return
        
        # 7. 执行倒计时
        print("\n⏱️  5秒后开始执行...")
        for i in range(5, 0, -1):
            print(f"  {i}...")
            time.sleep(1)
        
        # 8. 执行所有步骤
        print("\n" + "=" * 70)
        print("🚀 开始执行...")
        print("=" * 70)
        
        step_count = len(self.execution_sequence)
        for idx, (step_key, step_desc, should_click) in enumerate(self.execution_sequence, 1):
            print(f"\n📊 进度: {idx}/{step_count}")
            success = self._execute_step(step_key, step_desc, should_click)
            if not success:
                print(f"❌ 步骤失败: {step_desc}")
                print("⚠️ 自动化执行中断")
                return
        
        # 9. 完成
        print("\n" + "=" * 70)
        print("✅ 自动化执行完成！")
        print("=" * 70)
        
        print(f"\n🎉 执行结果:")
        print(f"  • 所有 {step_count} 个步骤执行成功")
        print(f"  • SCV文件已加载: {os.path.basename(self.scv_path)}")
        print(f"  • 测试已启动，正在运行中...")
        
        # 等待一段时间让测试开始
        print("\n⏳ 等待测试启动...")
        time.sleep(5)
        
        print("\n📢 重要提示：")
        print("  • 测试在后台运行中，请勿关闭Souren.ToolSet窗口")
        print("  • 等待测试完成（时间根据测试内容而定）")
        print("  • 测试完成后请手动查看结果")
    
    def _execute_step(self, step_key: str, step_desc: str, should_click: bool) -> bool:
        """执行单个步骤"""
        print(f"\n▶ {step_desc}")
        
        try:
            # 特殊步骤：输入SCV路径
            if step_key == 'input_scv_path':
                print(f"  📝 输入SCV文件路径")
                print(f"    文件: {self.scv_path}")
                
                # 等待文件对话框完全打开
                time.sleep(2.5)
                
                # 确保焦点在文件路径输入框
                # 先按Alt+N快捷键聚焦到文件名输入框（Windows文件对话框通用）
                pyautogui.hotkey('alt', 'n')
                time.sleep(0.5)
                
                # 清空可能存在的文本
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.3)
                pyautogui.press('delete')
                time.sleep(0.3)
                
                # 输入路径
                pyautogui.write(self.scv_path)
                time.sleep(1.5)  # 等待输入完成
                return True
            
            # 普通步骤：使用坐标
            if step_key in self.coordinates:
                # 计算绝对坐标
                rel_x, rel_y = self.coordinates[step_key]
                abs_x = self.window_offset[0] + rel_x
                abs_y = self.window_offset[1] + rel_y
                
                # 执行操作
                if should_click:
                    print(f"  👆 点击: ({abs_x}, {abs_y})")
                    pyautogui.click(abs_x, abs_y)
                    
                    # 根据步骤类型设置不同的等待时间
                    if step_key == 'click_toolset_menu':
                        time.sleep(1.2)  # 等待菜单展开
                    elif step_key == 'hover_command_menu':
                        time.sleep(0.8)   # 悬停后等待
                    elif step_key == 'click_commander_visa':
                        time.sleep(2.0)   # 等待界面切换
                    elif step_key == 'click_open_file_menu':
                        time.sleep(1.5)   # 等待菜单展开
                    elif step_key == 'click_open_scv_option':
                        time.sleep(2.0)   # 等待文件对话框打开
                    elif step_key == 'click_open_dialog_button':
                        time.sleep(2.5)   # 等待文件加载
                    elif step_key == 'click_device_area':
                        time.sleep(1.5)   # 等待设备区域响应
                    elif step_key == 'click_device_list':
                        time.sleep(1.5)   # 等待设备选择
                    elif step_key == 'click_run_all':
                        time.sleep(3.0)   # 等待测试开始
                    else:
                        time.sleep(1.5)
                        
                else:
                    print(f"  🖱️  悬停: ({abs_x}, {abs_y})")
                    pyautogui.moveTo(abs_x, abs_y, duration=0.5)
                    time.sleep(1.0)  # 悬停时间
                
                return True
            else:
                print(f"  ⚠ 未找到位置配置: {step_key}")
                return False
                
        except Exception as e:
            print(f"  ❌ 执行失败: {e}")
            return False
    
    def quick_run(self):
        """快速运行（假设程序已打开）"""
        print("=" * 70)
        print("⚡ 快速运行模式")
        print("=" * 70)
        
        print("⚠️ 确保Souren.ToolSet已打开并在最前面")
        
        # 加载配置
        if not self.load_configuration():
            return
        
        # 获取窗口
        try:
            windows = gw.getWindowsWithTitle('Souren.ToolSet')
            if windows:
                self.window = windows[0]
                self.window.activate()
                self.window_offset = (self.window.left, self.window.top)
                print(f"✅ 找到窗口: ({self.window.left}, {self.window.top})")
            else:
                print("❌ 找不到窗口")
                return
        except:
            print("❌ 无法获取窗口")
            return
        
        # 立即执行
        print("\n🚀 开始快速执行...")
        
        step_count = len(self.execution_sequence)
        for idx, (step_key, step_desc, should_click) in enumerate(self.execution_sequence, 1):
            print(f"\n📊 进度: {idx}/{step_count}")
            self._execute_step(step_key, step_desc, should_click)
        
        print("\n✅ 快速执行完成！")

def main():
    """主函数"""
    print("=" * 80)
    print("🎯 Souren.ToolSet 完整版自动化系统")
    print("=" * 80)
    print("完整执行顺序：")
    print("  1. 👆 点击工具集菜单")
    print("  2. 🖱️  悬停命令菜单（不点击）")
    print("  3. 👆 点击Commander.Visa")
    print("  4. 👆 点击打开文件")
    print("  5. 👆 点击打开SCV")
    print("  6. 📝 输入SCV文件路径")
    print("  7. 👆 点击打开按钮")
    print("  8. 👆 点击设备区域")
    print("  9. 👆 点击设备列表")
    print("  10. 👆 点击全部运行")
    print("=" * 80)
    
    # 检查依赖
    try:
        import pyautogui
        import pygetwindow
    except ImportError:
        print("\n❌ 缺少必要的库")
        print("请运行: pip install pyautogui pygetwindow")
        input("按 Enter 退出...")
        return
    
    # 创建执行器
    automator = SourenCompleteAutomator()
    
    while True:
        print("\n" + "=" * 70)
        print("📋 主菜单")
        print("=" * 70)
        print("\n请选择操作:")
        print("  1. 🎯 校准所有位置（首次使用）")
        print("  2. 🤖 执行完整自动化")
        print("  3. ⚡ 快速运行（程序已打开时）")
        print("  4. 📋 查看当前配置")
        print("  5. 🚪 退出")
        
        choice = input("\n请选择 (1-5): ").strip()
        
        if choice == '1':
            automator.calibrate_all_positions()
            
        elif choice == '2':
            automator.execute_full_sequence()
            
        elif choice == '3':
            automator.quick_run()
            
        elif choice == '4':
            config_file = "souren_complete_coords.py"
            if os.path.exists(config_file):
                print(f"\n📄 配置文件: {config_file}")
                print("-" * 60)
                with open(config_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                    print(content)
                print("-" * 60)
            else:
                print(f"❌ 找不到配置文件")
            
        elif choice == '5':
            print("\n👋 退出程序")
            break
            
        else:
            print("❌ 无效选择")
        
        # 继续或退出
        cont = input("\n按 Enter 返回菜单，或输入 'q' 退出: ").lower()
        if cont == 'q':
            print("👋 退出程序")
            break

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n🛑 用户中断")
    except Exception as e:
        print(f"\n❌ 程序错误: {e}")
        import traceback
        traceback.print_exc()
    
    input("\n按 Enter 键退出...")
    