import os
import sys
import argparse

def check_dependencies():
    try:
        import pyautogui
        import pygetwindow
        return True
    except ImportError:
        print("缺少必要的库")
        print("请运行: pip install pyautogui pygetwindow")
        return False

def check_config():
    config_file = "souren_config.py"
    
    if not os.path.exists(config_file):
        print(f"找不到配置文件: {config_file}")
        
        create = input("是否创建配置文件？(y/n): ").lower()
        if create == 'y':
            config_content = '''TOOLSET_PATH = r"D:\\SourenToolset\\Souren.ToolSet.exe"
SCV_FOLDER = r"D:\\feifei\\2025_12_26\\auto_test\\save_scv"
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

DEFAULT_EXECUTION_MODE = 'run_all'
DEFAULT_LOOP_COUNT = 1'''
            
            with open(config_file, 'w', encoding='utf-8') as f:
                f.write(config_content)
            
            print(f"配置文件已创建: {config_file}")
            print("请从以下位置复制固定配置项到配置文件中:")
            print("1. CALIBRATION_STEPS 列表")
            print("2. EXECUTION_SEQUENCE_RUN_ALL 列表")
            print("3. EXECUTION_SEQUENCE_LOOP_INFO 列表")
        else:
            print("请手动创建配置文件")
            return False
    
    return True

def show_menu(automator):
    while True:
        print("\n" + "="*50)
        print("Souren.ToolSet 自动化系统")
        print("="*50)
        
        current_scv = os.path.basename(automator.scv_path) if automator.scv_path else "未选择"
        print(f"\n当前执行模式: {'全部执行' if automator.execution_mode == 'run_all' else '循环信息'}")
        print(f"当前SCV文件: {current_scv}")
        
        scv_files = automator.list_scv_files()
        if scv_files:
            print(f"SCV文件夹内容 ({len(scv_files)}个文件):")
            for i, file_info in enumerate(scv_files, 1):
                size_kb = file_info['size'] / 1024
                current = "✓" if automator.scv_path == file_info['path'] else " "
                print(f"  {i}. {file_info['name']} ({size_kb:.1f} KB) {current}")
        else:
            print("SCV文件夹: 无可用文件")
        
        print("\n执行计划:")
        for i, (_, step_desc, _) in enumerate(automator.execution_sequence, 1):
            print(f"  {i}. {step_desc}")
        
        print("\n请选择操作:")
        print("  1. 校准位置")
        print("  2. 执行自动化")
        print("  3. 选择SCV文件")
        print("  4. 切换执行模式")
        print("  5. 查看配置")
        print("  6. 退出")
        
        choice = input("\n请选择 (1-6): ").strip()
        
        if choice == '1':
            automator.calibrate()
        elif choice == '2':
            automator.execute(skip_confirmation=True)
        elif choice == '3':
            automator.select_scv_file()
        elif choice == '4':
            automator = switch_execution_mode(automator)
        elif choice == '5':
            show_config_info(automator)
        elif choice == '6':
            print("\n退出程序")
            break
        else:
            print("无效选择")
        
        cont = input("\n按 Enter 继续，或输入 q 退出: ").lower()
        if cont == 'q':
            print("退出程序")
            break

def switch_execution_mode(automator):
    print("\n切换执行模式")
    print("  1. 全部执行模式 (run_all)")
    print("  2. 循环信息模式 (loop_info)")
    
    choice = input("请选择模式 (1/2): ").strip()
    
    if choice == '1':
        from souren_main import SourenAutomator
        automator = SourenAutomator(
            custom_toolset_path=automator.toolset_path,
            custom_scv_path=automator.scv_path,
            execution_mode='run_all'
        )
        print("已切换到全部执行模式")
        return automator
    elif choice == '2':
        from souren_main import SourenAutomator
        automator = SourenAutomator(
            custom_toolset_path=automator.toolset_path,
            custom_scv_path=automator.scv_path,
            execution_mode='loop_info'
        )
        print("已切换到循环信息模式")
        return automator
    else:
        print("无效选择，保持当前模式")
        return automator

def show_config_info(automator):
    print("\n配置信息")
    print("="*40)
    
    print("路径配置:")
    print(f"  程序路径: {automator.toolset_path}")
    print(f"  SCV文件夹: {automator.scv_folder}")
    if automator.scv_path:
        print(f"  当前SCV: {os.path.basename(automator.scv_path)}")
    print(f"  坐标文件: {automator.coordinates_file}")
    
    print(f"\n执行模式: {'全部执行' if automator.execution_mode == 'run_all' else '循环信息'}")
    if automator.execution_mode == 'loop_info':
        print(f"循环次数: {automator.loop_count}")
    
    print("\n文件状态:")
    check_file = lambda path, label: f"  ✅ {label}" if os.path.exists(path) else f"  ❌ {label}"
    
    print(check_file(automator.toolset_path, "程序"))
    print(check_file(automator.scv_folder, "SCV文件夹"))
    if automator.scv_path:
        print(check_file(automator.scv_path, "当前SCV文件"))
    print(check_file(automator.coordinates_file, "坐标文件"))
    
    scv_files = automator.list_scv_files()
    if scv_files:
        print(f"\nSCV文件列表 ({len(scv_files)}个):")
        for i, file_info in enumerate(scv_files, 1):
            size_kb = file_info['size'] / 1024
            current = "✓" if automator.scv_path == file_info['path'] else " "
            print(f"  {i}. {file_info['name']} ({size_kb:.1f} KB) {current}")
    else:
        print("\nSCV文件夹: 无可用文件")

def main():
    print("="*60)
    print("Souren.ToolSet 自动化系统")
    print("="*60)
    
    if not check_dependencies():
        input("按 Enter 退出...")
        return
    
    if not check_config():
        input("按 Enter 退出...")
        return
    
    try:
        from souren_main import SourenAutomator
    except ImportError as e:
        print(f"导入主程序失败: {e}")
        input("按 Enter 退出...")
        return
    
    parser = argparse.ArgumentParser(description='Souren.ToolSet 自动化系统')
    parser.add_argument('--toolset', type=str, help='指定Souren.ToolSet.exe路径')
    parser.add_argument('--scv', type=str, help='指定SCV文件路径')
    parser.add_argument('--mode', type=str, choices=['run_all', 'loop_info'], 
                       help='执行模式: run_all 或 loop_info')
    parser.add_argument('--calibrate', action='store_true', help='运行校准模式')
    parser.add_argument('--run', action='store_true', help='运行自动化模式')
    
    args = parser.parse_args()
    
    automator = SourenAutomator(
        custom_toolset_path=args.toolset,
        custom_scv_path=args.scv,
        execution_mode=args.mode
    )
    
    scv_files = automator.list_scv_files()
    if scv_files:
        print(f"\n📁 SCV文件夹: {automator.scv_folder}")
        print(f"发现 {len(scv_files)} 个SCV文件:")
        for file_info in scv_files:
            size_kb = file_info['size'] / 1024
            print(f"  • {file_info['name']} ({size_kb:.1f} KB)")
    
    if args.calibrate:
        automator.calibrate()
    elif args.run:
        automator.execute(skip_confirmation=True)
    else:
        show_menu(automator)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n用户中断")
    except Exception as e:
        print(f"\n程序错误: {e}")
    
    input("\n按 Enter 键退出...")