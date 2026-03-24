from lib.var import *

def debug_cell_command():
    """调试CELL命令"""
    print("\n" + "="*60)
    print("🔧 调试CELL命令...")
    print("="*60)
    
    try:
        from souren_core import InstrumentController
    except ImportError as e:
        print(f"❌ 导入模块失败: {e}")
        return
    
    controller = InstrumentController()
    
    print("🔌 连接仪器...")
    success, message = controller.connect()
    if not success:
        print(f"❌ 连接失败: {message}")
        return
    
    print("✅ 连接成功，开始测试命令...\n")
    

    test_commands = [
        "CELL1 ON",
        "CELL1 OFF", 
        "CELL1?",
        "CALL:CELL1 ON",
        "CALL:CELL1 OFF",
        "CALL:CELL1?",
    ]
    
    for i, cmd in enumerate(test_commands, 1):
        print(f"{'='*40}")
        print(f"测试 {i}/{len(test_commands)}: {cmd}")
        
        if "CALL:" in cmd.upper():
            success, result = controller.execute_call_command(cmd)
        else:
            success, result = controller.execute_scpi_command(cmd)
        
        status = "✅ 成功" if success else "❌ 失败"
        print(f"结果: {status}")
        if result:
            print(f"返回: {result}")
        
        if i < len(test_commands):
            time.sleep(1) 
    
    controller.disconnect()
    print(f"\n{'='*60}")
    print("🔧 调试完成")
    print("="*60)

def find_scv_file(scv_folder, scv_name):
    """查找SCV文件支持多种搜索方式"""
    print(f"🔍 查找SCV文件: {scv_name}")
    
    if os.path.isabs(scv_name) and os.path.exists(scv_name):
        print(f"✅ 找到绝对路径文件: {scv_name}")
        return scv_name
    
    if scv_folder:
        scv_path = os.path.join(scv_folder, scv_name)
        if os.path.exists(scv_path):
            print(f"✅ 找到SCV文件夹中的文件: {scv_path}")
            return scv_path
        
        try:
            files = [f for f in os.listdir(scv_folder) if f.endswith('.scv')]
            if files:
                first_scv = os.path.join(scv_folder, files[0])
                print(f"📄 未找到指定的 '{scv_name}', 使用第一个SCV文件: {files[0]}")
                return first_scv
        except Exception as e:
            print(f"⚠️ 无法列出SCV文件夹内容: {e}")
    
    if os.path.exists(scv_name):
        print(f"✅ 找到当前目录中的文件: {scv_name}")
        return scv_name
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_scv_path = os.path.join(script_dir, scv_name)
    if os.path.exists(script_scv_path):
        print(f"✅ 找到脚本同目录中的文件: {script_scv_path}")
        return script_scv_path
    
    return None

def main():
    parser = argparse.ArgumentParser(description='Souren.ToolSet 自动化系统')
    parser.add_argument('--list', '-l', action='store_true', help='列出SCV文件')
    parser.add_argument('--run', '-r', action='store_true', help='执行完整工作流程')
    parser.add_argument('--debug-cell', action='store_true', help='调试CELL命令')
    parser.add_argument('--dynamic-scv', action='store_true', help='使用动态SCV生成')
    parser.add_argument('--static-scv', action='store_true', help='使用静态SCV文件')
    parser.add_argument('--scv-format', choices=['standard', 'table'], default='standard', 
                       help='动态SCV输出格式 (standard/table)')
    parser.add_argument('--scv-file', type=str, help='指定SCV文件路径(覆盖默认配置）')
    
    args = parser.parse_args()
    
    if args.debug_cell:
        debug_cell_command()
        return 0
    
    current_dir = os.getcwd()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"📁 当前工作目录: {current_dir}")
    print(f"📁 Python执行目录: {script_dir}")
    
    try:
        from souren_config import (
            SCV_FOLDER, DEFAULT_SCV_NAME, display_config_info, 
            DEFAULT_IP, USE_DYNAMIC_SCV, SCV_OUTPUT_FORMAT,
            PYTHON_SCRIPT_NAME, get_scv_file_path, LOG_ENABLED
        )
        
        use_dynamic = USE_DYNAMIC_SCV
        scv_format = SCV_OUTPUT_FORMAT
        
        if args.dynamic_scv:
            print("🔄 使用动态SCV生成模式(命令行参数覆盖）")
            use_dynamic = True
            if args.scv_format:
                scv_format = args.scv_format
                print(f"📄 SCV格式: {scv_format}")
        
        if args.static_scv:
            print("📄 使用静态SCV文件模式(命令行参数覆盖)")
            use_dynamic = False
        
        display_config_info()
        print(f"✅ 配置文件导入成功")
        
        scv_file_path = None
        
        if args.scv_file:
            scv_file_path = args.scv_file
            if not os.path.exists(scv_file_path):
                print(f"❌ 指定的SCV文件不存在: {scv_file_path}")
                # 尝试查找文件
                print("🔍 尝试查找指定的SCV文件...")
                scv_file_path = find_scv_file(SCV_FOLDER, os.path.basename(scv_file_path))
                if not scv_file_path:
                    return 1
            print(f"📄 使用命令行指定的SCV文件: {os.path.basename(scv_file_path)}")
        else:
            if use_dynamic:
                print(f"🔄 动态SCV模式激活")
            else:
                print(f"📄 静态SCV模式激活")
            
            try:
                config_scv_path = get_scv_file_path()
                if config_scv_path and os.path.exists(config_scv_path):
                    scv_file_path = config_scv_path
                    print(f"✅ 从配置找到SCV文件: {config_scv_path}")
                else:
                    print(f"⚠️ 配置中的SCV文件不存在: {config_scv_path}")
                    
                    scv_file_path = find_scv_file(SCV_FOLDER, DEFAULT_SCV_NAME)
                    
                    if not scv_file_path:
                        scv_file_path = find_scv_file(SCV_FOLDER, "*.scv")
            except Exception as e:
                print(f"⚠️ 获取配置文件路径失败: {e}")
                scv_file_path = find_scv_file(SCV_FOLDER, DEFAULT_SCV_NAME)
        
        if not scv_file_path or not os.path.exists(scv_file_path):
            print(f"❌ 无法找到可用的SCV文件")
            print(f"💡 请尝试以下方法:")
            print(f"   1. 使用 --scv-file 参数指定文件路径")
            print(f"   2. 将SCV文件放在以下位置之一:")
            print(f"      - 当前目录: {current_dir}")
            print(f"      - SCV文件夹: {SCV_FOLDER}")
            print(f"      - 脚本同目录: {script_dir}")
            
            print(f"\n📁 当前目录中的SCV文件:")
            current_scv_files = [f for f in os.listdir(current_dir) if f.endswith('.scv')]
            if current_scv_files:
                for f in current_scv_files:
                    print(f"   - {f}")
            else:
                print("   (无)")
            
            if os.path.exists(SCV_FOLDER):
                print(f"\n📁 {SCV_FOLDER} 中的SCV文件:")
                try:
                    folder_files = [f for f in os.listdir(SCV_FOLDER) if f.endswith('.scv')]
                    if folder_files:
                        for f in folder_files:
                            print(f"   - {f}")
                    else:
                        print("   (无)")
                except Exception as e:
                    print(f"   无法访问文件夹: {e}")
            
            return 1
        
        scv_file_name = os.path.basename(scv_file_path)
        print(f"📂 使用的SCV文件: {scv_file_name}")
        print(f"📁 完整路径: {scv_file_path}")
        
    except ImportError as e:
        print(f"❌ 导入配置文件失败: {e}")
        print(f"当前目录: {current_dir}")
        print(f"脚本目录: {script_dir}")
        print(f"请确保 souren_config.py 文件存在于以下位置之一:")
        print(f"  1. 脚本目录: {script_dir}")
        print(f"  2. 当前目录: {current_dir}")
        print(f"  3. Python路径")
        return 1
    
    scv_folder_abs = os.path.abspath(SCV_FOLDER)
    print(f"📂 SCV文件夹路径: {scv_folder_abs}")
    
    if not os.path.exists(scv_folder_abs):
        print(f"⚠️ SCV文件夹不存在: {scv_folder_abs}")
        print(f"💡 将在其他位置查找SCV文件")
    
    try:
        from souren_manager import SourenManager, display_execution_result
        print(f"✅ 所有模块导入成功")
    except ImportError as e:
        print(f"❌ 导入模块失败: {e}")
        print(f"请确保以下文件存在于脚本目录 ({script_dir}):")
        print(f"  - souren_core.py")
        print(f"  - souren_manager.py")
        print(f"  - souren_monitor.py")
        print(f"  - souren_config.py")
        print(f"  - souren_exporter.py")
        return 1
    

    manager = SourenManager()
    
    print("=" * 60)
    print(f"启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"脚本目录: {script_dir}")
    print(f"SCV文件夹: {scv_folder_abs}")
    print(f"仪器IP: {DEFAULT_IP}")
    print(f"SCV模式: {'动态生成' if use_dynamic else '使用现有SCV文件'}")
    if use_dynamic:
        from souren_config import PYTHON_SCRIPT_PATH
        print(f"动态源文件: {os.path.basename(PYTHON_SCRIPT_PATH)}")
        print(f"输出格式: {scv_format}")
    print(f"执行文件: {scv_file_name}")
    print(f"日志功能: {'启用' if LOG_ENABLED else '禁用'}")
    print("=" * 60)
    
    print("\n📋 使用说明:")
    print(f"1. 程序将自动连接到仪器: {DEFAULT_IP}")
    if use_dynamic:
        print(f"2. 动态生成SCV文件并加载执行")
        from souren_config import PYTHON_SCRIPT_PATH
        print(f"3. 源文件: {os.path.basename(PYTHON_SCRIPT_PATH)}")
        print(f"4. 格式: {scv_format}")
    else:
        print(f"2. 加载现有的SCV文件")
    print(f"3. 执行测试步骤并实时获取仪器数据")
    print(f"4. 测试完成后自动保存结果并导出Excel文件")
    print(f"5. 按 Ctrl+C 可以中断测试，已执行的数据会被保存")
    print("=" * 60)
    
    # 处理命令行参数
    if args.list:
        success, result = manager.list_scv_files()
        if success and result.get("success"):
            files = result.get("files", [])
            print(f"\n📁 SCV文件列表 ({result.get('count', 0)} 个):")
            
            for i, file_info in enumerate(files, 1):
                size_mb = file_info.get("size", 0) / (1024 * 1024)
                modified = file_info.get("modified", "未知时间")
                
                if file_info['name'] == DEFAULT_SCV_NAME:
                    default_mark = " (默认静态文件)"
                elif "dynamic" in file_info['name'].lower() or "test_steps" in file_info['name']:
                    default_mark = " (动态/生成文件)"
                else:
                    default_mark = ""
                
                print(f"  {i}. {file_info['name']}{default_mark} ({size_mb:.2f} MB, 修改于: {modified})")
        else:
            print(f"❌ 获取文件列表失败: {result.get('error', '未知错误')}")
        return 0
    
    if args.run:
        print(f"\n🚀 开始执行自动化工作流程...")
    else:
        print(f"\n🚀 开始执行自动化工作流程（默认运行）...")
    
    print(f"📄 执行文件: {scv_file_name}")
    
    print("\n⚠ 注意:")
    print(f"1. 确保仪器已开机并连接到网络")
    print(f"2. 仪器IP: {DEFAULT_IP}")
    print(f"4. SCV文件路径: {scv_file_path}")
    if use_dynamic:
        from souren_config import PYTHON_SCRIPT_PATH
        print(f"5. 动态生成源: {os.path.basename(PYTHON_SCRIPT_PATH)}")
        print(f"6. 输出格式: {scv_format}")
    print(f"7. 程序将实时获取仪器数据并在测试完成后保存结果")
    print(f"8. 测试完成后会自动生成JSON和Excel结果文件")
    print(f"9. 按 Ctrl+C 可以中断测试，已执行的数据会被保存")
    print("=" * 60)
    
    confirm = input("开始执行? (y/n): ").strip().lower()
    if confirm != 'y':
        print("❌ 用户取消")
        return 0
    
    try:
        success, result = manager.run_complete_workflow(scv_file_path)
        
        display_execution_result(result)
        
        print("\n📝 实现说明:")
        print("✅ 实时版本特点:")
        print("  1. 直接与仪器通信，无模拟数据")
        print("  2. 实时获取仪器返回结果")
        print("  3. 测试完成后自动保存JSON结果文件")
        print("  4. 自动导出Excel文件便于数据分析")
        print("  5. 支持循环测试")
        print("  6. 支持用户中断(Ctrl+C),已执行数据会被保存")
        if use_dynamic:
            print("  7. 支持从Python脚本动态生成SCV文件")
        print("=" * 60)
        
        if success or result.get("interrupted"):
            if result.get("interrupted"):
                print("⏹️  测试被用户中断")
            else:
                print("✅ 工作流程执行完成")
            print(f"💾 结果文件已保存到脚本目录下的log文件夹:")
            print(f"   📄 日志文件: souren_execution_*.log")
            print(f"   📄 JSON结果: souren_results_*.json")
            print(f"   📊 Excel文件: souren_results_*.xlsx")
            print(f"   📁 脚本目录: {script_dir}")
            
            print(f"\n📊 Excel导出说明:")
            print(f"  1. 自动提取详细执行记录中所有查询命令（包含?的命令）的数字结果")
            print(f"  2. 按循环轮次和步骤编号整理数据")
            print(f"  3. 创建三维折线图展示数据趋势")
            print(f"  4. 不同颜色代表不同循环轮次的数据")
            return 0
        else:
            print("❌ 工作流程执行失败")
            return 1
            
    except KeyboardInterrupt:
        print("\n\n⏹️  用户中断执行")
        
        # ⭐⭐⭐ 关键：即使中断也要尝试保存和导出数据
        print(f"\n📊 正在尝试保存已执行的数据...")
        try:
            # 尝试找到最新的JSON结果文件并导出Excel
            from souren_exporter import ResultExporter
            exporter = ResultExporter()
            json_file = exporter.find_latest_json_result()
            if json_file:
                print(f"📁 找到结果文件: {os.path.basename(json_file)}")
                excel_success = exporter.convert_to_excel(json_file)
                if excel_success:
                    print(f"✅ Excel文件导出成功!")
                else:
                    print(f"⚠️  Excel文件导出失败")
            else:
                print("❌ 未找到JSON结果文件")
        except Exception as e:
            print(f"⚠️  导出数据时出错: {e}")
        
        return 0
    except Exception as e:
        print(f"\n❌ 程序执行异常: {e}")
        import traceback
        traceback.print_exc()
        return 1
    finally:
        # 清理资源
        try:
            manager.cleanup()
            print("🧹 资源已清理")
        except:
            pass

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n👋 用户中断")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 程序错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)