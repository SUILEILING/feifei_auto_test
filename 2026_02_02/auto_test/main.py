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

def create_script_subdirectory(main_execution_dir, script_name, params=None):
    """为每个脚本创建子目录"""
    script_base = os.path.splitext(os.path.basename(script_name))[0]
    
    if params:
        param_str = ""
        if 'lineLoss' in params:
            lineLoss_value = params['lineLoss']
            if isinstance(lineLoss_value, (int, float)):
                lineLoss_int = int(lineLoss_value)
                param_str += f"_ll{lineLoss_int}"
            else:
                try:
                    lineLoss_int = int(float(lineLoss_value))
                    param_str += f"_ll{lineLoss_int}"
                except:
                    param_str += f"_ll{lineLoss_value}"
        if 'band' in params:
            param_str += f"_b{params['band']}"
        if 'bw' in params:
            param_str += f"_bw{params['bw']}"
        if 'scs' in params:
            param_str += f"_scs{params['scs']}"
        if 'range' in params:
            param_str += f"_{params['range']}"
        
        sub_dir = os.path.join(main_execution_dir, f"{script_base}{param_str}")
    else:
        sub_dir = os.path.join(main_execution_dir, f"{script_base}")
    
    if not os.path.exists(sub_dir):
        os.makedirs(sub_dir, exist_ok=True)
        print(f"📁 为脚本 {script_name} 创建子目录: {sub_dir}")
    
    return sub_dir

def main():
    parser = argparse.ArgumentParser(description='Souren.ToolSet 自动化系统')
    parser.add_argument('--list', '-l', action='store_true', help='列出Python脚本文件')
    parser.add_argument('--run', '-r', action='store_true', help='执行完整工作流程')
    parser.add_argument('--debug-cell', action='store_true', help='调试CELL命令')
    parser.add_argument('--py-file', type=str, help='指定Python脚本文件路径')
    
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
            display_config_info, 
            DEFAULT_IP, 
            PYTHON_SCRIPT_NAME, LOG_ENABLED,
            LOOP_COUNT, EXECUTION_MODE, set_execution_dir,
            PARAMETER_COMMAND_MAPPINGS
        )
    except ImportError as e:
        print(f"❌ 导入配置文件失败: {e}")
        print(f"当前目录: {current_dir}")
        print(f"脚本目录: {script_dir}")
        return 1
    
    main_execution_dir = None
    if LOG_ENABLED:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        main_execution_dir = os.path.join(script_dir, "log", f"execution_{timestamp}")
        if not os.path.exists(main_execution_dir):
            os.makedirs(main_execution_dir, exist_ok=True)
            print(f"📁 创建主执行目录: {main_execution_dir}")
    
    if main_execution_dir:
        set_execution_dir(main_execution_dir)
    
    display_config_info()
    print(f"✅ 配置文件导入成功")
    
    script_configs = []
    
    if args.py_file:
        py_file_path = args.py_file
        if not os.path.exists(py_file_path):
            print(f"❌ 指定的Python脚本文件不存在: {py_file_path}")
            return 1
        script_configs = [{"script_path": py_file_path, "params": None}]
        print(f"📄 使用命令行指定的Python文件: {os.path.basename(py_file_path)}")
    else:
        if not PYTHON_SCRIPT_NAME:
            print(f"❌ 未配置Python脚本文件")
            return 1
            
        print(f"📄 使用配置文件中的Python脚本配置")
        
        for script_config in PYTHON_SCRIPT_NAME:
            script_name = None
            params = None
            
            if isinstance(script_config, str):
                script_name = script_config
                params = None
            elif isinstance(script_config, dict):
                script_name = script_config.get('script', '')
                params = {k: v for k, v in script_config.items() if k != 'script'}
            
            if not script_name:
                print(f"  ❌ 脚本配置中未找到脚本名: {script_config}")
                continue
            
            script_path = None
            save_scv_dir = os.path.join(current_dir, "save_scv")
            save_scv_path = os.path.join(save_scv_dir, script_name)
            if os.path.exists(save_scv_path):
                script_path = save_scv_path
                print(f"  ✅ 在save_scv目录中找到脚本: {os.path.basename(script_name)}")
            
            if not script_path:
                script_save_scv_dir = os.path.join(script_dir, "save_scv")
                script_save_scv_path = os.path.join(script_save_scv_dir, script_name)
                if os.path.exists(script_save_scv_path):
                    script_path = script_save_scv_path
                    print(f"  ✅ 在脚本目录的save_scv中找到脚本: {os.path.basename(script_name)}")
            
            if not script_path:
                current_path = os.path.join(current_dir, script_name)
                if os.path.exists(current_path):
                    script_path = current_path
                    print(f"  ✅ 在当前目录中找到脚本: {os.path.basename(script_name)}")
            
            if not script_path:
                script_dir_path = os.path.join(script_dir, script_name)
                if os.path.exists(script_dir_path):
                    script_path = script_dir_path
                    print(f"  ✅ 在脚本目录中找到脚本: {os.path.basename(script_name)}")
            
            if script_path:
                script_configs.append({
                    "script_path": script_path,
                    "script_name": script_name,
                    "params": params
                })
                
                if params:
                    print(f"    参数配置: {params}")
            else:
                print(f"  ❌ 未找到脚本: {script_name}")
        
    if not script_configs:
        print(f"❌ 无法找到可用的Python脚本文件")
        return 1
    
    print(f"\n✅ 找到 {len(script_configs)} 个Python脚本配置:")
    for i, config in enumerate(script_configs, 1):
        script_name = os.path.basename(config["script_path"])
        params = config["params"]
        if params:
            param_str = " | "
            for key, value in params.items():
                param_str += f"{key}:{value} "
            print(f"  {i}. {script_name}{param_str}")
        else:
            print(f"  {i}. {script_name} (无参数)")
    
    print(f"📁 主执行目录: {main_execution_dir}")
    
    try:
        from souren_manager import SourenManager, display_execution_result
        print(f"✅ 所有模块导入成功")
    except ImportError as e:
        print(f"❌ 导入模块失败: {e}")
        return 1
    
    if args.list:
        print(f"\n📁 Python脚本文件列表 ({len(script_configs)} 个):")
        
        for i, config in enumerate(script_configs, 1):
            script_path = config["script_path"]
            try:
                size_mb = os.path.getsize(script_path) / (1024 * 1024)
                modified = datetime.fromtimestamp(
                    os.path.getmtime(script_path)
                ).strftime('%Y-%m-%d %H:%M:%S')
                
                script_name = os.path.basename(script_path)
                params = config["params"]
                param_info = ""
                if params:
                    param_info = " | "
                    for key, value in params.items():
                        param_info += f"{key}:{value} "
                
                print(f"  {i}. {script_name} ({size_mb:.2f} MB, 修改于: {modified}){param_info}")
            except:
                print(f"  {i}. {os.path.basename(script_path)}")
        return 0
    
    if args.run:
        print(f"\n🚀 开始执行自动化工作流程...")
    else:
        print(f"\n🚀 开始执行自动化工作流程（默认运行）...")
    
    all_results = []
    interrupted = False
    
    for config_idx, config in enumerate(script_configs, 1):
        py_file_path = config["script_path"]
        params = config["params"]
        script_name = config["script_name"]
        
        if not os.path.exists(py_file_path):
            print(f"❌ Python文件不存在: {py_file_path}")
            continue
        
        py_file_name = os.path.basename(py_file_path)
        print(f"\n{'='*60}")
        print(f"📂 处理第 {config_idx}/{len(script_configs)} 个配置: {py_file_name}")
        
        if params:
            print(f"📋 参数配置:")
            for key, value in params.items():
                print(f"    {key}: {value}")
        
        print(f"📁 完整路径: {py_file_path}")
        
        script_sub_dir = None
        if main_execution_dir:
            script_sub_dir = create_script_subdirectory(main_execution_dir, py_file_name, params)
            print(f"📁 脚本专属目录: {script_sub_dir}")
        
        print(f"\n⚠ 注意:")
        print(f"1. 确保仪器已开机并连接到网络")
        print(f"2. 仪器IP: {DEFAULT_IP}")
        print(f"3. Python脚本路径: {py_file_path}")
        if params:
            print(f"4. 参数配置: {params}")
        if EXECUTION_MODE == 'loop_info':
            print(f"5. 循环模式: 每个脚本执行 {LOOP_COUNT} 次")
        print(f"6. 程序将实时获取仪器数据并在测试完成后保存结果")
        print(f"7. 测试完成后会自动生成JSON和Excel结果文件")
        print(f"8. 按 Ctrl+C 可以中断测试，已执行的数据会被保存")
        print("="*60)
        
        if config_idx == 1:
            confirm = input("开始执行? (y/n): ").strip().lower()
            if confirm != 'y':
                print("❌ 用户取消")
                return 0
        
        try:
            manager = SourenManager(script_sub_dir, params)
            success, result = manager.run_complete_workflow(py_file_path)
            
            result['script_file'] = py_file_name
            result['script_name'] = py_file_name
            result['script_directory'] = script_sub_dir
            if params:
                result['parameters'] = params
            all_results.append(result)
            
            display_execution_result(result)
            
            print("\n📝 实现说明:")
            print("✅ 实时版本特点:")
            print("  1. 直接与仪器通信，无模拟数据")
            print("  2. 实时获取仪器返回结果")
            print("  3. 支持参数化配置，动态生成命令")
            print("  4. 测试完成后自动保存JSON结果文件")
            print("  5. 自动导出Excel文件便于数据分析")
            print("  6. 支持循环测试")
            print("  7. 支持用户中断(Ctrl+C),已执行数据会被保存")
            print("  8. 直接从Python脚本中读取测试步骤")
            print("="*60)
            
            if success or result.get("interrupted"):
                if result.get("interrupted"):
                    print("⏹️  测试被用户中断")
                else:
                    print("✅ 工作流程执行完成")
                print(f"💾 结果文件已保存:")
                if script_sub_dir:
                    print(f"   📁 脚本专属目录: {script_sub_dir}")
                    
                    result_file = manager.get_result_file()
                    excel_file = manager.get_excel_file()
                    
                    if result_file and os.path.exists(result_file):
                        print(f"   📄 JSON结果: {os.path.basename(result_file)}")
                    else:
                        print(f"   📄 JSON结果: (未生成或不存在)")
                    
                    if excel_file and os.path.exists(excel_file):
                        print(f"   📊 Excel文件: {os.path.basename(excel_file)}")
                    else:
                        print(f"   📊 Excel文件: (未生成或不存在)")
                
                print(f"\n📊 Excel导出说明:")
                print(f"  1. 自动提取详细执行记录中所有查询命令（包含?的命令）的数字结果")
                print(f"  2. 按循环轮次和步骤编号整理数据")
                print(f"  3. 创建三维折线图展示数据趋势")
                print(f"  4. 不同颜色代表不同循环轮次的数据")
            else:
                print("❌ 工作流程执行失败")
            
            manager.cleanup()
            print(f"\n✅ 第 {config_idx}/{len(script_configs)} 个配置处理完成")
            
            if config_idx < len(script_configs):
                print(f"\n⏱️  准备处理下一个配置...")
                time.sleep(2)
                
        except KeyboardInterrupt:
            print(f"\n\n⏹️  用户中断执行 (配置 {config_idx}/{len(script_configs)})")
            interrupted = True
            
            print(f"\n📊 正在尝试保存已执行的数据...")
            try:
                if 'manager' in locals():
                    result_file = manager.get_result_file()
                    if result_file and os.path.exists(result_file):
                        print(f"📁 找到结果文件: {os.path.basename(result_file)}")
                        excel_file = manager.export_to_excel()
                        if excel_file:
                            print(f"✅ Excel文件导出成功!")
                        else:
                            print(f"⚠️  Excel文件导出失败")
                    else:
                        print("❌ 未找到JSON结果文件")
                    
                    if 'result' in locals():
                        all_results.append(result)
                        
            except Exception as e:
                print(f"⚠️  导出数据时出错: {e}")
            
            break
        except Exception as e:
            print(f"\n❌ 程序执行异常: {e}")
            import traceback
            traceback.print_exc()
            continue
    
    print(f"\n{'='*60}")
    
    if interrupted:
        print("⏹️  执行被用户中断，正在生成已执行数据的汇总...")
    else:
        print("🎉 所有配置处理完成")
    
    print(f"{'='*60}")
    
    total_files = len(script_configs)
    success_files = sum(1 for r in all_results if r.get("overall_success", False))
    interrupted_files = sum(1 for r in all_results if r.get("interrupted", False))
    failed_files = total_files - success_files - interrupted_files
    
    print(f"📊 执行统计:")
    print(f"  总配置数: {total_files}")
    print(f"  成功完成: {success_files}")
    print(f"  用户中断: {interrupted_files}")
    print(f"  执行失败: {failed_files}")
    
    if main_execution_dir and all_results:
        print(f"\n📁 所有结果文件保存在:")
        print(f"  {main_execution_dir}")
        print(f"  每个配置有自己的子目录:")
        for result in all_results:
            if result.get('script_directory'):
                script_name = result.get('script_name', '未知脚本')
                params_info = ""
                if result.get('parameters'):
                    params_info = " (参数: "
                    for key, value in result['parameters'].items():
                        params_info += f"{key}:{value} "
                    params_info += ")"
                print(f"  - {os.path.basename(result['script_directory'])}: {script_name}{params_info}")
        
        print(f"\n{'='*60}")
        print("📊 开始创建汇总Excel文件...")
        if interrupted:
            print("📌 注意：当前为中断后生成的汇总，仅包含已执行的数据")
        print(f"{'='*60}")
        
        try:
            from souren_exporter import create_summary_excel
            summary_file = create_summary_excel(main_execution_dir, all_results)
            
            if summary_file:
                print(f"✅ 汇总Excel文件已创建: {os.path.basename(summary_file)}")
                print(f"📁 路径: {summary_file}")
                
                if interrupted:
                    print(f"\n💡 提示：由于执行被中断，汇总图表仅包含已执行的 {len(all_results)} 个配置的数据")
                    print(f"     未执行的配置: {total_files - len(all_results)} 个")
            else:
                print("⚠️  汇总Excel文件创建失败")
        except Exception as e:
            print(f"❌ 创建汇总Excel文件失败: {e}")
            
    if interrupted:
        return 2
    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n👋 用户中断")
        sys.exit(2)
    except Exception as e:
        print(f"\n❌ 程序错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)