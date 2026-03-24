#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Souren.ToolSet 自动化执行系统 - 增强版（支持动态SCV生成）
"""

import sys
import argparse
import os
import time
from datetime import datetime

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
    
    # 创建仪器控制器
    controller = InstrumentController()
    
    # 连接
    print("🔌 连接仪器...")
    success, message = controller.connect()
    if not success:
        print(f"❌ 连接失败: {message}")
        return
    
    print("✅ 连接成功，开始测试命令...\n")
    
    # 测试不同的命令格式
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
            time.sleep(1)  # 等待1秒
    
    controller.disconnect()
    print(f"\n{'='*60}")
    print("🔧 调试完成")
    print("="*60)

def main():
    parser = argparse.ArgumentParser(description='Souren.ToolSet 自动化系统')
    parser.add_argument('--list', '-l', action='store_true', help='列出SCV文件')
    parser.add_argument('--run', '-r', action='store_true', help='执行完整工作流程')
    parser.add_argument('--debug-cell', action='store_true', help='调试CELL命令')
    parser.add_argument('--dynamic-scv', action='store_true', help='使用动态SCV生成')
    parser.add_argument('--static-scv', action='store_true', help='使用静态SCV文件')
    parser.add_argument('--scv-format', choices=['standard', 'table'], default='standard', 
                       help='动态SCV输出格式 (standard/table)')
    parser.add_argument('--scv-file', type=str, help='指定SCV文件路径（覆盖默认配置）')
    
    args = parser.parse_args()
    
    # 调试CELL命令
    if args.debug_cell:
        debug_cell_command()
        return 0
    
    # 首先显示当前工作目录
    current_dir = os.getcwd()
    print(f"📁 当前工作目录: {current_dir}")
    print(f"📁 Python执行目录: {os.path.dirname(os.path.abspath(__file__))}")
    
    try:
        # 先导入配置
        from souren_config import (
            SCV_FOLDER, DEFAULT_SCV_NAME, display_config_info, 
            DEFAULT_IP, DEFAULT_PORT, USE_DYNAMIC_SCV, SCV_OUTPUT_FORMAT,
            PYTHON_SCRIPT_NAME, get_scv_file_path
        )
        
        # 根据命令行参数覆盖配置
        use_dynamic = USE_DYNAMIC_SCV
        scv_format = SCV_OUTPUT_FORMAT
        
        if args.dynamic_scv:
            print("🔄 使用动态SCV生成模式（命令行参数覆盖）")
            use_dynamic = True
            if args.scv_format:
                scv_format = args.scv_format
                print(f"📄 SCV格式: {scv_format}")
        
        if args.static_scv:
            print("📄 使用静态SCV文件模式（命令行参数覆盖）")
            use_dynamic = False
        
        display_config_info()
        print(f"✅ 配置文件导入成功")
        
        # 获取SCV文件路径
        if args.scv_file:
            # 使用命令行指定的文件
            scv_file_path = args.scv_file
            if not os.path.exists(scv_file_path):
                print(f"❌ 指定的SCV文件不存在: {scv_file_path}")
                return 1
            print(f"📄 使用命令行指定的SCV文件: {os.path.basename(scv_file_path)}")
        else:
            # 根据配置获取文件
            if use_dynamic:
                print(f"🔄 动态SCV模式激活")
            else:
                print(f"📄 静态SCV模式激活")
            
            scv_file_path = get_scv_file_path()
            if not scv_file_path or not os.path.exists(scv_file_path):
                print(f"❌ SCV文件不存在: {scv_file_path}")
                return 1
        
        # 获取文件名用于显示
        scv_file_name = os.path.basename(scv_file_path)
        print(f"📂 使用的SCV文件: {scv_file_name}")
        print(f"📁 完整路径: {scv_file_path}")
        
    except ImportError as e:
        print(f"❌ 导入配置文件失败: {e}")
        print(f"当前目录: {os.getcwd()}")
        return 1
    
    # 检查SCV文件夹是否存在
    scv_folder_abs = os.path.abspath(SCV_FOLDER)
    print(f"📂 SCV文件夹路径: {scv_folder_abs}")
    
    if not os.path.exists(scv_folder_abs):
        print(f"❌ SCV文件夹不存在: {scv_folder_abs}")
        print(f"💡 请确保以下目录存在:")
        print(f"   相对路径: {SCV_FOLDER}")
        print(f"   绝对路径: {scv_folder_abs}")
        return 1
    
    try:
        # 然后导入其他模块
        from souren_manager import SourenManager, display_execution_result
        print(f"✅ 所有模块导入成功")
    except ImportError as e:
        print(f"❌ 导入模块失败: {e}")
        print(f"请确保以下文件存在于当前目录:")
        print(f"  - souren_core.py")
        print(f"  - souren_manager.py")
        print(f"  - souren_monitor.py")
        print(f"  - souren_config.py")
        print(f"  - souren_exporter.py")
        return 1
    
    # 检查Excel导出依赖
    try:
        import pandas as pd
        import openpyxl
        print("✅ Excel导出依赖已安装")
    except ImportError:
        print("⚠️  Excel导出依赖未完全安装")
        print("💡 如需Excel导出功能，请运行: pip install pandas openpyxl")
        print("   测试仍可正常执行，但不会生成Excel文件")
    
    # 创建管理器
    manager = SourenManager()
    
    print("=" * 60)
    print(f"启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"SCV文件夹: {os.path.abspath(SCV_FOLDER)}")
    print(f"仪器IP: {DEFAULT_IP}")
    print(f"仪器端口: {DEFAULT_PORT}")
    print(f"SCV模式: {'动态生成' if use_dynamic else '使用现有SCV文件'}")
    if use_dynamic:
        print(f"动态源文件: {PYTHON_SCRIPT_NAME}")
        print(f"输出格式: {scv_format}")
    print(f"执行文件: {scv_file_name}")
    print("=" * 60)
    
    print("\n📋 使用说明:")
    print(f"1. 程序将自动连接到仪器: {DEFAULT_IP}:{DEFAULT_PORT}")
    if use_dynamic:
        print(f"2. 动态生成SCV文件并加载执行")
        print(f"3. 源文件: {PYTHON_SCRIPT_NAME}")
        print(f"4. 格式: {scv_format}")
    else:
        print(f"2. 加载现有的SCV文件")
    print(f"3. 执行测试步骤并实时获取仪器数据")
    print(f"4. 测试完成后自动保存结果并导出Excel文件")  
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
    print(f"3. 仪器端口: {DEFAULT_PORT}")
    print(f"4. SCV文件路径: {scv_file_path}")
    if use_dynamic:
        print(f"5. 动态生成源: {PYTHON_SCRIPT_NAME}")
        print(f"6. 输出格式: {scv_format}")
    print(f"7. 程序将实时获取仪器数据并在测试完成后保存结果")
    print(f"8. 测试完成后会自动生成JSON和Excel结果文件")  
    print("=" * 60)
    
    confirm = input("开始执行? (y/n): ").strip().lower()
    if confirm != 'y':
        print("❌ 用户取消")
        return 0
    
    try:
        # 执行完整工作流程，传入SCV文件路径
        success, result = manager.run_complete_workflow(scv_file_path)
        
        # 显示详细结果
        display_execution_result(result)
        
        print("\n📝 实现说明:")
        print("✅ 实时版本特点:")
        print("  1. 直接与仪器通信，无模拟数据")
        print("  2. 实时获取仪器返回结果")
        print("  3. 测试完成后自动保存JSON结果文件")
        print("  4. 自动导出Excel文件便于数据分析")
        print("  5. 支持循环测试")
        if use_dynamic:
            print("  6. 支持从Python脚本动态生成SCV文件")
        print("=" * 60)
        
        if success:
            print("✅ 工作流程执行完成")
            print(f"💾 结果文件已保存到执行目录:")
            print(f"   📄 日志文件: souren_execution_*.log")
            print(f"   📄 JSON结果: souren_results_*.json")
            print(f"   📊 Excel文件: souren_results_*.xlsx")
            # 添加执行目录路径显示
            from souren_config import EXECUTION_DIR
            print(f"   📁 完整路径: {os.path.abspath(EXECUTION_DIR)}")
            
            # 显示Excel导出说明
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