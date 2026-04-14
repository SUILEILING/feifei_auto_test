from lib.var import *
from souren_monitor import SourenEngine, ExecutionStatus

class SourenManager():
    
    def __init__(self, result_dir=None, params=None):
        from souren_config import EXECUTION_MODE, LOOP_COUNT
        
        self.execution_mode = EXECUTION_MODE
        self.loop_count = LOOP_COUNT
        self.result_dir = result_dir
        self.params = params
        
        self.script_name = None
        if result_dir:
            dir_name = os.path.basename(result_dir)
            if "_" in dir_name:
                parts = dir_name.split("_")
                if "b" in parts[1] or "bw" in parts[1] or "scs" in parts[1]:
                    self.script_name = f"{parts[0]}.py"
                else:
                    self.script_name = f"{dir_name}.py"
        
        self.engine = SourenEngine(result_dir, self.script_name, params)
        self.logger = self.engine.logger
        self.result_saver = self.engine.result_saver
        
        self.current_file = None
        self.selected_device_info = None
        self.execution_history = []
        self.interrupted = False
        self.excel_file = None
    
    def get_result_file(self):
        return self.engine.get_result_file()
    
    def get_excel_file(self):
        return self.excel_file
    
    def cleanup(self):
        self.logger.info("开始清理资源")
        self.engine.cleanup()
        self.logger.info("资源清理完成")
    
    def initialize_system(self) -> Tuple[bool, str]:
        self.logger.info("开始初始化系统")
        
        success, message = self.engine.initialize()
        if not success:
            return False, message
        
        self.logger.info(f"系统初始化完成")
        return True, f"系统初始化完成"
    
    def load_python_file(self, file_path: str) -> Tuple[bool, Dict]:
        if not file_path:
            return False, {"error": "未提供文件路径"}
        
        self.current_file = os.path.basename(file_path)
        
        self.logger.info(f"加载Python文件: {self.current_file}")
        self.logger.info(f"文件完整路径: {file_path}")
        
        if not os.path.exists(file_path):
            print(f"❌ 文件不存在: {file_path}")
            return False, {"error": f"文件不存在: {file_path}"}
        
        if self.params:
            self.logger.info(f"使用参数配置: {self.params}")
        
        success, result = self.engine.execute_command("load_python_file", file_path=file_path)
        return success, result
    
    def get_instrument_info(self) -> Tuple[bool, Dict]:
        self.logger.info("获取仪器信息")
        
        from souren_config import INSTRUMENT_ADDRESS
        
        device_info = {
            "id": INSTRUMENT_ADDRESS,
            "name": INSTRUMENT_ADDRESS,
            "type": "TCPIP",
            "status": "online",
            "address": INSTRUMENT_ADDRESS
        }
        
        self.selected_device_info = device_info
        
        print(f"✅ 获取仪器信息成功: {INSTRUMENT_ADDRESS}")
        
        return True, {
            "success": True,
            "device": device_info,
            "message": f"成功获取仪器信息: {INSTRUMENT_ADDRESS}"
        }
    
    def execute_tests(self, loop_iteration: int = 1) -> Tuple[bool, Dict]:
        if not self.current_file:
            return False, {"error": "未加载Python文件"}
        
        self.logger.info(f"开始执行测试 - 模式: {self.execution_mode}, 循环次数: {loop_iteration}")
        
        try:
            if self.execution_mode == 'loop_info':
                success, result = self.engine.execute_command("run_tests", 
                                                          mode=self.execution_mode, 
                                                          loop_count=self.loop_count,
                                                          loop_iteration=loop_iteration)
            else:
                success, result = self.engine.execute_command("run_tests", mode=self.execution_mode)
            
            if result and result.get("interrupted"):
                self.interrupted = True
                print(f"⚠️ 测试被用户中断，但已保存已执行的数据")
            
            if success and result.get("success"):
                execution_data = {
                    "status": result.get("status", ExecutionStatus.SUCCESS.value),
                    "file": self.current_file,
                    "device": self.selected_device_info,
                    "mode": self.execution_mode,
                    "result": result,
                    "timestamp": datetime.now().isoformat(),
                    "loop_iteration": loop_iteration
                }
                
                if self.params:
                    execution_data["parameters"] = self.params
                
                if self.execution_mode == 'loop_info':
                    execution_data["loop_count"] = self.loop_count
                
                self.result_saver.save_result(execution_data)
                self.execution_history.append(execution_data)
            
            return success, result
            
        except KeyboardInterrupt:
            print(f"\n⏹️  测试执行过程中被用户中断")
            self.interrupted = True
            return False, {"error": "用户中断测试执行", "interrupted": True}
    
    def run_complete_workflow(self, python_file_path: str = None) -> Tuple[bool, Dict]:
        self.logger.info("开始完整工作流程执行")
        
        workflow_result = {
            "workflow": "complete",
            "steps": [],
            "overall_success": False,
            "start_time": datetime.now().isoformat(),
            "loop_count": self.loop_count if self.execution_mode == 'loop_info' else 1,
            "loop_results": []
        }
        
        if self.params:
            workflow_result["parameters"] = self.params
        
        success, message = self.initialize_system()
        workflow_result["steps"].append({
            "step": 1,
            "name": "initialize_system",
            "success": success,
            "message": message
        })
        
        if not success:
            workflow_result["overall_success"] = False
            workflow_result["end_time"] = datetime.now().isoformat()
            workflow_result["error"] = message
            return False, workflow_result
        
        if not python_file_path:
            workflow_result["overall_success"] = False
            workflow_result["end_time"] = datetime.now().isoformat()
            workflow_result["error"] = "未提供Python文件路径"
            return False, workflow_result
        
        success, result = self.load_python_file(python_file_path)
        workflow_result["steps"].append({
            "step": 2,
            "name": "load_file",
            "success": success,
            "result": result
        })
        
        if not success:
            workflow_result["overall_success"] = False
            workflow_result["end_time"] = datetime.now().isoformat()
            workflow_result["error"] = result.get("error", "加载文件失败")
            return False, workflow_result
        
        success, device_result = self.get_instrument_info()
        workflow_result["steps"].append({
            "step": 3,
            "name": "get_instrument_info",
            "success": success,
            "result": device_result
        })
        
        if not success:
            workflow_result["overall_success"] = False
            workflow_result["end_time"] = datetime.now().isoformat()
            workflow_result["error"] = device_result.get("error", "获取仪器信息失败")
            return False, workflow_result
        
        print(f"\n🔄 开始循环执行测试，总循环次数: {self.loop_count}")
        all_test_results = []
        
        for loop_idx in range(1, self.loop_count + 1):
            if self.interrupted:
                print(f"⏹️ 循环执行被用户中断")
                workflow_result["interrupted"] = True
                break
            
            print(f"\n{'='*60}")
            print(f"🔄 开始第 {loop_idx}/{self.loop_count} 次循环")
            print(f"{'='*60}")
            
            if loop_idx > 1:
                interval = 3  
                print(f"⏰ 等待 {interval} 秒后开始下一次循环...")
                time.sleep(interval)
            
            success, test_result = self.execute_tests(loop_iteration=loop_idx)
            
            loop_result = {
                "loop_index": loop_idx,
                "success": success,
                "result": test_result
            }
            
            if self.params:
                loop_result["parameters"] = self.params
            
            workflow_result["loop_results"].append(loop_result)
            
            workflow_result["steps"].append({
                "step": 4 + loop_idx - 1,
                "name": f"execute_tests_loop_{loop_idx}",
                "success": success,
                "result": test_result
            })
            
            if success and test_result.get("success"):
                print(f"✅ 第 {loop_idx} 次循环执行成功")
                all_test_results.append(test_result)
            else:
                print(f"❌ 第 {loop_idx} 次循环执行失败")
                if test_result.get("error"):
                    print(f"   错误信息: {test_result.get('error')}")
        
        if self.interrupted:
            workflow_result["overall_success"] = True
            workflow_result["interrupted"] = True
            workflow_result["status"] = "interrupted"
            workflow_result["message"] = "测试被用户中断，但已保存已执行的数据"
            print(f"\n⏹️  测试被用户中断，已执行 {len(workflow_result['loop_results'])} 次循环")
        else:
            all_success = all(result.get("success", False) for result in workflow_result["loop_results"])
            workflow_result["overall_success"] = all_success
            workflow_result["status"] = "completed" if all_success else "failed"
            workflow_result["message"] = f"完成所有 {self.loop_count} 次循环执行"
            print(f"\n✅ 完成所有 {self.loop_count} 次循环执行")
        
        workflow_result["end_time"] = datetime.now().isoformat()
        workflow_result["total_loops_executed"] = len(workflow_result["loop_results"])
        
        if all_test_results or self.interrupted:
            workflow_result["final_result"] = all_test_results if all_test_results else workflow_result["loop_results"]
            
            if self.interrupted:
                self.logger.info(f"完整工作流程被用户中断，已执行 {len(workflow_result['loop_results'])} 次循环")
            else:
                self.logger.info(f"完整工作流程执行成功，完成 {len(workflow_result['loop_results'])} 次循环")
            
            print("\n" + "=" * 60)
            print("📤 正在自动导出结果到Excel文件...")
            print("=" * 60)
            
            excel_success = self._auto_export_to_excel(workflow_result)
            if excel_success:
                workflow_result["excel_export"] = "成功"
                print("✅ Excel文件导出成功!")
            else:
                workflow_result["excel_export"] = "失败"
                print("⚠️  Excel文件导出失败,但测试已完成")
            
            print("=" * 60)
        else:
            workflow_result["error"] = "所有循环执行都失败"
            self.logger.error("完整工作流程执行失败，所有循环都失败")
        
        self.result_saver.save_result(workflow_result)
        
        return workflow_result["overall_success"], workflow_result
    
    def _auto_export_to_excel(self, test_result: Dict) -> bool:
        try:
            from souren_exporter import ResultExporter
            
            exporter = ResultExporter()
            
            json_file = self.get_result_file()
            if json_file is None or not os.path.exists(json_file):
                print("❌ 未找到JSON结果文件,无法导出Excel")
                return False
            
            print(f"📁 找到结果文件: {os.path.basename(json_file)}")
            print(f"📁 保存目录: {os.path.dirname(json_file)}")
            
            success = exporter.convert_to_excel(json_file)
            
            if success:
                excel_file = json_file.replace('.json', '.xlsx')
                if os.path.exists(excel_file):
                    self.excel_file = excel_file
                    print(f"✅ Excel文件已保存: {os.path.basename(excel_file)}")
            
            return success
            
        except ImportError as e:
            print(f"❌ 导出模块导入失败: {e}")
            return False
        except Exception as e:
            print(f"❌ 自动导出Excel失败: {e}")
            return False
    
    def export_to_excel(self):
        return self._auto_export_to_excel({})

def display_execution_result(result: Dict):
    print("\n" + "=" * 60)
    
    if 'parameters' in result:
        print(f"📋 参数配置:")
        for key, value in result['parameters'].items():
            print(f"   {key}: {value}")
        print("-" * 40)
    
    if result.get("interrupted"):
        print("⏹️  执行状态: 用户中断")
    elif result.get("overall_success"):
        print("✅ 执行状态: 成功")
    else:
        print("❌ 执行状态: 失败")
        if "error" in result:
            print(f"   错误信息: {result['error']}")
    
    print("=" * 60)
    
    print(f"📅 开始时间: {result.get('start_time', '未知')}")
    print(f"📅 结束时间: {result.get('end_time', '未知')}")
    
    if "loop_count" in result:
        print(f"🔄 配置循环次数: {result['loop_count']}")
        print(f"🔄 实际执行次数: {result.get('total_loops_executed', 0)}")
    
    if "excel_export" in result:
        excel_status = "✅ 成功" if result["excel_export"] == "成功" else "❌ 失败"
        print(f"📤 Excel导出: {excel_status}")
    
    steps = result.get("steps", [])
    if steps:
        print("\n🔧 执行步骤详情:")
        for step in steps:
            status = "✅" if step.get("success") else "❌"
            step_name = step.get("name", "未知步骤")
            print(f"  {status} {step_name}")
    
    loop_results = result.get("loop_results", [])
    if loop_results:
        print(f"\n🔄 循环执行结果:")
        for loop_result in loop_results:
            loop_idx = loop_result.get("loop_index", 0)
            success = loop_result.get("success", False)
            status = "✅" if success else "❌"
            print(f"  第{loop_idx}次循环: {status}")
    
    print("=" * 60)