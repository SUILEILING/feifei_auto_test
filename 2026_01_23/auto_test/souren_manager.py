from lib.var import *

from souren_monitor import SourenEngine, ExecutionStatus
from souren_core import SourenResultSaver


# ==============================================
# ↓↓↓ 自动化执行管理器
# ==============================================
class SourenManager():
    
    def __init__(self):
        from souren_config import SCV_FOLDER, DEFAULT_SCV_NAME, EXECUTION_MODE, LOOP_COUNT
        
        self.scv_folder = os.path.abspath(SCV_FOLDER)
        self.default_scv_name = DEFAULT_SCV_NAME
        self.execution_mode = EXECUTION_MODE
        self.loop_count = LOOP_COUNT
        
        self.engine = SourenEngine()
        self.logger = self.engine.logger
        self.result_saver = SourenResultSaver()
        
        self.current_file = None
        self.selected_device_info = None
        self.execution_history = []
        self.interrupted = False  # ⭐⭐⭐ 新增：中断标志
        
    def cleanup(self):
        """清理资源"""
        self.logger.info("开始清理资源")
        self.engine.cleanup()
        self.logger.info("资源清理完成")
    
    def initialize(self) -> Tuple[bool, str]:
        self.logger.info("开始初始化系统")
        
        success, message = self.engine.connect()
        if not success:
            return False, message
        
        try:
            from souren_config import DEVICE_TYPE
            if DEVICE_TYPE.lower() == 'board':
                print(f"📱 BOARD模式: 将在需要时初始化AT控制器")
        except:
            pass
        
        success, result = self.engine.execute_command("get_scv_files", folder_path=self.scv_folder)
        if not success:
            return False, f"检查SCV文件夹失败: {result.get('error', '未知错误')}"
        
        if not result.get("success"):
            return False, f"SCV文件夹错误: {result.get('error', '未知错误')}"
        
        file_count = result.get("count", 0)
        self.logger.info(f"系统初始化完成，找到 {file_count} 个SCV文件")
        return True, f"系统初始化完成，找到 {file_count} 个SCV文件"
    
    def list_scv_files(self) -> Tuple[bool, Dict]:
        self.logger.info("获取SCV文件列表")
        
        success, result = self.engine.execute_command("get_scv_files", folder_path=self.scv_folder)
        return success, result
    
    def load_scv_file(self, file_path: str) -> Tuple[bool, Dict]:
        if not file_path:
            return False, {"error": "未提供文件路径"}
        
        self.current_file = os.path.basename(file_path)
        
        self.logger.info(f"加载SCV文件: {self.current_file}")
        self.logger.info(f"文件完整路径: {file_path}")
        
        if not os.path.exists(file_path):
            print(f"❌ 文件不存在: {file_path}")
            return False, {"error": f"文件不存在: {file_path}"}
        
        success, result = self.engine.execute_command("load_scv_file", file_path=file_path)
        return success, result
    
    def connect_device(self) -> Tuple[bool, Dict]:
        """动态获取并连接设备"""
        self.logger.info("动态获取并连接设备")
        
        success, result = self.engine.execute_command("select_and_connect_device")
        
        if success and result.get("success"):
            self.selected_device_info = result.get("device")
        
        return success, result
    
    def execute_tests(self) -> Tuple[bool, Dict]:
        """执行测试"""
        if not self.current_file:
            return False, {"error": "未加载SCV文件"}
        
        if not self.selected_device_info:
            return False, {"error": "未连接设备"}
        
        self.logger.info(f"开始执行测试 - 模式: {self.execution_mode}")
        
        try:
            if self.execution_mode == 'loop_info':
                success, result = self.engine.execute_command("run_tests", 
                                                          mode=self.execution_mode, 
                                                          loop_count=self.loop_count)
            else:
                success, result = self.engine.execute_command("run_tests", mode=self.execution_mode)
            
            # ⭐⭐⭐ 检查是否被中断
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
                    "timestamp": datetime.now().isoformat()
                }
                
                if self.execution_mode == 'loop_info':
                    execution_data["loop_count"] = self.loop_count
                
                self.result_saver.save_result(execution_data)
                self.execution_history.append(execution_data)
            
            return success, result
            
        except KeyboardInterrupt:
            print(f"\n⏹️  测试执行过程中被用户中断")
            self.interrupted = True
            return False, {"error": "用户中断测试执行", "interrupted": True}
        
    def run_complete_workflow(self, scv_file_path: str = None) -> Tuple[bool, Dict]:
        self.logger.info("开始完整工作流程执行")
        
        workflow_result = {
            "workflow": "complete",
            "steps": [],
            "overall_success": False,
            "start_time": datetime.now().isoformat()
        }
        
        # 步骤1: 初始化
        success, message = self.initialize()
        workflow_result["steps"].append({
            "step": 1,
            "name": "initialize",
            "success": success,
            "message": message
        })
        
        if not success:
            workflow_result["overall_success"] = False
            workflow_result["end_time"] = datetime.now().isoformat()
            workflow_result["error"] = message
            return False, workflow_result
        
        # 步骤2: 加载文件（使用提供的文件路径）
        if not scv_file_path:
            workflow_result["overall_success"] = False
            workflow_result["end_time"] = datetime.now().isoformat()
            workflow_result["error"] = "未提供SCV文件路径"
            return False, workflow_result
        
        success, result = self.load_scv_file(scv_file_path)
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
        
        # 步骤3: 连接设备
        success, device_result = self.connect_device()
        workflow_result["steps"].append({
            "step": 3,
            "name": "connect_device",
            "success": success,
            "result": device_result
        })
        
        if not success:
            workflow_result["overall_success"] = False
            workflow_result["end_time"] = datetime.now().isoformat()
            workflow_result["error"] = device_result.get("error", "连接设备失败")
            return False, workflow_result
        
        if not device_result.get("success"):
            workflow_result["overall_success"] = False
            workflow_result["end_time"] = datetime.now().isoformat()
            workflow_result["error"] = device_result.get("error", "未找到可用设备")
            return False, workflow_result
        
        # 步骤4: 执行测试
        success, test_result = self.execute_tests()
        workflow_result["steps"].append({
            "step": 4,
            "name": "execute_tests",
            "success": success,
            "result": test_result
        })
        
        # ⭐⭐⭐ 修改：即使中断也要设置整体成功标志，确保数据能保存
        if self.interrupted:
            # 用户中断，但可能有部分数据
            workflow_result["overall_success"] = True  # 设置为True让数据能保存
            workflow_result["interrupted"] = True
            workflow_result["status"] = "interrupted"
            workflow_result["message"] = "测试被用户中断，但已保存已执行的数据"
        else:
            workflow_result["overall_success"] = success
        
        workflow_result["end_time"] = datetime.now().isoformat()
        
        if success or self.interrupted:  # ⭐⭐⭐ 修改：中断时也处理结果
            workflow_result["final_result"] = test_result
            if self.interrupted:
                self.logger.info("完整工作流程被用户中断，但已保存部分数据")
            else:
                self.logger.info("完整工作流程执行成功")
            
            print("\n" + "=" * 60)
            print("📤 正在自动导出结果到Excel文件...")
            print("=" * 60)
            
            # ⭐⭐⭐ 关键：即使中断也要导出Excel ⭐⭐⭐
            excel_success = self._auto_export_to_excel(test_result)
            if excel_success:
                workflow_result["excel_export"] = "成功"
                print("✅ Excel文件导出成功!")
            else:
                workflow_result["excel_export"] = "失败"
                print("⚠️  Excel文件导出失败,但测试已完成")
            
            print("=" * 60)
        else:
            workflow_result["error"] = test_result.get("error", "执行测试失败")
            self.logger.error("完整工作流程执行失败")
        
        # ⭐⭐⭐ 关键：即使中断也要保存结果 ⭐⭐⭐
        self.result_saver.save_result(workflow_result)
        
        return success, workflow_result
    
    def _auto_export_to_excel(self, test_result: Dict) -> bool:
        """自动将结果导出到Excel"""
        try:
            from souren_exporter import ResultExporter
            
            exporter = ResultExporter()
            
            json_file = exporter.find_latest_json_result()
            if json_file is None:
                print("❌ 未找到JSON结果文件，无法导出Excel")
                return False
            
            print(f"📁 找到结果文件: {os.path.basename(json_file)}")
            
            success = exporter.convert_to_excel(json_file)
            return success
            
        except ImportError as e:
            print(f"❌ 导出模块导入失败: {e}")
            print("💡 请安装所需依赖: pip install pandas openpyxl")
            return False
        except Exception as e:
            print(f"❌ 自动导出Excel失败: {e}")
            import traceback
            traceback.print_exc()
            return False

# ==============================================
# ↓↓↓ 显示工具函数
# ==============================================
def display_execution_result(result: Dict):
    """显示执行结果"""
    print("\n" + "=" * 60)
    
    # ⭐⭐⭐ 修改：支持显示中断状态
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
    
    print("=" * 60)