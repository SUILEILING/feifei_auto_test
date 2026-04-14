from lib.var import *
from souren_core import (
    DirectCommandExecutor,
    SourenLogger,
    SourenResultSaver,
    PythonScriptExecutor
)

try:
    import common
    print("✅ common模块导入成功")
except ImportError as e:
    print(f"⚠️  导入common模块失败: {e}")
    common = None

class ExecutionStatus(Enum):
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    STOPPED = "stopped"
    CANCELLED = "cancelled"
    INTERRUPTED = "interrupted"

class TestMonitor:
    def __init__(self):
        self.should_stop = False
        self.is_cancelled = False
        self.is_interrupted = False
        self.start_time = None
    def start(self):
        self.start_time = time.time()
        self.should_stop = False
        self.is_cancelled = False
        self.is_interrupted = False
    def stop(self): self.should_stop = True
    def cancel(self): self.is_cancelled = True; self.should_stop = True
    def interrupt(self): self.is_interrupted = True; self.should_stop = True; print("\n⏹️  用户中断测试执行")

class SourenEngine:
    def __init__(self, result_dir=None, script_name=None, params=None):
        self.result_dir = result_dir
        self.script_name = script_name
        self.params = params
        self.logger = SourenLogger()
        self.result_saver = SourenResultSaver(result_dir, script_name)
        self.script_executor = PythonScriptExecutor()
        self.connection_status = False
        self.last_error = ""
        self.selected_device = None
        self.test_monitor = None
        self.current_file = None
        self.current_file_path = None
        self.execution_start_time = None
        self.interrupted = False
        self.current_loop_iteration = 1
        self.total_loop_count = 1
        self.extracted_data = []
        self.execution_context = {}

    def get_result_file(self): return self.result_saver.get_result_file()
    def get_result_dir(self): return self.result_saver.get_result_dir()

    def initialize(self) -> Tuple[bool, str]:
        print("初始化仪器连接...")
        try:
            success = DirectCommandExecutor.initialize()
            if success:
                self.connection_status = True
                print("✅ 连接成功")
                return True, "连接成功"
            else:
                self.connection_status = False
                self.last_error = "仪器连接失败"
                print("❌ 仪器连接失败")
                return False, "仪器连接失败"
        except Exception as e:
            error_msg = f"连接失败: {str(e)}"
            self.connection_status = False
            self.last_error = error_msg
            print(f"❌ {error_msg}")
            return False, error_msg

    def execute_command(self, command: str, **kwargs) -> Tuple[bool, Dict]:
        if not self.connection_status:
            return False, {"error": "未连接到仪器"}
        try:
            if command == "load_python_file":
                result = self._load_python_file(**kwargs)
                return True, result
            elif command == "run_tests":
                mode = kwargs.get('mode', 'run_all')
                loop_count = kwargs.get('loop_count', 1)
                loop_iteration = kwargs.get('loop_iteration', 1)
                result = self._run_tests_with_souren(mode, loop_count, loop_iteration)
                return True, result
            else:
                return False, {"error": f"未知命令: {command}"}
        except Exception as e:
            error_msg = f"执行命令失败: {str(e)}"
            print(f"❌ {error_msg}")
            return False, {"error": error_msg}

    def _load_python_file(self, file_path: str) -> Dict:
        print(f"加载Python文件: {os.path.basename(file_path)}")
        try:
            if not os.path.exists(file_path):
                return {"success": False, "error": f"文件不存在: {file_path}"}
            self.current_file = os.path.basename(file_path)
            self.current_file_path = file_path
            return {"success": True, "file": file_path, "message": "文件加载成功", "size": os.path.getsize(file_path)}
        except Exception as e:
            error_msg = f"加载文件失败: {str(e)}"
            print(f"❌ {error_msg}")
            return {"success": False, "error": error_msg}

    def _run_tests_with_souren(self, mode: str = 'run_all', loop_count: int = 1, loop_iteration: int = 1) -> Dict:
        self.logger.info(f"开始运行测试 - 模式: {mode}, 循环: {loop_iteration}/{loop_count}")
        if not self.current_file_path:
            return {"success": False, "error": "未加载Python文件"}
        from souren_config import LOOP_COUNT as CFG_LOOP_COUNT, EXECUTION_MODE, INSTRUMENT_ADDRESS
        if loop_count == 1 and CFG_LOOP_COUNT > 1 and EXECUTION_MODE == 'loop_info':
            loop_count = CFG_LOOP_COUNT
        current_file_name = os.path.basename(self.current_file_path)
        self.execution_start_time = time.time()
        self.test_monitor = TestMonitor()
        try:
            device_name = INSTRUMENT_ADDRESS
            print(f"\n{'='*60}")
            print(f"🚀 开始运行测试 (第 {loop_iteration}/{loop_count} 次循环)")
            print(f"{'='*60}")
            print(f"📄 文件: {current_file_name}")
            print(f"🔌 设备: {device_name}")
            print(f"🎯 模式: {mode}")
            print(f"🔄 循环: {loop_iteration}/{loop_count}")
            print(f"📋 参数配置: {self.params}")
            print(f"📊 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"{'='*60}")
            self.test_monitor.start()
            self.script_executor.reset()
            self.script_executor.set_loop_info(loop_iteration, loop_count)
            success, result = self.script_executor.execute_script(
                self.current_file_path, self.params,
                loop_iteration=loop_iteration, total_loop_count=loop_count
            )
            if success:
                execution_details = result.get("execution_details", [])
                extracted_data = result.get("extracted_data", [])
                print(f"📊 本次循环执行详情数量: {len(execution_details)}")
                print(f"📊 本次循环提取数据数量: {len(extracted_data)}")
                self.extracted_data = extracted_data
                final_result = self._generate_final_result(
                    execution_details, device_name, "normal", loop_iteration, loop_count
                )
                final_result["success"] = True
                final_result["extracted_data"] = extracted_data
                final_result["file"] = current_file_name
                final_result["parameters"] = self.params
                return final_result
            else:
                return {"success": False, "error": result.get("error", "执行脚本失败"),
                        "loop_iteration": loop_iteration, "loop_count": loop_count}
        except Exception as e:
            error_msg = f"运行测试失败: {str(e)}"
            print(f"❌ {error_msg}")
            self.logger.error(f"运行测试失败: {error_msg}")
            if self.test_monitor:
                self.test_monitor.stop()
            return {"success": False, "error": error_msg,
                    "loop_iteration": loop_iteration, "loop_count": loop_count}

    def _generate_final_result(self, execution_details: List, device_name: str, mode: str,
                               loop_iteration: int, loop_count: int) -> Dict:
        execution_time = time.time() - self.execution_start_time if self.execution_start_time else 0
        executed_steps = len(execution_details)
        passed_steps = len([d for d in execution_details if d.get('status') == 'success'])
        failed_steps = len([d for d in execution_details if d.get('status') == 'failed'])
        success_rate = round((passed_steps / executed_steps) * 100, 2) if executed_steps > 0 else 0
        result = {
            "success": True, "mode": mode, "device": device_name,
            "file": self.current_file, "status": "completed",
            "message": f"测试执行完成，模式: {mode}",
            "executed_steps": executed_steps, "passed": passed_steps, "failed": failed_steps,
            "communication_failed": 0, "success_rate": success_rate,
            "execution_time": execution_time, "duration": execution_time,
            "start_time": datetime.fromtimestamp(self.execution_start_time).strftime('%Y-%m-%d %H:%M:%S') if self.execution_start_time else "",
            "end_time": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "execution_details": execution_details, "extracted_data": self.extracted_data,
            "interrupted": self.interrupted,
            "loop_iteration": loop_iteration, "loop_count": loop_count
        }
        return result

    def cleanup(self):
        DirectCommandExecutor.cleanup()
        self.connection_status = False
        self.current_file = None
        self.current_file_path = None
        self.selected_device = None
        if hasattr(self, 'execution_context'):
            self.execution_context.clear()
        else:
            self.execution_context = {}
        self.extracted_data.clear()
        self.current_loop_iteration = 1
        self.total_loop_count = 1