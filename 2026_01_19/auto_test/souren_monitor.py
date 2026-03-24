from lib.var import *

# 从 souren_core 导入必要的类和函数
from souren_core import (
    DirectCommandExecutor, 
    SourenLogger, 
    SourenResultSaver, 
    SCVParser
)

# ==============================================
# ↓↓↓ 执行状态枚举
# ==============================================
class ExecutionStatus(Enum):
    """执行状态"""
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    STOPPED = "stopped"
    CANCELLED = "cancelled"

# ==============================================
# ↓↓↓ 简化版测试监控器（移除超时机制）
# ==============================================
class TestMonitor:
    """测试监控器 """
    
    def __init__(self):
        self.should_stop = False
        self.is_cancelled = False
        self.start_time = None
        
    def start(self):
        """启动监控器"""
        self.start_time = time.time()
        self.should_stop = False
        self.is_cancelled = False
        
        print(f"⏰ 监控器已启动（无超时限制）")
    
    def stop(self):
        """停止监控器"""
        self.should_stop = True
    
    def cancel(self):
        """取消测试"""
        self.is_cancelled = True
        self.should_stop = True
    
    def should_test_stop(self) -> Tuple[bool, str]:
        """检查测试是否应该停止"""
        if self.should_stop:
            if self.is_cancelled:
                return True, "用户取消"
            else:
                return True, "用户停止"
        
        return False, ""
    
    def get_monitor_info(self) -> Dict:
        """获取监控器信息"""
        if not self.start_time:
            return {
                "status": "not_started",
                "should_stop": self.should_stop,
                "is_cancelled": self.is_cancelled
            }
        
        elapsed = time.time() - self.start_time
        
        status = "running"
        if self.is_cancelled:
            status = "cancelled"
        elif self.should_stop:
            status = "stopped"
        
        return {
            "start_time": datetime.fromtimestamp(self.start_time).strftime('%Y-%m-%d %H:%M:%S'),
            "elapsed_seconds": round(elapsed, 1),
            "should_stop": self.should_stop,
            "is_cancelled": self.is_cancelled,
            "status": status
        }

# ==============================================
# ↓↓↓ Souren.ToolSet 核心功能类
# ==============================================
class SourenEngine:
    """Souren.ToolSet 核心引擎"""
    
    def __init__(self):
        self.logger = SourenLogger()
        self.result_saver = SourenResultSaver()
        self.scv_parser = SCVParser()
        self.connection_status = False
        self.last_error = ""
        self.selected_device = None
        self.test_monitor = None
        self.current_file = None
        self.current_file_path = None
        self.execution_start_time = None
    
    def connect(self) -> Tuple[bool, str]:
        print("尝试连接到仪器")
        
        try:
            success = DirectCommandExecutor.initialize()
            
            if success:
                self.connection_status = True
                message = "连接成功"
                print(f"✅ {message}")
                return True, message
            else:
                error_msg = "仪器连接失败"
                self.connection_status = False
                self.last_error = error_msg
                print(f"❌ {error_msg}")
                return False, error_msg
            
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
            if command == "get_scv_files":
                result = self._get_scv_files(**kwargs)
            elif command == "load_scv_file":
                result = self._load_scv_file(**kwargs)
            elif command == "select_and_connect_device":
                result = self._select_and_connect_device(**kwargs)
            elif command == "run_tests":
                mode = kwargs.get('mode', 'run_all')
                loop_count = kwargs.get('loop_count', 1)
                print(f"🔧 执行命令: run_tests - mode={mode}, loop_count={loop_count}")
                print(f"📁 当前文件: {self.current_file_path}")
                result = self._run_tests_with_souren(mode=mode, loop_count=loop_count)
            elif command == "parse_scv_file":
                result = self._parse_scv_file(**kwargs)
            else:
                return False, {"error": f"未知命令: {command}"}
            
            return True, result
            
        except Exception as e:
            error_msg = f"执行命令失败: {str(e)}"
            print(f"❌ {error_msg}")
            return False, {"error": error_msg}
        
    def _get_scv_files(self, folder_path: str) -> Dict:
        print(f"读取SCV文件夹: {folder_path}")
        
        try:
            import os
            
            if not os.path.exists(folder_path):
                return {
                    "success": False,
                    "error": f"文件夹不存在: {folder_path}",
                    "files": []
                }
            
            scv_files = []
            for filename in os.listdir(folder_path):
                if filename.lower().endswith('.scv'):
                    filepath = os.path.join(folder_path, filename)
                    if os.path.isfile(filepath):
                        file_info = {
                            "name": filename,
                            "path": filepath,
                            "size": os.path.getsize(filepath),
                            "modified": datetime.fromtimestamp(
                                os.path.getmtime(filepath)
                            ).strftime('%Y-%m-%d %H:%M:%S')
                        }
                        
                        scv_files.append(file_info)
            
            return {
                "success": True,
                "folder": folder_path,
                "count": len(scv_files),
                "files": scv_files
            }
            
        except Exception as e:
            error_msg = f"读取文件列表失败: {str(e)}"
            print(f"❌ {error_msg}")
            return {
                "success": False,
                "error": error_msg,
                "files": []
            }
    
    def _load_scv_file(self, file_path: str) -> Dict:
        """加载SCV文件"""
        print(f"加载SCV文件: {os.path.basename(file_path)}")
        
        try:
            if not os.path.exists(file_path):
                return {
                    "success": False,
                    "error": f"文件不存在: {file_path}"
                }
            
            self.current_file = os.path.basename(file_path)
            self.current_file_path = file_path
            
            return {
                "success": True,
                "file": file_path,
                "message": "文件加载成功",
                "size": os.path.getsize(file_path)
            }
            
        except Exception as e:
            error_msg = f"加载文件失败: {str(e)}"
            print(f"❌ {error_msg}")
            return {
                "success": False,
                "error": error_msg
            }
    
    def _parse_scv_file(self, file_path: str) -> Dict:
        """解析SCV文件"""
        return self.scv_parser.parse_file(file_path)
    
    def _select_and_connect_device(self) -> Dict:
        print("动态获取并连接设备")
        
        try:
            from souren_config import INSTRUMENT_ADDRESS
            device_address = INSTRUMENT_ADDRESS
            print(f"连接设备: {device_address}")
            
            device_info = {
                "id": device_address,
                "name": device_address,
                "type": "TCPIP",
                "status": "online",
                "address": device_address
            }
            
            self.selected_device = device_info
            
            print(f"✅ 设备连接成功: {device_address}")
            
            return {
                "success": True,
                "device": device_info,
                "message": f"成功连接设备: {device_address}"
            }
            
        except Exception as e:
            error_msg = f"连接设备失败: {str(e)}"
            print(f"❌ {error_msg}")
            return {
                "success": False,
                "error": error_msg
            }
    
    def _run_tests_with_souren(self, mode: str = 'run_all', loop_count: int = 1) -> Dict:
        """使用仪器运行测试"""
        self.logger.info(f"开始运行测试 - 模式: {mode}")
        
        if not self.selected_device:
            return {
                "success": False,
                "error": "未连接设备，请先连接设备"
            }
        
        if not self.current_file_path:
            return {
                "success": False,
                "error": "未加载SCV文件"
            }
        
        from souren_config import SHOW_STEP_TYPE, SHOW_STEP_CONTENT, LOOP_COUNT, EXECUTION_MODE
        
        # 优先使用传入的 loop_count 参数
        if loop_count == 1 and LOOP_COUNT > 1 and EXECUTION_MODE == 'loop_info':
            loop_count = LOOP_COUNT
        
        self.execution_start_time = time.time()
        self.test_monitor = TestMonitor()
        
        try:
            device_name = self.selected_device.get("name")

            current_file_name = os.path.basename(self.current_file_path) if self.current_file_path else self.current_file
            print(f"\n{'='*60}")
            print(f"🚀 开始运行测试")
            print(f"{'='*60}")
            print(f"📄 SCV文件: {current_file_name}")
            print(f"🔌 设备: {device_name}")
            print(f"🎯 模式: {mode}")
            
            if mode == 'loop_info' and loop_count > 1:
                print(f"🔄 全局循环模式: 整个流程执行 {loop_count} 次")
                self.logger.info(f"全局循环模式激活: LOOP_COUNT = {loop_count}")
            else:
                print(f"🔄 单次执行模式")
                
            print(f"📊 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"{'='*60}")
            
            # 启动测试监控器
            self.test_monitor.start()
            
            # 解析SCV文件
            print(f"\n📋 解析SCV文件...")
            parse_result = self.scv_parser.parse_file(self.current_file_path)
            
            if not parse_result.get("success"):
                return {
                    "success": False,
                    "error": f"解析文件失败: {parse_result.get('error', '未知错误')}"
                }
            
            steps = parse_result.get("steps", [])
            total_steps = len(steps)
            
            if total_steps == 0:
                return {
                    "success": False,
                    "error": "文件中没有找到可执行的测试步骤"
                }
            
            print(f"✅ 解析完成，找到 {total_steps} 个测试步骤")
            self.logger.info(f"解析SCV文件完成,共 {total_steps} 个步骤")
            
            # 显示分析信息
            analysis = parse_result.get("analysis", {})
            if analysis:
                print(f"📊 步骤分析:")
                print(f"   Normal步骤: {analysis.get('normal_steps', 0)}")
                print(f"   CALL步骤: {analysis.get('call_steps', 0)}")
                print(f"   SLEEP步骤: {analysis.get('sleep_steps', 0)}")
            
            # 计算总执行次数
            total_executions = total_steps
            if mode == 'loop_info' and loop_count > 1:
                total_executions = total_steps * loop_count
                print(f"\n🔄 开始执行{total_executions}个测试步骤 ({total_steps}步骤 x {loop_count}次全局循环)...")
            else:
                print(f"\n🔄 开始执行{total_steps}个测试步骤...")
            
            print(f"{'='*60}")
            
            executed_steps = 0
            passed_steps = 0
            failed_steps = 0
            communication_failed_steps = 0
            execution_details = []
            
            if mode == 'loop_info' and loop_count > 1:
                # 全局循环模式
                for loop_index in range(loop_count):
                    loop_start_time = time.time()
                    print(f"\n🔄 第 {loop_index + 1}/{loop_count} 轮全局循环:")
                    print(f"{'-'*50}")
                    self.logger.info(f"开始第 {loop_index + 1}/{loop_count} 轮全局循环")
                    
                    for step in steps:
                        step_num = step.get("step", 0)
                        step_content = step.get("content", "")
                        step_type = step.get("type", "Normal")
                        has_loop = step.get("has_loop", False)
                        loop_config = step.get("loop_config", {})
                        
                        # 计算全局步骤编号
                        global_step_num = loop_index * total_steps + step_num
                        
                        # 检查是否应该停止
                        should_stop, stop_reason = self.test_monitor.should_test_stop()
                        if should_stop:
                            print(f"\n⏹️  测试被{stop_reason}")
                            self.logger.warning(f"测试被停止: {stop_reason}")
                            break
                        
                        # 显示步骤信息
                        step_info_msg = f"步骤 {global_step_num}/{total_executions}"
                        loop_info = f"第{loop_index + 1}/{loop_count}轮循环 - 原始步骤{step_num}"
                        step_info_msg += f" ({loop_info})"
                        
                        print(f"\n📝 {step_info_msg}:")
                        self.logger.info(f"开始执行: {step_info_msg} - 内容: {step_content}")
                        
                        if SHOW_STEP_TYPE:
                            print(f"   🔧 类型: {step_type}")
                        
                        if SHOW_STEP_CONTENT:
                            print(f"   📄 内容: {step_content}")
                        
                        # 记录步骤开始时间
                        step_start_time = time.time()
                        
                        try:
                            # 执行步骤
                            if has_loop and loop_config.get('enable', 'false').lower() == 'true':
                                step_success, step_result, loop_details = self._execute_loop_step_with_expected_result(
                                    step, device_name, loop_config
                                )
                            else:
                                step_success, step_result = self._execute_single_step(step, device_name)
                                loop_details = []
                            
                            from souren_config import COMMUNICATION_ERROR_PASS_COMMANDS
                            is_special_command = step_content in COMMUNICATION_ERROR_PASS_COMMANDS
                            
                            is_communication_error = self._is_communication_error(step_result)
                            
                            if is_communication_error:
                                print(f"   ⚠️  仪器通信错误")
                                self.logger.warning(f"仪器通信错误: {step_result}")
                                
                                # 如果是特殊命令，通信错误视为通过
                                if is_special_command:
                                    print(f"   ⭐  '{step_content}' 是特殊命令，通信错误视为通过")
                                    self.logger.info(f"特殊命令 '{step_content}' 通信错误视为通过")
                                    passed_steps += 1
                                    step_status = "success"
                                    status_icon = "✅"
                                    status_text = "通过"
                                    step_execution_success = True
                                    step_success = True
                                else:
                                    communication_failed_steps += 1
                                    step_status = "communication_error"
                                    status_icon = "⚠️"
                                    status_text = "通信错误"
                                    step_execution_success = True
                                    step_success = True
                            elif step_success:
                                passed_steps += 1
                                step_status = "success"
                                status_icon = "✅"
                                status_text = "通过"
                                step_execution_success = True
                                self.logger.info(f"步骤执行成功: {step_info_msg} - 结果: {step_result[:100] if step_result else '无结果'}")
                            else:
                                failed_steps += 1
                                step_status = "failed"
                                status_icon = "❌"
                                status_text = "失败"
                                step_execution_success = False
                                self.logger.warning(f"步骤执行失败: {step_info_msg} - 错误: {step_result}")
                            
                            # 记录步骤结束时间
                            step_end_time = time.time()
                            step_duration = step_end_time - step_start_time
                            
                            executed_steps += 1
                            
                            print(f"   {status_icon} 执行{status_text} - 耗时: {step_duration:.2f}秒")
                            
                            # 显示结果
                            if step_result and step_result != "命令执行成功":
                                print(f"   📊 结果: {step_result}")
                            
                            # 保存步骤执行详情
                            execution_detail = {
                                "step": global_step_num,
                                "original_step": step_num,
                                "content": step_content,
                                "type": step_type,
                                "status": step_status,
                                "duration": step_duration,
                                "result": step_result,
                                "start_time": step_start_time,
                                "end_time": step_end_time,
                                "has_loop": has_loop,
                                "loop_index": loop_index + 1,
                                "total_loops": loop_count
                            }
                            
                            execution_details.append(execution_detail)
                            
                            # 如果步骤失败且配置了失败停止（通信错误除外）
                            if not step_execution_success and step.get("unloop_pass_result_stop") == "True":
                                print(f"   ⚠️  检测到失败停止配置，停止后续步骤执行")
                                self.logger.warning(f"检测到失败停止配置，停止后续执行")
                                break
                            
                        except Exception as e:
                            step_end_time = time.time()
                            step_duration = step_end_time - step_start_time
                            
                            failed_steps += 1
                            executed_steps += 1
                            
                            print(f"   ❌ 执行异常 - 耗时: {step_duration:.2f}秒")
                            print(f"     错误: {str(e)}")
                            self.logger.error(f"步骤执行异常: {step_info_msg} - 异常: {str(e)}")
                            
                            execution_details.append({
                                "step": global_step_num,
                                "original_step": step_num,
                                "content": step_content,
                                "type": step_type,
                                "status": "error",
                                "duration": step_duration,
                                "result": f"执行异常: {str(e)}",
                                "start_time": step_start_time,
                                "end_time": step_end_time,
                                "loop_index": loop_index + 1,
                                "total_loops": loop_count
                            })
                        
                        # 更新进度
                        elapsed_total = time.time() - self.execution_start_time
                        progress = (global_step_num / total_executions) * 100
                        
                        print(f"   📊 进度: {global_step_num}/{total_executions} ({progress:.1f}%)")
                        print(f"     已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
                        print(f"     总耗时: {elapsed_total:.1f}秒")
                        print(f"{'-'*60}")
                        
                        self.logger.info(f"进度更新: {global_step_num}/{total_executions} ({progress:.1f}%) - 已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
                    
                    # 检查是否应该停止
                    should_stop, stop_reason = self.test_monitor.should_test_stop()
                    if should_stop:
                        break
                    
                    loop_end_time = time.time()
                    loop_duration = loop_end_time - loop_start_time
                    print(f"✅ 第 {loop_index + 1}/{loop_count} 轮循环完成，耗时: {loop_duration:.1f}秒")
                    self.logger.info(f"第 {loop_index + 1}/{loop_count} 轮循环完成，耗时: {loop_duration:.1f}秒")
            
            else:
                # 单次执行模式
                for step in steps:
                    step_num = step.get("step", 0)
                    step_content = step.get("content", "")
                    step_type = step.get("type", "Normal")
                    has_loop = step.get("has_loop", False)
                    loop_config = step.get("loop_config", {})
                    
                    # 检查是否应该停止
                    should_stop, stop_reason = self.test_monitor.should_test_stop()
                    if should_stop:
                        print(f"\n⏹️  测试被{stop_reason}")
                        self.logger.warning(f"测试被停止: {stop_reason}")
                        break
                    
                    # 显示步骤信息
                    step_info_msg = f"步骤 {step_num}/{total_steps}"
                    print(f"\n📝 {step_info_msg}:")
                    self.logger.info(f"开始执行: {step_info_msg} - 内容: {step_content}")
                    
                    if SHOW_STEP_TYPE:
                        print(f"   🔧 类型: {step_type}")
                    
                    if SHOW_STEP_CONTENT:
                        print(f"   📄 内容: {step_content}")
                    
                    # 记录步骤开始时间
                    step_start_time = time.time()
                    
                    try:
                        # 执行步骤
                        if has_loop and loop_config.get('enable', 'false').lower() == 'true':
                            step_success, step_result, loop_details = self._execute_loop_step_with_expected_result(
                                step, device_name, loop_config
                            )
                        else:
                            step_success, step_result = self._execute_single_step(step, device_name)
                            loop_details = []
                        
                        # 检查是否是特殊命令
                        from souren_config import COMMUNICATION_ERROR_PASS_COMMANDS
                        is_special_command = step_content in COMMUNICATION_ERROR_PASS_COMMANDS
                        
                        # 检查是否为通信错误
                        is_communication_error = self._is_communication_error(step_result)
                        
                        if is_communication_error:
                            print(f"   ⚠️  仪器通信错误")
                            self.logger.warning(f"仪器通信错误: {step_result}")
                            
                            # 如果是特殊命令，通信错误视为通过
                            if is_special_command:
                                print(f"   ⭐  '{step_content}' 是特殊命令，通信错误视为通过")
                                self.logger.info(f"特殊命令 '{step_content}' 通信错误视为通过")
                                passed_steps += 1
                                step_status = "success"
                                status_icon = "✅"
                                status_text = "通过"
                                step_execution_success = True
                                step_success = True
                            else:
                                communication_failed_steps += 1
                                step_status = "communication_error"
                                status_icon = "⚠️"
                                status_text = "通信错误"
                                step_execution_success = True
                                step_success = True
                        elif step_success:
                            passed_steps += 1
                            step_status = "success"
                            status_icon = "✅"
                            status_text = "通过"
                            step_execution_success = True
                            self.logger.info(f"步骤执行成功: {step_info_msg} - 结果: {step_result[:100] if step_result else '无结果'}")
                        else:
                            failed_steps += 1
                            step_status = "failed"
                            status_icon = "❌"
                            status_text = "失败"
                            step_execution_success = False
                            self.logger.warning(f"步骤执行失败: {step_info_msg} - 错误: {step_result}")
                        
                        # 记录步骤结束时间
                        step_end_time = time.time()
                        step_duration = step_end_time - step_start_time
                        
                        executed_steps += 1
                        
                        print(f"   {status_icon} 执行{status_text} - 耗时: {step_duration:.2f}秒")
                        
                        # 显示结果
                        if step_result and step_result != "命令执行成功":
                            print(f"   📊 结果: {step_result}")
                    
                        # 保存步骤执行详情
                        execution_detail = {
                            "step": step_num,
                            "content": step_content,
                            "type": step_type,
                            "status": step_status,
                            "duration": step_duration,
                            "result": step_result,
                            "start_time": step_start_time,
                            "end_time": step_end_time,
                            "has_loop": has_loop,
                        }
                        
                        execution_details.append(execution_detail)
                        
                        # 如果步骤失败且配置了失败停止（通信错误除外）
                        if not step_execution_success and step.get("unloop_pass_result_stop") == "True":
                            print(f"   ⚠️  检测到失败停止配置，停止后续步骤执行")
                            self.logger.warning(f"检测到失败停止配置，停止后续执行")
                            break
                        
                    except Exception as e:
                        step_end_time = time.time()
                        step_duration = step_end_time - step_start_time
                        
                        failed_steps += 1
                        executed_steps += 1
                        
                        print(f"   ❌ 执行异常 - 耗时: {step_duration:.2f}秒")
                        print(f"     错误: {str(e)}")
                        self.logger.error(f"步骤执行异常: {step_info_msg} - 异常: {str(e)}")
                        
                        execution_details.append({
                            "step": step_num,
                            "content": step_content,
                            "type": step_type,
                            "status": "error",
                            "duration": step_duration,
                            "result": f"执行异常: {str(e)}",
                            "start_time": step_start_time,
                            "end_time": step_end_time
                        })
                    
                    # 更新进度
                    elapsed_total = time.time() - self.execution_start_time
                    progress = (step_num / total_steps) * 100
                    
                    print(f"   📊 进度: {step_num}/{total_steps} ({progress:.1f}%)")
                    print(f"     已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
                    print(f"     总耗时: {elapsed_total:.1f}秒")
                    print(f"{'-'*60}")
                    
                    self.logger.info(f"进度更新: {step_num}/{total_steps} ({progress:.1f}%) - 已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
            
            # 停止监控器
            self.test_monitor.stop()
            
            # 计算总执行时间
            execution_time = time.time() - self.execution_start_time
            
            # 确定最终状态
            should_stop, stop_reason = self.test_monitor.should_test_stop()
            if should_stop:
                if stop_reason == "用户取消":
                    status = ExecutionStatus.CANCELLED
                    message = f"测试被用户取消 (设备: {device_name})"
                else:
                    status = ExecutionStatus.STOPPED
                    message = f"测试已结束 (设备: {device_name})"
            elif executed_steps == 0:
                status = ExecutionStatus.FAILED
                message = f"没有步骤被执行 (设备: {device_name})"
            elif failed_steps == 0:
                status = ExecutionStatus.SUCCESS
                message = f"测试执行完成，全部通过 (设备: {device_name})"
                if communication_failed_steps > 0:
                    message = f"测试执行完成，{passed_steps}个步骤通过，{communication_failed_steps}个步骤通信错误 (设备: {device_name})"
            else:
                status = ExecutionStatus.FAILED
                message = f"测试执行完成，有{failed_steps}个步骤失败 (设备: {device_name})"
                if communication_failed_steps > 0:
                    message += f",{communication_failed_steps}个步骤通信错误"
            
            success = (status == ExecutionStatus.SUCCESS)
            
            # 构建结果
            result = {
                "success": success,
                "mode": mode,
                "device": device_name,
                "file": self.current_file,
                "status": status.value,
                "message": message,
                "stop_reason": stop_reason if should_stop else "正常完成",
                "original_total_steps": total_steps,
                "total_executions": total_executions,
                "executed_steps": executed_steps,
                "passed": passed_steps,
                "failed": failed_steps,
                "communication_failed": communication_failed_steps,
                "execution_time": execution_time,
                "duration": execution_time,
                "start_time": datetime.fromtimestamp(self.execution_start_time).strftime('%Y-%m-%d %H:%M:%S'),
                "end_time": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "monitor_info": self.test_monitor.get_monitor_info(),
                "execution_details": execution_details,
                "parsed_steps": steps
            }
            
            if mode == 'loop_info':
                result["loop_count"] = loop_count
                result["execution_mode"] = "全局循环"
            
            # 记录最终结果到日志
            self.logger.info(f"测试执行完成 - 状态: {status.value}, 消息: {message}")
            self.logger.info(f"统计信息: 原始步骤{total_steps}, 总执行次数{total_executions}, 已执行{executed_steps}, 通过{passed_steps}, 失败{failed_steps}, 通信错误{communication_failed_steps}")
            
            # 保存结果到JSON文件
            self.result_saver.save_result(result)
            
            # 显示最终结果
            print(f"\n{'='*60}")
            print(f"🎯 测试执行完成")
            print(f"{'='*60}")
            print(f"📅 开始时间: {result['start_time']}")
            print(f"📅 结束时间: {result['end_time']}")
            print(f"⏱️  总耗时: {execution_time:.1f}秒 ({execution_time/60:.1f}分钟)")
            print(f"📊 状态: {message}")
            print(f"📈 步骤统计:")
            print(f"   原始步骤: {total_steps}")
            print(f"   总执行次数: {total_executions}")
            print(f"   已执行: {executed_steps}")
            print(f"   通过: {passed_steps}")
            print(f"   失败: {failed_steps}")
            print(f"   通信错误: {communication_failed_steps}")
            if executed_steps > 0:
                success_rate = (passed_steps/executed_steps*100)
                print(f"   成功率: {success_rate:.1f}%")
            
            if mode == 'loop_info' and loop_count > 1:
                print(f"🔄 循环信息: {loop_count}次全局循环，共执行 {executed_steps} 次")
                self.logger.info(f"循环模式统计: {loop_count}次全局循环，共执行 {executed_steps} 次，成功率: {success_rate:.1f}%")
            
            print(f"{'='*60}")
            
            return result
            
        except Exception as e:
            error_msg = f"运行测试失败: {str(e)}"
            print(f"❌ {error_msg}")
            self.logger.error(f"运行测试失败: {error_msg}")
            
            # 确保监控器停止
            if self.test_monitor:
                self.test_monitor.stop()
            
            return {
                "success": False,
                "error": error_msg
            }
    
    def _execute_single_step(self, step: Dict, device_name: str) -> Tuple[bool, str]:
        """执行单个测试步骤 - 修复版本，确保所有命令都执行"""
        step_content = step.get("content", "").strip()
        
        if not step_content:
            return False, "空的命令内容"
        
        # 使用DirectCommandExecutor执行命令
        success, result = DirectCommandExecutor.execute_command(step_content)
        
        # 记录执行日志
        if success:
            self.logger.info(f"命令执行成功: '{step_content}' - 结果: {result[:100] if result else '无结果'}")
        else:
            self.logger.warning(f"命令执行失败: '{step_content}' - 错误: {result}")
        
        return success, result
    
    def _execute_loop_step_with_expected_result(self, step: Dict, device_name: str, loop_config: Dict) -> Tuple[bool, str, List]:
        """执行循环步骤 - 支持预期结果检查"""
        step_content = step.get("content", "")
        
        # 使用 loop_config 中的配置
        loop_times = int(loop_config.get('looptimes', 1))
        loop_sleep_ms = int(loop_config.get('sleepms', 0))
        expected_result = loop_config.get('looppassresult', '')
        
        # 清理期望结果中的引号
        if expected_result:
            expected_result = expected_result.replace('"', '').replace('&quot;', '').replace("'", '').strip()
        
        print(f"   🔄 开始执行循环，最大次数: {loop_times}, 期望结果: '{expected_result}'")
        self.logger.info(f"开始执行循环步骤: 最大次数={loop_times}, 期望结果='{expected_result}'")
        
        loop_details = []
        found_expected = False
        communication_error_occurred = False
        final_result = ""
        actual_iterations = 0
        
        # 检查是否是特殊命令（通信错误也视为通过）
        from souren_config import COMMUNICATION_ERROR_PASS_COMMANDS
        is_special_command = step_content in COMMUNICATION_ERROR_PASS_COMMANDS
        
        for i in range(loop_times):
            # 检查是否应该停止
            should_stop, stop_reason = self.test_monitor.should_test_stop()
            if should_stop:
                self.logger.warning(f"循环被停止: {stop_reason}")
                return True, f"循环被{stop_reason}", loop_details
            
            actual_iterations += 1
            iteration_start = time.time()
            print(f"     迭代 {i+1}/{loop_times}: ", end="")
            self.logger.info(f"循环迭代 {i+1}/{loop_times}")
            
            # 执行单个迭代
            success, result = self._execute_single_step(step, device_name)
            
            iteration_end = time.time()
            iteration_duration = iteration_end - iteration_start
            
            loop_details.append({
                "iteration": i + 1,
                "success": success,
                "result": result,
                "duration": iteration_duration,
                "start_time": iteration_start,
                "end_time": iteration_end
            })
            
            # 检查是否为通信错误
            is_communication_error = self._is_communication_error(result)
            
            if is_communication_error:
                print(f"⚠️  仪器通信错误 ({iteration_duration:.2f}秒)")
                print(f"      错误: {result}")
                self.logger.warning(f"循环迭代 {i+1} 仪器通信错误 - 错误: {result}")
                
                final_result = result
                communication_error_occurred = True
                
                # 如果是特殊命令，通信错误也视为通过
                if is_special_command:
                    print(f"      ⭐  '{step_content}' 是特殊命令，通信错误视为通过")
                    self.logger.info(f"特殊命令 '{step_content}' 通信错误视为通过")
                    found_expected = True
                    break
                else:
                    print(f"      ⚠️  检测到仪器通信错误，停止循环")
                    self.logger.info(f"检测到仪器通信错误，停止循环")
                    break
            
            elif success:
                print(f"✅ 命令执行成功 ({iteration_duration:.2f}秒)")
                self.logger.info(f"循环迭代 {i+1} 命令执行成功 - 耗时: {iteration_duration:.2f}秒")
                
                # 检查是否符合期望结果
                if expected_result:
                    # 清理结果中的引号
                    clean_result = result.replace('"', '').replace('&quot;', '').replace("'", '').strip()
                    clean_expected = expected_result.replace('"', '').replace('&quot;', '').replace("'", '').strip()
                    
                    print(f"      实际结果: '{clean_result}'")
                    print(f"      期望结果: '{clean_expected}'")
                    
                    if clean_expected and clean_expected in clean_result:
                        print(f"      🎯 符合期望结果: '{clean_expected}'")
                        self.logger.info(f"循环迭代 {i+1} 符合期望结果: '{clean_expected}'")
                        found_expected = True
                        final_result = result
                        
                        # 找到期望结果，停止循环
                        print(f"      ⏹  找到期望结果 '{clean_expected}'，停止循环")
                        self.logger.info(f"找到期望结果 '{clean_expected}'，停止循环")
                        break
                    else:
                        final_result = result
                        print(f"      ❌ 不符合期望结果，继续循环...")
                        self.logger.info(f"循环迭代 {i+1} 不符合期望结果: 实际='{clean_result[:50]}...', 期望='{clean_expected}'")
                else:
                    # 没有期望结果时，执行一次就成功
                    final_result = result
                    print(f"      结果: {result[:50]}..." if len(result) > 50 else f"      结果: {result}")
                    self.logger.info(f"循环迭代 {i+1} 结果: {result[:100] if result else '无结果'}")
                    
                    found_expected = True
                    break
            else:
                # 非通信错误的执行失败
                print(f"❌ 命令执行失败 ({iteration_duration:.2f}秒)")
                print(f"      错误: {result}")
                self.logger.warning(f"循环迭代 {i+1} 命令执行失败 - 错误: {result}")
                final_result = result
                
                # 如果配置了失败停止
                if loop_config.get('unlooppassresultstop', 'false').lower() == 'true':
                    print(f"      ⚠️  检测到失败停止配置，停止循环")
                    self.logger.warning(f"检测到失败停止配置，停止循环")
                    break
            
            # 如果不是最后一次迭代且没有找到期望结果，等待间隔时间
            if i < loop_times - 1 and loop_sleep_ms > 0 and not found_expected and not communication_error_occurred:
                print(f"     等待 {loop_sleep_ms}ms...")
                self.logger.info(f"循环等待: {loop_sleep_ms}ms")
                time.sleep(loop_sleep_ms / 1000)
        
        # 确定最终结果消息
        if communication_error_occurred and is_special_command:
            # 特殊命令的通信错误视为通过
            result_msg = f"循环完成，特殊命令 '{step_content}' 通信错误视为通过"
            self.logger.info(result_msg)
            return True, result_msg, loop_details
        elif communication_error_occurred:
            result_msg = f"循环检测到仪器通信错误，执行了{actual_iterations}次迭代"
            self.logger.warning(result_msg)
            return True, result_msg, loop_details
        elif found_expected:
            result_msg = f"循环完成，在第{actual_iterations}次迭代找到期望结果: '{expected_result}'"
            self.logger.info(result_msg)
            return True, result_msg, loop_details
        else:
            result_msg = f"循环完成，执行了{actual_iterations}次迭代但未找到期望结果: '{expected_result}'"
            self.logger.warning(result_msg)
            return False, result_msg, loop_details
    
    def _is_communication_error(self, result: str) -> bool:
        """检查是否为通信错误"""
        error_keywords = [
            "仪器通信错误",
            "VI_ERROR_TMO",
            "Timeout",
            "timeout",
            "通信失败",
            "通信错误",
            "VI_ERROR_RSRC_LOCKED",
            "VI_ERROR_RSRC_NFOUND"
        ]
        
        if not result:
            return False
        
        result_lower = result.lower()
        for keyword in error_keywords:
            if keyword.lower() in result_lower:
                return True
        
        return False
    
    def cleanup(self):
        """清理资源"""
        # 清理仪器连接
        DirectCommandExecutor.cleanup()
        
        self.connection_status = False
        self.current_file = None
        self.current_file_path = None
        self.selected_device = None
