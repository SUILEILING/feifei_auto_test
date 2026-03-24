from lib.var import *

from souren_core import (
    DirectCommandExecutor, 
    SourenLogger, 
    SourenResultSaver, 
    PythonScriptParser
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
    INTERRUPTED = "interrupted"  # 新增：用户中断状态

# ==============================================
# ↓↓↓ 简化版测试监控器（移除超时机制）
# ==============================================
class TestMonitor:
    """测试监控器 """
    
    def __init__(self):
        self.should_stop = False
        self.is_cancelled = False
        self.is_interrupted = False  # 新增：中断标志
        self.start_time = None
        
    def start(self):
        """启动监控器"""
        self.start_time = time.time()
        self.should_stop = False
        self.is_cancelled = False
        self.is_interrupted = False
        
        print(f"⏰ 监控器已启动（无超时限制）")
    
    def stop(self):
        """停止监控器"""
        self.should_stop = True
    
    def cancel(self):
        """取消测试"""
        self.is_cancelled = True
        self.should_stop = True
    
    def interrupt(self):
        """中断测试（用户Ctrl+C）"""
        self.is_interrupted = True
        self.should_stop = True
        print(f"\n⏹️  用户中断测试执行")
    
    def should_test_stop(self) -> Tuple[bool, str]:
        """检查测试是否应该停止"""
        if self.should_stop:
            if self.is_interrupted:
                return True, "用户中断"
            elif self.is_cancelled:
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
                "is_cancelled": self.is_cancelled,
                "is_interrupted": self.is_interrupted
            }
        
        elapsed = time.time() - self.start_time
        
        status = "running"
        if self.is_interrupted:
            status = "interrupted"
        elif self.is_cancelled:
            status = "cancelled"
        elif self.should_stop:
            status = "stopped"
        
        return {
            "start_time": datetime.fromtimestamp(self.start_time).strftime('%Y-%m-%d %H:%M:%S'),
            "elapsed_seconds": round(elapsed, 1),
            "should_stop": self.should_stop,
            "is_cancelled": self.is_cancelled,
            "is_interrupted": self.is_interrupted,
            "status": status
        }

# ==============================================
# ↓↓↓ Souren.ToolSet 核心功能类
# ==============================================
class SourenEngine:
    """Souren.ToolSet 核心引擎"""
    
    def __init__(self, result_dir=None, script_name=None, params=None):
        self.result_dir = result_dir
        self.script_name = script_name
        self.params = params  # 新增：参数配置
        self.logger = SourenLogger()
        self.result_saver = SourenResultSaver(result_dir, script_name)
        self.python_parser = PythonScriptParser()
        self.connection_status = False
        self.last_error = ""
        self.selected_device = None
        self.test_monitor = None
        self.current_file = None
        self.current_file_path = None
        self.execution_start_time = None
        self.interrupted = False  
    
    def get_result_file(self):
        """获取结果文件路径"""
        return self.result_saver.get_result_file()
    
    def get_result_dir(self):
        """获取结果目录"""
        return self.result_saver.get_result_dir()
    
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
            if command == "load_python_file":
                result = self._load_python_file(**kwargs)
                success = True
            elif command == "select_and_connect_device":
                result = self._select_and_connect_device(**kwargs)
                success = True
            elif command == "run_tests":
                mode = kwargs.get('mode', 'run_all')
                loop_count = kwargs.get('loop_count', 1)
                print(f"🔧 执行命令: run_tests - mode={mode}, loop_count={loop_count}")
                print(f"📁 当前文件: {self.current_file_path}")
                result = self._run_tests_with_souren(mode=mode, loop_count=loop_count)
                success = True
            else:
                return False, {"error": f"未知命令: {command}"}
            
            return success, result
            
        except Exception as e:
            error_msg = f"执行命令失败: {str(e)}"
            print(f"❌ {error_msg}")
            return False, {"error": error_msg}
        
    def _load_python_file(self, file_path: str) -> Dict:
        """加载Python文件"""
        print(f"加载Python文件: {os.path.basename(file_path)}")
        
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
    
    def _parse_python_file(self, file_path: str) -> Dict:
        """解析Python文件"""
        result = self.python_parser.parse_file(file_path)
        
        # 如果有参数配置，修改步骤中的命令
        if self.params and result.get("success"):
            steps = result.get("steps", [])
            modified_steps = self._modify_steps_with_params(steps, self.params)
            result["steps"] = modified_steps
        
        return result
    
    def _modify_steps_with_params(self, steps: List[Dict], params: Dict) -> List[Dict]:
        """根据参数修改步骤中的命令"""
        if not params:
            return steps
        
        print(f"🔄 根据参数修改测试步骤:")
        print(f"   参数配置: {params}")
        
        from souren_config import PARAMETER_COMMAND_MAPPINGS
        
        modified_steps = []
        
        for step in steps:
            command = step.get("content", "")
            original_command = command
            
            # 检查并替换每个参数
            for param_name, param_value in params.items():
                if param_name in PARAMETER_COMMAND_MAPPINGS:
                    mapping = PARAMETER_COMMAND_MAPPINGS[param_name]
                    pattern = mapping["pattern"]
                    
                    # 如果命令匹配模式，则替换
                    import re
                    if re.search(pattern, command):
                        # 生成新的命令
                        new_command = mapping["command"].format(**{param_name: param_value})
                        command = new_command
                        print(f"     替换 {param_name}: {original_command} -> {command}")
                        original_command = command  # 更新原始命令，以便下一次替换
            
            # 创建修改后的步骤
            modified_step = step.copy()
            modified_step["content"] = command
            
            # 记录原始命令和参数
            if command != step.get("content", ""):
                modified_step["original_content"] = step.get("content", "")
                modified_step["parameters_applied"] = params
            
            modified_steps.append(modified_step)
        
        print(f"✅ 步骤修改完成，共修改 {len(modified_steps)} 个步骤")
        return modified_steps
    
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
        """使用仪器运行测试 - 从Python脚本读取测试步骤"""
        self.logger.info(f"开始运行测试 - 模式: {mode}")
        
        if not self.selected_device:
            return {
                "success": False,
                "error": "未连接设备，请先连接设备"
            }
        
        if not self.current_file_path:
            return {
                "success": False,
                "error": "未加载Python文件"
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
            print(f"📄 文件: {current_file_name}")
            print(f"🔌 设备: {device_name}")
            print(f"🎯 模式: {mode}")
            
            # 显示参数信息
            if self.params:
                print(f"📋 参数配置:")
                for key, value in self.params.items():
                    print(f"    {key}: {value}")
            
            if mode == 'loop_info' and loop_count > 1:
                print(f"🔄 全局循环模式: 整个流程执行 {loop_count} 次")
                self.logger.info(f"全局循环模式激活: LOOP_COUNT = {loop_count}")
            else:
                print(f"🔄 单次执行模式")
            
            # 显示脚本信息
            if self.script_name:
                print(f"📝 脚本: {self.script_name}")
            
            print(f"📊 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"{'='*60}")
            
            # 启动测试监控器
            self.test_monitor.start()
            
            # 解析Python文件获取测试步骤和跳过命令
            print(f"\n📖 解析Python文件: {current_file_name}")
            parse_result = self._parse_python_file(self.current_file_path)
            
            if not parse_result.get("success"):
                return {
                    "success": False,
                    "error": f"解析Python文件失败: {parse_result.get('error', '未知错误')}"
                }
            
            steps = parse_result.get("steps", [])
            skip_commands = parse_result.get("skip_commands", set())
            total_steps = len(steps)
            
            if total_steps == 0:
                return {
                    "success": False,
                    "error": "文件中没有找到可执行的测试步骤"
                }
            
            print(f"✅ 解析完成，找到 {total_steps} 个测试步骤")
            self.logger.info(f"解析文件完成,共 {total_steps} 个步骤")
            
            # 显示分析信息
            print(f"📊 步骤分析:")
            print(f"   总步骤数: {total_steps}")
            if skip_commands:
                print(f"   跳过命令数: {len(skip_commands)}")
            
            # 检查循环步骤
            loop_steps = [step for step in steps if step.get("has_loop", False)]
            if loop_steps:
                print(f"   循环步骤数: {len(loop_steps)}")
                for step in loop_steps:
                    loop_config = step.get("raw_loop_config", {})
                    print(f"     - 步骤{step['step']}: 循环{loop_config.get('times', 1)}次, 期望结果='{loop_config.get('expected_result', '')}'")
            
            # 计算总执行次数（考虑跳过）
            if mode == 'loop_info' and loop_count > 1 and skip_commands:
                # 第一次循环执行所有步骤，后续循环执行过滤后的步骤
                filtered_steps_count = len([step for step in steps if step.get("content", "").strip() not in skip_commands])
                total_executions = total_steps + filtered_steps_count * (loop_count - 1)
                print(f"\n📊 执行计划（考虑跳过）:")
                print(f"   第1次循环: 执行所有{total_steps}个步骤")
                print(f"   第2-{loop_count}次循环: 跳过{len(skip_commands)}个步骤，每次执行{filtered_steps_count}个步骤")
                print(f"   总执行次数: {total_executions}")
            elif mode == 'loop_info' and loop_count > 1:
                total_executions = total_steps * loop_count
                print(f"\n🔄 开始执行{total_executions}个测试步骤 ({total_steps}步骤 x {loop_count}次全局循环)...")
            else:
                total_executions = total_steps
                print(f"\n🔄 开始执行{total_steps}个测试步骤...")
            
            print(f"{'='*60}")
            
            executed_steps = 0
            passed_steps = 0
            failed_steps = 0
            communication_failed_steps = 0
            execution_details = []
            
            # 添加键盘中断处理
            import signal
            
            def signal_handler(sig, frame):
                """处理键盘中断信号"""
                if self.test_monitor:
                    self.test_monitor.interrupt()
                    self.interrupted = True
                print(f"\n⏹️  检测到Ctrl+C，正在保存已执行的数据...")
            
            # 注册信号处理
            original_signal = signal.getsignal(signal.SIGINT)
            signal.signal(signal.SIGINT, signal_handler)
            
            try:
                if mode == 'loop_info' and loop_count > 1:
                    # 全局循环模式 - 添加跳过逻辑
                    for loop_index in range(loop_count):
                        # 检查中断
                        if self.interrupted:
                            print(f"\n⏹️  测试被用户中断")
                            break
                            
                        loop_start_time = time.time()
                        print(f"\n🔄 第 {loop_index + 1}/{loop_count} 轮全局循环:")
                        print(f"{'-'*50}")
                        self.logger.info(f"开始第 {loop_index + 1}/{loop_count} 轮全局循环")
                        
                        # 关键修改：根据循环索引处理跳过逻辑
                        if loop_index == 0:
                            # 第一次循环：执行所有步骤
                            print(f"  第{loop_index + 1}次循环：执行所有{len(steps)}个步骤")
                        else:
                            # 后续循环：跳过指定的步骤
                            if skip_commands:
                                print(f"  第{loop_index + 1}次循环：跳过{len(skip_commands)}个步骤")
                            else:
                                print(f"  第{loop_index + 1}次循环：未配置跳过步骤，执行所有{len(steps)}个步骤")
                        
                        # 遍历所有步骤，根据循环索引决定是否跳过
                        for step_idx, step in enumerate(steps):
                            # 检查中断
                            if self.interrupted:
                                print(f"\n⏹️  测试被用户中断")
                                break
                                
                            step_num = step.get("step", 0)
                            step_content = step.get("content", "").strip()
                            step_type = step.get("type", "Normal")
                            has_loop = step.get("has_loop", False)
                            loop_config = step.get("loop_config", {})
                            raw_loop_config = step.get("raw_loop_config", {})
                            
                            # 计算全局步骤编号
                            global_step_num = executed_steps + 1
                            
                            # 检查是否应该停止
                            should_stop, stop_reason = self.test_monitor.should_test_stop()
                            if should_stop:
                                print(f"\n⏹️  测试被{stop_reason}")
                                self.logger.warning(f"测试被停止: {stop_reason}")
                                break
                            
                            # 关键：检查是否需要跳过此步骤
                            if loop_index > 0 and step_content in skip_commands:
                                # 后续循环中，跳过配置的命令
                                print(f"\n⏭️  步骤 {global_step_num}/{total_executions} (第{loop_index + 1}/{loop_count}轮循环): 跳过")
                                print(f"   命令: {step_content}")
                                print(f"   原因: 配置为后续循环跳过")
                                
                                # 记录跳过信息
                                execution_detail = {
                                    "step": global_step_num,
                                    "original_step": step_num,
                                    "content": step_content,
                                    "type": step_type,
                                    "status": "skipped",
                                    "duration": 0,
                                    "result": "配置为后续循环跳过",
                                    "start_time": time.time(),
                                    "end_time": time.time(),
                                    "has_loop": has_loop,
                                    "loop_index": loop_index + 1,
                                    "total_loops": loop_count,
                                    "skipped": True
                                }
                                
                                execution_details.append(execution_detail)
                                executed_steps += 1
                                
                                # 更新进度
                                elapsed_total = time.time() - self.execution_start_time
                                progress = (executed_steps / total_executions) * 100 if total_executions > 0 else 0
                                
                                print(f"   📊 进度: {executed_steps}/{total_executions} ({progress:.1f}%)")
                                print(f"     已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
                                print(f"     总耗时: {elapsed_total:.1f}秒")
                                print(f"{'-'*60}")
                                
                                continue
                            
                            # 显示步骤信息
                            step_info_msg = f"步骤 {global_step_num}/{total_executions}"
                            loop_info = f"第{loop_index + 1}/{loop_count}轮循环"
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
                                # 执行步骤 - 关键修复：使用raw_loop_config检查是否有循环
                                if has_loop and raw_loop_config.get("enable", False):
                                    print(f"   🔄 检测到循环配置，开始执行循环步骤")
                                    print(f"      循环配置: 次数={raw_loop_config.get('times', 1)}, 期望结果='{raw_loop_config.get('expected_result', '')}'")
                                    step_success, step_result, loop_details = self._execute_loop_step_with_expected_result(
                                        step, device_name, raw_loop_config
                                    )
                                else:
                                    step_success, step_result = self._execute_single_step(step, device_name)
                                    loop_details = []
                                
                                is_communication_error = self._is_communication_error(step_result)
                                
                                if is_communication_error:
                                    print(f"   ⚠️  仪器通信错误")
                                    self.logger.warning(f"仪器通信错误: {step_result}")
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
                                    "total_loops": loop_count,
                                    "skipped": False
                                }
                                
                                execution_details.append(execution_detail)
                                
                                # 如果步骤失败且配置了失败停止（通信错误除外）
                                if not step_execution_success and step.get("unloop_pass_result_stop") == "True":
                                    print(f"   ⚠️  检测到失败停止配置，停止后续步骤执行")
                                    self.logger.warning(f"检测到失败停止配置，停止后续执行")
                                    break
                                
                            except KeyboardInterrupt:
                                # 处理步骤执行中的中断
                                raise
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
                            progress = (executed_steps / total_executions) * 100 if total_executions > 0 else 0
                            
                            print(f"   📊 进度: {executed_steps}/{total_executions} ({progress:.1f}%)")
                            print(f"     已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
                            print(f"     总耗时: {elapsed_total:.1f}秒")
                            print(f"{'-'*60}")
                            
                            self.logger.info(f"进度更新: {executed_steps}/{total_executions} ({progress:.1f}%) - 已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
                        
                        # 检查中断
                        if self.interrupted:
                            break
                            
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
                        # 检查中断
                        if self.interrupted:
                            print(f"\n⏹️  测试被用户中断")
                            break
                            
                        step_num = step.get("step", 0)
                        step_content = step.get("content", "").strip()
                        step_type = step.get("type", "Normal")
                        has_loop = step.get("has_loop", False)
                        loop_config = step.get("loop_config", {})
                        raw_loop_config = step.get("raw_loop_config", {})
                        
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
                            # 执行步骤 - 关键修复：使用raw_loop_config检查是否有循环
                            if has_loop and raw_loop_config.get("enable", False):
                                print(f"   🔄 检测到循环配置，开始执行循环步骤")
                                print(f"      循环配置: 次数={raw_loop_config.get('times', 1)}, 期望结果='{raw_loop_config.get('expected_result', '')}'")
                                step_success, step_result, loop_details = self._execute_loop_step_with_expected_result(
                                    step, device_name, raw_loop_config
                                )
                            else:
                                step_success, step_result = self._execute_single_step(step, device_name)
                                loop_details = []
                            
                            # 检查是否为通信错误
                            is_communication_error = self._is_communication_error(step_result)
                            
                            if is_communication_error:
                                print(f"   ⚠️  仪器通信错误")
                                self.logger.warning(f"仪器通信错误: {step_result}")
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
                            
                        except KeyboardInterrupt:
                            # 处理步骤执行中的中断
                            raise
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
                        progress = (executed_steps / total_steps) * 100 if total_steps > 0 else 0
                        
                        print(f"   📊 进度: {executed_steps}/{total_steps} ({progress:.1f}%)")
                        print(f"     已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
                        print(f"     总耗时: {elapsed_total:.1f}秒")
                        print(f"{'-'*60}")
                        
                        self.logger.info(f"进度更新: {executed_steps}/{total_steps} ({progress:.1f}%) - 已执行: {executed_steps}, 通过: {passed_steps}, 失败: {failed_steps}, 通信错误: {communication_failed_steps}")
            
            except KeyboardInterrupt:
                # 处理测试过程中的中断
                print(f"\n⏹️  测试被用户中断，正在保存已执行的数据...")
                self.logger.warning(f"测试被用户中断")
                self.interrupted = True
            finally:
                # 恢复原来的信号处理
                signal.signal(signal.SIGINT, original_signal)
            
            # 停止监控器
            if self.test_monitor:
                self.test_monitor.stop()
            
            # 计算总执行时间
            execution_time = time.time() - self.execution_start_time
            
            # 确定最终状态
            should_stop, stop_reason = self.test_monitor.should_test_stop()
            if should_stop:
                if self.interrupted or stop_reason == "用户中断":
                    status = ExecutionStatus.INTERRUPTED
                    message = f"测试被用户中断 (设备: {device_name})"
                elif stop_reason == "用户取消":
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
                "monitor_info": self.test_monitor.get_monitor_info() if self.test_monitor else {},
                "execution_details": execution_details,
                "parsed_steps": steps,
                "skip_commands": list(skip_commands) if skip_commands else [],  # 记录跳过的命令
                "interrupted": self.interrupted  # 添加中断标志
            }
            
            # 添加参数信息
            if self.params:
                result["parameters"] = self.params
            
            if mode == 'loop_info':
                result["loop_count"] = loop_count
                result["execution_mode"] = "全局循环"
                result["skip_config_used"] = len(skip_commands) > 0 if skip_commands else False
                result["skipped_steps_count"] = len(skip_commands) if skip_commands else 0
            
            # 记录最终结果到日志
            if self.interrupted:
                self.logger.info(f"测试被用户中断 - 状态: {status.value}, 消息: {message}")
                self.logger.info(f"中断前统计: 原始步骤{total_steps}, 总执行次数{total_executions}, 已执行{executed_steps}, 通过{passed_steps}, 失败{failed_steps}, 通信错误{communication_failed_steps}")
            else:
                self.logger.info(f"测试执行完成 - 状态: {status.value}, 消息: {message}")
                self.logger.info(f"统计信息: 原始步骤{total_steps}, 总执行次数{total_executions}, 已执行{executed_steps}, 通过{passed_steps}, 失败{failed_steps}, 通信错误{communication_failed_steps}")
            
            # 关键：即使中断也要保存结果
            print(f"\n📊 正在保存测试结果...")
            save_success = self.result_saver.save_result(result)
            if save_success:
                print(f"✅ 测试结果已保存到JSON文件")
                print(f"   📁 保存目录: {self.result_dir}")
                print(f"   📄 文件: {os.path.basename(self.result_saver.get_result_file())}")
                if self.interrupted:
                    print(f"💡 用户中断，已保存 {executed_steps} 个步骤的执行数据")
            else:
                print(f"❌ 保存测试结果失败")
            
            # 显示最终结果
            print(f"\n{'='*60}")
            if self.interrupted:
                print(f"⏹️  测试被用户中断")
            else:
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
            
            # 显示参数信息
            if self.params:
                print(f"📋 使用的参数:")
                for key, value in self.params.items():
                    print(f"   {key}: {value}")
            
            if mode == 'loop_info' and loop_count > 1:
                completed_loops = min(loop_count, (executed_steps // total_steps) + (1 if executed_steps % total_steps > 0 else 0))
                print(f"🔄 循环信息: 计划 {loop_count} 次，完成 {completed_loops} 次全局循环")
                
                if skip_commands:
                    print(f"   跳过配置: 后续循环跳过 {len(skip_commands)} 个步骤")
                    if loop_count > 1:
                        print(f"   第二次及以后循环执行步骤数: {total_steps - len(skip_commands)}")
                
                self.logger.info(f"循环模式统计: 计划 {loop_count} 次，完成 {completed_loops} 次，共执行 {executed_steps} 次，成功率: {success_rate:.1f}%")
            
            print(f"{'='*60}")
            
            # 返回结果，确保Excel导出也能获取到
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
        step_content = step.get("content", "").strip()
        
        # 使用 loop_config 中的配置
        loop_times = int(loop_config.get("times", 1))
        loop_sleep_ms = int(loop_config.get("sleep_ms", 0))
        expected_result = loop_config.get("expected_result", "")
        
        # 清理期望结果中的引号
        if expected_result:
            expected_result = expected_result.replace('"', '').replace('&quot;', '').replace("'", '').strip()
        
        print(f"   🔄 开始执行循环，最大次数: {loop_times}, 期望结果: '{expected_result}', 每次等待: {loop_sleep_ms}ms")
        self.logger.info(f"开始执行循环步骤: 最大次数={loop_times}, 期望结果='{expected_result}', 等待={loop_sleep_ms}ms")
        
        loop_details = []
        found_expected = False
        communication_error_occurred = False
        final_result = ""
        actual_iterations = 0
        
        for i in range(loop_times):
            # 检查中断
            if self.interrupted:
                self.logger.warning(f"循环被用户中断")
                return True, f"循环被用户中断", loop_details
            
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
                print(f"      ⚠️  检测到仪器通信错误，停止循环")
                self.logger.info(f"检测到仪器通信错误，停止循环")
                break
            
            elif success:
                print(f"✅ 命令执行成功 ({iteration_duration:.2f}秒)")
                self.logger.info(f"循环迭代 {i+1} 命令执行成功 - 耗时: {iteration_duration:.2f}秒")
                
                # 检查是否符合期望结果 - 修复：严格比较
                if expected_result:
                    # 清理结果中的引号和空格
                    clean_result = result.replace('"', '').replace('&quot;', '').replace("'", '').strip()
                    clean_expected = expected_result.replace('"', '').replace('&quot;', '').replace("'", '').strip()
                    
                    print(f"      实际结果: '{clean_result}'")
                    print(f"      期望结果: '{clean_expected}'")
                    
                    if clean_expected and clean_expected == clean_result:
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
                        self.logger.info(f"循环迭代 {i+1} 不符合期望结果: 实际='{clean_result}', 期望='{clean_expected}'")
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
        if communication_error_occurred:
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