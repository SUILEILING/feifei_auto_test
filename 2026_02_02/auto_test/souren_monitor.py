from lib.var import *

from souren_core import (
    DirectCommandExecutor, 
    SourenLogger, 
    SourenResultSaver, 
    PythonScriptParser
)

class ExecutionStatus(Enum):
    """执行状态"""
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    STOPPED = "stopped"
    CANCELLED = "cancelled"
    INTERRUPTED = "interrupted"

class LoopConfig:
    """循环配置类"""
    
    def __init__(self, times: int = 1, power_decrement: float = 1.0, loop_type: str = "fixed_times"):
        self.times = times  # 循环次数
        self.power_decrement = power_decrement  # 功率递减值
        self.loop_type = loop_type  # 循环类型: "fixed_times" 或 "until_stop_condition"
        self.current_iteration = 0
        self.is_active = False
    
    def reset(self):
        """重置循环状态"""
        self.current_iteration = 0
        self.is_active = True
    
    def next_iteration(self) -> bool:
        """进入下一次迭代"""
        if not self.is_active:
            return False
        
        self.current_iteration += 1
        
        if self.loop_type == "fixed_times":
            return self.current_iteration <= self.times
        else:
            return True  # 直到停止条件的循环，外部控制结束
    
    def get_iteration_info(self) -> str:
        """获取迭代信息"""
        if self.loop_type == "fixed_times":
            return f"迭代 {self.current_iteration}/{self.times}"
        else:
            return f"迭代 {self.current_iteration}"
    
    def is_completed(self) -> bool:
        """检查循环是否完成"""
        if self.loop_type == "fixed_times":
            return self.current_iteration >= self.times
        return False

class ConditionProcessor:
    """条件处理器 - 处理if-else条件分支"""
    
    def __init__(self):
        self.execution_context = {}
        self.condition_result_cache = {}  # 缓存条件判断结果
        self.last_condition_value = None
    
    def evaluate_condition(self, condition: str, context: Dict = None) -> bool:
        """评估条件表达式"""
        if not condition:
            return True
        
        # 更新上下文
        if context:
            self.execution_context.update(context)
        
        try:
            # 简化条件判断，支持常见的条件格式
            condition = condition.strip()
            
            # 如果是简单的字符串比较
            if condition == "Connected":
                last_ue_state = self.execution_context.get("last_ue_state", "")
                return str(last_ue_state).strip('"').strip("'") == "Connected"
            
            elif condition == "Disconnected":
                last_ue_state = self.execution_context.get("last_ue_state", "")
                return str(last_ue_state).strip('"').strip("'") != "Connected"
            
            # 处理复杂的条件表达式
            elif "==" in condition:
                parts = condition.split("==")
                if len(parts) == 2:
                    left = parts[0].strip()
                    right = parts[1].strip().strip('"').strip("'")
                    
                    if left == "last_result":
                        last_result = self.execution_context.get("last_result", "")
                        return str(last_result).strip('"').strip("'") == right
                    elif left == "last_ue_state":
                        last_ue_state = self.execution_context.get("last_ue_state", "")
                        return str(last_ue_state).strip('"').strip("'") == right
            
            # 缓存结果
            self.last_condition_value = condition
            
            # 默认返回False
            return False
                
        except Exception as e:
            print(f"❌ 条件评估失败: {condition}, 错误: {e}")
            return False
    
    def check_condition_result(self, command: str, result: str) -> bool:
        """检查条件判断命令的结果"""
        # 检查是否是条件检查命令
        if "CONFigure:CELL1:NR:SIGN:UE:STATe?" in command and "is_condition_check" in command:
            expected_value = None
            condition_type = None
            
            # 尝试从上下文中获取期望值
            if hasattr(self, 'condition_check_expected'):
                expected_value = self.condition_check_expected
            
            # 清理结果字符串
            result_clean = result.strip().strip('"').strip("'")
            
            # 检查条件类型
            if "condition" in command:
                if "==" in command:
                    condition_type = "=="
                elif "!=" in command:
                    condition_type = "!="
            
            # 执行条件检查
            if condition_type == "==" and expected_value:
                return result_clean == str(expected_value).strip('"').strip("'")
            elif condition_type == "!=" and expected_value:
                return result_clean != str(expected_value).strip('"').strip("'")
            
            # 默认检查是否是Connected
            return result_clean == "Connected"
        
        return False
    
    def set_condition_check(self, expected_value: str):
        """设置条件检查的期望值"""
        self.condition_check_expected = expected_value
    
    def update_context(self, key: str, value):
        """更新执行上下文"""
        self.execution_context[key] = value
        
        # 如果是UE状态查询，特别记录
        if key == "last_result" and "CONFigure:CELL1:NR:SIGN:UE:STATe?" in value:
            self.execution_context["last_ue_state"] = value
    
    def get_context(self, key: str, default=None):
        """获取执行上下文中的值"""
        return self.execution_context.get(key, default)
    
    def clear_context(self):
        """清空执行上下文"""
        self.execution_context.clear()
        self.condition_result_cache.clear()
        self.last_condition_value = None

class LoopStepParser:
    """循环步骤解析器 - 支持多个循环段和条件分支"""
    
    def __init__(self):
        self.condition_processor = ConditionProcessor()
        self.in_if_branch = False
        self.in_else_branch = False
        self.current_condition_result = False
    
    def extract_loop_steps(self, steps: List[Dict]) -> Dict:
        """
        从步骤中提取循环部分，处理if-else条件分支
        返回: {"loop_blocks": List[List[Dict]], "pre_steps": List[Dict], "post_steps": List[Dict]}
        """
        loop_blocks = []  # 多个循环块
        pre_steps = []
        post_steps = []
        
        i = 0
        while i < len(steps):
            step = steps[i]
            step_type = step.get("type", "Normal")
            
            # 检查是否为循环开始标记
            if step_type == "ForLoop" and step.get("for") == "on":
                # 提取循环配置
                loop_config = self._extract_loop_config(step)
                loop_start_index = i
                i += 1  # 跳过开始标记
                
                # 收集循环内的步骤
                loop_steps = []
                condition_blocks = []
                current_if_condition = None
                if_steps = []
                else_steps = []
                in_if_branch = False
                in_else_branch = False
                
                # 遍历直到找到循环结束标记
                while i < len(steps):
                    inner_step = steps[i]
                    inner_step_type = inner_step.get("type", "Normal")
                    
                    # 检查是否为循环结束标记
                    if inner_step_type == "ForLoop" and inner_step.get("for") == "off":
                        # 如果还有未处理的if-else分支，添加到循环步骤
                        if current_if_condition:
                            condition_blocks.append({
                                "type": "ConditionalBlock",
                                "condition": current_if_condition,
                                "if_steps": if_steps.copy(),
                                "else_steps": else_steps.copy()
                            })
                        
                        # 将条件块转换为步骤
                        for condition_block in condition_blocks:
                            loop_steps.append(condition_block)
                        
                        # 创建循环块
                        loop_block = {
                            "steps": loop_steps.copy(),
                            "config": loop_config,
                            "start_index": loop_start_index,
                            "end_index": i,
                            "has_conditional_blocks": len(condition_blocks) > 0
                        }
                        loop_blocks.append(loop_block)
                        i += 1  # 跳过结束标记
                        break
                    
                    # 检查是否为条件判断标记
                    elif inner_step_type == "IfCondition":
                        # 如果之前有条件块，先保存
                        if current_if_condition:
                            condition_blocks.append({
                                "type": "ConditionalBlock",
                                "condition": current_if_condition,
                                "if_steps": if_steps.copy(),
                                "else_steps": else_steps.copy()
                            })
                        
                        # 开始新的条件块
                        current_if_condition = inner_step.get("condition", "")
                        if_steps = []
                        else_steps = []
                        in_if_branch = True
                        in_else_branch = False
                    
                    # 检查是否为else分支标记
                    elif inner_step_type == "ElseCondition":
                        in_if_branch = False
                        in_else_branch = True
                    
                    # 处理条件分支中的步骤
                    elif in_if_branch or in_else_branch:
                        if in_if_branch:
                            if_steps.append(inner_step)
                        elif in_else_branch:
                            else_steps.append(inner_step)
                    
                    # 普通步骤
                    else:
                        loop_steps.append(inner_step)
                    
                    i += 1
                
                continue
            
            # 非循环步骤
            if not loop_blocks:
                pre_steps.append(step)
            i += 1
        
        # 如果有循环块，提取循环后的步骤
        if loop_blocks:
            last_end_index = loop_blocks[-1]["end_index"] + 1
            if last_end_index < len(steps):
                post_steps = steps[last_end_index:]
        
        return {
            "loop_blocks": loop_blocks,
            "pre_steps": pre_steps,
            "post_steps": post_steps,
            "has_loop": len(loop_blocks) > 0,
            "has_conditional_blocks": any(block.get("has_conditional_blocks", False) for block in loop_blocks)
        }
    
    def _extract_loop_config(self, step: Dict) -> Dict:
        """从循环开始步骤中提取循环配置"""
        config = {
            "times": 1,
            "power_decrement": 1.0,
            "loop_type": "fixed_times"
        }
        
        # 从步骤中提取times参数
        if "times" in step:
            config["times"] = step["times"]
        else:
            # 如果没有指定times，则认为是条件循环
            config["loop_type"] = "until_stop_condition"
        
        # 从步骤内容中提取power_decrement
        content = step.get("content", "")
        if content.startswith("for:"):
            parts = content.split(",")
            for part in parts:
                if "power_decrement=" in part:
                    try:
                        config["power_decrement"] = float(part.split("=")[1].strip())
                    except:
                        pass
        
        return config
    
    def process_conditional_block(self, block: Dict, context: Dict = None) -> List[Dict]:
        """处理条件块，返回需要执行的步骤"""
        condition = block.get("condition", "")
        if_steps = block.get("if_steps", [])
        else_steps = block.get("else_steps", [])
        
        # 评估条件
        condition_result = self.condition_processor.evaluate_condition(condition, context)
        
        print(f"🔍 条件评估: '{condition}' -> {condition_result}")
        
        # 根据条件结果返回对应的步骤
        if condition_result:
            return if_steps
        else:
            return else_steps

class PowerLoopManager:
    """功率循环管理器"""
    
    def __init__(self, loop_blocks: List[List[Dict]], power_decrement: int = 1, initial_power: float = -85):
        self.loop_blocks = loop_blocks
        self.power_decrement = power_decrement
        self.initial_power = initial_power  # 保存初始功率值
        self.current_power = initial_power
        self.enabled = len(loop_blocks) > 0
        self.stop_condition_met = False  # 停止条件是否满足
        self.stop_conditions_blocks = []  # 每个循环块的停止条件
        
        # 为每个循环块提取停止条件
        for block in loop_blocks:
            self.stop_conditions_blocks.append(self._extract_stop_conditions(block))
        
        if self.enabled:
            print(f"⚡ 功率循环已启用:")
            print(f"   循环块数量: {len(self.loop_blocks)}")
            print(f"   功率递减值: {self.power_decrement}")
            print(f"   初始功率值: {self.initial_power}")
    
    def reset_to_initial_power(self):
        """重置为初始功率值"""
        self.current_power = self.initial_power
        self.stop_condition_met = False
        print(f"⚡ 重置功率值为初始值: {self.initial_power}")
    
    def _extract_stop_conditions(self, loop_steps: List[Dict]) -> Dict[str, List[Dict]]:
        """从循环步骤中提取停止条件"""
        stop_conditions = {}
        
        for step in loop_steps:
            # 处理条件分支
            if isinstance(step, dict) and step.get("type") == "ConditionalBlock":
                # 递归提取条件分支中的停止条件
                if_steps = step.get("if_steps", [])
                else_steps = step.get("else_steps", [])
                
                for branch_step in if_steps + else_steps:
                    command = branch_step.get("content", "").strip()
                    if not command:
                        continue
                    
                    self._add_stop_condition(stop_conditions, branch_step)
            elif isinstance(step, dict):
                # 普通步骤
                command = step.get("content", "").strip()
                if not command:
                    continue
                
                self._add_stop_condition(stop_conditions, step)
        
        return stop_conditions
    
    def _add_stop_condition(self, stop_conditions: Dict, step: Dict):
        """添加停止条件到字典"""
        command = step.get("content", "").strip()
        
        # 检查是否有停止条件字段
        has_check_index = "check_index" in step
        has_condition = "condition" in step
        
        if has_check_index and has_condition:
            condition_config = {
                "check_index": step.get("check_index"),
                "condition": step.get("condition"),
                "data_type": step.get("data_type", "str")
            }
            
            # 根据条件类型添加额外参数
            condition_type = step.get("condition")
            if condition_type in [">", "<", ">=", "<="]:
                if "threshold" in step:
                    condition_config["threshold"] = step.get("threshold")
            elif condition_type in ["!=", "=="]:
                if "expected_value" in step:
                    condition_config["expected_value"] = step.get("expected_value")
                elif "expected_result" in step:
                    condition_config["expected_value"] = step.get("expected_result")
            
            if command not in stop_conditions:
                stop_conditions[command] = []
            stop_conditions[command].append(condition_config)
    
    def _clean_string_value(self, value: str) -> str:
        """清理字符串值"""
        if not isinstance(value, str):
            return value
        
        value = value.strip()
        
        # 去除外层的单引号或双引号
        if (value.startswith('"') and value.endswith('"')) or \
           (value.startswith("'") and value.endswith("'")):
            value = value[1:-1]
        
        return value.strip()
    
    def set_initial_power(self, power: float):
        """设置初始功率"""
        self.current_power = power
        self.initial_power = power
    
    def modify_power_command(self, command: str, power_value: float) -> str:
        """修改功率命令中的功率值"""
        # 处理占位符 {current_power}
        if "{current_power}" in command:
            return command.replace("{current_power}", str(power_value))
        else:
            command = command.rstrip()
            return f"{command} {power_value}"
    
    def get_loop_block_count(self):
        """获取循环块数量"""
        return len(self.loop_blocks)
    
    def get_loop_block(self, index: int) -> List[Dict]:
        """获取指定索引的循环块"""
        if 0 <= index < len(self.loop_blocks):
            return self.loop_blocks[index]
        return []
    
    def get_stop_conditions_for_block(self, block_index: int) -> Dict[str, List[Dict]]:
        """获取指定循环块的停止条件"""
        if 0 <= block_index < len(self.stop_conditions_blocks):
            return self.stop_conditions_blocks[block_index]
        return {}
    
    def check_stop_conditions(self, command: str, result: str, block_index: int) -> Tuple[bool, str]:
        """检查是否满足停止条件"""
        stop_conditions = self.get_stop_conditions_for_block(block_index)
        conditions = stop_conditions.get(command, [])
        if not conditions:
            return False, "无停止条件配置"
        
        # 清理结果字符串
        result_clean = self._clean_string_value(result)
        
        for condition in conditions:
            check_result, message = self._check_single_condition(condition, result_clean, command)
            if check_result:
                return True, message
        
        return False, "未满足任何停止条件"
    
    def _check_single_condition(self, condition: Dict, result: str, command: str) -> Tuple[bool, str]:
        """检查单个停止条件"""
        try:
            check_index = condition.get("check_index", 1) - 1  # 转换为0-based索引
            condition_type = condition.get("condition", "")
            data_type = condition.get("data_type", "str")
            
            # 清理结果字符串
            result = self._clean_string_value(result)
            
            # 解析结果
            values = result.split(',')
            
            # 处理可能的空值
            for i in range(len(values)):
                values[i] = values[i].strip()
            
            # 检查索引是否有效
            if len(values) <= check_index:
                # 如果只有一个值，可能是字符串结果
                if len(values) == 1 and check_index == 0:
                    check_value = values[0]
                else:
                    return False, f"结果格式不正确，索引{check_index+1}超出范围(总共{len(values)}个值)"
            else:
                check_value = values[check_index]
            
            # 清理检查值
            check_value = self._clean_string_value(check_value)
            
            # 根据数据类型转换值
            if data_type == "float":
                try:
                    # 处理可能的非数字字符
                    check_value_str = str(check_value).strip()
                    check_value_str = re.sub(r'[^\d\.\-]', '', check_value_str)
                    if check_value_str:
                        check_value = float(check_value_str)
                    else:
                        check_value = 0.0
                except ValueError:
                    try:
                        check_value = float(check_value)
                    except:
                        return False, f"无法将值 '{check_value}' 转换为浮点数"
            elif data_type == "int":
                try:
                    check_value = int(float(check_value))
                except ValueError:
                    return False, f"无法将值 '{check_value}' 转换为整数"
            
            # 获取期望值并清理
            expected_value = condition.get("expected_value", "")
            if expected_value:
                expected_value = self._clean_string_value(expected_value)
            
            # 获取阈值
            threshold = condition.get("threshold", 0.0)
            
            # 检查条件
            if condition_type == ">":
                if data_type in ["float", "int"]:
                    if check_value > threshold:
                        return True, f"值 {check_value} > 阈值 {threshold}，满足停止条件"
                    else:
                        return False, f"值 {check_value} <= 阈值 {threshold}，不满足停止条件"
            elif condition_type == "<":
                if data_type in ["float", "int"]:
                    if check_value < threshold:
                        return True, f"值 {check_value} < 阈值 {threshold}，满足停止条件"
                    else:
                        return False, f"值 {check_value} >= 阈值 {threshold}，不满足停止条件"
            elif condition_type == ">=":
                if data_type in ["float", "int"]:
                    if check_value >= threshold:
                        return True, f"值 {check_value} >= 阈值 {threshold}，满足停止条件"
                    else:
                        return False, f"值 {check_value} < 阈值 {threshold}，不满足停止条件"
            elif condition_type == "<=":
                if data_type in ["float", "int"]:
                    if check_value <= threshold:
                        return True, f"值 {check_value} <= 阈值 {threshold}，满足停止条件"
                    else:
                        return False, f"值 {check_value} > 阈值 {threshold}，不满足停止条件"
            elif condition_type == "!=":
                if str(check_value).strip() != str(expected_value).strip():
                    return True, f"值 '{check_value}' != 期望值 '{expected_value}'，满足停止条件"
                else:
                    return False, f"值 '{check_value}' == 期望值 '{expected_value}'，不满足停止条件"
            elif condition_type == "==":
                if str(check_value).strip() == str(expected_value).strip():
                    return True, f"值 '{check_value}' == 期望值 '{expected_value}'，满足停止条件"
                else:
                    return False, f"值 '{check_value}' != 期望值 '{expected_value}'，不满足停止条件"
            
            return False, f"不满足停止条件 (值='{check_value}', 期望='{expected_value}', 阈值={threshold})"
            
        except Exception as e:
            return False, f"检查停止条件失败: {str(e)}"

class TestMonitor:
    """测试监控器"""
    
    def __init__(self):
        self.should_stop = False
        self.is_cancelled = False
        self.is_interrupted = False
        self.start_time = None
        
    def start(self):
        """启动监控器"""
        self.start_time = time.time()
        self.should_stop = False
        self.is_cancelled = False
        self.is_interrupted = False
        
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

class SourenEngine:
    """Souren.ToolSet 核心引擎"""
    
    def __init__(self, result_dir=None, script_name=None, params=None):
        self.result_dir = result_dir
        self.script_name = script_name
        self.params = params
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
        
        # 执行上下文
        self.execution_context = {}
        
        # 条件处理器
        self.condition_processor = ConditionProcessor()
        
        # 功率循环管理器
        self.power_loop_manager = None
        
        # 当前功率值
        self.current_power_value = None
        
        # 循环配置
        self.loop_configs = []
        self.current_loop_index = 0
        self.current_loop_config = None
        
        # 条件分支状态
        self.in_if_branch = False
        self.in_else_branch = False
        self.condition_passed = False
        
        # 停止条件标志
        self.stop_condition_met = False
        
        # 第二个循环的初始功率
        self.second_loop_initial_power = None
        
        # SKIP_IN_NEXT_CYCLES 支持
        self.skip_commands = set()
        self.current_loop_iteration = 1
        self.total_loop_count = 1
    
    def get_result_file(self):
        """获取结果文件路径"""
        return self.result_saver.get_result_file()
    
    def get_result_dir(self):
        """获取结果目录"""
        return self.result_saver.get_result_dir()
    
    def initialize(self) -> Tuple[bool, str]:
        """初始化连接"""
        print("初始化仪器连接...")
        
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
            elif command == "run_tests":
                mode = kwargs.get('mode', 'run_all')
                loop_count = kwargs.get('loop_count', 1)
                loop_iteration = kwargs.get('loop_iteration', 1)
                result = self._run_tests_with_souren(mode=mode, loop_count=loop_count, loop_iteration=loop_iteration)
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
        
        if result.get("success"):
            # 保存skip_commands
            self.skip_commands = result.get("skip_commands", set())
            print(f"✅ 解析到 {len(self.skip_commands)} 个需要在后续循环中跳过的命令")
            for cmd in self.skip_commands:
                print(f"   ⏭️  {cmd}")
        
        if self.params and result.get("success"):
            steps = result.get("steps", [])
            modified_steps = self._modify_steps_with_params(steps, self.params)
            result["steps"] = modified_steps
        
        return result
    
    def _modify_steps_with_params(self, steps: List[Dict], params: Dict) -> List[Dict]:
        """根据参数修改步骤中的命令"""
        if not params:
            return steps
        
        from souren_config import PARAMETER_COMMAND_MAPPINGS
        
        modified_steps = []
        
        for step in steps:
            command = step.get("content", "")
            original_command = command
            
            for param_name, param_value in params.items():
                if param_name in PARAMETER_COMMAND_MAPPINGS:
                    mapping = PARAMETER_COMMAND_MAPPINGS[param_name]
                    pattern = mapping["pattern"]
                    
                    if pattern and re.search(pattern, command):
                        if mapping["command"]:
                            new_command = mapping["command"].format(**{param_name: param_value})
                            command = new_command
                            original_command = command
            
            modified_step = step.copy()
            modified_step["content"] = command
            
            if command != step.get("content", ""):
                modified_step["original_content"] = step.get("content", "")
                modified_step["parameters_applied"] = params
            
            modified_steps.append(modified_step)
        
        return modified_steps
    
    def _run_tests_with_souren(self, mode: str = 'run_all', loop_count: int = 1, loop_iteration: int = 1) -> Dict:
        """使用仪器运行测试"""
        self.logger.info(f"开始运行测试 - 模式: {mode}, 循环: {loop_iteration}/{loop_count}")
        
        if not self.current_file_path:
            return {
                "success": False,
                "error": "未加载Python文件"
            }
        
        from souren_config import SHOW_STEP_TYPE, SHOW_STEP_CONTENT, LOOP_COUNT, EXECUTION_MODE, INSTRUMENT_ADDRESS
        
        if loop_count == 1 and LOOP_COUNT > 1 and EXECUTION_MODE == 'loop_info':
            loop_count = LOOP_COUNT
        
        current_file_name = os.path.basename(self.current_file_path) if self.current_file_path else self.current_file
        
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
            
            if self.params:
                print(f"📋 参数配置:")
                for key, value in self.params.items():
                    print(f"    {key}: {value}")
            
            print(f"📊 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            # 显示skip信息
            if loop_iteration > 1 and self.skip_commands:
                print(f"⏭️  本次循环将跳过 {len(self.skip_commands)} 个步骤")
            
            print(f"{'='*60}")
            
            self.test_monitor.start()
            
            parse_result = self._parse_python_file(self.current_file_path)
            
            if not parse_result.get("success"):
                return {
                    "success": False,
                    "error": f"解析Python文件失败: {parse_result.get('error', '未知错误')}"
                }
            
            base_steps = parse_result.get("steps", [])
            
            # 设置当前循环信息
            self.current_loop_iteration = loop_iteration
            self.total_loop_count = loop_count
            
            # 解析循环步骤
            loop_parser = LoopStepParser()
            loop_info = loop_parser.extract_loop_steps(base_steps)
            
            if loop_info["has_loop"]:
                print(f"\n🔄 检测到循环步骤，启用功率循环模式")
                print(f"   循环前步骤: {len(loop_info['pre_steps'])}")
                print(f"   循环块数量: {len(loop_info['loop_blocks'])}")
                print(f"   循环后步骤: {len(loop_info['post_steps'])}")
                
                return self._run_power_loop_mode(loop_info, device_name, mode, loop_parser)
            else:
                print(f"\n🔧 无循环步骤，启用正常执行模式")
                return self._run_normal_mode(base_steps, device_name, mode)
                
        except Exception as e:
            error_msg = f"运行测试失败: {str(e)}"
            print(f"❌ {error_msg}")
            self.logger.error(f"运行测试失败: {error_msg}")
            
            if self.test_monitor:
                self.test_monitor.stop()
            
            return {
                "success": False,
                "error": error_msg,
                "loop_iteration": loop_iteration,
                "loop_count": loop_count
            }
    
    def _run_power_loop_mode(self, loop_info: Dict, device_name: str, mode: str, loop_parser: LoopStepParser) -> Dict:
        """运行功率循环模式"""
        execution_details = []
        executed_steps = 0
        passed_steps = 0
        failed_steps = 0
        communication_failed_steps = 0
        
        # 从参数获取初始功率和功率递减值
        initial_power = self.params.get("current_power", -85) if self.params else -85
        if isinstance(initial_power, str):
            try:
                initial_power = float(initial_power)
            except ValueError:
                initial_power = -85
        
        power_decrement = self.params.get("power_decrement", 1) if self.params else 1
        
        # 保存第二个循环的初始功率
        self.second_loop_initial_power = initial_power
        
        # 初始化功率循环管理器
        self.power_loop_manager = PowerLoopManager(
            [block["steps"] for block in loop_info["loop_blocks"]],
            power_decrement,
            initial_power
        )
        
        # 设置初始功率
        current_power = initial_power
        self.current_power_value = current_power
        
        print(f"⚡ 初始功率设置为: {current_power}")
        
        # 执行循环前的步骤
        print(f"\n🔧 执行循环前的初始化步骤...")
        for step in loop_info["pre_steps"]:
            if self.interrupted:
                print(f"\n⏹️  测试被用户中断")
                break
                
            step_result = self._execute_step_with_tracking(step, device_name, execution_details, 
                                                        executed_steps, passed_steps, 
                                                        failed_steps, communication_failed_steps)
            executed_steps, passed_steps, failed_steps, communication_failed_steps = step_result
        
        # 处理每个循环块
        for loop_index, loop_block in enumerate(loop_info["loop_blocks"]):
            if self.interrupted:
                break
                
            print(f"\n{'='*60}")
            print(f"🔄 开始执行第 {loop_index + 1} 个循环块")
            print(f"{'='*60}")
            
            # 获取循环配置
            loop_config = loop_block.get("config", {})
            loop_times = loop_config.get("times", 1)
            loop_power_decrement = loop_config.get("power_decrement", power_decrement)
            loop_steps = loop_block.get("steps", [])
            loop_type = loop_config.get("loop_type", "fixed_times")
            
            print(f"🔧 循环配置: 类型={loop_type}, 次数={loop_times}, 功率递减={loop_power_decrement}")
            
            if loop_index == 0:
                # 第一个循环：条件循环
                iteration = 1
                self.stop_condition_met = False
                loop_condition_met = False
                
                print(f"\n{'='*60}")
                print(f"🔄 条件循环 {loop_index + 1} - 迭代 {iteration}: 功率={current_power}")
                print(f"{'='*60}")
                
                while not self.stop_condition_met and not self.interrupted and not loop_condition_met:
                    # 更新当前功率值
                    self.current_power_value = current_power
                    
                    # 执行循环内的步骤
                    for step in loop_steps:
                        if self.interrupted or self.stop_condition_met or loop_condition_met:
                            break
                        
                        step_result = self._execute_loop_step(
                            step, device_name, execution_details, 
                            executed_steps, passed_steps, failed_steps, 
                            communication_failed_steps, iteration, current_power,
                            loop_index=loop_index, loop_times=loop_times
                        )
                        executed_steps, passed_steps, failed_steps, communication_failed_steps = step_result
                        
                        # 检查是否满足停止条件
                        if self.stop_condition_met:
                            loop_condition_met = True
                            print(f"\n🛑 满足停止条件，退出条件循环 {loop_index + 1}")
                            break
                    
                    # 如果没有满足停止条件，递减功率并继续下一次迭代
                    if not self.stop_condition_met and not loop_condition_met:
                        iteration += 1
                        new_power = current_power - loop_power_decrement
                        print(f"\n📉 功率递减: {current_power} -> {new_power}")
                        current_power = new_power
                        self.current_power_value = current_power
                        
                        print(f"\n{'='*60}")
                        print(f"🔄 条件循环 {loop_index + 1} - 迭代 {iteration}: 功率={current_power}")
                        print(f"{'='*60}")
                
                print(f"✅ 条件循环 {loop_index + 1} 完成，共执行 {iteration-1} 次迭代")
                
            elif loop_index == 1:
                # 第二个循环：固定次数循环
                # 重置为第二个循环的初始功率
                current_power = self.second_loop_initial_power
                self.current_power_value = current_power
                print(f"⚡ 第二个循环重置功率为初始值: {current_power}")
                
                for iteration in range(1, loop_times + 1):
                    if self.interrupted:
                        break
                        
                    print(f"\n{'='*60}")
                    print(f"🔄 固定循环 {loop_index + 1} - 迭代 {iteration}/{loop_times}: 功率={current_power}")
                    print(f"{'='*60}")
                    
                    # 更新当前功率值
                    self.current_power_value = current_power
                    
                    # 执行循环内的步骤
                    for step in loop_steps:
                        if self.interrupted:
                            break
                        
                        # 处理条件分支
                        if isinstance(step, dict) and step.get("type") == "ConditionalBlock":
                            # 处理条件分支
                            branch_steps = loop_parser.process_conditional_block(step, self.execution_context)
                            
                            # 执行分支中的步骤
                            for branch_step in branch_steps:
                                if self.interrupted:
                                    break
                                
                                step_result = self._execute_loop_step(
                                    branch_step, device_name, execution_details, 
                                    executed_steps, passed_steps, failed_steps, 
                                    communication_failed_steps, iteration, current_power,
                                    loop_index=loop_index, loop_times=loop_times
                                )
                                executed_steps, passed_steps, failed_steps, communication_failed_steps = step_result
                            
                            continue
                        
                        # 普通步骤
                        step_result = self._execute_loop_step(
                            step, device_name, execution_details, 
                            executed_steps, passed_steps, failed_steps, 
                            communication_failed_steps, iteration, current_power,
                            loop_index=loop_index, loop_times=loop_times
                        )
                        executed_steps, passed_steps, failed_steps, communication_failed_steps = step_result
                    
                    # 如果不是最后一次迭代，递减功率
                    if iteration < loop_times:
                        new_power = current_power - loop_power_decrement
                        print(f"\n📉 功率递减: {current_power} -> {new_power}")
                        current_power = new_power
                        self.current_power_value = current_power
        
        # 执行循环后的步骤
        if not self.interrupted:
            print(f"\n🔧 执行循环后的清理步骤...")
            for step in loop_info["post_steps"]:
                if self.interrupted:
                    print(f"\n⏹️  测试被用户中断")
                    break
                    
                step_result = self._execute_step_with_tracking(step, device_name, execution_details, 
                                                            executed_steps, passed_steps, 
                                                            failed_steps, communication_failed_steps)
                executed_steps, passed_steps, failed_steps, communication_failed_steps = step_result
        
        # 生成最终结果
        result = self._generate_final_result(execution_details, executed_steps, passed_steps, 
                                        failed_steps, communication_failed_steps, 
                                        device_name, "power_loop")
        
        # 保存结果
        save_success = self.result_saver.save_result(result)
        
        return result
    
    def _execute_loop_step(self, step: Dict, device_name: str, execution_details: List,
                          executed_steps: int, passed_steps: int, 
                          failed_steps: int, communication_failed_steps: int,
                          iteration: int, current_power: float,
                          loop_index: int = 0, loop_times: int = 1) -> Tuple[int, int, int, int]:
        """执行循环中的步骤（包含功率值替换）"""
        step_content = step.get("content", "").strip()
        
        # 跳过控制标记和注释
        if step_content in ["for:on", "for:off", "if", "else"] or step_content.startswith("if:") or step_content.startswith("else:") or step_content.startswith("#"):
            print(f"⏭️  跳过控制标记: {step_content}")
            return executed_steps, passed_steps, failed_steps, communication_failed_steps
        
        # 检查是否需要跳过（后续循环且命令在skip_commands中）
        if self.current_loop_iteration > 1 and step_content in self.skip_commands:
            print(f"⏭️  跳过步骤（后续循环跳过）: {step_content}")
            
            # 记录跳过状态
            executed_steps += 1
            execution_detail = {
                "step": executed_steps,
                "content": step_content,
                "status": "skipped",
                "duration": 0,
                "result": "配置为后续循环跳过",
                "start_time": time.time(),
                "end_time": time.time(),
                "iteration": iteration,
                "current_power": current_power,
                "loop_index": loop_index,
                "loop_iteration": self.current_loop_iteration,
                "loop_count": self.total_loop_count,
                "is_conditional": False,
                "is_skipped": True
            }
            
            execution_details.append(execution_detail)
            return executed_steps, passed_steps, failed_steps, communication_failed_steps
        
        # 替换功率值占位符
        if "{current_power}" in step_content:
            if self.current_power_value is not None:
                step_content = step_content.replace("{current_power}", str(self.current_power_value))
                step = step.copy()
                step["content"] = step_content
                print(f"   ⚡ 替换功率值: {self.current_power_value}")
            else:
                print(f"   ⚠️  警告: 未设置当前功率值，跳过功率值替换")
        
        # 调用原有的步骤执行方法
        return self._execute_step_with_tracking(step, device_name, execution_details,
                                              executed_steps, passed_steps,
                                              failed_steps, communication_failed_steps,
                                              iteration=iteration, current_power=current_power,
                                              loop_index=loop_index, loop_times=loop_times)
    
    def _execute_step_with_tracking(self, step: Dict, device_name: str, execution_details: List,
                                executed_steps: int, passed_steps: int, 
                                failed_steps: int, communication_failed_steps: int,
                                **kwargs) -> Tuple[int, int, int, int]:
        """执行步骤并跟踪结果"""
        step_content = step.get("content", "").strip()
        
        # 检查是否需要跳过（后续循环且命令在skip_commands中）
        if self.current_loop_iteration > 1 and step_content in self.skip_commands:
            print(f"⏭️  跳过步骤（后续循环跳过）: {step_content}")
            
            # 记录跳过状态
            executed_steps += 1
            execution_detail = {
                "step": executed_steps,
                "content": step_content,
                "status": "skipped",
                "duration": 0,
                "result": "配置为后续循环跳过",
                "start_time": time.time(),
                "end_time": time.time(),
                "iteration": kwargs.get("iteration", 0),
                "current_power": kwargs.get("current_power", None),
                "loop_index": kwargs.get("loop_index", 0),
                "loop_iteration": self.current_loop_iteration,
                "loop_count": self.total_loop_count,
                "is_conditional": False,
                "is_skipped": True
            }
            
            execution_details.append(execution_detail)
            return executed_steps, passed_steps, failed_steps, communication_failed_steps
        
        # 添加功率值替换逻辑
        if "{current_power}" in step_content:
            if self.current_power_value is not None:
                step_content = step_content.replace("{current_power}", str(self.current_power_value))
                step = step.copy()
                step["content"] = step_content
                print(f"   ⚡ 替换功率值: {self.current_power_value}")
            else:
                print(f"   ⚠️  警告: 未设置当前功率值，跳过功率值替换")
        
        # 处理条件分支标记
        if step_content == "if:Connected":
            # 检查上一次UE状态查询结果
            last_ue_state = self.condition_processor.get_context("last_ue_state", "")
            last_ue_state_clean = str(last_ue_state).strip('"').strip("'")
            self.condition_passed = (last_ue_state_clean == "Connected")
            self.in_if_branch = self.condition_passed
            self.in_else_branch = not self.condition_passed
            
            print(f"🔍 条件判断: UE状态='{last_ue_state_clean}', 条件通过={self.condition_passed}")
            
            if self.condition_passed:
                print(f"   ✅ 进入if分支")
            else:
                print(f"   🔄 进入else分支")
            
            # 跳过这个标记步骤
            return executed_steps, passed_steps, failed_steps, communication_failed_steps
        
        elif step_content == "else:":
            # 重置条件分支状态
            self.in_if_branch = False
            self.in_else_branch = False
            self.condition_passed = False
            
            print(f"🔚 结束条件分支")
            
            # 跳过这个标记步骤
            return executed_steps, passed_steps, failed_steps, communication_failed_steps
        
        # 如果在条件分支中，检查是否应该执行当前步骤
        if self.in_if_branch or self.in_else_branch:
            # 只执行对应分支的步骤
            if (self.in_if_branch and not self.condition_passed) or (self.in_else_branch and self.condition_passed):
                print(f"⏭️  跳过非当前分支的步骤: {step_content[:50]}...")
                return executed_steps, passed_steps, failed_steps, communication_failed_steps
        
        # 跳过其他控制标记和注释
        if step_content in ["for:on", "for:off"] or step_content.startswith("#"):
            print(f"⏭️  跳过控制标记: {step_content}")
            return executed_steps, passed_steps, failed_steps, communication_failed_steps
        
        # 检查是否有循环配置
        has_loop = step.get("has_loop", False)
        loop_times = int(step.get("loop_times", 1))
        sleep_ms = int(step.get("sleep_ms", 0))
        loop_pass_result = step.get("loop_pass_result", "")
        
        executed_steps += 1
        step_info_msg = f"步骤 {executed_steps}"
        
        # 获取参数
        iteration = kwargs.get("iteration", 0)
        current_power = kwargs.get("current_power", None)
        loop_index = kwargs.get("loop_index", 0)
        loop_times_config = kwargs.get("loop_times", 1)
        
        if iteration > 0:
            step_info_msg += f" (循环{loop_index+1} 迭代{iteration}/{loop_times_config}, 功率={current_power})"
        
        print(f"\n📝 {step_info_msg}:")
        print(f"   📄 内容: {step_content}")
        
        self.logger.info(f"开始执行: {step_info_msg} - 内容: {step_content}")
        
        step_start_time = time.time()
        
        try:
            # 如果是Sleep命令，特殊处理
            if step_content.upper().startswith("SLEEP"):
                import re
                sleep_match = re.search(r'SLEEP\s+(\d+)', step_content.upper())
                if sleep_match:
                    sleep_ms = int(sleep_match.group(1))
                    print(f"😴 睡眠 {sleep_ms} 毫秒...")
                    time.sleep(sleep_ms / 1000)
                    
                    step_end_time = time.time()
                    step_duration = step_end_time - step_start_time
                    
                    passed_steps += 1
                    step_status = "success"
                    status_icon = "✅"
                    status_text = "通过"
                    
                    print(f"   {status_icon} 执行{status_text} - 耗时: {step_duration:.2f}秒")
                    
                    execution_detail = {
                        "step": executed_steps,
                        "content": step_content,
                        "status": step_status,
                        "duration": step_duration,
                        "result": f"睡眠完成 ({sleep_ms}毫秒)",
                        "start_time": step_start_time,
                        "end_time": step_end_time,
                        "iteration": iteration,
                        "current_power": current_power,
                        "loop_index": loop_index,
                        "loop_iteration": self.current_loop_iteration,
                        "loop_count": self.total_loop_count,
                        "is_conditional": False,
                        "is_skipped": False
                    }
                    
                    execution_details.append(execution_detail)
                    
                    return executed_steps, passed_steps, failed_steps, communication_failed_steps
            
            # 处理有循环配置的步骤
            if has_loop and loop_times > 1:
                print(f"   🔄 检测到循环配置，循环次数: {loop_times}, 每次等待: {sleep_ms}毫秒")
                final_success = False
                final_result = ""
                
                for loop_iter in range(1, loop_times + 1):
                    if self.interrupted:
                        break
                    
                    print(f"   🔄 循环执行第 {loop_iter}/{loop_times} 次")
                    success, result = self._execute_single_step(step, device_name)
                    
                    # 更新执行上下文和条件处理器
                    if result:
                        self.execution_context["last_result"] = result
                        self.condition_processor.update_context("last_result", result)
                        
                        # 如果是UE状态查询，特别记录
                        if "CONFigure:CELL1:NR:SIGN:UE:STATe?" in step_content:
                            self.condition_processor.update_context("last_ue_state", result)
                            self.execution_context["last_ue_state"] = result
                    
                    # 如果期望结果不为空，且结果匹配，则退出循环
                    if loop_pass_result and result == loop_pass_result:
                        print(f"   ✅ 达到期望结果，退出循环")
                        final_success = success
                        final_result = result
                        break
                    
                    if sleep_ms > 0 and loop_iter < loop_times:
                        time.sleep(sleep_ms / 1000)
                    
                    final_success = success
                    final_result = result
                
                success = final_success
                result = final_result
            else:
                # 执行其他命令
                success, result = self._execute_single_step(step, device_name)
                
                # 更新执行上下文和条件处理器
                if result:
                    self.execution_context["last_result"] = result
                    self.condition_processor.update_context("last_result", result)
                    
                    # 如果是UE状态查询，特别记录
                    if "CONFigure:CELL1:NR:SIGN:UE:STATe?" in step_content:
                        self.condition_processor.update_context("last_ue_state", result)
                        self.execution_context["last_ue_state"] = result
            
            # 检查是否为通信错误
            is_communication_error = self._is_communication_error(result)
            
            if is_communication_error:
                communication_failed_steps += 1
                step_status = "communication_error"
                status_icon = "⚠️"
                status_text = "通信错误"
            elif success:
                passed_steps += 1
                step_status = "success"
                status_icon = "✅"
                status_text = "通过"
            else:
                failed_steps += 1
                step_status = "failed"
                status_icon = "❌"
                status_text = "失败"
            
            step_end_time = time.time()
            step_duration = step_end_time - step_start_time
            
            print(f"   {status_icon} 执行{status_text} - 耗时: {step_duration:.2f}秒")
            
            # 检查停止条件
            if self.power_loop_manager and not self.stop_condition_met:
                condition_met, message = self.power_loop_manager.check_stop_conditions(
                    step_content, result, kwargs.get("loop_index", 0)
                )
                if condition_met:
                    self.stop_condition_met = True
                    print(f"   🛑 满足停止条件: {message}")
                else:
                    # 如果不满足停止条件，打印原因
                    if "BLER(索引8)" in locals() and "FETCh:NR:BLER:DL:RESult?" in step_content:
                        try:
                            result_clean = result.strip().strip('"').strip("'")
                            parts = result_clean.split(',')
                            if len(parts) > 8:
                                bler_value = float(parts[8].strip())
                                if bler_value <= 0.5:
                                    print(f"   📊 BLER值 {bler_value} <= 0.5，不满足停止条件")
                        except:
                            pass
            
            if result and result != "命令执行成功":
                # 如果是BLER结果，简化显示
                if "FETCh:NR:BLER:DL:RESult?" in step_content:
                    try:
                        # 清理结果字符串
                        result_clean = result.strip().strip('"').strip("'")
                        print(f"   📊 结果: {result_clean}")
                        
                        # 解析第9个值（索引8，从0开始）
                        parts = result_clean.split(',')
                        if len(parts) > 8:
                            try:
                                bler_value = float(parts[7].strip())
                                print(f"      BLER(索引8): {bler_value}")
                            except:
                                pass
                    except Exception as e:
                        print(f"   📊 结果: {result}")
                else:
                    result_display = result[:100] + "..." if len(result) > 100 else result
                    print(f"   📊 结果: {result_display}")
            
            execution_detail = {
                "step": executed_steps,
                "content": step_content,
                "status": step_status,
                "duration": step_duration,
                "result": result,
                "start_time": step_start_time,
                "end_time": step_end_time,
                "iteration": iteration,
                "current_power": current_power,
                "loop_index": loop_index,
                "loop_iteration": self.current_loop_iteration,
                "loop_count": self.total_loop_count,
                "is_conditional": False,
                "is_skipped": False
            }
            
            execution_details.append(execution_detail)
            
        except Exception as e:
            step_end_time = time.time()
            step_duration = step_end_time - step_start_time
            
            failed_steps += 1
            
            print(f"   ❌ 执行异常 - 耗时: {step_duration:.2f}秒")
            print(f"     错误: {str(e)}")
            
            execution_detail = {
                "step": executed_steps,
                "content": step_content,
                "status": "error",
                "duration": step_duration,
                "result": f"执行异常: {str(e)}",
                "start_time": step_start_time,
                "end_time": step_end_time,
                "iteration": iteration,
                "current_power": current_power,
                "loop_index": loop_index,
                "loop_iteration": self.current_loop_iteration,
                "loop_count": self.total_loop_count,
                "is_conditional": False,
                "is_skipped": False
            }
            
            execution_details.append(execution_detail)
        
        return executed_steps, passed_steps, failed_steps, communication_failed_steps
    
    def _execute_single_step(self, step: Dict, device_name: str) -> Tuple[bool, str]:
        """执行单个测试步骤"""
        step_content = step.get("content", "").strip()
        
        if not step_content:
            return False, "空的命令内容"
        
        success, result = DirectCommandExecutor.execute_command(step_content)
        
        if success:
            self.logger.info(f"命令执行成功: '{step_content}' - 结果: {result[:100] if result else '无结果'}")
        else:
            self.logger.warning(f"命令执行失败: '{step_content}' - 错误: {result}")
        
        return success, result
    
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
    
    def _generate_final_result(self, execution_details: List, executed_steps: int, 
                            passed_steps: int, failed_steps: int, communication_failed_steps: int,
                            device_name: str, mode: str) -> Dict:
        """生成最终结果"""
        execution_time = time.time() - self.execution_start_time if self.execution_start_time else 0
        
        result = {
            "success": True,
            "mode": mode,
            "device": device_name,
            "file": self.current_file,
            "status": "completed",
            "message": f"测试执行完成，模式: {mode}",
            "executed_steps": executed_steps,
            "passed": passed_steps,
            "failed": failed_steps,
            "communication_failed": communication_failed_steps,
            "execution_time": execution_time,
            "duration": execution_time,
            "start_time": datetime.fromtimestamp(self.execution_start_time).strftime('%Y-%m-%d %H:%M:%S') if self.execution_start_time else "",
            "end_time": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "execution_details": execution_details,
            "interrupted": self.interrupted,
            "loop_iteration": self.current_loop_iteration,
            "loop_count": self.total_loop_count,
            "skip_commands": list(self.skip_commands) if self.skip_commands else []
        }
        
        if self.params:
            result["parameters"] = self.params
        
        if self.power_loop_manager and self.power_loop_manager.enabled:
            result["power_loop_config"] = {
                "power_decrement": self.power_loop_manager.power_decrement,
                "initial_power": self.power_loop_manager.initial_power,
                "loop_block_count": self.power_loop_manager.get_loop_block_count()
            }
        
        return result
    
    def _run_normal_mode(self, base_steps: List[Dict], device_name: str, mode: str) -> Dict:
        """运行正常模式"""
        print(f"\n🔧 进入正常执行模式")
        
        execution_details = []
        executed_steps = 0
        passed_steps = 0
        failed_steps = 0
        communication_failed_steps = 0
        
        # 执行所有步骤
        for step in base_steps:
            if self.interrupted:
                print(f"\n⏹️  测试被用户中断")
                break
                
            step_result = self._execute_step_with_tracking(step, device_name, execution_details, 
                                                        executed_steps, passed_steps, 
                                                        failed_steps, communication_failed_steps)
            executed_steps, passed_steps, failed_steps, communication_failed_steps = step_result
        
        # 生成最终结果
        return self._generate_final_result(execution_details, executed_steps, passed_steps, 
                                        failed_steps, communication_failed_steps, 
                                        device_name, "normal")
    
    def cleanup(self):
        """清理资源"""
        DirectCommandExecutor.cleanup()
        
        self.connection_status = False
        self.current_file = None
        self.current_file_path = None
        self.selected_device = None
        self.execution_context.clear()
        self.condition_processor.clear_context()
        self.current_power_value = None
        self.in_if_branch = False
        self.in_else_branch = False
        self.condition_passed = False
        self.stop_condition_met = False
        self.second_loop_initial_power = None
        self.skip_commands.clear()
        self.current_loop_iteration = 1
        self.total_loop_count = 1