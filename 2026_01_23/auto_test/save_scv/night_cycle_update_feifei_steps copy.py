import os
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field
from enum import Enum


# ==============================================
# 数据类型定义
# ==============================================

class StepType(Enum):
    """步骤类型"""
    NORMAL = "Normal"
    CALL = "Call"
    SLEEP = "Sleep"
    QUERY = "Query"
    CONFIGURE = "Configure"


@dataclass
class LoopConfig:
    """循环配置"""
    enable: bool = False
    times: int = 1
    sleep_ms: int = 0
    expected_result: str = ""
    stop_on_fail: bool = False
    
    def to_dict(self) -> Dict:
        """转换为字典"""
        return {
            "enable": self.enable,
            "times": self.times,
            "sleep_ms": self.sleep_ms,
            "expected_result": self.expected_result,
            "stop_on_fail": self.stop_on_fail
        }


@dataclass
class TestStep:
    """测试步骤"""
    command: str  # 命令内容
    expected_result: str = ""  # 预期结果
    loop_config: LoopConfig = field(default_factory=LoopConfig)  # 循环配置
    step_type: StepType = StepType.NORMAL  # 步骤类型
    run_status: str = "Pass"  # 运行状态
    description: str = ""  # 步骤描述
    timeout_ms: int = 10000  # 超时时间(毫秒)
    
    def __post_init__(self):
        """初始化后自动确定步骤类型"""
        cmd = self.command.upper()
        if cmd.startswith("CALL:"):
            self.step_type = StepType.CALL
        elif cmd.startswith("SLEEP"):
            self.step_type = StepType.SLEEP
        elif "?" in cmd:
            self.step_type = StepType.QUERY
        elif cmd.startswith("CONFIGURE:"):
            self.step_type = StepType.CONFIGURE
    
    def to_dict(self) -> Dict:
        """转换为字典"""
        return {
            "command": self.command,
            "expected_result": self.expected_result,
            "loop_config": self.loop_config.to_dict(),
            "step_type": self.step_type.value,
            "run_status": self.run_status,
            "description": self.description,
            "timeout_ms": self.timeout_ms
        }
    
    def __str__(self) -> str:
        """字符串表示"""
        result = f"命令: {self.command}"
        if self.expected_result:
            result += f"\n  预期: {self.expected_result}"
        if self.loop_config.enable:
            result += f"\n  循环: {self.loop_config.times}次"
            if self.loop_config.sleep_ms > 0:
                result += f" (间隔{self.loop_config.sleep_ms}ms)"
            if self.loop_config.expected_result:
                result += f" -> {self.loop_config.expected_result}"
        if self.description:
            result += f"\n  描述: {self.description}"
        return result


# ==============================================
# 测试步骤配置 - 按顺序执行
# ==============================================

TEST_STEPS: List[TestStep] = [


    TestStep(
        command="CALL:CELL1 ON",
        description="开启CELL1",
        step_type=StepType.CALL,
        run_status="Pass"
    ),
    
    TestStep(
        command="Sleep 10000",
        description="等待CELL1启动",
        step_type=StepType.SLEEP,
        run_status="Pass"
    ),
    
    TestStep(
        command="CONFigure:CELL1:NR:SIGN:UE:STATe?",
        expected_result="Connected",
        loop_config=LoopConfig(
            enable=True,
            times=20,
            sleep_ms=1000,
            expected_result="Connected"
        ),
        description="查询UE连接状态,最多尝试20次",
        step_type=StepType.QUERY,
        run_status="Pass"
    ),
    

    TestStep(
        command="CONFigure:CELL1:NR:SIGN:SLOT:CLEar",
        description="清除时隙配置",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    TestStep(
        command="CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh",
        description="配置时隙3为PDSCh",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    TestStep(
        command="CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh",
        description="配置时隙4为PDSCh",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    TestStep(
        command="CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 3",
        description="配置时隙4为PDSCh tind 3",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    TestStep(
        command="CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh",
        description="配置时隙8为PUSCh",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    TestStep(
        command="CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh",
        description="配置时隙9为PUSCh",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    TestStep(
        command="CONFigure:CELL1:NR:SIGN:SLOT:UPDate",
        description="时隙配置更新",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    TestStep(
        command="CONFigure:CELL1:NR:SIGN:SLOT:APPLy",
        description="应用时隙配置",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),


    TestStep(
        command="CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF",
        description="配置ME评估结果",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    TestStep(
        command="CONFigure:NR:MEValuation:REPetition SINGLESHOT",
        description="配置ME评估为单次模式",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    
    TestStep(
        command="CONFigure:NR:BLER:REPetition SINGLESHOT",
        description="配置BLER为单次模式",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    
    TestStep(
        command="INITiate:NR:BLER",
        description="初始化BLER测试",
        step_type=StepType.NORMAL,
        run_status="Pass"
    ),
    
    TestStep(
        command="INITiate:NR:MEValuation",
        description="初始化ME评估",
        step_type=StepType.NORMAL,
        run_status="Pass"
    ),
    
    TestStep(
        command="FETCh:NR:MEValuation:STATe?",
        expected_result="RDY",
        loop_config=LoopConfig(
            enable=True,
            times=10,
            sleep_ms=1000,
            expected_result="RDY"
        ),
        description="查询ME评估状态,等待就绪",
        step_type=StepType.QUERY,
        run_status="Pass"
    ),
    
    TestStep(
        command="FETCh:NR:BLER:STATe?",
        expected_result="RUN",
        loop_config=LoopConfig(
            enable=True,
            times=10,
            sleep_ms=1000,
            expected_result="RDY"
        ),
        description="查询BLER状态,等待就绪",
        step_type=StepType.QUERY,
        run_status="Pass"  
    ),
    
    TestStep(
        command="Sleep 5000",
        description="等待数据稳定",
        step_type=StepType.SLEEP,
        run_status="Pass"
    ),
    
    TestStep(
        command="FETCh:NR:BLER:DL:RESult?",
        expected_result="0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
        description="查询BLER下行结果",
        step_type=StepType.QUERY,
        run_status="Pass"
    ),
    
    TestStep(
        command="FETCh:NR:BLER:UL:RESult?",
        expected_result="0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
        description="查询BLER上行结果",
        step_type=StepType.QUERY,
        run_status="Pass"
    ),
    
    TestStep(
        command="CONFigure:NR:MEValuation:REPetition CONTINUOUS",
        description="配置ME评估为连续模式",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    
    TestStep(
        command="CONFigure:NR:BLER:REPetition CONTINUOUS",
        description="配置BLER为连续模式",
        step_type=StepType.CONFIGURE,
        run_status="Pass"
    ),
    
    TestStep(
        command="Sleep 20000",
        description="等待连续模式运行",
        step_type=StepType.SLEEP,
        run_status="Pass"
    ),
    
    TestStep(
        command="FETCh:NR:MEValuation:TXP:AVG?",
        expected_result="0,15.700",
        description="再次查询发射功率平均值",
        step_type=StepType.QUERY,
        run_status="Pass"
    ),
    
    TestStep(
        command="ABORt:NR:BLER ",
        description="中止BLER测试",
        step_type=StepType.NORMAL,
        run_status="Pass"
    ),
    
    TestStep(
        command="ABORt:NR:MEValuation",
        description="中止ME评估",
        step_type=StepType.NORMAL,
        run_status="Pass"
    ),
    
    TestStep(
        command="CALL:CELL1 OFF",
        description="关闭CELL1",
        step_type=StepType.CALL,
        run_status="Pass"
    ),
    
    TestStep(
        command="Sleep 5000",
        description="等待CELL1关闭",
        step_type=StepType.SLEEP,
        run_status="Pass"
    ),
    
    TestStep(
        command="CALL:CELL1?",
        expected_result="OFF",
        loop_config=LoopConfig(
            enable=True,
            times=20,
            sleep_ms=1000,
            expected_result="OFF"
        ),
        description="确认CELL1已关闭",
        step_type=StepType.QUERY,
        run_status="Pass"
    ),
]


# ==============================================
# 统计和分析函数
# ==============================================

def get_statistics() -> Dict:
    """获取步骤统计信息"""
    stats = {
        "total_steps": len(TEST_STEPS),
        "step_types": {},
        "commands_with_loop": 0,
        "commands_with_expected": 0,
        "estimated_duration_ms": 0
    }
    
    # 统计步骤类型
    for step in TEST_STEPS:
        step_type = step.step_type.value
        stats["step_types"][step_type] = stats["step_types"].get(step_type, 0) + 1
        
        # 统计循环命令
        if step.loop_config.enable:
            stats["commands_with_loop"] += 1
        
        # 统计有预期结果的命令
        if step.expected_result or step.loop_config.expected_result:
            stats["commands_with_expected"] += 1
        
        # 估算总时长（仅统计SLEEP命令）
        if step.step_type == StepType.SLEEP:
            import re
            sleep_match = re.search(r'SLEEP\s+(\d+)', step.command.upper())
            if sleep_match:
                stats["estimated_duration_ms"] += int(sleep_match.group(1))
        
        # 加上循环等待时间
        if step.loop_config.enable:
            stats["estimated_duration_ms"] += step.loop_config.times * step.loop_config.sleep_ms
    
    stats["estimated_duration_sec"] = stats["estimated_duration_ms"] / 1000
    return stats


def print_statistics() -> None:
    """打印统计信息"""
    stats = get_statistics()
    
    print("=" * 60)
    print("📊 测试步骤统计")
    print("=" * 60)
    print(f"总步骤数: {stats['total_steps']}")
    print(f"预计总时长: {stats['estimated_duration_sec']:.1f}秒")
    print()
    print("步骤类型分布:")
    for step_type, count in stats["step_types"].items():
        print(f"  {step_type}: {count}")
    print()
    print(f"循环步骤: {stats['commands_with_loop']}")
    print(f"有预期结果的步骤: {stats['commands_with_expected']}")
    print("=" * 60)


def print_all_steps(show_details: bool = False) -> None:
    """打印所有步骤"""
    print("=" * 60)
    print("📋 所有测试步骤")
    print("=" * 60)

    for i, step in enumerate(TEST_STEPS, 1):
        print(f"\n[{i}] {step.command}")
        if show_details:
            if step.description:
                print(f"   描述: {step.description}")
            if step.expected_result:
                print(f"   预期: {step.expected_result}")
            if step.loop_config.enable:
                loop_info = f"   循环: {step.loop_config.times}次"
                if step.loop_config.sleep_ms > 0:
                    loop_info += f" (间隔{step.loop_config.sleep_ms}ms)"
                if step.loop_config.expected_result:
                    loop_info += f" -> {step.loop_config.expected_result}"
                print(loop_info)
            print(f"   类型: {step.step_type.value}")
            print(f"   状态: {step.run_status}")
    
    print("=" * 60)




# ==============================================
# 主程序
# ==============================================

if __name__ == "__main__":
    """主程序 - 演示如何使用"""
    
    print("🔧 测试步骤配置系统")
    print("=" * 60)
    
    while True:
        print("\n请选择操作:")
        print("  1. 显示所有步骤")
        print("  2. 显示详细步骤")
        print("  3. 显示统计信息")
        print("  0. 退出")
        
        choice = input("请输入选项: ").strip()
        
        if choice == "1":
            print_all_steps(show_details=False)
        elif choice == "2":
            print_all_steps(show_details=True)
        elif choice == "3":
            print_statistics()
        elif choice == "0":
            print("再见！")
            break
        else:
            print("无效选项，请重新输入")
