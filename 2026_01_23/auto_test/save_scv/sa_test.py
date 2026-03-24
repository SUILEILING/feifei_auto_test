from typing import List, Dict


TEST_STEPS: List[Dict] = [
    {
        "command": "CALL:CELL1 ON",
        "description": "开启CELL1",
        "step_type": "CALL",
        "run_status": "Pass"
    },
    
    {
        "command": "Sleep 10000",
        "description": "等待CELL1启动",
        "step_type": "SLEEP",
        "run_status": "Pass"
    },
    
    {
        "command": "CONFigure:CELL1:NR:SIGN:UE:STATe?",
        "expected_result": "Connected",
        "loop_config": {
            "enable": True,
            "times": 20,
            "sleep_ms": 1000,
            "expected_result": "Connected"
        },
        "description": "查询UE连接状态,最多尝试20次",
        "step_type": "QUERY",
        "run_status": "Pass"
    },
    
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT:CLEar",
        "description": "清除时隙配置",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh",
        "description": "配置时隙3为PDSCh",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh",
        "description": "配置时隙4为PDSCh",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 3",
        "description": "配置时隙4为PDSCh tind 3",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh",
        "description": "配置时隙8为PUSCh",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh",
        "description": "配置时隙9为PUSCh",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT:UPDate",
        "description": "时隙配置更新",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT:APPLy",
        "description": "应用时隙配置",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },

    {
        "command": "CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF",
        "description": "配置ME评估结果",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:NR:MEValuation:REPetition SINGLESHOT",
        "description": "配置ME评估为单次模式",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    
    {
        "command": "CONFigure:NR:BLER:REPetition SINGLESHOT",
        "description": "配置BLER为单次模式",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    
    {
        "command": "INITiate:NR:BLER",
        "description": "初始化BLER测试",
        "step_type": "NORMAL",
        "run_status": "Pass"
    },
    
    {
        "command": "INITiate:NR:MEValuation",
        "description": "初始化ME评估",
        "step_type": "NORMAL",
        "run_status": "Pass"
    },
    
    {
        "command": "FETCh:NR:MEValuation:STATe?",
        "expected_result": "RDY",
        "loop_config": {
            "enable": True,
            "times": 10,
            "sleep_ms": 1000,
            "expected_result": "RDY"
        },
        "description": "查询ME评估状态,等待就绪",
        "step_type": "QUERY",
        "run_status": "Pass"
    },
    
    {
        "command": "FETCh:NR:BLER:STATe?",
        "expected_result": "RUN",
        "loop_config": {
            "enable": True,
            "times": 10,
            "sleep_ms": 1000,
            "expected_result": "RDY"
        },
        "description": "查询BLER状态,等待就绪",
        "step_type": "QUERY",
        "run_status": "Pass"  
    },
    
    {
        "command": "Sleep 5000",
        "description": "等待数据稳定",
        "step_type": "SLEEP",
        "run_status": "Pass"
    },
    
    {
        "command": "FETCh:NR:BLER:DL:RESult?",
        "expected_result": "0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
        "description": "查询BLER下行结果",
        "step_type": "QUERY",
        "run_status": "Pass"
    },
    
    {
        "command": "FETCh:NR:BLER:UL:RESult?",
        "expected_result": "0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
        "description": "查询BLER上行结果",
        "step_type": "QUERY",
        "run_status": "Pass"
    },
    
    {
        "command": "CONFigure:NR:MEValuation:REPetition CONTINUOUS",
        "description": "配置ME评估为连续模式",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    
    {
        "command": "CONFigure:NR:BLER:REPetition CONTINUOUS",
        "description": "配置BLER为连续模式",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    
    {
        "command": "Sleep 20000",
        "description": "等待连续模式运行",
        "step_type": "SLEEP",
        "run_status": "Pass"
    },
    
    {
        "command": "FETCh:NR:MEValuation:TXP:AVG?",
        "expected_result": "0,15.700",
        "description": "再次查询发射功率平均值",
        "step_type": "QUERY",
        "run_status": "Pass"
    },
    
    {
        "command": "ABORt:NR:BLER ",
        "description": "中止BLER测试",
        "step_type": "NORMAL",
        "run_status": "Pass"
    },
    
    {
        "command": "ABORt:NR:MEValuation",
        "description": "中止ME评估",
        "step_type": "NORMAL",
        "run_status": "Pass"
    },
    
    {
        "command": "CALL:CELL1 OFF",
        "description": "关闭CELL1",
        "step_type": "CALL",
        "run_status": "Pass"
    },
    
    {
        "command": "Sleep 5000",
        "description": "等待CELL1关闭",
        "step_type": "SLEEP",
        "run_status": "Pass"
    },
    
    {
        "command": "CALL:CELL1?",
        "expected_result": "OFF",
        "loop_config": {
            "enable": True,
            "times": 20,
            "sleep_ms": 1000,
            "expected_result": "OFF"
        },
        "description": "确认CELL1已关闭",
        "step_type": "QUERY",
        "run_status": "Pass"
    },
]

# ==============================================
# 需要在后续循环中跳过的步骤（第一次执行后不再执行）
# ==============================================

SKIP_IN_NEXT_CYCLES: List[Dict] = [
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT:CLEar",
        "description": "清除时隙配置",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh",
        "description": "配置时隙3为PDSCh",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh",
        "description": "配置时隙4为PDSCh",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 3",
        "description": "配置时隙4为PDSCh tind 3",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh",
        "description": "配置时隙8为PUSCh",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh",
        "description": "配置时隙9为PUSCh",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT:UPDate",
        "description": "时隙配置更新",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT:APPLy",
        "description": "应用时隙配置",
        "step_type": "CONFIGURE",
        "run_status": "Pass"
    },
]



def get_all_test_steps():
    """获取所有测试步骤"""
    return TEST_STEPS.copy()


def get_steps_for_cycle(cycle_index: int = 0):

    if cycle_index == 0:
        # 第一次执行：所有步骤
        return TEST_STEPS.copy()
    else:
        # 后续循环：跳过指定的步骤
        skip_commands = {step["command"] for step in SKIP_IN_NEXT_CYCLES}
        
        # 过滤掉需要跳过的步骤
        filtered_steps = [
            step for step in TEST_STEPS 
            if step["command"] not in skip_commands
        ]
        
        return filtered_steps


def get_skipped_commands():
    """获取需要跳过的命令列表"""
    return [step["command"] for step in SKIP_IN_NEXT_CYCLES]


def get_executed_commands(cycle_index: int = 0):
    """获取实际执行的命令列表"""
    steps = get_steps_for_cycle(cycle_index)
    return [step["command"] for step in steps]


# ==============================================
# 主程序（仅用于测试）
# ==============================================

if __name__ == "__main__":

    # 显示统计信息
    total_steps = len(TEST_STEPS)
    skip_steps = len(SKIP_IN_NEXT_CYCLES)
    
    print(f"📊 总步骤数: {total_steps}")
    print(f"📊 需要跳过的步骤数: {skip_steps}")
    print(f"📊 后续循环执行的步骤数: {total_steps - skip_steps}")
    
    # 演示不同循环的执行情况
    print("\n🔁 循环执行演示:")
    print("-" * 60)
    
    for cycle in range(3):  # 演示3次循环
        steps = get_steps_for_cycle(cycle)
        print(f"\n第 {cycle + 1} 次循环:")
        print(f"  执行步骤数: {len(steps)}")
        
        if cycle > 0:
            skipped = get_skipped_commands()
            print(f"  跳过的步骤: {len(skipped)} 个")
            print(f"  跳过的命令示例: {skipped[:3]}...")
        
        # 显示前几个执行的命令
        executed = get_executed_commands(cycle)
        print(f"  执行的命令示例: {executed[:3]}...")
    
    print("\n📋 需要跳过的完整命令列表:")
    for i, cmd in enumerate(get_skipped_commands(), 1):
        print(f"  {i:2d}. {cmd}")
    
    print("\n📋 第一次循环执行的所有命令:")
    for i, cmd in enumerate(get_executed_commands(0), 1):
        print(f"  {i:2d}. {cmd}")