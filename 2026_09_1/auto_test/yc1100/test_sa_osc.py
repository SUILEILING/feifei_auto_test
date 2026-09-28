from lib.var import *
from common import *

DEFAULT_PARAMETER = {
    'lineLoss1': 25.00,
    'nr_band': 1,
    'nr_bw': 20,
    'scs': 15,
    'range': 'LOW',
    'duty_cycle': 10,               
    'duty_cycle_List': [5, 10, 20],  
    'rb_mode': 'Inner_Full',
}


VALID_DUTY_CYCLES = [5, 10, 20, 25, 30, 40, 50]

parameter = DEFAULT_PARAMETER.copy()

def update_parameters(external_params=None):
    global parameter
    if external_params:
        for key, value in external_params.items():
            if key in parameter:
                parameter[key] = value


def case_start():
    remote_gnb_start()
    remote_diag_start()

    ## line loss configuration
    config_line_loss(parameter)

    ## band bw scs range configuration
    config_cell_band(parameter)

    duty = parameter.get('duty_cycle', 10)
    if duty not in VALID_DUTY_CYCLES:
        print(f"⚠️ duty_cycle={duty} 不在合法档位 {VALID_DUTY_CYCLES}，仍尝试设置")
    duty_set_ok = remote_set_duty_cycle(duty)
    ap.check(f"初始 TDD Uplink Duty Cycle = {duty}%", bool(duty_set_ok),
                  detail=f"TDD Uplink Duty Cycle={duty}%",
                 status_msg=f"set TDD Uplink Duty Cycle {duty}% {'pass' if duty_set_ok else 'fail'}")


def case_body():
    ap.send("CALL:CELL1 ON")
    check_phone_at()
    my_sleep(5)

    connected = False
    for i in range(20):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            connected = True
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接")
            my_sleep(2)

    if not connected:
        print("❌ UE 多次未连接，跳过后续测量，直接进入 case_clear")
        ap.check("UE连接", False, detail="多次查询UE未连接", status_msg="UE未连接,测试判定失败")
        return

    init_duty = parameter.get('duty_cycle', 10)
    ap.send(" CONFigure:NR:MEValuation:REPetition CONTINUOUS ")
    ap.send(" CONFigure:GPRF:OSC:REPetition CONTINUOUS ")
    ap.send(" INITiate:GPRF:OSC ")
    my_sleep(2)
    osc_file = remote_osc_screenshot(tag=f"{init_duty}_init")
    ap.check(f"OSC 截图 初始(duty={init_duty}%)", osc_file is not None,
             detail=f"OSC duty{init_duty}%截图{'成功' if osc_file else '失败'}",
             status_msg=f"OSC screenshot init duty{init_duty}% {'pass' if osc_file else 'fail'}")
    ap.send(" ABORt:GPRF:OSC ")
    my_sleep(2)

    duty_list = parameter.get('duty_cycle_List', [10])
    for idx, duty_var in enumerate(duty_list, 1):
        if duty_var not in VALID_DUTY_CYCLES:
            print(f"⚠️ duty_cycle={duty_var} 不在合法档位 {VALID_DUTY_CYCLES}，仍尝试设置")

        duty_set_ok = remote_set_duty_cycle(duty_var)
        ap.check(f"TDD Uplink Duty Cycle #{idx} = {duty_var}%", bool(duty_set_ok),
                 detail=f"TDD Uplink Duty Cycle={duty_var}%",
                 status_msg=f"set TDD Uplink Duty Cycle {duty_var}% {'pass' if duty_set_ok else 'fail'}")

        ap.send(" CONFigure:NR:MEValuation:REPetition CONTINUOUS ")
        ap.send(" CONFigure:GPRF:OSC:REPetition CONTINUOUS ")
        ap.send(" INITiate:GPRF:OSC ")
        my_sleep(2)

        osc_file = remote_osc_screenshot(tag=f"{duty_var}_{idx}")
        ap.check(f"OSC 截图 #{idx} (duty={duty_var}%)", osc_file is not None,
                 detail=f"OSC duty{duty_var}%截图{'成功' if osc_file else '失败'}",
                 status_msg=f"OSC screenshot #{idx}")

        ap.send(" ABORt:GPRF:OSC ")
        my_sleep(2)


def case_clear():
    ap.send(" ABORt:GPRF:OSC ")
    ap.send(" CALL:CELL1 OFF ")
    my_sleep(1)

    remote_diag_stop()

    for i in range(5):
        result = ap.query("CALL:CELL1?")
        if "OFF" == result:
            print(f"✅ CELL已关闭")
            break
        else:
            print(f"⏳ 等待CELL关闭...")
            my_sleep(2)

    remote_gnb_stop()
    remote_restart()
