from lib.var import *
from common import *

DEFAULT_PARAMETER = {
    'lineLoss1': 25.00,
    'lineLoss3': None,
    'nr_band': 1,
    'lte_band': None,
    'nr_bw': 20,
    'lte_bw': None,
    'scs': 15,
    'range': 'LOW',
    "nr_start_power": -40,
    "nsa_start_power": -40,
    "end_power": -100,
    "step": -2,
    "fallback_delta": 10,
    "nr_slots": None,
    "mcs": None,
    "lte_slots": None,
    "rb_mode": "Inner_Full",
}

NR_SLOTS = {
    "DL": [
        {3:  {"TIND": 5, "MCS1": 4}},
        {4:  {"TIND": 4, "MCS1": 4}},
        {5:  {"TIND": 3, "MCS1": 4}},
        {6:  {"TIND": 2, "MCS1": 4}},
        {10: {"TIND": 8, "MCS1": 4}},
        {11: {"TIND": 7, "MCS1": 4}},
        {12: {"TIND": 6, "MCS1": 4}},
        {13: {"TIND": 5, "MCS1": 4}},
        {14: {"TIND": 4, "MCS1": 4}},
        {15: {"TIND": 3, "MCS1": 4}},
        {16: {"TIND": 2, "MCS1": 4}},
    ],
    "UL": [
        {8:  {"MCS1": 2}},
        {9:  {"MCS1": 2}},
        {18: {"MCS1": 2}},
        {19: {"MCS1": 2}},
    ],
}


LTE_SLOTS = {
    "DL": [
        {3: {}},
        {4: {}},
    ],
    "UL": [
        {8: {}},
        {9: {}},
    ],
}

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

def case_body():
    ap.send("CALL:CELL1 ON")
    check_phone_at()
    my_sleep(5)

    connected = False
    for i in range(10):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            connected = True
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接")
            my_sleep(2)

    if not connected:
        print("⚠️ 首轮未连接，重新触发注册后重试...")
        check_phone_at()
        my_sleep(5)
        for i in range(50):
            result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
            if '"Connected"' == result:
                print(f"✅ 重试第 {i+1} 次查询: UE已连接")
                connected = True
                break
            else:
                my_sleep(0.25)
                print(f"⏳ 重试第 {i+1} 次查询: UE未连接")

    if not connected:
        print("❌ UE 多次未连接，跳过后续功率扫描，直接进入 case_clear")
        ap.check("UE连接", False, detail="多次查询UE未连接", status_msg="UE未连接,测试判定失败")
        return

    calibrate_line_loss(parameter, tolerance=3.0)

    # config nr
    config_nr_slots(parameter, NR_SLOTS)

    # config lte
    config_lte_subframes(parameter, LTE_SLOTS)

    def wait_ready(state_query, label):
        for i in range(5):
            result = ap.query(state_query)
            if "RDY" == result:
                print(f"✅ 第 {i+1} 次查询: {label}已准备好")
                break
            else:
                print(f"⏳ 第 {i+1} 次查询: {label}未准备好")

    nr_start_power = parameter['nr_start_power']
    nsa_start_power = parameter['nsa_start_power']
    end_power = parameter['end_power']
    step = parameter['step']
    fallback_delta = parameter.get('fallback_delta', 10)

    def run_power_scan(tech, start_power, end_power, step, fallback_delta, ap):
        power_cmd = f"CONFigure:CELL1:{tech}:SIGN:POWer"
        meas_init_cmd = f"INITiate:{tech}:MEValuation" if tech == "NR" else f"INITiate:{tech}:TXP"
        bler_dl_title = f"DL {tech}_BLER"
        bler_ul_title = f"UL {tech}_BLER"
        txp_title = f"{tech} TXP AVG"
        txp_query = "FETCh:NR:MEValuation:TXP:AVG?" if tech == "NR" else "FETCh:LTE:TXP:AVG?"

        def _single_measure(power):
            ap.send(f"{power_cmd} {power}")
            ap.send(f"INITiate:{tech}:BLER")
            ap.send(meas_init_cmd)
            my_sleep(0.25) # ap.send("*OPC?")

            if tech == "NR":
                ap.send("CONFigure:CELL1:NR:UEReport:RSRP?", 1, True, "Reasonable Line Loss", f"{power}dBm", True, True)
            dl_bler_str = ap.send(f"FETCh:{tech}:BLER:DL:RESult?", 7, True, bler_dl_title, f"{power}dBm", True, True)
            ul_bler_str = ap.send(f"FETCh:{tech}:BLER:UL:RESult?", 7, True, bler_ul_title, f"{power}dBm", True, True)
            dl_bler = float(dl_bler_str.split(',')[7])
            ul_bler = float(ul_bler_str.split(',')[7])
            UE_status = ap.send("CONFigure:CELL1:NR:SIGN:UE:STATe?")
            return dl_bler, ul_bler, UE_status

        def measure(power):
            dl_bler, ul_bler, UE_status = _single_measure(power)

            bler_fail = (dl_bler > 0.2) or (ul_bler > 0.2)
            if not bler_fail and UE_status != '"Connected"':
                print(f"[{tech}] 功率: {power} dBm, BLER正常但UE掉线({UE_status})，等待重连后复测...")
                recovered = False
                for _ in range(5):
                    my_sleep(2)
                    if ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?") == '"Connected"':
                        recovered = True
                        break
                if recovered:
                    dl_bler, ul_bler, UE_status = _single_measure(power)

            print(f"[{tech}] 功率: {power} dBm, UE状态: {UE_status}, DL BLER: {dl_bler}  UL BLER: {ul_bler}")
            return (dl_bler > 0.2) or (ul_bler > 0.2) or (UE_status != '"Connected"')

        def mark(status):
            ap.tag_last(status, 2)

        failed_powers = set()
        current_power = start_power

        while current_power >= end_power:
            if not measure(current_power):
                ap.send(txp_query, 1, True, txp_title, f"{current_power}dBm", True, True)
                current_power += step
                continue

            if current_power in failed_powers:
                print(f"✅ [{tech}] 功率 {current_power} dBm 第二次出现异常，判定为稳定边界，记录并停止")
                mark("normal")
                ap.check(f"{tech} 功率扫描", True,
                         detail=f"稳定边界 {current_power}dBm",
                         status_msg=f"{tech} power scan 正常 pass")
                return True

            failed_powers.add(current_power)
            new_power = current_power + fallback_delta
            if new_power > start_power:
                new_power = start_power
            print(f"🔄 [{tech}] 功率 {current_power} dBm 首次异常，回退到 {new_power} dBm 复测")
            my_sleep(1.5)
             

            if measure(new_power):
                ue_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
                recovered = False

                if ue_status != '"Connected"':
                    print(f"🔁 [{tech}] 回退到 {new_power} dBm 后 UE 掉线,先每秒检查状态,最多等待200秒...")
                    reconnected = False
                    start_time = time.time()
                    while time.time() - start_time < 200:
                        my_sleep(1)
                        ue_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
                        if ue_status == '"Connected"':
                            print(f"✅ [{tech}] 200秒内重连成功")
                            reconnected = True
                            break
                    if not reconnected:
                        print(f"🔁 [{tech}] 200秒内未连上,触发check_phone_at后再每秒检查状态,最多等待100秒...")
                        check_phone_at()
                        start_time = time.time()
                        while time.time() - start_time < 100:
                            my_sleep(1)
                            ue_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
                            if ue_status == '"Connected"':
                                print(f"✅ [{tech}] check_phone_at后重连成功")
                                reconnected = True
                                break
                    if reconnected:
                        if not measure(new_power):
                            recovered = True
                            print(f"✅ [{tech}] 重连后 {new_power} dBm 复测通过")
                        else:
                            print(f"❌ [{tech}] 重连后 {new_power} dBm 仍异常")
                    else:
                        print(f"❌ [{tech}] 触发check_phone_at后仍未连上")

                if not recovered:
                    print(f"❌ [{tech}] 回退到 {new_power} dBm 后仍异常，记录并停止")
                    mark("abnormal")
                    ap.check(f"{tech} 功率扫描", False,
                            detail=f"回退到 {new_power}dBm 后仍异常",
                            status_msg=f"{tech} power scan 回退后异常 fail")
                    return True

            current_power = new_power + step

        ap.check(f"{tech} 功率扫描", True,
                 detail=f"扫描完成 {start_power}~{end_power}dBm 无异常",
                 status_msg=f"{tech} power scan 扫描完成 pass")
        return False

    # ========== SA  ==========
    ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")
    ap.send("INITiate:NR:BLER")
    ap.send("INITiate:NR:MEValuation")
    my_sleep(0.5)

    wait_ready("FETCh:NR:MEValuation:STATe?", "MEValuation")
    wait_ready("FETCh:NR:BLER:STATe?", "BLER测试")

    nr_should_stop = run_power_scan("NR", nr_start_power, end_power, step, fallback_delta, ap)
    if nr_should_stop:
        print("NR测试因连续失败而终止")

    ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {nr_start_power}")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")

    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    ap.send("CONFigure:NR:MEValuation:RESult OFF,OFF,OFF,OFF,OFF,OFF")

    # ========== NSA/LTE  ==========
    if parameter.get('lte_band') is not None:
        ap.send("CONFigure:LTE:TXP:REPetition SINGLESHOT")
        ap.send("CONFigure:LTE:BLER:REPetition SINGLESHOT")
        ap.send("INITiate:LTE:BLER")
        ap.send("INITiate:LTE:TXP")
        my_sleep(0.5)

        wait_ready("FETCh:LTE:TXP:STATe?", "LTE TXP")
        wait_ready("FETCh:LTE:BLER:STATe?", "LTE BLER测试")

        nsa_should_stop = run_power_scan("LTE", nsa_start_power, end_power, step, fallback_delta, ap)
        if nsa_should_stop:
            print("NSA测试因连续失败而终止")

        ap.send(f"CONFigure:CELL1:LTE:SIGN:POWer {nsa_start_power}")
        ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame:APPLy")

        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")

def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    if parameter.get('lte_band') is not None:
        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")

    ap.send("CALL:CELL1 OFF")
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
