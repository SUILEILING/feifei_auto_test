from lib.var import *
from common import *

DEFAULT_PARAMETER  = {
    'lineLoss1': 25.00,
    'lineLoss3': None,
    'nr_band': 1,
    'lte_band': None,
    'nr_bw': 20,
    'lte_bw': None,
    'scs': 15,
    'range': 'LOW',
    "TD":600,
    "resource_allocation_type":None,
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

parameter = DEFAULT_PARAMETER.copy()

def update_parameters(external_params=None):
    global parameter
    
    if external_params:
        for key, value in external_params.items():
            if key in parameter:
                parameter[key] = value

def monitor_bler(ap, tech="NR", duration=60, threshold=0.2, fixed_power="-49.847 dBm"):
    meas = "MEValuation" if tech == "NR" else "TXP"

    start_time = time.time()
    end_time = start_time + duration

    while time.time() < end_time:
        try:
            ap.send(f"CONFigure:{tech}:{meas}:REPetition SINGLESHOT")
            ap.send(f"CONFigure:{tech}:BLER:REPetition SINGLESHOT")

            ap.send(f"INITiate:{tech}:BLER")
            ap.send(f"INITiate:{tech}:{meas}")
            my_sleep(0.25)

            bler_ready = False
            for i in range(5):
                result = ap.query(f"FETCh:{tech}:BLER:STATe?")
                print(f"{tech} BLER state: {result}")
                if "RDY" == result:
                    print(f"✅ 第 {i+1} 次查询: {tech} BLER测试已准备好")
                    bler_ready = True
                    break
                else:
                    print(f"⏳ 第 {i+1} 次查询: {tech} BLER测试未准备好,状态={result}")
                    if i < 4:  
                        my_sleep(2)

            if not bler_ready:
                print(f"❌ {tech} BLER测试超时未完成,跳过本次测量")
                continue

            dl_bler_str = ap.send(f"FETCh:{tech}:BLER:DL:RESult?", 7, True, f"DL {tech}_BLER", fixed_power, False, False)
            ul_bler_str = ap.send(f"FETCh:{tech}:BLER:UL:RESult?", 7, True, f"UL {tech}_BLER", fixed_power, False, False)

            dl_bler = float(dl_bler_str.split(',')[7])
            ul_bler = float(ul_bler_str.split(',')[7])

        except Exception as e:
            print(f"读取 {tech} BLER 失败: {e}")
            time.sleep(0.5)
            continue

        if dl_bler >= threshold or ul_bler >= threshold:
            print(f"{tech} BLER 超过阈值！ dl_bler={dl_bler:.4f}, ul_bler={ul_bler:.4f}")
            return False

        time.sleep(0.25)

    print(f"{tech} 监控完成，{duration} 秒内 BLER 均未超过 {threshold}")
    return True


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

    for i in range(20):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?") 
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接")
            my_sleep(2)
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:CLEar")

    # config nr
    config_nr_slots(parameter, NR_SLOTS)

    # config lte
    config_lte_subframes(parameter)

    fixed_power = ap.send(f"CONFigure:CELL1:NR:SIGN:POWer?")

    # ========== NR ==========
    ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")
    monitor_bler(ap, tech="NR", duration=parameter['TD'], threshold=0.2, fixed_power=fixed_power)
    nr_txp = ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True, "NR TXP AVG")

    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    ap.send("CONFigure:NR:MEValuation:RESult OFF,OFF,OFF,OFF,OFF,OFF")

    # ========== NSA ==========
    if parameter['lte_band'] is not None:
        monitor_bler(ap, tech="LTE", duration=parameter['TD'], threshold=0.2, fixed_power=fixed_power)
        lte_txp = ap.send("FETCh:LTE:TXP:AVG?", 1, True, "LTE TXP AVG")

        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")


def case_clear():

    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    if parameter['lte_band'] is not None:
        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")
    # if parameter['nr_band']<40 and  parameter['lineLoss1']<15 :
        # ap.send("CONFigure:CELL1:NR:SIGN:PDUSession:REQuest DISABLE ")
    
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