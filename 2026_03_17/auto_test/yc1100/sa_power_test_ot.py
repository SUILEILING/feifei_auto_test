from lib.var import *
from common import ap, check_phone_at, my_sleep

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

DEFAULT_PARAMETER = {
    'lineLoss': 25.00,
    'band': 1,
    'bw': 20,
    'scs': 15,
    'range': 'LOW',
    "start_power": -40,         
    "end_power": -100,          
    "step": -2,                  
    "fallback_delta": 10,       
}

parameter = DEFAULT_PARAMETER.copy()

def update_parameters(external_params=None):
    global parameter
    if external_params:
        for key, value in external_params.items():
            if key in parameter:
                parameter[key] = value

def case_start():
    # line loss configuration
    ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,{parameter['lineLoss']},6000000000,{parameter['lineLoss']}")
    ap.send("CONFigure:BASE:FDCorrection:SAVE")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX")

    # band bw scs range configuration
    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {parameter['band']}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{parameter['bw']}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{parameter['scs']}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:CONFig:RANGe {parameter['range']}")

def case_body():
    ap.send("CALL:CELL1 ON")
    check_phone_at()
    my_sleep(5)

    for i in range(10):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")  # '"Connected"'
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接")
            my_sleep(2)

    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:CLEar")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT3:DL:TIND 5")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 4")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT5:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT5:DL:TIND 3")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT6:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT6:DL:TIND 2")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT10:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT10:DL:TIND 8")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT11:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT11:DL:TIND 7")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT12:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT12:DL:TIND 6")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT13:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT13:DL:TIND 5")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT14:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT14:DL:TIND 4")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT15:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT15:DL:TIND 3")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT16:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT16:DL:TIND 2")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
    my_sleep(2)

    ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")
    ap.send("INITiate:NR:BLER")
    ap.send("INITiate:NR:MEValuation")

    for i in range(5):
        result = ap.query("FETCh:NR:MEValuation:STATe?")   # 'RDY'
        if "RDY" == result:
            print(f"✅ 第 {i+1} 次查询: MEValuation已准备好")
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: MEValuation未准备好")
            my_sleep(2)

    for i in range(5):
        result = ap.query("FETCh:NR:BLER:STATe?")   #'RDY'
        if "RDY" == result:
            print(f"✅ 第 {i+1} 次查询: BLER测试已准备好")
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: BLER测试未准备好")
            my_sleep(2)

    start_power = parameter['start_power']
    end_power = parameter['end_power']
    step = parameter['step']
    fallback_delta = parameter.get('fallback_delta', 10)

    for attempt in range(3):
        current_power = start_power
        consecutive_failures = 0
        failed_powers = set()   

        while True:
            if current_power < end_power:
                break

            ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {current_power}")
            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
            ap.send("INITiate:NR:BLER")
            ap.send("INITiate:NR:MEValuation")
            my_sleep(0.05)   

            dl_bler_str = ap.send("FETCh:NR:BLER:DL:RESult?", 7, True, "DL BLER", f"{current_power}")
            dl_bler=float(dl_bler_str.split(',')[7])
            UE_status = ap.send("CONFigure:CELL1:NR:SIGN:UE:STATe?")

            print(f"功率: {current_power} dBm, UE状态: {UE_status}, DL BLER: {dl_bler:.4f}")

            is_failure = (dl_bler > 0.2) or (UE_status != '"Connected"')

            if is_failure:
                if current_power in failed_powers:
                    break
                failed_powers.add(current_power)

                consecutive_failures += 1

                if consecutive_failures >= 2:
                    print(f"❌ 连续两次失败，结束第 {attempt+1} 次大循环")
                    break

                new_power = current_power + fallback_delta
                if new_power > start_power:
                    new_power = start_power
                print(f"🔄 回退功率: {current_power} -> {new_power} dBm")

                ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {new_power}")
                ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
                ap.send("INITiate:NR:BLER")
                ap.send("INITiate:NR:MEValuation")
                my_sleep(0.05)

                dl_bler_retry_str = ap.send("FETCh:NR:BLER:DL:RESult?", 7, True, "DL BLER", f"{new_power}")
                dl_bler_retry=float(dl_bler_retry_str.split(',')[7])
                UE_status_retry = ap.send("CONFigure:CELL1:NR:SIGN:UE:STATe?")

                is_failure_retry = (dl_bler_retry > 0.2) or (UE_status_retry != '"Connected"')

                if is_failure_retry:
                    if new_power in failed_powers:
                        print(f"❌ 回退后功率 {new_power} dBm 已经失败过，结束第 {attempt+1} 次大循环")
                    else:
                        failed_powers.add(new_power)
                        print(f"❌ 回退后仍失败，结束第 {attempt+1} 次大循环")
                    break
                else:
                    current_power = new_power
                    consecutive_failures = 0
                    current_power += step
                    continue
            else:
                consecutive_failures = 0
                current_power += step


def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    ap.send(f"CONFigure:CELL1:NR:SIGN:POWer -49.847")

    ap.send("CALL:CELL1 OFF")
    my_sleep(5)

    for i in range(5):
        result = ap.query("CALL:CELL1?") 
        if "OFF" == result:
            print(f"✅ CELL已关闭")
            break
        else:
            print(f"⏳ 等待CELL关闭...")
            my_sleep(2)

    ap.send("*rst")
    my_sleep(5)