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

    current_power = start_power
    consecutive_failures = 0            

    while True:
        if current_power < end_power:
            break

        ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {current_power}")
        ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
        ap.send("INITiate:NR:BLER")
        ap.send("INITiate:NR:MEValuation")
        my_sleep(0.5)

        dl_bler_str=ap.send("FETCh:NR:BLER:DL:RESult?", 7, True, "DL BLER", f"{current_power}") # 0,67,67,0,0,57955,57955,0.000,0.000,173.000,173.000
        dl_bler=float(dl_bler_str.split(',')[7])
        UE_status = ap.send("CONFigure:CELL1:NR:SIGN:UE:STATe?") # '"Connected"'

        print(f"当前功率: {current_power} dBm, UE状态: {UE_status}, DL BLER: {dl_bler}")

        is_failure = (dl_bler > 0.5) or (UE_status != '"Connected"')

        if is_failure:
            consecutive_failures += 1
            if consecutive_failures >= 2:
                print(f"⚠️ 连续两次失败，复位连接...")
                ap.send("CALL:CELL1 OFF")
                my_sleep(5)
                ap.send("CALL:CELL1 ON")
                check_phone_at()
                my_sleep(5)
                consecutive_failures = 0   # 复位后清除计数，但功率点不变
                continue
            else:
                current_power += fallback_delta  

                if current_power > start_power:
                    current_power = start_power
                continue
        else:
            consecutive_failures = 0
            current_power += step


def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    ap.send(f"CONFigure:CELL1:NR:SIGN:POWer -49.847 ")

    ap.send("CALL:CELL1 OFF")
    my_sleep(5)

    for i in range(5):
        result = ap.query("CALL:CELL1?") #  'OFF'
        if "OFF" == result:
            print(f"✅ CELL已关闭")
            break
        else:
            print(f"⏳ 等待CELL关闭...")
            my_sleep(2)
            
    ap.send("*rst")
    my_sleep(5)