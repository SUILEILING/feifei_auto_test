from regex import F

from lib.var import *
from common import ap, check_phone_at, my_sleep

DEFAULT_PARAMETER  = {
    'lineLoss': 25.00,
    'band': 1,
    'bw': 20,
    'scs': 15,
    'range': 'LOW',
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
    ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {parameter['range']}")


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
    
    # ap.send("CONFigure:CELL1:NR:SIGN:SLOT:CLEar")
    # ap.send("CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh")
    # ap.send("CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh")
    # ap.send("CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 3")
    # ap.send("CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh")
    # ap.send("CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh")
    # ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
    # ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")

    ap.send('CONFigure:CELL1:NR:Sign:SLOT:Clear')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT5:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT6:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT3:DL:TIND 5')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT5:DL:TIND 3')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT6:DL:TIND 2')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT3:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT4:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT5:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT6:DL:MCS1 4')

    ap.send('CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT8:UL:MCS1 2')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT9:UL:MCS1 2')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT8:UL:RB 0,135')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT9:UL:RB 0,135')

    ap.send('CONFigure:CELL1:NR:SIGN:SLOT10:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT11:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT12:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT13:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT14:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT15:CTYPe PDSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT16:CTYPe PDSCh')

    ap.send('CONFigure:CELL1:NR:SIGN:SLOT10:DL:TIND 8')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT11:DL:TIND 7')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT12:DL:TIND 6')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT13:DL:TIND 5')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT14:DL:TIND 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT15:DL:TIND 3')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT16:DL:TIND 2')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT10:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT11:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT12:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT13:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT14:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT15:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT16:DL:MCS1 4')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT18:CTYPe PUSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT19:CTYPe PUSCh')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT18:UL:MCS1 2')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT19:UL:MCS1 2')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT18:UL:RB 0,135')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT19:UL:RB 0,135')
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT:APPLy')

    my_sleep(1)
    ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")
    my_sleep(2)
    ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")
    ap.send("INITiate:NR:BLER")
    ap.send("INITiate:NR:MEValuation")
    my_sleep(2)
    
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
    
    my_sleep(3)
    
    dl_bler = ap.send("FETCh:NR:BLER:DL:RESult?", 7, True,"DL BLER") # 0,67,67,0,0,57955,57955,0.000,0.000,173.000,173.000
    
    ul_bler = ap.send("FETCh:NR:BLER:UL:RESult?", 7, True,"UL BLER") # 0,67,67,0,0,57955,57955,0.000,0.000,173.000,173.000
    
    ap.send("CONFigure:NR:MEValuation:REPetition CONTINUOUS")
    ap.send("CONFigure:NR:BLER:REPetition CONTINUOUS")
    
    my_sleep(10)
    txp = ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,"TXP AVG") # 0,10.873


    # ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,"feifei") 
    # for _ in range(3):
    #     ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,"TXP AVG")


def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    
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