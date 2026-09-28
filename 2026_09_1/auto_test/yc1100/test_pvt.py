from math import e

from lib.var import *
from common import *

DEFAULT_PARAMETER  = {
    'lineLoss1': 25.00,
    'nr_band': 1,
    'nr_bw': 20,
    'scs': 15,
    'range': 'LOW',
    "TD":600,
}

parameter = DEFAULT_PARAMETER.copy()

def update_parameters(external_params=None):
    global parameter
    
    if external_params:
        for key, value in external_params.items():
            if key in parameter:
                parameter[key] = value


def case_start():
    remote_diag_start()
    # line loss configuration
    ap.send(f"CONFigure:NSASa:SWITch SA")
    ap.send(f"CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,{parameter['lineLoss1']},6000000000,{parameter['lineLoss1']}")
    ap.send("CONFigure:BASE:FDCorrection:SAVE")

    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX")
    ap.send("CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX")
    ap.send(f"CONFigure:RFINdex:CLear:ALL")

    ap.send(f"CONFigure:RFINdex:DL NR,1")
    ap.send(f"CONFigure:RFINdex:UL NR,1")
    ap.send(f"CONFigure:RFINdex1:CONNector IO")
    ap.send(f"CONFigure:RFINdex3:CONNector IO")

    ap.send("CONFigure:RFINdex:apply")
    my_sleep(6)

    # band bw scs range configuration
    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {parameter['nr_band']}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{parameter['nr_bw']}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{parameter['scs']}")
    ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {parameter['range']}")

def case_body():
    ap.send("CALL:CELL1 ON")
    check_phone_at()
    my_sleep(5)

    for i in range(20):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?") 
        print("UE state",result) 
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接")
            my_sleep(2)
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:CLEar")

    ## --------------------------PUSCh config ---------------------------
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT8:UL:MCS1 2")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT9:UL:MCS1 2")
    ## -----------------------------------------------------
    
    if parameter['nr_bw']==100 and parameter['scs']==30:
        inner_full="67,135" 
    elif parameter['nr_bw']==20 and parameter['scs']==30:
        inner_full="25,50"
    else:
        inner_full="0,270"
    ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT8:UL:RB {inner_full}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:SLOT9:UL:RB {inner_full}")
    ap.send('CONFigure:CELL1:NR:SIGN:SLOT:APPLy')
    my_sleep(2)

    ap.send("INITiate:NR:GOOTmask")

    start_time = time.time()
    end_time = start_time + parameter['TD']

    while time.time() < end_time:
        try:
            ap.send("CONFigure:NR:GOOTmask:REPetition CONTINUOUS")  # CONTINUOUS SINGLESHOT
            ap.send("FETCh:NR:GOOTmask:RESult?",1, True,"OFF Power(before)")
            ap.send("FETCh:NR:GOOTmask:RESult?",2, True,"ON Power")
            ap.send("FETCh:NR:GOOTmask:RESult?",3, True,"OFF Power(after)")

        except Exception as e:
            time.sleep(0.5)
            continue
        time.sleep(0.25)

    remote_pvt_screenshot()
    
    ap.send("CONFigure:NR:GOOTmask:SLOT:STARt?")
    ap.send("CONFigure:NR:GOOTmask:SLOT:LENGth?")

def case_clear():
    ap.send("ABORt:NR:GOOTmask")
    ap.send("CALL:CELL1 OFF")
    my_sleep(1)

    remote_diag_stop()

    for i in range(5):
        result = ap.query("CALL:CELL1?")
        print("CALL result",result)
        if "OFF" == result:
            print(f"✅ CELL已关闭")
            break
        else:
            print(f"⏳ 等待CELL关闭...")
            my_sleep(2)
    
    remote_restart()