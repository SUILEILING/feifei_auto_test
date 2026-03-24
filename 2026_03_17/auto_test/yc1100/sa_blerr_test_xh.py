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
def rmc_config():
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:CLEar")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT3:DL:TIND 5")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT3:DL:MCS1 4")
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 4")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT4:DL:MCS1 4")

    ap.send("CONFigure:CELL1:NR:SIGN:SLOT5:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT5:DL:TIND 3")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT5:DL:MCS1 4")

    ap.send("CONFigure:CELL1:NR:SIGN:SLOT6:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT6:DL:TIND 2")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT6:DL:MCS1 4")
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT10:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT10:DL:TIND 8")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT10:DL:MCS1 4")
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT11:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT11:DL:TIND 7")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT11:DL:MCS1 4")
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT12:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT12:DL:TIND 6")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT12:DL:MCS1 4")
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT13:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT13:DL:TIND 5")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT13:DL:MCS1 4")
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT14:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT14:DL:TIND 4")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT14:DL:MCS1 4")
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT15:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT15:DL:TIND 3")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT15:DL:MCS1 4")
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT16:CTYPe PDSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT16:DL:TIND 2")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT16:DL:MCS1 4")
    
    
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT8:UL:MCS1 2")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT9:UL:MCS1 2")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT18:CTYPe PUSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT18:UL:MCS1 2")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT19:CTYPe PUSCh")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT19:UL:MCS1 2")

    ap.send("CONFigure:CELL1:NR:SIGN:SLOT8:UL:RB 0,135 ")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT9:UL:RB 0,135 ")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT18:UL:RB 0,135 ")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT19:UL:RB 0,135 ")
      
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")

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

    ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {parameter['start_power']}")

    ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")

    ap.send("CONFigure:NR:BLER:TEST:DIRection DL")   #测试bler的方向为下行方向
    ap.send("CONFigure:NR:BLER:EJUDgment ON")        #开启 NR bler判断
    ap.send("CONFigure:NR:BLER:DTXFlag ENABLE")      #使能  NR bler的DTX标志
    ap.send("CONFigure:NR:BLER:MEASlength 100 ")     #配置  NR bler的测试长度为100帧
   # ap.send(f"CONFigure:CELL1:NR:SIGN:REDCap:MODe REDCAP")

def case_body():

    ap.send("CALL:CELL1 OFF")
    ap.send("CALL:CELL1 ON")
    #ap.send("CONFigure:CELL1:NR:SIGN:DDETection:SWITch ON,1 ")
    ap.send(f"CONFigure:CELL1:NR:SIGN:POWer -60")
    check_phone_at()
    my_sleep(5)
    rmc_config()
    for i in range(1000):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")  # '"Connected"'
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")            
            ap.send("CONFigure:CELL1:NR:SIGN:DDETection:SWITch ON,1 ")
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接")
            my_sleep(1)

    
    start_power = parameter['start_power']
    end_power = parameter['end_power']
    step = parameter['step']
    fallback_delta = parameter.get('fallback_delta', 10) 

    current_power = start_power
    consecutive_failures = 0            
    current_count =0
    while True:
        if current_power != start_power:
            ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {current_power}")      
        UE_status = ap.send("CONFigure:CELL1:NR:SIGN:UE:STATe?") # '"Connected"'
        ap.send("INITiate:NR:BLER")
     
        for i in range(500):
            result = ap.query("FETCh:NR:BLER:STATe?")   #'RDY'
            if "RDY" == result:
                #print(f"✅ 第 {i+1} 次查询: BLER测试已准备好")
                break
            else:
               # print(f"⏳ 第 {i+1} 次查询: BLER测试未准备好")
                my_sleep(0.005)

        dl_bler_str=ap.send("FETCh:NR:BLER:DL:RESult?", 7, True, "DL BLER", f"{current_power}") # 0,67,67,0,0,57955,57955,0.000,0.000,173.000,173.000
        
        dl_bler_total=float(dl_bler_str.split(',')[1])
        dl_bler_ack=float(dl_bler_str.split(',')[2])
        dl_bler_nack=float(dl_bler_str.split(',')[3])
        dl_bler_dtx=float(dl_bler_str.split(',')[4])
        dl_bler=float(dl_bler_str.split(',')[7])        

        now_time = datetime.now()
        # 格式1: 年月日 时分秒毫秒
        format1 = now_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        print(f" {format1} 当前功率: {current_power:.3f} dBm, UE状态: {UE_status}, DL TOTAL: {dl_bler_total},ACK: {dl_bler_ack},NACK: {dl_bler_nack},DTX: {dl_bler_dtx},BLER: {dl_bler}")
        ap.send("ABORt:NR:BLER")
        #ap.send("ABORt:NR:MEValuation")
            
        if  UE_status != '"Connected"':
            ap.send("CALL:CELL1 OFF")
            my_sleep(5)
            ap.send("CALL:CELL1 ON")
           
            while current_power < -80:
               current_power += fallback_delta
               current_power = int(current_power)
            print(f"当前发送功率: {current_power} dBm")
            ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {current_power}")
            
            for i in range(1000):
                result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")  # '"Connected"'
                if '"Connected"' == result:
                    print(f"✅ 第 {i+1} 次查询: UE已连接")
                    #ap.send("CONFigure:CELL1:NR:SIGN:DDETection:SWITch ON,1 ")
                    rmc_config()           
                    break
                else:
                    print(f"⏳ 第 {i+1} 次查询: UE未连接")
                    my_sleep(1)
            continue 

        if consecutive_failures == 0:
            is_failure = not((dl_bler < 0.05) and (UE_status == '"Connected"'))
        else:
            is_failure = not((dl_bler == 0) and (UE_status == '"Connected"'))
        #print(f"is_failure1 {is_failure}")
        if is_failure:
            #print(f"is_failure {is_failure}")
            consecutive_failures += 1
           # print(f"is_failure {is_failure}  consecutive_failures {consecutive_failures} ")
            if consecutive_failures >= 2 and (step==-2):
                '''
                
                if  dl_bler != 1 :      #欧铊
                    UE_status = ap.send("CONFigure:CELL1:NR:SIGN:UE:STATe?") # '"Connected"'
                    ap.send("INITiate:NR:BLER")
                # ap.send("INITiate:NR:MEValuation")
                    #my_sleep(0.4)
                    for i in range(500):
                        result = ap.query("FETCh:NR:BLER:STATe?")   #'RDY'
                        if "RDY" == result:
                            #print(f"✅ 第 {i+1} 次查询: BLER测试已准备好")
                            break
                        else:
                        # print(f"⏳ 第 {i+1} 次查询: BLER测试未准备好")
                            my_sleep(0.005)

                    dl_bler_str=ap.send("FETCh:NR:BLER:DL:RESult?", 7, True, "DL BLER", f"{current_power}") # 0,67,67,0,0,57955,57955,0.000,0.000,173.000,173.000
                    
                    dl_bler_total=float(dl_bler_str.split(',')[1])
                    dl_bler_ack=float(dl_bler_str.split(',')[2])
                    dl_bler_nack=float(dl_bler_str.split(',')[3])
                    dl_bler_dtx=float(dl_bler_str.split(',')[4])
                    dl_bler=float(dl_bler_str.split(',')[7])     
                

                    now_time = datetime.now()
                    # 格式1: 年月日 时分秒毫秒
                    format1 = now_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
                    print(f" {format1}第 {i+1} 次查询BLER测试已准备好, 当前功率: {current_power} dBm, UE状态: {UE_status}, DL TOTAL: {dl_bler_total},ACK: {dl_bler_ack},NACK: {dl_bler_nack},DTX: {dl_bler_dtx},BLER: {dl_bler}")
                    ap.send("ABORt:NR:BLER")
                    consecutive_failures = 0   # 复位后清除计数，但功率点不变
                    ap.send("CALL:CELL1 OFF")
                    break
                    '''
                if consecutive_failures == 3:  #星航
                    consecutive_failures = 0   # 复位后清除计数，但功率点不变
                    ap.send("CALL:CELL1 OFF")
                    break
                print(f"⚠️ 连续两次失败，复位连接...")
                ap.send("CALL:CELL1 OFF")
                my_sleep(5)
                ap.send("CALL:CELL1 ON")
                ap.send("CONFigure:CELL1:NR:SIGN:DDETection:SWITch ON,1 ")
                current_power += fallback_delta
                current_power = int(current_power)
                if current_power < -80:
                    current_power = -50
                print(f"当前发送功率: {current_power:.3f} dBm")
                ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {current_power}")
                #ap.send(f"CONFigure:CELL1:NR:SIGN:POWer -60")
                ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
                ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")
                ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")

                ap.send("CONFigure:NR:BLER:TEST:DIRection DL")   #测试bler的方向为下行方向
                ap.send("CONFigure:NR:BLER:EJUDgment ON")        #开启 NR bler判断
                ap.send("CONFigure:NR:BLER:DTXFlag ENABLE")      #使能  NR bler的DTX标志
                ap.send("CONFigure:NR:BLER:MEASlength 100 ")     #配置  NR bler的测试长度为100帧
                #check_phone_at()
                #my_sleep(5)
                #consecutive_failures = 0   # 复位后清除计数，但功率点不变
                for i in range(1000):
                    result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")  # '"Connected"'
                    if '"Connected"' == result:
                        print(f"✅ 第 {i+1} 次查询: UE已连接")
                        rmc_config()
                        break
                    else:
                        print(f"⏳ 第 {i+1} 次查询: UE未连接")
                        my_sleep(1)
                continue           
            else:
                #print(f"当修改前功率: {current_power:.3f} dBm, dl_bler: {dl_bler}, step: {step}")
                if dl_bler==1:
                   current_power += fallback_delta 
                   current_power = int(current_power)
                   step=-2
                   print("="*60)
                elif (dl_bler > 0.050) and (step==-2):
                    current_power += 1
                    step=-1    
                elif (dl_bler > 0.050) and (step==-1):
                    current_power += 1
                    step=-0.3    
                elif (dl_bler > 0.050) and (step==-0.3):
                    current_power += fallback_delta
                    current_power = int(current_power)
                    step=-2   
                    print("="*60)    

                #current_power = -60  
                #print(f"当修改后功率: {current_power:.3f} dBm, fallback_delta: {fallback_delta}, step: {step}")
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
    #my_sleep(5)

    for i in range(5):
        result = ap.query("CALL:CELL1?") #  'OFF'
        if "OFF" == result:
            print(f"✅ CELL已关闭")
            break
        else:
            print(f"⏳ 等待CELL关闭...")
            my_sleep(2)
            
    #ap.send("*rst")
    my_sleep(5)