from lib.var import *
from common import *

DEFAULT_PARAMETER = {
    'lineLoss1': 25.00,
    'lineLoss3': None,
    'nr_band': 1,    
    'nr_bw': 20,    
    'scs': 15,
    'dl_arfcn': None,    
    'range': 'LOW',
    "nr_start_power": -40,
    "nsa_start_power": -40,
    "end_power": -100,
    "step": -2,
    "fallback_delta": 10,
    "nr_slots": None,    
    "mcs": None,
    "rb_mode": "Inner_Full",
    "disconnected_detect_time":1,   

    'lte_band': None,
    'lte_bw': None,
    "lte_dl_arfcn": None,
    "lte_ul_arfcn": None,
    "lte_slots": None,
    "resource_allocation_type": 2,
    
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
dl_bler = 0
ul_bler = 0
def update_parameters(external_params=None):
    global parameter
    print("update_parameters》》》")
    if external_params:
        for key, value in external_params.items():
            if key in parameter:
                parameter[key] = value
def opc_fun():                
    while ap.query("*OPC?") != "1":                  
        my_sleep(0.05) 
    UE_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")    
    return UE_status

def case_start():
    remote_gnb_start()
   
    my_sleep(2)
    remote_diag_start()
   
    ## line loss configuration
    config_line_loss(parameter)
    
    ## band bw scs range configuration
    config_cell_band(parameter)
    
def case_body():
    for i in range(5):
        cell_current_state = ap.query("CALL:CELL1?")
        if  cell_current_state == "OFF": 
            ap.send("CALL:CELL1 ON")
            check_phone_at()
            my_sleep(2)
        elif cell_current_state == "ON": 
            break 
    
    if  cell_current_state == "OFF":
        print("❌仪表程序异常")
        ap.send("CONFigure:RFINdex:apply")
        my_sleep(20)
        ap.send("CALL:CELL1 ON")
        check_phone_at()
    
   # check_phone_at()
    my_sleep(1)
    
    connected = False
    for i in range(180):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            connected = True            
            modify_config_line_loss(parameter)
            # calibrate_line_loss(parameter, tolerance=3.0)
            # restore_line_loss(parameter)
            break
        else:
            result = opc_fun() 
            if '"DisConnected"' == result:
                ap.send("CALL:CELL1 OFF") 
                my_sleep(1)
                ap.send("CALL:CELL1 ON") 
            print(f"⏳ 第 {i+1} 次查询: UE未连接{result}")
            my_sleep(2)

    if not connected:
        print("❌ UE 多次未连接，跳过后续功率扫描，直接进入 case_clear")
        ap.check("UE连接", False, detail="多次查询UE未连接", status_msg="UE未连接,测试判定失败")
        return
   

    
    ap.send(f"CONFigure:CELL1:NR:SIGN:DDETection:SWITch ON,{parameter['disconnected_detect_time']}")

    def wait_ready(state_query, label):
        for i in range(100):
            result = ap.query(state_query)
            if "RDY" == result:
                print(f"✅ 第 {i+1} 次查询: {label}已准备好")               
                break
            else:
                print(f"⏳ 第 {i+1} 次查询: {label}未准备好")
                my_sleep(1)

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
        bler_state_cmd =f"FETCh:{tech}:BLER:STATe?"
        band_query = "CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator?" if tech == "NR" else "CONFigure:CELL1:LTE:SIGN:BAND:DL?"

        if tech == "NR":
            dl_arfcn_str = ap.query(f"CONFigure:CELL1:{tech}:SIGN:ARFCn?")
            dl_arfcn = (dl_arfcn_str.split(',')[0])
        else:
            dl_arfcn = ap.query("CONFigure:CELL1:LTE:SIGN:ARFCn:DL?")

        bandID = ap.query(f"{band_query}")
        def mark(status):
            ap.tag_last(status, 2)
        def _single_measure(power):
            ap.send(f"{power_cmd} {power}") 
            opc_fun()           
            ap.send(f"INITiate:{tech}:BLER")
           # ap.send(meas_init_cmd)
            #my_sleep(0.25) # ap.send("*OPC?")
            UE_RSRP_val = ap.send(f"CONFigure:CELL1:{tech}:UEReport:RSRP?", 1, True, "UEReport RSRP", True, True, True)
            TX_RSRP_val = ap.send(f"CONFigure:CELL1:{tech}:SIGN:RSRP?", 0, True, "TX RSRP", True, True, True)            
            opc_fun()                        
            UE_status =   ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
            
            while ap.query(bler_state_cmd) != "RDY" and UE_status == '"Connected"':
                UE_status = opc_fun() 
            #if tech == "NR":
                #ap.send("CONFigure:CELL1:NR:UEReport:RSRP?", 1, True, "Reasonable Line Loss", f"{power}dBm", True, True)
            if UE_status == '"Connected"':
                dl_bler_str = ap.send(f"FETCh:{tech}:BLER:DL:RESult?", 7, True, "bler_dl_title", f"{power}dBm", True, True)                
                #ul_bler_str = ap.send(f"FETCh:{tech}:BLER:UL:RESult?", 7, True, "bler_ul_title", f"{power}dBm", True, True)                
                dl_bler = float(dl_bler_str.split(',')[7])
                #ul_bler = float(ul_bler_str.split(',')[7])              
                ul_bler = 0 
            else:
                ul_bler = 1
                dl_bler = 1
            ap.send(f"ABORt:{tech}:BLER")
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} [{tech}]  频点[{dl_arfcn}]发送的cell power功率: {float(power):.2f} dBm,UE状态: {UE_status} RSRP功率: {float(TX_RSRP_val):.2f} dBm，UE上报的RSRP功率: {UE_RSRP_val} dBm DL BLER: {dl_bler:.2f}  UL BLER: {ul_bler:.2f}")
            
            return dl_bler, ul_bler, UE_status

        def measure(power):
            global dl_bler, ul_bler
            dl_bler, ul_bler, UE_status = _single_measure(power)
            bler_fail = (dl_bler > 0.20) or (ul_bler > 0.20)
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
           
            #print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} [{tech}] 频点[{dl_arfcn}] 功率: {power:.2f} dBm, UE状态: {UE_status}, DL BLER: {dl_bler:.2f}  UL BLER: {ul_bler:.2f}")
            return (dl_bler > 0.20) or (ul_bler > 0.20) or (UE_status != '"Connected"')

        

        failed_powers = set()
        current_power = start_power         
        loop_count = 0 
       
        
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} band[{bandID}] 频点[{dl_arfcn}]  dBm,开始测试>>>>>>")
        while current_power >= end_power:            
            if not measure(current_power):                
                current_power += step
                continue
            loop_count = loop_count + 1
                      
            if current_power in failed_powers or ((dl_bler == 1.0) or (ul_bler  == 1.0) and (loop_count > 1 or  current_power != start_power )):
                print(f"✅ [{tech}] 功率 {current_power} dBm 第{loop_count}次出现异常，判定为稳定边界，记录并停止")
                mark("normal")
                ap.check(f"{tech} 功率扫描", True,
                         detail=f"稳定边界 {current_power}dBm",
                         status_msg=f"{tech} power scan 正常 pass")
                return True
            
            failed_powers.add(current_power)
            
            ue_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")            
            if  ue_status != '"Connected"':
                ap.send("CALL:CELL1 OFF")
                my_sleep(1)  
                ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {nr_start_power}")             
                ap.send("CALL:CELL1 ON")
                for i in range(120):
                        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
                        if '"Connected"' == result:
                            print(f"✅ 第 {i+1} 次查询: UE已连接")
                            ap.send('CONFigure:CELL1:NR:SIGN:DDETection:SWITch ON,5')
                            connected = True
                            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
                            break
                        else:
                            print(f"⏳ 第 {i+1} 次查询: UE未连接 {result} 功率: {nr_start_power:.2f} dBm")
                            my_sleep(2)

            new_power = current_power + fallback_delta
            if new_power > start_power:
                new_power = start_power
            step = step/2
           # fallback_delta = fallback_delta/2 if (fallback_delta/2) > abs(step *3) else abs(step *3)
            #for i in range(fallback_delta + 1):
                #ap.send(f"CONFigure:CELL1:{tech}:SIGN:POWer {current_power + i*5}")  
                #UE_status = opc_fun()              
                #my_sleep(0.2)
                #tmeTX_RSRP_val = ap.query(f"CONFigure:CELL1:{tech}:SIGN:RSRP?")
                #tmeUE_RSRP_val = ap.query(f"CONFigure:CELL1:{tech}:UEReport:RSRP?")                
                #print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} [{tech}] 发送的cell power功率: {current_power + i*5} dBm,UE状态: {UE_status} RSRP功率: {float(tmeTX_RSRP_val):.2f} dBm UE上报的RSRP功率: {tmeUE_RSRP_val} dBm ")
                #measure(current_power + 1 + i*5)           

            if loop_count >= 3 or step == 0.25:
                print(f"✅ [{tech}] 功率 {current_power} dBm 第{loop_count}次出现异常，判定为稳定边界，记录并停止")
                mark("normal")
                ap.check(f"{tech} 功率扫描", True,
                detail=f"稳定边界 {current_power}dBm",status_msg=f"{tech} power scan 正常 pass")
                return True
            
            print(f"🔄 [{tech}] 功率 {current_power} dBm 第{loop_count}异常，回退到 {new_power} dBm 复测")
                 

            if measure(new_power):
                ue_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
                recovered = False
                print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} [{tech}] 频点[{dl_arfcn}] 功率功率: {new_power:.2f} dBm, UE状态: {ue_status}")
                           
                if ue_status != '"Connected"':
                    ap.send("CALL:CELL1 OFF")
                    print(f"🔁 [{tech}] 回退到 {new_power} dBm 后 UE 掉线,先每秒检查状态,最多等待200秒...")
                    reconnected = False
                    my_sleep(1)
                    ap.send("CALL:CELL1 ON")
                    start_time = time.time()
                    while time.time() - start_time < 200:
                        my_sleep(1)
                        ue_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
                        if ue_status == '"Connected"':
                            print(f"✅ [{tech}] 200秒内重连成功")
                            reconnected = True
                            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
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
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} band[{bandID}] 频点[{dl_arfcn}]  dBm,测试结束")   
        ap.check(f"{tech} 功率扫描", True,
                 detail=f"扫描完成 {start_power}~{end_power}dBm 无异常",
                 status_msg=f"{tech} power scan 扫描完成 pass")
        return False

    # ========== SA  ==========
    ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:MEValuation:RESult OFF,OFF,OFF,OFF,OFF,OFF")
    ap.send("CONFigure:NR:BLER:TEST:DIRection DL")
    ap.send("CONFigure:NR:BLER:DTXFlag ENABLE")
    ap.send("CONFigure:NR:BLER:EJUDgment ON")
    ap.send("CONFigure:NR:BLER:MEASlength 200")    
    
    ap.send('CONFigure:CELL1:NR:Sign:SLOT:Clear')
    ap.send('CONFigure:CELL1:LTE:SIGN:SUBFrame:CLear')    
    # config nr
    config_lte_subframes(parameter, LTE_SLOTS)
    config_nr_slots(parameter, NR_SLOTS)
    my_sleep(2)
    ap.send("ABORt:LTE:BLER")
    ap.send("ABORt:LTE:TXP")
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")

    ap.send("INITiate:NR:BLER")
    
   # my_sleep(0.5)
    opc_fun()

    wait_ready("FETCh:NR:BLER:STATe?", "BLER测试")
    my_sleep(2)
    ap.send("ABORt:NR:BLER")
    for i in range(2):
        base = parameter['nr_start_power']
        if i % 2 == 1:
            nr_start_power = base + 1
        else:
            nr_start_power = base
        print(f"{i} % 2 ={i % 2} nr_start_power {nr_start_power}dBm  base:{base}dBm")
        nr_should_stop = run_power_scan("NR", nr_start_power, end_power, step, fallback_delta, ap)

        
          
    if nr_should_stop:
        dl_arfcn_str = ap.query(f"CONFigure:CELL1:NR:SIGN:ARFCn?")
        dl_arfcn = (dl_arfcn_str.split(',')[0])
        bandID = ap.query(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator?")
        ap.send("ABORt:NR:BLER")
        ap.send("ABORt:NR:MEValuation")
        print(f"N{bandID} {dl_arfcn} NR测试终止")        

    ap.send(f"CONFigure:CELL1:NR:SIGN:POWer {nr_start_power}")
      

    # ========== NSA/LTE  ==========
    if parameter.get('lte_band') is not None:
        ap.send('CONFigure:CELL1:NR:Sign:SLOT:Clear')
        ap.send('CONFigure:CELL1:LTE:SIGN:SUBFrame:CLear')
        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")
        ap.send("ABORt:NR:BLER")
        ap.send("ABORt:NR:MEValuation")

        ap.send("CONFigure:LTE:BLER:REPetition SINGLESHOT")
        ap.send("CONFigure:NR:MEValuation:RESult OFF,OFF,OFF,OFF,OFF,OFF")
        ap.send("CONFigure:LTE:BLER:TEST:DIRection DL")
        ap.send("CONFigure:LTE:BLER:EJUDgment ON")
        for i in range(180):
            result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
            if '"Connected"' == result:
                print(f"✅ 第 {i+1} 次查询: UE已连接")
                connected = True            
                modify_config_line_loss(parameter)
                # calibrate_line_loss(parameter, tolerance=3.0)
                # restore_line_loss(parameter)
                break
            else:
                result = opc_fun() 
                if '"DisConnected"' == result:
                    ap.send("CALL:CELL1 OFF") 
                    my_sleep(1)
                    ap.send("CALL:CELL1 ON")  
                print(f"⏳ 第 {i+1} 次查询: UE未连接{result}")
                my_sleep(2)
        # config lte
        config_lte_subframes(parameter, LTE_SLOTS)
        opc_fun() 
        my_sleep(2)

        ap.send("INITiate:LTE:BLER")  
        wait_ready("FETCh:LTE:BLER:STATe?", "LTE BLER测试")
       
        nsa_should_stop = run_power_scan("LTE", nsa_start_power, end_power, step, fallback_delta, ap)
        if nsa_should_stop:
            dl_arfcn = ap.send("CONFigure:CELL1:LTE:SIGN:ARFCn:DL?")
            lte_band = ap.send("CONFigure:CELL1:LTE:SIGN:BAND:DL?")
            print(f"{lte_band} {dl_arfcn} NSA_LTE测试终止")  
            
        ap.send(f"CONFigure:CELL1:LTE:SIGN:POWer {nsa_start_power}")
         

def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    if parameter.get('lte_band') is not None:
        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")     
    for i in range(parameter['disconnected_detect_time']*4):
        ue_status = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        my_sleep(0.5) 
        nr_start_power = ap.send(f"CONFigure:CELL1:NR:SIGN:POWer?") 
        UE_RSRP_val = ap.query(f"CONFigure:CELL1:NR:UEReport:RSRP?")
        TX_RSRP_val = ap.query(f"CONFigure:CELL1:NR:SIGN:RSRP?")   
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} UE状态: {ue_status} {nr_start_power} RSRP功率: {float(TX_RSRP_val):.2f} dBm，UE上报的RSRP功率: {UE_RSRP_val} dBm")
                         
    if  ue_status != '"Connected"':      
        ap.send("CALL:CELL1 OFF")  
    
    '''
    
    ap.send("CALL:CELL1 OFF")
    opc_fun() 

    remote_diag_stop()

    for i in range(5):
        result = ap.query("CALL:CELL1?")
        if "OFF" == result:
            print(f"✅ CELL已关闭")
            break
        else:
            print(f"⏳ 等待CELL关闭...")
            opc_fun() 

    remote_gnb_stop()
    #remote_restart()
    '''
