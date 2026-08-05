from lib.var import *
from common import *

DEFAULT_PARAMETER = {
    'lineLoss1': 5.00,
    'lineLoss3': 5.00,
    'lte_band': 3,
    'lte_bw': 20,
    'nr_band_list': [78, 77, 79, 41, 1, 5, 8, 28],
    'lte_band_list': [{78: "1,20:[300],10:[50]"}],
    'nr_range_list': ["Low", "Mid", "High", "Low"],
    'lte_nr_range': "Low",
    'rb_mode': 'Inner_Full',
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

def wait_for_ue_connected(ap, max_attempts=30, delay=1):
    for i in range(max_attempts):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            return True
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接,状态={result}")
            my_sleep(delay)
    print("❌ UE 连接超时")
    return False

def perform_nr_measurement(ap, band, range_str, connected=True):
    try:
        ap.send("ABORt:NR:BLER")
        ap.send("ABORt:NR:MEValuation")
        my_sleep(0.2)

        ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
        ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")

        ap.send("INITiate:NR:BLER")
        ap.send("INITiate:NR:MEValuation")
        my_sleep(0.5)

        max_wait = 10 if connected else 2
        for i in range(max_wait):
            result = ap.query("FETCh:NR:BLER:STATe?")
            if "RDY" == result:
                break
            else:
                my_sleep(1)

        base_x_label = f"{band}({range_str})"
        x_label = base_x_label if connected else f"{base_x_label}[失败]"

        dl_bler_str = ap.send("FETCh:NR:BLER:DL:RESult?", 7, True,
                              "DL NR_BLER", x_label, True, True)
        ul_bler_str = ap.send("FETCh:NR:BLER:UL:RESult?", 7, True,
                              "UL NR_BLER", x_label, True, True)

        dl_bler = float(dl_bler_str.split(',')[7])
        ul_bler = float(ul_bler_str.split(',')[7])

        txp = ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,
                      "NR TXP AVG", x_label, True, True)

        return dl_bler, ul_bler, txp
    except Exception as e:
        print(f"NR 测量失败: {e}")
        return None

def perform_lte_measurement(ap, nr_band,rng, ob, lte_bw, arfcn):
    try:
        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")
        my_sleep(0.2)

        ap.send("CONFigure:LTE:TXP:REPetition SINGLESHOT")
        ap.send("CONFigure:LTE:BLER:REPetition SINGLESHOT")

        ap.send("INITiate:LTE:BLER")
        ap.send("INITiate:LTE:TXP")
        my_sleep(0.5)

        bler_ready = False
        for i in range(5):
            result = ap.query("FETCh:LTE:BLER:STATe?")
            if "RDY" == result:
                bler_ready = True
                break
            else:
                my_sleep(2)

        base_x_label = f"NR:{nr_band} range:{rng} OB:{ob} BW:{lte_bw} ({arfcn})"
        if bler_ready:
            x_label = base_x_label
        else:
            print(f"❌ LTE BLER 未就绪，仍记录为失败数据点")
            x_label = f"{base_x_label}[失败]"

        dl_bler_str = ap.send("FETCh:LTE:BLER:DL:RESult?", 7, True,
                              "DL LTE_BLER", x_label, True, True)
        ul_bler_str = ap.send("FETCh:LTE:BLER:UL:RESult?", 7, True,
                              "UL LTE_BLER", x_label, True, True)

        dl_bler = float(dl_bler_str.split(',')[7])
        ul_bler = float(ul_bler_str.split(',')[7])

        txp = ap.send("FETCh:LTE:TXP:AVG?", 1, True,
                      "LTE TXP AVG", x_label, True, True)

        return dl_bler, ul_bler, txp
    except Exception as e:
        print(f"LTE 测量失败: {e}")
        return None

def case_start():
    remote_gnb_start()
    remote_diag_start()

    ## line loss configuration
    config_line_loss(parameter)

    nr_band_list = parameter.get('nr_band_list', [78])
    first_band = nr_band_list[0]
    init_bw = 100 if first_band >= 41 else 20
    init_scs = 30 if first_band >= 41 else 15
    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {first_band}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{init_bw}")
    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{init_scs}")
    ap.send("CONFigure:CELL1:NR:UE:MReport ON ")

    ap.send(f"CONFigure:CELL1:LTE:SIGN:BAND:DL OB{parameter['lte_band']}")
    ap.send(f"CONFigure:CELL1:LTE:SIGN:BWidth BW_{parameter['lte_bw']}")


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

    config_nr_slots(parameter, NR_SLOTS)

    config_lte_subframes(parameter)

    ap.send("CONFigure:CELL1:NR:SIGN:POWer?")

    # ==================== NR  ====================
    ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")

    nr_band_list = parameter.get('nr_band_list', [78, 77, 79, 41,])
    nr_range_list = parameter.get('nr_range_list', ["Low", "Mid", "High", "Low"])

    for band in nr_band_list:
        update_bw = 100 if band >= 41 else 20
        update_scs = 30 if band >= 41 else 15
        for idx, rng in enumerate(nr_range_list):
            first_in_band = idx == 0

            ap.send("CONFigure:CELL1:NR:CONFig:PMODe AUTO")
            if first_in_band:
                ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {band}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{update_bw}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{update_scs}")

            ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {rng}")
            arfcn_resp = ap.query("CONFigure:CELL1:NR:SIGN:ARFCn?")
            arfcn = arfcn_resp.split(',')[0].strip()
            my_sleep(1)

            ap.send("CONFigure:CELL1:NR:CONFig:PMODe MANUAL")
            if first_in_band:
                ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {band}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{update_bw}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{update_scs}")
            ap.send(f"CONFigure:CELL1:NR:SIGN:CFSCommand {arfcn},AUTO,AUTO")
            ap.send("CONFigure:CELL1:NR:SIGN:CHANnel:SWITch")
            my_sleep(5)

            if not wait_for_ue_connected(ap):
                print(f"NR 频段 {band} 范围 {rng} 连接失败，仍读取数据并标记为失败档")
                perform_nr_measurement(ap, band, rng, connected=False)
                continue

            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
            ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame:APPLy")
            my_sleep(2)

            perform_nr_measurement(ap, band, rng)
            my_sleep(1)

    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    ap.send("CONFigure:NR:MEValuation:RESult OFF,OFF,OFF,OFF,OFF,OFF")
    my_sleep(1)

    # ==================== LTE ====================
    lte_band_list = parameter.get('lte_band_list', [])

    for entry in lte_band_list:  
        nr_band = None
        ob = None
        configs = [] 
        for key, val in entry.items():
            if isinstance(val, list):
                configs.append((int(key), val))
            else:
                nr_band = int(key)
                ob = int(val)
        if nr_band is None or ob is None or not configs:
            print(f"⚠️ 跳过无效 entry: {entry}")
            continue

        configs.sort(key=lambda x: x[0], reverse=True)

        for lte_bw, arfcn_list in configs:
            for arfcn in arfcn_list:
                for rng in nr_range_list:
                    # ---- config NR ----
                    nr_bw = 100 if nr_band >= 41 else 20
                    nr_scs = 30 if nr_band >= 41 else 15

                    ap.send("CONFigure:CELL1:NR:CONFig:PMODe AUTO")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {nr_band}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{nr_bw}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{nr_scs}")
                    ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {rng}")
                    nr_arfcn_resp = ap.query("CONFigure:CELL1:NR:SIGN:ARFCn?")
                    nr_arfcn = nr_arfcn_resp.split(',')[0].strip()
                    my_sleep(1)

                    ap.send("CONFigure:CELL1:NR:CONFig:PMODe MANUAL")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {nr_band}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{nr_bw}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{nr_scs}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:CFSCommand {nr_arfcn},AUTO,AUTO")

                    # ---- config LTE ----
                    ap.send(f"CONFigure:CELL1:LTE:SIGN:BAND:DL OB{ob}")
                    ap.send(f"CONFigure:CELL1:LTE:SIGN:BWidth BW_{lte_bw}")
                    ap.send(f"CONFigure:CELL1:LTE:SIGN:ARFCn:DL {arfcn}")

                    ap.send("CONFigure:CELL1:NR:SIGN:CHANnel:SWITch")
                    my_sleep(5)

                    if not wait_for_ue_connected(ap):
                        print(f"⚠️ LTE 连接超时 (NR {nr_band}, range {rng}, ARFCN {arfcn})，仍执行 LTE 测量以记录该失败档位")
                        perform_lte_measurement(ap, nr_band, rng, ob, lte_bw, arfcn)
                        continue

                    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
                    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
                    ap.send("CONFigure:CELL1:LTE:SIGN:SUBFrame:APPLy")
                    my_sleep(2)
                    perform_lte_measurement(ap, nr_band, rng, ob, lte_bw, arfcn)
                    my_sleep(1)

def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    ap.send("ABORt:LTE:BLER")
    ap.send("ABORt:LTE:TXP")
    ap.send("CALL:CELL1 OFF")
    my_sleep(1)

    remote_diag_stop()

    for i in range(5):
        result = ap.query("CALL:CELL1?")
        if "OFF" == result:
            print("✅ CELL已关闭")
            break
        else:
            print("⏳ 等待CELL关闭...")
            my_sleep(2)

    remote_gnb_stop()
    remote_restart()