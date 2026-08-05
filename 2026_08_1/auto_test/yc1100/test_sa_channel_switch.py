from lib.var import *
from common import *

DEFAULT_PARAMETER = {
    'lineLoss1': 25.00,
    'lineLoss3': None,
    'lte_band': None,
    'nr_band_list': [78, 77, 79],
    'range_list': ["Low", "Mid", "High"],
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

def wait_for_ue_connected(ap, max_attempts=50, delay=1):
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

def perform_measurement(ap, band, range_str, connected=True):
    try:
        ap.send("ABORt:NR:BLER")
        ap.send("ABORt:NR:MEValuation")
        my_sleep(0.2)

        ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
        ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")

        ap.send("INITiate:NR:BLER")
        ap.send("INITiate:NR:MEValuation")
        my_sleep(0.5)

        # UE 没连上时 BLER 永远不会 RDY，快速探测即可，不必干等 10 秒
        bler_ready = False
        max_wait = 5 if connected else 2
        for i in range(max_wait):
            result = ap.query("FETCh:NR:BLER:STATe?")
            if "RDY" == result:
                bler_ready = True
                break
            else:
                my_sleep(2)

        # 没连上或未就绪也照常读取并记录，横坐标加 [失败] 标识，图表里会标红，
        # 与真实测得值区分（否则这一档直接消失在图里）
        base_x_label = f"{band}({range_str})"
        if connected and bler_ready:
            x_label = base_x_label
        else:
            print(f"❌ 未连接或 BLER 未就绪，仍记录为失败数据点")
            x_label = f"{base_x_label}[失败]"

        dl_bler_str = ap.send("FETCh:NR:BLER:DL:RESult?", 7, True,
                              "DL NR_BLER", x_label, True, True)
        ul_bler_str = ap.send("FETCh:NR:BLER:UL:RESult?", 7, True,
                              "UL NR_BLER", x_label, True, True)

        dl_bler = float(dl_bler_str.split(',')[7])
        ul_bler = float(ul_bler_str.split(',')[7])

        txp = ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,
                      "TXP AVG", x_label, True, True)

        return dl_bler, ul_bler, txp
    except Exception as e:
        print(f"测量失败: {e}")
        return None


def case_start():
    remote_gnb_start()
    remote_diag_start()

    ## line loss configuration 
    config_line_loss(parameter)


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

    nr_band_list = parameter.get('nr_band_list', [78, 77, 79])
    range_list = parameter.get('range_list', ["Low", "Mid", "High"])

    for band in nr_band_list:
        update_bw = 100 if band >= 41 else 20
        update_scs = 30 if band >= 41 else 15
        for idx, rng in enumerate(range_list):
            first_in_band = idx == 0

            ap.send("CONFigure:CELL1:NR:CONFig:PMODe AUTO")
            if first_in_band:
                ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {band}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{update_bw}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{update_scs}")

            ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {rng}")
            arfcn_resp = ap.query("CONFigure:CELL1:NR:SIGN:ARFCn?")
            arfcn = arfcn_resp.split(',')[0].strip()
            print(f"获取 ARFCN: {arfcn}")
            my_sleep(2)

            ap.send("CONFigure:CELL1:NR:CONFig:PMODe MANUAL")
            if first_in_band:
                ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {band}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{update_bw}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{update_scs}")
            ap.send(f"CONFigure:CELL1:NR:SIGN:CFSCommand {arfcn},AUTO,AUTO")
            ap.send("CONFigure:CELL1:NR:SIGN:CHANnel:SWITch")
            my_sleep(5)

            if not wait_for_ue_connected(ap):
                print(f"频段 {band} 范围 {rng} 连接失败，仍读取数据并标记为失败档")
                perform_measurement(ap, band, rng, connected=False)
                continue

            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:UPDate")
            my_sleep(2)
            ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")
            my_sleep(3)

            result = perform_measurement(ap, band, rng)
            my_sleep(1)


def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
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