from lib.var import *
from common import *


DEFAULT_PARAMETER = {
    'lineLoss1': 25.00,
    'nr_band_list': [78, 77, 79],
    'range_list': ["Low", "Mid", "High"],
    'nr_slots': None,
    'mcs': None,
    'rb_mode': 'Inner_Full',
    'waveform':'DFTS',
    'bw':  {'tdd': [100], 'fdd': [20]},
    'scs': {'tdd': [30],  'fdd': [15]},
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

def wait_for_ue_connected(ap, max_attempts=300, delay=1):
    for i in range(max_attempts):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            calibrate_line_loss(parameter, tolerance=3.0)
            return True
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接,状态={result}")
            my_sleep(delay)
    print("❌ UE 连接超时")
    return False

def perform_measurement(ap, band, range_str, connected=True, bw=None):
    try:
        ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
        ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")

        ap.send("INITiate:NR:BLER")
        ap.send("INITiate:NR:MEValuation")
        my_sleep(0.5)

        bler_ready = False
        max_wait = 5 if connected else 2
        for i in range(max_wait):
            result = ap.query("FETCh:NR:BLER:STATe?")
            if "RDY" == result:
                bler_ready = True
                break
            else:
                my_sleep(2)

        bw_tag = f"BW{bw}" if bw is not None else ""
        base_x_label = f"{band}({range_str}){bw_tag}"
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

        txp_power = ap.send("CONFigure:CELL1:NR:SIGN:POWer?")
        txp_x_label = f"{x_label} pwr:{str(txp_power).strip()}dBm"
        txp = ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,
                      "TXP AVG", txp_x_label, True, True)
        check_txp(f"NR TXP AVG ({txp_x_label})", txp)

        return dl_bler, ul_bler, txp
    except Exception as e:
        print(f"测量失败: {e}")
        return None


def case_start():
    remote_gnb_start()
    remote_diag_start()
    config_line_loss(parameter)
    config_cell_band(parameter)


def case_body():
    nr_band_list = parameter.get('nr_band_list', [78, 77, 79])

    ap.send("CALL:CELL1 ON")
    check_phone_at()
    my_sleep(5)

    connected = False
    for i in range(20):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            connected = True
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接")
            my_sleep(2)

    if not connected:
        print("❌ UE 多次未连接，跳过后续测量，直接进入 case_clear")
        ap.check("UE连接", False, detail="多次查询UE未连接", status_msg="UE未连接,测试判定失败")
        return

    calibrate_line_loss(parameter, tolerance=3.0, skip_next=True)

    ap.send("CONFigure:CELL1:NR:SIGN:POWer?")

    rb_mode = parameter.get('rb_mode', 'Inner_Full')
    waveform = parameter.get('waveform', 'DFTS')
    range_list = parameter.get('range_list', ["Low", "Mid", "High"])
    rounds = num_rounds(parameter)

    for round_idx in range(rounds):
        if rounds > 1:
            print(f"\n===== 第 {round_idx+1}/{rounds} 轮带宽扫描 =====")
        for band in nr_band_list:
            nr_bw, nr_scs = duplex_bw_scs(parameter, band, round_idx)
            ul_slots = [n for item in slots_for_band(band, NR_SLOTS).get("UL", []) for n in item.keys()]
            rmc_configured = False  

            for idx, rng in enumerate(range_list):
                first_in_band = idx == 0

                ap.send("CONFigure:CELL1:NR:CONFig:PMODe AUTO")
                if first_in_band:
                    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {band}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{nr_bw}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{nr_scs}")

                ap.send(f"CONFigure:CELL1:NR:CONFig:RANGe {rng}")
                arfcn_resp = ap.query("CONFigure:CELL1:NR:SIGN:ARFCn?")
                arfcn = arfcn_resp.split(',')[0].strip()
                print(f"获取 ARFCN: {arfcn}")
                my_sleep(2)

                ap.send("CONFigure:CELL1:NR:CONFig:PMODe MANUAL")
                if first_in_band:
                    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator {band}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:BWidth:DL BW{nr_bw}")
                    ap.send(f"CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz{nr_scs}")
                ap.send(f"CONFigure:CELL1:NR:SIGN:CFSCommand {arfcn},AUTO,AUTO")
                ap.send("CONFigure:CELL1:NR:SIGN:CHANnel:SWITch")
                my_sleep(3)

                if not wait_for_ue_connected(ap):
                    print(f"频段 {band} 范围 {rng} 连接失败，仍读取数据并标记为失败档")
                    perform_measurement(ap, band, rng, connected=False, bw=nr_bw)
                    continue

                if not rmc_configured:
                    config_nr_slots(parameter, NR_SLOTS, nr_bw=nr_bw, nr_scs=nr_scs, band=band)
                    rmc_configured = True
                else:
                    switch_ul_slot_rb(ul_slots, nr_bw, nr_scs, rb_mode, waveform, band=band)

                result = perform_measurement(ap, band, rng, bw=nr_bw)
                ap.send("ABORt:NR:BLER")
                ap.send("ABORt:NR:MEValuation")
                restore_line_loss(parameter)
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