from lib.var import *
from common import *


DEFAULT_PARAMETER = {
    'lineLoss1': 25.00,
    'lineLoss3': None,
    'band_list': [78, 77, 79, 41, 1, 5, 8, 28],
    'lte_band': None,
    'nr_bw': None,
    'lte_bw': None,
    'scs': None,
    'range': 'LOW',
    'waveform': 'DFTS',
    'rb_mode': 'Outer_Full',
    'resource_allocation_type': None,
    'nr_slots': None,
    'mcs': None,
    'lte_slots': None,
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


def update_parameters(external_params=None):
    global parameter
    if external_params:
        for key, value in external_params.items():
            parameter[key] = value


def _band_list():
    bl = parameter.get('band_list')
    return list(bl) if bl else []


def _bw_scs_for(band):
    if parameter.get('nr_bw') and parameter.get('scs'):
        return parameter['nr_bw'], parameter['scs']
    return (100, 30) if int(band) >= 41 else (20, 15)


def _band_param(band):
    bw, scs = _bw_scs_for(band)
    bp = dict(parameter)
    bp.pop('band_list', None) 
    bp['nr_band'] = band
    bp['nr_bw'] = bw
    bp['scs'] = scs
    return bp


def wait_for_ue_connected(ap, max_attempts=30, delay=2):
    for i in range(max_attempts):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            return True
        print(f"⏳ 第 {i+1} 次查询: UE未连接,状态={result}")
        my_sleep(delay)
    print("❌ UE 连接超时")
    return False

def perform_nr_measurement(ap, band, rng, connected=True):
    try:
        ap.send("ABORt:NR:BLER")
        ap.send("ABORt:NR:MEValuation")
        my_sleep(0.3)

        ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")
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
            my_sleep(1)

        base_x_label = f"n{band}({rng})"
        x_label = base_x_label if connected else f"{base_x_label}[失败]"

        dl_bler_str = ap.send("FETCh:NR:BLER:DL:RESult?", 7, True,
                              "DL NR_BLER", x_label, True, True)
        ul_bler_str = ap.send("FETCh:NR:BLER:UL:RESult?", 7, True,
                              "UL NR_BLER", x_label, True, True)
        try:
            dl_bler = float(dl_bler_str.split(',')[7])
            ul_bler = float(ul_bler_str.split(',')[7])
            print(f"n{band} {rng} NR BLER: DL={dl_bler}, UL={ul_bler}")
        except (AttributeError, ValueError, IndexError):
            pass

        txp_power = ap.send("CONFigure:CELL1:NR:SIGN:POWer?")
        txp_x_label = f"{x_label} pwr:{str(txp_power).strip()}dBm"
        txp = ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,
                      "NR TXP AVG", txp_x_label, True, True)
        check_txp(f"NR TXP AVG ({txp_x_label})", txp, tech="NR")

        ap.send("ABORt:NR:BLER")
        ap.send("ABORt:NR:MEValuation")
        ap.send("CONFigure:NR:MEValuation:RESult OFF,OFF,OFF,OFF,OFF,OFF")
    except Exception as e:
        print(f"NR 测量失败: {e}")


def perform_lte_measurement(ap, band, rng, connected=True):
    try:
        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")
        my_sleep(0.3)

        ap.send("CONFigure:LTE:TXP:REPetition SINGLESHOT")
        ap.send("CONFigure:LTE:BLER:REPetition SINGLESHOT")

        ap.send("INITiate:LTE:BLER")
        ap.send("INITiate:LTE:TXP")
        my_sleep(0.5)

        max_wait = 5 if connected else 2
        for i in range(max_wait):
            result = ap.query("FETCh:LTE:BLER:STATe?")
            if "RDY" == result:
                break
            my_sleep(2)

        base_x_label = f"NR:{band}({rng}) OB:{parameter.get('lte_band')} BW:{parameter.get('lte_bw')}"
        x_label = base_x_label if connected else f"{base_x_label}[失败]"

        dl_bler_str = ap.send("FETCh:LTE:BLER:DL:RESult?", 7, True,
                              "DL LTE_BLER", x_label, True, True)
        ul_bler_str = ap.send("FETCh:LTE:BLER:UL:RESult?", 7, True,
                              "UL LTE_BLER", x_label, True, True)
        try:
            dl_bler = float(dl_bler_str.split(',')[7])
            ul_bler = float(ul_bler_str.split(',')[7])
            print(f"n{band} {rng} LTE BLER: DL={dl_bler}, UL={ul_bler}")
        except (AttributeError, ValueError, IndexError):
            pass

        txp_power = ap.send("CONFigure:CELL1:LTE:SIGN:POWer?")
        txp_x_label = f"{x_label} pwr:{str(txp_power).strip()}dBm"
        txp = ap.send("FETCh:LTE:TXP:AVG?", 1, True,
                      "LTE TXP AVG", txp_x_label, True, True)
        check_txp(f"LTE TXP AVG ({txp_x_label})", txp, tech="lte")

        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")
    except Exception as e:
        print(f"LTE 测量失败: {e}")


def _cell_off():
    # if parameter.get('lte_band') is None:
    #     ap.send("ABORt:NR:BLER")
    #     ap.send("ABORt:NR:MEValuation")
    if parameter.get('lte_band') is not None:
        ap.send("ABORt:LTE:BLER")
        ap.send("ABORt:LTE:TXP")
    ap.send("CALL:CELL1 OFF")
    my_sleep(2)
    for i in range(5):
        result = ap.query("CALL:CELL1?")
        if "OFF" == result:
            print("✅ CELL已关闭")
            break
        print("⏳ 等待CELL关闭...")
        my_sleep(2)


def case_start():
    remote_gnb_start()
    remote_diag_start()
    ## line loss configuration
    config_line_loss(parameter)


def case_body():
    is_nsa = parameter.get('lte_band') is not None
    bands = _band_list()
    rng = parameter.get('range', 'LOW')

    if not bands:
        ap.check("参数", False, detail="未提供 band_list", status_msg="无频段可测 fail")
        return

    for band in bands:
        print(f"\n{'='*50}\n▶ 开始 band n{band} ({'NSA' if is_nsa else 'SA'}) 顺序开关小区流程\n{'='*50}")
        bp = _band_param(band)

        config_cell_band(bp)

        ap.send("CALL:CELL1 ON")
        check_phone_at()
        my_sleep(5)

        if not wait_for_ue_connected(ap):
            print(f"⚠️ n{band} 首轮未连接,重新触发注册后重试...")
            check_phone_at()
            my_sleep(5)
            connected = wait_for_ue_connected(ap)
        else:
            connected = True

        if not connected:
            print(f"❌ n{band} UE 未连接,仍记录为失败档并关小区")
            ap.check(f"n{band} UE连接", False, detail="多次查询UE未连接",
                     status_msg=f"n{band} UE未连接 fail")
            perform_nr_measurement(ap, band, rng, connected=False)
            if is_nsa:
                perform_lte_measurement(ap, band, rng, connected=False)
            _cell_off()
            my_sleep(1)
            continue

        calibrate_line_loss(parameter, tolerance=3.0)

        config_nr_slots(bp, NR_SLOTS)
        if is_nsa:
            config_lte_subframes(bp, LTE_SLOTS)

        perform_nr_measurement(ap, band, rng)
        if is_nsa:
            perform_lte_measurement(ap, band, rng)

        restore_line_loss(parameter)

        _cell_off()
        my_sleep(1)


def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    if parameter.get('lte_band') is not None:
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
        print("⏳ 等待CELL关闭...")
        my_sleep(2)

    remote_gnb_stop()
    remote_restart()
