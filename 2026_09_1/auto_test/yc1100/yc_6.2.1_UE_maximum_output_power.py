from lib.var import *
from common import *

DEFAULT_PARAMETER = {
    'lineLoss1': 25.00,
    'nr_band': 78,
    'nr_bw': 100,
    'scs': 30,
    'range': 'LOW',
    "nr_slots": None,
    "mcs": None,             # 覆盖 slot 表的 MCS1(软件 MCS 配置), None=用 NR_SLOTS 默认
    "rb_mode": "Inner_Full",
    'power_class': None,
    'meas_retries': 3,       # TXP 测量失败时的重试次数
    'waveform': 'DFTS',      # DFTS(DFT-s-OFDM) 或 CP(CP-OFDM)
}


# ── 调制方式(Modulation): 根据 UL MCS1 动态显示 ────────────────
MODULATION_TABLE = {
    range(0, 9):  "QPSK",
    range(10, 16): "16QAM",
    range(17, 27): "64QAM",
}

# ── Power Class 合格窗口 (LowLimit, UpLimit) ──────────────────
PC_LIMITS = {
    5:   (18, 22),
    4:   (19, 23),
    3:   (21, 25),
    2:   (23, 28),
    1.5: (26, 31),
    1:   (28, 33),
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
_external_params = None


def update_parameters(external_params=None):
    global parameter, _external_params
    _external_params = external_params
    if external_params:
        for key, value in external_params.items():
            if key in parameter:
                parameter[key] = value


def _power_limits():
    pc = parameter.get('power_class')
    if pc is None:
        return None
    return PC_LIMITS.get(pc)


def _limit_labels():
    limits = _power_limits()
    if not limits:
        return "无", "无"
    return limits


def _ul_modulation():
    def _name_for(mcs):
        for rng, name in MODULATION_TABLE.items():
            if mcs in rng:
                return name
        return "N/A"
    # 软件配置了 mcs 时实际下发以它为准, 调制显示同步
    if parameter.get('mcs') is not None:
        return _name_for(parameter['mcs'])
    for item in NR_SLOTS.get("UL", []):
        for _slot, fields in item.items():
            if "MCS1" in fields:
                return _name_for(fields["MCS1"])
    return "N/A"


def _report_title():
    pc = parameter.get('power_class')
    tag = f"pc{pc}" if pc is not None else "pcnone"
    return f"yc_6.2.1_UE_maximum_output_power_{tag}"



def case_start():
    remote_gnb_start()
    remote_diag_start()
    config_line_loss(parameter)
    config_cell_band(parameter)


def case_body():
    title = _report_title()
    limits = _power_limits()
    low_label, up_label = _limit_labels()
    modulation = _ul_modulation()
    if limits is None:
        print(f"📏 [6.2.1] 未指定 power_class(或未知等级),仅判定 TXP>0(check_txp) | 调制: {modulation}")
    else:
        lower, upper = limits
        print(f"📏 [6.2.1] {title}: 合格窗口 LowLimit={lower} / UpLimit={upper} dBm | 调制: {modulation}")

    ap.record_value("LowLimit", low_label)
    ap.record_value("UpLimit", up_label)
    ap.record_value("Modulation", modulation)

    ap.send("CALL:CELL1 ON")
    check_phone_at()
    my_sleep(5)

    connected = False
    for i in range(10):
        result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
        if '"Connected"' == result:
            print(f"✅ 第 {i+1} 次查询: UE已连接")
            connected = True
            break
        else:
            print(f"⏳ 第 {i+1} 次查询: UE未连接")
            my_sleep(2)

    if not connected:
        print("⚠️ 首轮未连接，重新触注册后重试...")
        check_phone_at()
        my_sleep(5)
        for i in range(50):
            result = ap.query("CONFigure:CELL1:NR:SIGN:UE:STATe?")
            if '"Connected"' == result:
                print(f"✅ 重试第 {i+1} 次查询: UE已连接")
                connected = True
                break
            else:
                print(f"⏳ 重试第 {i+1} 次查询: UE未连接")
                my_sleep(0.25)

    if not connected:
        print("❌ UE 多次未连接，跳过测量，直接进入 case_clear")
        ap.check(title, False,
                 detail="UE 未连接", status_msg="UE 未连接，无法测量最大功率 fail")
        return

    calibrate_line_loss(parameter, tolerance=3.0)

    config_nr_slots(parameter, NR_SLOTS)        

    ap.send("CONFigure:CELL1:NR:SIGN:UPC MAX")
    ap.send("CONFigure:CELL1:NR:SIGN:SLOT:APPLy")

    cell_power = ap.send("CONFigure:CELL1:NR:SIGN:POWer?")
    try:
        arfcn_raw = ap.send("CONFigure:CELL1:NR:SIGN:ARFCn?")
        nr_arfcn = str(arfcn_raw).strip().split(',')[0].strip()
        if nr_arfcn and isinstance(_external_params, dict):
            base_range = str(_external_params.get('range', parameter.get('range', ''))).strip()
            _external_params['range'] = f"{base_range}:{nr_arfcn}" if base_range else nr_arfcn
    except Exception as e:
        print(f"⚠️ [6.2.1] 获取 NR ARFCn 失败: {e}")

    ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")
    ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")

    def measure_pumax():
        ap.send("ABORt:NR:MEValuation")
        ap.send("ABORt:NR:BLER")
        my_sleep(0.25)
        ap.send("INITiate:NR:BLER")
        ap.send("INITiate:NR:MEValuation")
        ap.send("*OPC?")

        ready = False
        for _ in range(20):
            if "RDY" in str(ap.query("FETCh:NR:MEValuation:STATe?")):
                ready = True
                break
            my_sleep(0.5)
        if not ready:
            print("❌ TXP 测量超时未就绪")
            return None
        return ap.query("FETCh:NR:MEValuation:TXP:AVG?")

    raw = None
    pumax = None
    for attempt in range(1, parameter.get('meas_retries', 3) + 1):
        raw = measure_pumax()
        if raw is not None:
            try:
                pumax = float(str(raw).split(',')[1])
                print(f"📈 第 {attempt} 次测量: Pumax = {pumax:.2f} dBm")
                break
            except (ValueError, IndexError):
                print(f"⚠️ 第 {attempt} 次测量结果解析失败: {raw}")
        else:
            print(f"⏳ 第 {attempt} 次测量失败，重试...")
        my_sleep(1)

    x_label = f"nr_cell_power:{str(cell_power).strip()} dBm "
    if pumax is None:
        ap.check(title, False,
                 attempts=parameter.get('meas_retries', 3),
                 detail=f"TXP 读取失败: {raw}",
                 status_msg="Pumax 读取失败 fail")
        ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,
                "MeasuredValue", f"{x_label}[失败]", False, False, record_step=False)
    elif limits is None:
        check_txp(f"{title} {x_label}", raw)
        chart_x = x_label if pumax >= 0 else f"{x_label}[失败]"
        ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,
                "MeasuredValue", chart_x, False, False, record_step=False)
    else:
        lower, upper = limits
        passed = (lower <= pumax <= upper)
        status_msg = (f"Pumax={pumax:.2f}dBm 在 [{lower},{upper}] 内 pass"
                      if passed else
                      f"Pumax={pumax:.2f}dBm 超出 [{lower},{upper}] fail")
        ap.check(title, passed,
                 detail=f"Pumax={pumax:.2f}dBm, LowLimit={lower}, UpLimit={upper}",
                 status_msg=status_msg)

        chart_x = x_label if passed else f"{x_label}[失败]"
        ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True,
                "MeasuredValue", chart_x, False, False, record_step=False)
        print(f"{'✅' if passed else '❌'} 6.2.1 判定: Pumax={pumax:.2f}dBm -> "
              f"{'PASS' if passed else 'FAIL'}")

    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")
    ap.send("CONFigure:NR:MEValuation:RESult OFF,OFF,OFF,OFF,OFF,OFF")


def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")

    ap.send("CALL:CELL1 OFF")
    ap.send("*OPC?")

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
