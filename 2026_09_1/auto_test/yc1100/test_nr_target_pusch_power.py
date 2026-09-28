from lib.var import *
from common import *

DEFAULT_PARAMETER = {
    'lineLoss1': 25.00,
    'nr_band': 1,
    'nr_bw': 20,
    'scs': 15,
    'range': 'LOW',
    'min_power': -20,
    'max_power': None,
    'step': 2,
    'nr_slots': None,
    'mcs': None,
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


def wait_ready(state_query, label):
    for i in range(5):
        result = ap.query(state_query)
        if "RDY" == result:
            print(f"✅ 第 {i+1} 次查询: {label}已准备好")
            return True
        else:
            print(f"⏳ 第 {i+1} 次查询: {label}未准备好,状态={result}")
            my_sleep(2)
    print(f"❌ {label} 超时未就绪")
    return False


def wait_fresh_ready(state_query, label, rdy_timeout=10):
    for i in range(int(rdy_timeout / 0.5)):
        st = str(ap.query(state_query))
        if "RDY" in st:
            print(f"✅ 第 {i+1} 次查询: {label}本次测量已完成")
            return True
        my_sleep(0.5)
    print(f"❌ {label} 超时未就绪")
    return False


def measure_nr_txp_avg(x_label, target=None, tol=None):
    my_sleep(1)
    ap.send("INITiate:NR:BLER")
    ap.send("INITiate:NR:MEValuation")
    if not wait_fresh_ready("FETCh:NR:MEValuation:STATe?", "MEValuation"):
        return None

    raw = ap.query("FETCh:NR:MEValuation:TXP:AVG?")
    try:
        y = float(str(raw).split(',')[1])
    except Exception as e:
        print(f"⚠️ 解析 TXP:AVG 失败: raw={raw}, err={e}")
        y = None

    final_label = x_label
    if target is not None and tol is not None:
        if y is None or abs(y - target) > tol:
            final_label = f"{x_label}[失败]"

    ap.send("FETCh:NR:MEValuation:TXP:AVG?", 1, True, "NR Target TXP", final_label, True, True,
            record_step=False)
    return y



TXP_AVG_CMD = "FETCh:NR:MEValuation:TXP:AVG?"
TARGET_TOL = 1
TARGET_MAX_TRIES = 3


def check_target_power(target, x_label):
    last_y = None
    for i in range(TARGET_MAX_TRIES):
        y = measure_nr_txp_avg(x_label, target=target, tol=TARGET_TOL)
        if y is None:
            print(f"⏳ 第 {i+1} 次查询: {x_label} TXP:AVG 读取失败")
            my_sleep(1)
            continue
        last_y = y
        diff = y - target
        if abs(diff) <= TARGET_TOL:
            print(f"✅ 第 {i+1} 次查询: set target={target:.2f} dBm, 实测 y={y:.2f} dBm, "
                  f"diff={diff:+.2f} -> PASS")
            ap.check(TXP_AVG_CMD, True, attempts=i + 1,
                     status_msg=f"check tager power ≈≈ |cell power±TARGET_TOL| pass ",
                     detail=f"tager power={target:.2f},cell power={y:.2f}, diff={diff:+.2f}")
            return
        print(f"⏳ 第 {i+1} 次查询: set target={target:.2f} dBm, 实测 y={y:.2f} dBm, "
              f"diff={diff:+.2f} 超出容差±{TARGET_TOL}")
        if i < TARGET_MAX_TRIES - 1:
            my_sleep(2)
            ap.send("INITiate:NR:BLER")
            ap.send("INITiate:NR:MEValuation")

    if last_y is None:
        ap.check(TXP_AVG_CMD, False, attempts=TARGET_MAX_TRIES,
                 status_msg=f"check tager power !=|cell power±TARGET_TOL| fail ",
                 detail=f"tager power={target:.2f},cell power=None(读取失败)")
    else:
        ap.check(TXP_AVG_CMD, False, attempts=TARGET_MAX_TRIES,
                 status_msg=f"check tager power !=|cell power±TARGET_TOL| fail ",
                 detail=f"tager power={target:.2f},cell power={last_y:.2f}, diff={last_y - target:+.2f}")


def case_start():
    remote_gnb_start()
    remote_diag_start()

    ## line loss configuration
    config_line_loss(parameter)

    ## band bw scs range configuration
    config_cell_band(parameter)


def case_body():
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

    calibrate_line_loss(parameter, tolerance=3.0)

    config_nr_slots(parameter, NR_SLOTS)

    ap.send("CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF")
    ap.send("CONFigure:NR:MEValuation:REPetition SINGLESHOT")
    ap.send("CONFigure:NR:BLER:REPetition SINGLESHOT")
    ap.send("INITiate:NR:BLER")
    ap.send("INITiate:NR:MEValuation")
    my_sleep(1)
    wait_ready("FETCh:NR:MEValuation:STATe?", "MEValuation")
    wait_ready("FETCh:NR:BLER:STATe?", "BLER测试")

    x = measure_nr_txp_avg("baseline")
    if x is None:
        print("❌ 未能获取当前 txpower(x)，无法确定扫描上限，终止扫描")
        ap.check(TXP_AVG_CMD, False,
                 "未能获取当前 txpower(x)(基线 TXP:AVG 读取失败),扫描未执行")
        return
    print(f"📍 当前 txpower x = {x:.2f} dBm")

    ap.send("CONFigure:CELL1:NR:SIGN:UPC TARGET")

    min_power = parameter['min_power']
    max_power = parameter['max_power']
    step = parameter['step']
    upper = x if max_power is None else max_power

    if step <= 0:
        print(f"❌ step({step}) 必须为正数(往上加)，终止扫描")
        ap.check(TXP_AVG_CMD, False, f"step({step}) 必须为正数(往上加),扫描未执行")
        return

    if upper + 1e-9 < min_power:
        src = "当前txpower x" if max_power is None else "max_power"
        print(f"❌ 扫描上限({src}={upper:.2f}) 低于起点 min_power({min_power})，无点可扫，终止扫描")
        ap.check(TXP_AVG_CMD, False,
                 f"扫描上限({src}={upper:.2f}dBm) 低于 min_power({min_power}dBm),无点可扫"
                 f"(疑似上行异常/基站NR未正常发射)")
        return

    print(f"🎯 target 扫描范围: {min_power} -> {upper:.2f} dBm, step={step} "
          f"(上限来源: {'当前txpower x' if max_power is None else 'max_power'})")

    target = float(min_power)
    while target <= upper + 1e-9:
        ap.send(f"CONFigure:CELL1:NR:SIGN:UPC TARGET,0.00,{target:.2f},0.5")
        my_sleep(0.5)

        check_target_power(target, f"{target:.0f}dBm")

        target += step


def case_clear():
    ap.send("ABORt:NR:BLER")
    ap.send("ABORt:NR:MEValuation")

    ap.send("CALL:CELL1 OFF")
    my_sleep(1)

    remote_diag_stop()

    for i in range(5):
        result = ap.query("CALL:CELL1?")
        if "OFF" == result:
            print(f"✅ CELL已关闭")
            break
        else:
            print(f"⏳ 等待CELL关闭...")
            my_sleep(2)

    remote_gnb_stop()
    remote_restart()
