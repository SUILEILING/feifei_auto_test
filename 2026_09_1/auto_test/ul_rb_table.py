
RB_MODES = (
    "Edge_Full_Left",
    "Edge_Full_Right",
    "Edge_1RB_Left",
    "Edge_1RB_Right",
    "Outer_Full",
    "Inner_Full",
    "Inner_1RB_Left",
    "Inner_1RB_Right",
)


UL_RB_ALLOCATION = {
    # 3 MHz
    (3, 15):  {"Edge_Full_Left": "0,2", "Edge_Full_Right": "13,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "14,1",  "Outer_Full": "0,15",  "Inner_Full": "4,6",    "Inner_1RB_Left": "0,1", "Inner_1RB_Right": "13,1"},

    # 5 MHz  (60kHz: N/A)
    (5, 15):  {"Edge_Full_Left": "0,2", "Edge_Full_Right": "23,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "24,1",  "Outer_Full": "0,25",  "Inner_Full": "6,12",   "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "23,1"},
    (5, 30):  {"Edge_Full_Left": "0,2", "Edge_Full_Right": "9,2",   "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "10,1",  "Outer_Full": "0,10",  "Inner_Full": "2,5",    "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "9,1"},

    # 10 MHz
    (10, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "50,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "51,1",  "Outer_Full": "0,50",  "Inner_Full": "12,25",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "50,1"},
    (10, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "22,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "23,1",  "Outer_Full": "0,24",  "Inner_Full": "6,12",   "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "22,1"},
    (10, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "9,2",   "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "10,1",  "Outer_Full": "0,10",  "Inner_Full": "2,5",    "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "9,1"},

    # 15 MHz
    (15, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "77,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "78,1",  "Outer_Full": "0,75",  "Inner_Full": "18,36",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "77,1"},
    (15, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "36,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "37,1",  "Outer_Full": "0,36",  "Inner_Full": "9,18",   "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "36,1"},
    (15, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "16,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "17,1",  "Outer_Full": "0,18",  "Inner_Full": "4,9",    "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "16,1"},

    # 20 MHz
    (20, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "104,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "105,1", "Outer_Full": "0,100", "Inner_Full": "25,50",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "104,1"},
    (20, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "49,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "50,1",  "Outer_Full": "0,50",  "Inner_Full": "12,25",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "49,1"},
    (20, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "22,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "23,1",  "Outer_Full": "0,24",  "Inner_Full": "6,12",   "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "22,1"},

    # 25 MHz
    (25, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "131,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "132,1", "Outer_Full": "0,128", "Inner_Full": "32,64",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "131,1"},
    (25, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "63,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "64,1",  "Outer_Full": "0,64",  "Inner_Full": "16,32",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "63,1"},
    (25, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "29,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "30,1",  "Outer_Full": "0,30",  "Inner_Full": "7,15",   "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "29,1"},

    # 30 MHz
    (30, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "158,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "159,1", "Outer_Full": "0,160", "Inner_Full": "40,80",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "158,1"},
    (30, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "76,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "77,1",  "Outer_Full": "0,75",  "Inner_Full": "18,36",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "76,1"},
    (30, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "36,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "37,1",  "Outer_Full": "0,36",  "Inner_Full": "9,18",   "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "36,1"},

    # 35 MHz
    (35, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "186,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "187,1", "Outer_Full": "0,180", "Inner_Full": "45,90",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "186,1"},
    (35, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "90,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "91,1",  "Outer_Full": "0,90",  "Inner_Full": "22,45",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "90,1"},
    (35, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "42,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "43,1",  "Outer_Full": "0,40",  "Inner_Full": "10,20",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "42,1"},

    # 40 MHz
    (40, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "214,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "215,1", "Outer_Full": "0,216", "Inner_Full": "54,108", "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "214,1"},
    (40, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "104,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "105,1", "Outer_Full": "0,100", "Inner_Full": "25,50",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "104,1"},
    (40, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "49,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "50,1",  "Outer_Full": "0,50",  "Inner_Full": "12,25",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "49,1"},

    # 45 MHz
    (45, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "240,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "241,1", "Outer_Full": "0,240", "Inner_Full": "60,120", "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "240,1"},
    (45, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "117,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "118,1", "Outer_Full": "0,108", "Inner_Full": "27,54",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "117,1"},
    (45, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "56,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "57,1",  "Outer_Full": "0,54",  "Inner_Full": "13,27",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "56,1"},

    # 50 MHz
    (50, 15): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "268,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "269,1", "Outer_Full": "0,270", "Inner_Full": "67,135", "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "268,1"},
    (50, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "131,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "132,1", "Outer_Full": "0,128", "Inner_Full": "32,64",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "131,1"},
    (50, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "63,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "64,1",  "Outer_Full": "0,64",  "Inner_Full": "16,32",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "63,1"},

    # 60 MHz  (15kHz: N/A)
    (60, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "160,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "161,1", "Outer_Full": "0,162", "Inner_Full": "40,81",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "160,1"},
    (60, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "77,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "78,1",  "Outer_Full": "0,75",  "Inner_Full": "18,36",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "77,1"},

    # 70 MHz  (15kHz: N/A)
    (70, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "187,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "188,1", "Outer_Full": "0,180", "Inner_Full": "45,90",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "187,1"},
    (70, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "91,2",  "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "92,1",  "Outer_Full": "0,90",  "Inner_Full": "22,45",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "91,1"},

    # 80 MHz  (15kHz: N/A)
    (80, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "215,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "216,1", "Outer_Full": "0,216", "Inner_Full": "54,108", "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "215,1"},
    (80, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "105,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "106,1", "Outer_Full": "0,100", "Inner_Full": "25,50",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "105,1"},

    # 90 MHz  (15kHz: N/A)
    (90, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "243,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "244,1", "Outer_Full": "0,243", "Inner_Full": "60,120", "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "243,1"},
    (90, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "119,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "120,1", "Outer_Full": "0,120", "Inner_Full": "30,60",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "119,1"},

    # 100 MHz  (15kHz: N/A)
    (100, 30): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "271,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "272,1", "Outer_Full": "0,270", "Inner_Full": "67,135", "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "271,1"},
    (100, 60): {"Edge_Full_Left": "0,2", "Edge_Full_Right": "133,2", "Edge_1RB_Left": "0,1", "Edge_1RB_Right": "134,1", "Outer_Full": "0,135", "Inner_Full": "32,64",  "Inner_1RB_Left": "1,1", "Inner_1RB_Right": "133,1"},
}

 
# ---------------------------------------------------------------------------
# CP-OFDM 与 DFT-s-OFDM 的差异
# ---------------------------------------------------------------------------
# 对照 Table 6.1-1 的 CP 行与 DFT-s 行:只有 Outer_Full / Inner_Full 两列不同

_CP_FULL_OVERRIDES = {
    (3, 15):   {"Outer_Full": "0,15",  "Inner_Full": "4,7"},
    (5, 15):   {"Outer_Full": "0,25",  "Inner_Full": "6,13"},
    (5, 30):   {"Outer_Full": "0,11",  "Inner_Full": "2,5"},
    (10, 15):  {"Outer_Full": "0,52",  "Inner_Full": "13,26"},
    (10, 30):  {"Outer_Full": "0,24",  "Inner_Full": "6,12"},
    (10, 60):  {"Outer_Full": "0,11",  "Inner_Full": "2,5"},
    (15, 15):  {"Outer_Full": "0,79",  "Inner_Full": "19,39"},
    (15, 30):  {"Outer_Full": "0,38",  "Inner_Full": "9,19"},
    (15, 60):  {"Outer_Full": "0,18",  "Inner_Full": "4,9"},
    (20, 15):  {"Outer_Full": "0,106", "Inner_Full": "26,53"},
    (20, 30):  {"Outer_Full": "0,51",  "Inner_Full": "12,25"},
    (20, 60):  {"Outer_Full": "0,24",  "Inner_Full": "6,12"},
    (25, 15):  {"Outer_Full": "0,133", "Inner_Full": "33,67"},
    (25, 30):  {"Outer_Full": "0,65",  "Inner_Full": "16,33"},
    (25, 60):  {"Outer_Full": "0,31",  "Inner_Full": "7,15"},
    (30, 15):  {"Outer_Full": "0,160", "Inner_Full": "40,80"},
    (30, 30):  {"Outer_Full": "0,78",  "Inner_Full": "19,39"},
    (30, 60):  {"Outer_Full": "0,36",  "Inner_Full": "9,19"},
    (35, 15):  {"Outer_Full": "0,188", "Inner_Full": "47,94"},
    (35, 30):  {"Outer_Full": "0,92",  "Inner_Full": "23,46"},
    (35, 60):  {"Outer_Full": "0,44",  "Inner_Full": "11,22"},
    (40, 15):  {"Outer_Full": "0,216", "Inner_Full": "54,108"},
    (40, 30):  {"Outer_Full": "0,106", "Inner_Full": "26,53"},
    (40, 60):  {"Outer_Full": "0,51",  "Inner_Full": "12,25"},
    (45, 15):  {"Outer_Full": "0,242", "Inner_Full": "60,121"},
    (45, 30):  {"Outer_Full": "0,119", "Inner_Full": "29,59"},
    (45, 60):  {"Outer_Full": "0,58",  "Inner_Full": "14,29"},
    (50, 15):  {"Outer_Full": "0,270", "Inner_Full": "67,135"},
    (50, 30):  {"Outer_Full": "0,133", "Inner_Full": "33,67"},
    (50, 60):  {"Outer_Full": "0,65",  "Inner_Full": "16,33"},
    (60, 30):  {"Outer_Full": "0,162", "Inner_Full": "40,81"},
    (60, 60):  {"Outer_Full": "0,79",  "Inner_Full": "19,39"},
    (70, 30):  {"Outer_Full": "0,189", "Inner_Full": "47,95"},
    (70, 60):  {"Outer_Full": "0,93",  "Inner_Full": "23,47"},
    (80, 30):  {"Outer_Full": "0,217", "Inner_Full": "54,109"},
    (80, 60):  {"Outer_Full": "0,107", "Inner_Full": "26,53"},
    (90, 30):  {"Outer_Full": "0,245", "Inner_Full": "61,123"},
    (90, 60):  {"Outer_Full": "0,121", "Inner_Full": "30,61"},
    (100, 30): {"Outer_Full": "0,273", "Inner_Full": "68,137"},
    (100, 60): {"Outer_Full": "0,135", "Inner_Full": "33,67"},
}

UL_RB_ALLOCATION_CP = {}
for _key, _modes in UL_RB_ALLOCATION.items():
    _cp = dict(_modes)
    _cp.update(_CP_FULL_OVERRIDES.get(_key, {}))
    UL_RB_ALLOCATION_CP[_key] = _cp


_ALLOCATION_BY_WAVEFORM = {
    "DFTS": UL_RB_ALLOCATION,
    "CP": UL_RB_ALLOCATION_CP,
}


def _normalize_waveform(waveform):
    wf = str(waveform or "DFTS").upper()
    return "DFTS" if wf.startswith("DFT") else "CP"


def get_ul_rb(nr_bw, scs, rb_mode, waveform="DFTS"):
    table = _ALLOCATION_BY_WAVEFORM[_normalize_waveform(waveform)]
    modes = table.get((nr_bw, scs))
    if not modes:
        return None
    return modes.get(rb_mode)


UL_RB_TABLE = {mode: {} for mode in RB_MODES}
UL_RB_TABLE_CP = {mode: {} for mode in RB_MODES}
for _key, _modes in UL_RB_ALLOCATION.items():
    for _mode, _val in _modes.items():
        UL_RB_TABLE[_mode][_key] = _val
for _key, _modes in UL_RB_ALLOCATION_CP.items():
    for _mode, _val in _modes.items():
        UL_RB_TABLE_CP[_mode][_key] = _val
del _key, _modes, _mode, _val, _cp
