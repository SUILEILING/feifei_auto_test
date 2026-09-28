ALL_BANDS = [1, 2, 3, 5, 7, 8, 12, 20, 25, 28, 34, 38, 39, 40, 41, 50, 51,
             66, 70, 71, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 86]
ALL_BWS = [5, 10, 15, 20, 25, 30, 40, 50, 60, 80, 90, 100]
ALL_SCS = [15, 30]


BAND_BWS = {
    1:  [5, 10, 15, 20],
    2:  [5, 10, 15, 20],
    3:  [5, 10, 15, 20, 25, 30],
    5:  [5, 10, 15, 20],
    7:  [5, 10, 15, 20],
    8:  [5, 10, 15, 20],
    12: [5, 10, 15],
    20: [5, 10, 15, 20],
    25: [5, 10, 15, 20],
    28: [5, 10, 15, 20],
    34: [5, 10, 15],
    38: [5, 10, 15, 20],
    39: [5, 10, 15, 20, 25, 30, 40],
    40: [5, 10, 15, 20, 25, 30, 40, 50, 60, 80],
    41: [10, 15, 20, 40, 50, 60, 80, 90, 100],
    50: [5, 10, 15, 20, 40, 50, 60, 80],
    51: [5],
    66: [5, 10, 15, 20, 40],
    70: [5, 10, 15, 20, 25],
    71: [5, 10, 15, 20],
    74: [5, 10, 15, 20],
    75: [5, 10, 15, 20],
    76: [5],
    77: [10, 15, 20, 40, 50, 60, 80, 90, 100],
    78: [10, 15, 20, 40, 50, 60, 80, 90, 100],
    79: [40, 50, 60, 80, 100],
    80: [5, 10, 15, 20, 25, 30],
    81: [5, 10, 15, 20],
    82: [5, 10, 15, 20],
    83: [5, 10, 15, 20],
    84: [5, 10, 15, 20],
    86: [5, 10, 15, 20, 40],
}


BAND_CAPABILITY = {
    band: {
        15: [b for b in bws if b <= 50],
        30: [b for b in bws if b >= 10],
    }
    for band, bws in BAND_BWS.items()
}

RANGE_OPTIONS = ["LOW", "MID", "HIGH"]
# “无”: 仅DL配置时使用(无上行 RB), 兼容占位, 生成时折叠为单条用例
RB_MODE_NONE = "无"
RB_MODE_OPTIONS = [RB_MODE_NONE,
                   "Edge_Full_Left", "Edge_Full_Right", "Edge_1RB_Left", "Edge_1RB_Right",
                   "Outer_Full", "Inner_Full", "Inner_1RB_Left", "Inner_1RB_Right"]
WAVEFORM_OPTIONS = ["DFTS", "CP"]

POWER_CLASS_OPTIONS = ["None", "1", "1.5", "2", "3"]

POWER_CLASS_CASE_PREFIX = "yc_6.2.1_UE_maximum_output_power"


def is_supported(band, scs, bw):

    return bw in BAND_CAPABILITY.get(band, {}).get(scs, [])


def supported_bws(band, scs):
    return list(BAND_CAPABILITY.get(band, {}).get(scs, []))
