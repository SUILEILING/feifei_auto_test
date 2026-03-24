from typing import List, Dict


TEST_STEPS: List[Dict] = [
    # line loss configuration
    {"command": "CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,25.00,6000000000,25.00",},
    {"command": "CONFigure:BASE:FDCorrection:SAVE",},
    {"command": "CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX",},
    {"command": "CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX",},
    # band bw scs range configuration
    {"command": "CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator 77",},
    {"command": "CONFigure:CELL1:NR:SIGN:BWidth:DL BW100",},
    {"command": "CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz30",},
    {"command": "CONFigure:CELL1:NR:CONFig:RANGe LOW",},

    {"command": "CALL:CELL1 ON",},
    {"command": "Sleep 10000",},
    {
        "command": "CONFigure:CELL1:NR:SIGN:UE:STATe?",
        "expected_result": "Connected",
        "loop_config": {
            "enable": True,
            "times": 20,
            "sleep_ms": 1000,
            "expected_result": "Connected"
        },
    },
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT:CLEar",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 3",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT:UPDate",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT:APPLy",},
    {"command": "CONFigure:NR:MEValuation:REPetition SINGLESHOT",},
    {"command": "CONFigure:NR:BLER:REPetition SINGLESHOT",},
    {"command": "CONFigure:NR:MEValuation:RESult ON,OFF,OFF,OFF,OFF,OFF",},
    {"command": "INITiate:NR:BLER",},
    {"command": "INITiate:NR:MEValuation",},
    {
        "command": "FETCh:NR:MEValuation:STATe?",
        "expected_result": "RDY",
        "loop_config": {
            "enable": True,
            "times": 10,
            "sleep_ms": 1000,
            "expected_result": "RDY"
        },
    },
    {
        "command": "FETCh:NR:BLER:STATe?",
        "expected_result": "RUN",
        "loop_config": {
            "enable": True,
            "times": 10,
            "sleep_ms": 1000,
            "expected_result": "RDY"
        },
    },
    {"command": "Sleep 5000",},
    {"command": "FETCh:NR:BLER:DL:RESult?",
     "expected_result": "0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
    },
    {"command": "FETCh:NR:BLER:UL:RESult?",
     "expected_result": "0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
    },
    {"command": "CONFigure:NR:MEValuation:REPetition CONTINUOUS",},
    {"command": "CONFigure:NR:BLER:REPetition CONTINUOUS",},
    {"command": "Sleep 20000",},
    {"command": "FETCh:NR:MEValuation:TXP:AVG?",
     "expected_result": "0,15.700",
    },
    {"command": "ABORt:NR:BLER ",},
    {"command": "ABORt:NR:MEValuation",},
    {"command": "CALL:CELL1 OFF",},
    {"command": "Sleep 5000",},
    {
        "command": "CALL:CELL1?",
        "expected_result": "OFF",
        "loop_config": {
            "enable": True,
            "times": 20,
            "sleep_ms": 1000,
            "expected_result": "OFF"
        },
    },
]


# ==============================================
# 需要在后续循环中跳过的步骤（第一次执行后不再执行）
# ==============================================

SKIP_IN_NEXT_CYCLES: List[Dict] = [
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT:CLEar",
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh",
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh",
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 3",
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT8:CTYPe PUSCh",
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:SLOT9:CTYPe PUSCh",
    },

]



