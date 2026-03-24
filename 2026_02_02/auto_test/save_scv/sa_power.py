from typing import List, Dict


TEST_STEPS: List[Dict] = [
    # line loss configuration
    {"command": "CONFigure:BASE:FDCorrection:CTABle:CREate LineLossTable_1,100000000,25.00,6000000000,25.00",},
    {"command": "CONFigure:BASE:FDCorrection:SAVE",},
    {"command": "CONFigure:FDCorrection:ACTivate LineLossTable_1,1,IO,RXTX",},
    {"command": "CONFigure:FDCorrection:ACTivate LineLossTable_1,1,OUT,TX",},
    # band bw scs range configuration
    {"command": "CONFigure:CELL1:NR:SIGN:COMMon:FBANd:INDCator 79",},
    {"command": "CONFigure:CELL1:NR:SIGN:BWidth:DL BW100",},
    {"command": "CONFigure:CELL1:NR:SIGN:COMMon:FBANd:DL:SCSList:SCSPacing kHz30",},
    {"command": "CONFigure:CELL1:NR:CONFig:RANGe LOW",},

    {"command": "CALL:CELL1 ON",},
    {"command": "Sleep 5000",},
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
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT3:DL:TIND 5",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT4:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT4:DL:TIND 4",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT5:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT5:DL:TIND 3",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT6:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT6:DL:TIND 2",},


    {"command": "CONFigure:CELL1:NR:SIGN:SLOT10:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT10:DL:TIND 8",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT11:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT11:DL:TIND 7",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT12:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT12:DL:TIND 6",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT13:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT13:DL:TIND 5",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT14:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT14:DL:TIND 4",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT15:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT15:DL:TIND 3",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT16:CTYPe PDSCh",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT16:DL:TIND 2",},

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
    {"command": "Sleep 3000",},
    {"command": "FETCh:NR:BLER:DL:RESult?",
     "expected_result": "0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
    },
    {"command": "CONFigure:NR:MEValuation:REPetition CONTINUOUS",},
    {"command": "CONFigure:NR:BLER:REPetition CONTINUOUS",},
    {"command": "Sleep 10000",},
    {"command": "FETCh:NR:MEValuation:TXP:AVG?",
     "expected_result": "0,15.700",
    },
    {"command": "CONFigure:NR:MEValuation:REPetition SINGLESHOT",},
    {"command": "CONFigure:NR:BLER:REPetition SINGLESHOT",},
    

    # ===================== 第一个循环（条件循环）=====================
    {"for": "on"},
    {"command": "CONFigure:CELL1:NR:SIGN:POWer {current_power}",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT:APPLy",},
    {"command": "INITiate:NR:BLER ",},
    {"command": "INITiate:NR:MEValuation",},
    {"command": "Sleep 2000",},
    {
        "command": "FETCh:NR:BLER:DL:RESult?",
        "expected_result": "0,67,67,0,0,70752,70752,0.000,0.000,211.200,211.200",
        "check_index": 8,
        "condition": ">",
        "threshold": 0.5,
        "data_type": "float",
    },
    {
        "command": "CONFigure:CELL1:NR:SIGN:UE:STATe?",
        "expected_result": "Connected",
        "check_index": 1,
        "condition": "!=",
        "data_type": "str",
        "expected_value": "Connected",
    },
    {"for": "off"},


    # ===================== 第二个循环（固定4次）=====================
    {"for": "on", "times": 3},
    {"command": "CONFigure:CELL1:NR:SIGN:POWer {current_power}",},
    {"command": "CONFigure:CELL1:NR:SIGN:SLOT:APPLy",},
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
    {"if": "Connected"},
        {"command": "INITiate:NR:BLER ",},
        {"command": "INITiate:NR:MEValuation",},
        {"command": "Sleep 5000",},
        {"command": "FETCh:NR:BLER:DL:RESult?",
        "expected_result": "0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
        },
        {"command": "CONFigure:CELL1:NR:SIGN:UE:STATe?",
        "expected_result": "Connected",
        },
    {"else": ""},
        {"command": "CALL:CELL1 ON",},
        {"command": "Sleep 10000",},
        {"command": "CONFigure:CELL1:NR:SIGN:POWer {current_power}",},
        {"command": "CONFigure:CELL1:NR:SIGN:SLOT:APPLy",},
        {"command": "INITiate:NR:BLER ",},
        {"command": "INITiate:NR:MEValuation",},
        {"command": "Sleep 20000",},
        {"command": "FETCh:NR:BLER:DL:RESult?",
        "expected_result": "0,0,0,0,0,0,0,0.000,0.000,0.000,0.000",
        },
        {"command": "CONFigure:CELL1:NR:SIGN:UE:STATe?",
        "expected_result": "Connected",
        },
    {"for": "off"},
    
    


    {"command": "ABORt:NR:BLER ",},
    {"command": "ABORt:NR:MEValuation",},
    {"command": "CALL:CELL1 OFF",},
    {"command": "Sleep 10000",},
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


]