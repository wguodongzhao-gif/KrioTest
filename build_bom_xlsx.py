#!/usr/bin/env python3
"""
Build a production-grade multi-sheet BOM xlsx for the Looki-class device,
using only the Python standard library (zipfile + xml strings).

Output: Looki_Class_HW_BOM_v1.0.xlsx
"""

import zipfile
import os
from datetime import date
from xml.sax.saxutils import escape

OUTFILE = "Looki_Class_HW_BOM_v1.0.xlsx"

# ---------------------------------------------------------------------------
# DATA: Main BOM
# Columns:
# Item, Category, Function, Manufacturer, MPN, Package, Spec/Notes,
# Qty, Unit Cost USD (10K MOQ), Ext Cost USD, Source/Distributor, Lead Time, Tier, Risk
# ---------------------------------------------------------------------------

BOM_HEADERS = [
    "#", "Category", "Function / Sub-Block", "Manufacturer", "MPN",
    "Package", "Key Spec / Notes",
    "Qty", "Unit Cost USD\n@10K MOQ", "Ext Cost USD",
    "Distributor / Source", "Lead Time (wk)", "Tier", "Risk Level"
]

# Each row: tuple of values matching BOM_HEADERS (#, Cat, Func, Mfr, MPN, Pkg, Spec, Qty, Unit, Ext (formula will be set), Dist, LT, Tier, Risk)
BOM_ROWS = [
    # ---- 1. SoC / Application Processor ----
    (1,  "1. Main SoC",        "Application SoC (ISP+NPU+WiFi/BLE)", "SigmaStar",        "SSC338Q",            "LFBGA-260, 9.5x9.5mm", "2x A7@1GHz, 0.5T NPU, 5MP ISP, ext Wi-Fi", 1, 4.50,  "Sertek / WPI", 12, "T1", "Med"),
    (2,  "1. Main SoC",        "Wi-Fi/BLE Combo (ext)",              "Realtek",          "RTL8821CS",          "QFN-40",               "Wi-Fi 5 + BT 5.0, SDIO/UART",              1, 1.80,  "WPG / Realtek", 10, "T1", "Low"),
    (3,  "1. Main SoC",        "DDR2 SDRAM (PoP/side)",              "ESMT",             "M14D5121632A-2.5BG", "FBGA-90",              "512Mb DDR2 @ 333MHz",                       1, 1.20,  "ESMT direct",   8, "T1", "Low"),
    (4,  "1. Main SoC",        "SPI NOR Flash (boot+kernel)",        "Winbond",          "W25Q128JVSIQ",       "SOIC-8 208mil",        "128Mbit, 1.8/3.3V, QSPI",                   1, 0.65,  "Avnet",         8, "T2", "Low"),
    (5,  "1. Main SoC",        "eMMC 5.1 (storage)",                 "Kioxia",           "THGBMNG5D1LBAIL",    "BGA-153",              "8GB eMMC 5.1, HS400",                       1, 3.40,  "Kioxia / WPI",  14, "T1", "Med"),

    # ---- 2. Always-On MCU (heterogeneous compute) ----
    (6,  "2. AON MCU",         "Always-On / sensor hub",             "Nordic",           "nRF5340-QKAA",       "aQFN-94 7x7mm",        "Dual-core M33, BLE 5.4, < 3uA sleep",        1, 4.20,  "Nordic / Mouser", 10, "T1", "Low"),
    (7,  "2. AON MCU",         "AON external 32.768kHz XTAL",        "Epson",            "FC-135 32.768kHz",   "3.2x1.5mm",            "+/-20ppm, CL=12.5pF",                       1, 0.18,  "Epson / Digikey", 6, "T2", "Low"),

    # ---- 3. Optics / Imaging (Main path) ----
    (8,  "3. Imaging Main",    "CMOS Image Sensor (Main)",           "OmniVision",       "OV13B10",            "CSP",                  "13MP, 1.12um BSI, MIPI 4-lane",             1, 4.00,  "OmniVision / WPG", 14, "T1", "Med"),
    (9,  "3. Imaging Main",    "Lens Module 6P (Main)",              "Sunny Optical",    "SY-Custom-6P-F18",   "Custom",               "6P, F/1.8, 105 deg DFOV, TTL 4.3mm",        1, 6.50,  "Sunny / OFilm",  16, "T1", "High"),
    (10, "3. Imaging Main",    "AA Active-Alignment assembly",       "Sunny Optical",    "SY-AA-Service",      "Service fee",          "Active Alignment, MTF >= 0.45@Nyquist/2",   1, 1.80,  "Sunny",          0, "T1", "Med"),
    (11, "3. Imaging Main",    "VCM (optional, fixed-focus skip)",   "N/A",              "-",                  "-",                    "Fixed focus, no VCM (cost saving)",          0, 0.00,  "-",              0, "T3", "Low"),
    (12, "3. Imaging Main",    "Sensor decoupling caps (set)",       "Murata",           "GRM155 series",      "0402",                 "10x 100nF + 4x 4.7uF X5R",                  1, 0.08,  "Murata / Digikey", 4, "T2", "Low"),

    # ---- 4. Optics / Imaging (AON sub-sensor for differentiation #1) ----
    (13, "4. Imaging AON",     "Sub CMOS (event-trigger)",           "GalaxyCore",       "GC02M2",             "CSP",                  "2MP, 1.75um, low-power MIPI/DVP",            1, 1.20,  "GalaxyCore",     8, "T1", "Med"),
    (14, "4. Imaging AON",     "Sub Lens Module",                    "Q-Tech",           "QT-Custom-5P1G",     "Custom",               "5P1G hybrid, F/2.0, 90 deg",                1, 2.50,  "Q-Tech",        12, "T2", "Med"),

    # ---- 5. Audio / Voice ----
    (15, "5. Audio",           "MEMS Mic (digital PDM) - left",      "Knowles",          "SPH0641LU4H-1",      "LGA 3.5x2.65",         "PDM, AOP=120dB, ALD anti-liquid",           1, 1.60,  "Knowles / Avnet", 10, "T1", "Low"),
    (16, "5. Audio",           "MEMS Mic (digital PDM) - right",     "Knowles",          "SPH0641LU4H-1",      "LGA 3.5x2.65",         "Beamforming pair",                          1, 1.60,  "Knowles / Avnet", 10, "T1", "Low"),
    (17, "5. Audio",           "Bone-conduction transducer",         "Sanwa",            "BCT-2818",           "18x10x2.8mm",          "1.5W, 200~7kHz, IPX7-friendly",             1, 1.80,  "Sanwa direct",  10, "T2", "Med"),
    (18, "5. Audio",           "Audio amp for BCT",                  "Maxim",            "MAX98357AETE+",      "TQFN-16",              "Class-D 3.2W, I2S in",                      1, 0.55,  "Mouser",         6, "T2", "Low"),
    (19, "5. Audio",           "Mic acoustic mesh (waterproof)",     "Gore",             "PMF200373",          "3.0mm dia",            "Acoustic vent IP67, 0.05dB insertion",      2, 0.45,  "Gore direct",    8, "T2", "Med"),

    # ---- 6. Sensors ----
    (20, "6. Sensors",         "6-axis IMU",                         "Bosch",            "BMI270",             "LGA-14 2.5x3",         "Accel+Gyro, < 4uA low-power, double-tap",    1, 1.10,  "Bosch / Mouser", 8, "T1", "Low"),
    (21, "6. Sensors",         "Ambient Light Sensor",               "Vishay",           "VEML7700-TT",        "OPLGA 6.8x2.35",       "I2C, lux range 0-120k",                     1, 0.55,  "Vishay / Digikey", 6, "T2", "Low"),
    (22, "6. Sensors",         "ToF / proximity",                    "STMicro",          "VL53L1X",            "OLGA-12 4.9x2.5",      "Up to 4m, 50Hz, occlusion detect",          1, 1.40,  "ST / Arrow",     8, "T2", "Low"),
    (23, "6. Sensors",         "Hall switch (mag accessory ID)",     "TI",               "DRV5032FBDBZR",      "SOT-23-3",             "Omnipolar, 1.3uA avg",                      1, 0.18,  "TI / Mouser",    6, "T2", "Low"),
    (24, "6. Sensors",         "Capacitive touch slider IC",         "Microchip",        "AT42QT1110-AU",      "TSSOP-20",             "11-key/slider, < 30uA",                     1, 0.65,  "Microchip / Avnet", 8, "T2", "Low"),

    # ---- 7. Power / PMIC ----
    (25, "7. Power",           "PMIC (charger+SIMO+LDO+gauge IF)",   "Maxim/ADI",        "MAX77654AEWX+",      "WLP-30 2.7x3",         "1x SIMO, 3x LDO, 500mA charger, < 3uA Iq",  1, 1.85,  "Maxim / WPI",    12, "T1", "Med"),
    (26, "7. Power",           "Fuel gauge",                         "Maxim/ADI",        "MAX17048G+T10",      "TDFN-8",               "ModelGauge m5, +/-2% SOC",                  1, 1.20,  "Maxim / Mouser", 10, "T2", "Low"),
    (27, "7. Power",           "Li-Po battery 3.7V 350mAh",          "ATL / EVE",        "ATL-353038-350",     "30x22x3.8mm pouch",    "JEITA, NTC, PCM, UN38.3",                    1, 2.40,  "ATL Tier-1",    10, "T1", "High"),
    (28, "7. Power",           "Battery NTC (10k B3380)",            "Murata",           "NCP15XH103J03RC",    "0402",                 "Internal to battery pack (incl above)",     0, 0.00,  "(in pack)",      0, "T3", "Low"),
    (29, "7. Power",           "TVS / OVP for VBUS",                 "Nexperia",         "PESD5V0X1BT",        "SOD-523",              "5V working, 6.5V clamp",                     2, 0.06,  "Nexperia / LCSC", 4, "T3", "Low"),
    (30, "7. Power",           "Boost for BCT/LED rail",             "TI",               "TPS61291DRVR",       "WSON-6",               "3.3-5V boost, 1.2uA Iq",                    1, 0.55,  "TI / Mouser",    8, "T2", "Low"),

    # ---- 8. Connectivity / IO ----
    (31, "8. IO",              "Pogo Pin connector (6-pin)",         "Mill-Max",         "0850-0-15-20-82-14-11-0", "6P 2.5mm pitch", "Au plated, 10k cycles",                       1, 2.40,  "Mill-Max / Digikey", 14, "T1", "Med"),
    (32, "8. IO",              "Pogo Pin alt (cost-down)",           "Xinyi (信音)",     "XY-PG-6P-Au076",     "6P 2.5mm pitch",       "Au 0.76um, qual'd 10k cycles",              0, 1.20,  "Xinyi direct",   10, "T2", "Med"),
    (33, "8. IO",              "Accessory auth IC (1-Wire)",         "Maxim/ADI",        "DS28E15P+T",         "TSOC-6",               "SHA-256 auth, anti-clone accessory",         1, 0.85,  "Maxim / Mouser", 10, "T2", "Med"),
    (34, "8. IO",              "Tact switch (shutter)",              "C&K",              "PTS526SMG15SMTR2LFS", "SMT 5x5",             "IP67 rated, 200gf, 500k cycles",            1, 0.32,  "C&K / Digikey",  6, "T2", "Low"),
    (35, "8. IO",              "Status LED (RGB)",                   "Everlight",        "19-217/RSGHBHC-A01", "PLCC-4 0606",          "Common-anode RGB, dome lens",               1, 0.12,  "Everlight",     6, "T3", "Low"),
    (36, "8. IO",              "Magnet (N52 disc)",                  "Generic NdFeB",    "N52-D6x1.5",         "D6x1.5 mm",            "Nickel-plated, axially magnetized x2",       2, 0.18,  "Ningbo Jingci", 6, "T3", "Low"),

    # ---- 9. Antenna ----
    (37, "9. RF",              "Wi-Fi/BLE chip antenna",             "Yageo",            "ANT-W63WS5-T",       "0603 chip",            "2.4/5GHz dual-band, 50ohm",                 1, 0.35,  "Yageo / LCSC",   6, "T2", "Med"),
    (38, "9. RF",              "Pi-match RLC kit",                   "Murata",           "GRM/LQG combo",      "0402",                 "0R/0R/NM placeholder, tuned per DVT",        1, 0.10,  "Murata",         4, "T3", "Low"),

    # ---- 10. Passives & Misc EE ----
    (39, "10. Passives",       "MLCC bulk pack (0402/0201)",         "Samsung EM",       "CL-series",          "0402/0201",            "approx 180 caps mixed",                     1, 1.10,  "Samsung EM",     6, "T2", "Low"),
    (40, "10. Passives",       "Resistors bulk (0402)",              "Yageo",            "RC0402-series",      "0402",                 "approx 120 resistors, 1%",                  1, 0.35,  "Yageo / LCSC",   4, "T3", "Low"),
    (41, "10. Passives",       "Inductors (DCDC)",                   "Murata",           "DFE201610E-2R2M",    "2.0x1.6x1.0",          "2.2uH, 1.5A, low-DCR",                      3, 0.18,  "Murata",         6, "T2", "Low"),
    (42, "10. Passives",       "Crystal 24/40 MHz (SoC)",            "TXC",              "7M-40.000MEEQ",      "3.2x2.5",              "+/-10ppm",                                   1, 0.22,  "TXC / Digikey",  6, "T2", "Low"),
    (43, "10. Passives",       "ESD diode array (USB/Pogo)",         "ON Semi",          "ESD7104MUTAG",       "uDFN-10",              "4-line 2.5pF, 8kV contact",                 1, 0.22,  "Onsemi / Avnet", 6, "T3", "Low"),

    # ---- 11. PCB / SMT / Mech ----
    (44, "11. PCB & SMT",      "Main PCB (Rigid-Flex 6L HDI)",       "Compeq / KCE",     "Custom-6L-HDI-RF",   "Rigid-Flex",           "6L, 2-stage HDI, ENIG, 1 flex layer",       1, 5.20,  "Compeq",         18, "T1", "High"),
    (45, "11. PCB & SMT",      "SMT assembly + test labor",          "OEM (TBD)",        "SMT-Service",        "Service",              "0201-capable line, AOI+X-ray",              1, 3.80,  "OEM (Hongqun / Daiwa)", 0, "T1", "Med"),
    (46, "11. PCB & SMT",      "EMI shielding can (sensor+SoC)",     "Laird",            "Custom-shield-2pc",  "Stamped + frame",      "0.2mm tin-plated, 2-piece",                 1, 0.35,  "Laird / local", 8, "T2", "Med"),

    # ---- 12. Mechanical / Enclosure ----
    (47, "12. Mech",           "Front housing PC+ABS",               "TBD molder",       "FH-001-PCABS",       "Custom",               "Bayer T65, ETA-1 wear",                     1, 0.85,  "TBD molder",    12, "T1", "Med"),
    (48, "12. Mech",           "Rear housing aluminum (anodized)",   "TBD CNC",          "RH-002-ALU6063",     "CNC + anodize",        "T6, ETA-2 finish, 1 antenna window",        1, 1.20,  "TBD CNC",       14, "T1", "Med"),
    (49, "12. Mech",           "O-ring silicone IP67",               "TBD seal",         "OR-0.6mm-Sil70",     "Custom profile",       "Shore A70, perimeter seal",                 1, 0.12,  "TBD seal",       6, "T3", "Low"),
    (50, "12. Mech",           "Acrylic lens window (AR coated)",    "TBD optics",       "LW-001-AR",          "Sapphire alt avail",   "AR coat both sides, R<0.4%@VIS",            1, 0.45,  "TBD optics",     8, "T2", "Med"),
    (51, "12. Mech",           "Clip mechanism (steel + spring)",    "TBD hw",           "CL-003-SS304",       "SS304",                "Sprung clip, 5N retention",                 1, 0.95,  "TBD hw",        10, "T2", "Med"),
    (52, "12. Mech",           "Foam gaskets / acoustic damp",       "3M",               "VHB-5915 + foam",    "Die-cut",              "VHB tape + acoustic foam",                  1, 0.18,  "3M / local",     6, "T3", "Low"),
    (53, "12. Mech",           "Screws (M1.4 SS torx)",              "Generic",          "M1.4x3-T5",          "SS304",                "Black-oxide, x6",                           6, 0.02,  "Local fastener", 4, "T3", "Low"),

    # ---- 13. Packaging ----
    (54, "13. Pack",           "Retail box (printed)",               "TBD print",        "BOX-001",            "Litho 4C+matte",       "FSC paper, magnetic flap",                  1, 0.65,  "TBD print",      8, "T3", "Low"),
    (55, "13. Pack",           "USB-C charge cable (1m)",            "Generic",          "CBL-USB-C-1m",       "TPE jacket",           "USB-C to Pogo magnetic, eMark optional",    1, 0.55,  "TBD",            6, "T3", "Low"),
    (56, "13. Pack",           "Quick-start manual",                 "TBD print",        "QSG-001",            "Folded paper",         "5 languages",                               1, 0.08,  "TBD print",      6, "T3", "Low"),
    (57, "13. Pack",           "Inner tray (molded pulp)",           "TBD pulp",         "TRAY-001",           "Pulp",                 "Molded pulp, FSC",                          1, 0.18,  "TBD pulp",       6, "T3", "Low"),

    # ---- 14. Certifications & SW (amortized over 10K) ----
    (58, "14. Cert / Amort",   "FCC ID + CE-RED amortized",          "Cert house",       "Cert-bundle",        "Service",              "FCC/CE/SRRC, amortized /10K units",          1, 0.35,  "Bureau Veritas / SGS", 16, "T1", "High"),
    (59, "14. Cert / Amort",   "UN38.3 + IEC62133 (battery)",        "Cert house",       "BattCert-bundle",    "Service",              "Battery transport+safety, amortized /10K", 1, 0.18,  "TUV / SGS",     12, "T1", "High"),
    (60, "14. Cert / Amort",   "SW / firmware NRE amortized",        "Internal",         "FW-NRE-amort",       "NRE",                  "FW + cloud NRE, amortized /50K planned",    1, 0.80,  "Internal",       0, "T1", "Med"),
]

# ---------------------------------------------------------------------------
# DATA: Alternates / Risk Mitigation matrix
# ---------------------------------------------------------------------------
ALT_HEADERS = [
    "BOM #", "Item / Function", "Primary MPN", "Primary Mfr", "Primary $",
    "Alt-1 MPN", "Alt-1 Mfr", "Alt-1 $", "Alt-1 Pin-compat?",
    "Alt-2 MPN", "Alt-2 Mfr", "Alt-2 $", "Alt-2 Pin-compat?",
    "Switching Effort", "Notes"
]

ALT_ROWS = [
    (1,  "Application SoC",           "SSC338Q",            "SigmaStar",   4.50,
        "CV28M",              "Ambarella",   12.00, "No (PCB respin)",
        "V853",               "Allwinner",   3.80,  "No (SDK rework)",
        "High",
        "EVT 阶段 A/B 双版主板并行验证；正式选型在 EVT 末 freeze。"),
    (2,  "Wi-Fi/BLE Combo",           "RTL8821CS",          "Realtek",     1.80,
        "AIC8800DC",          "AICSemi",      1.30, "No",
        "ESP32-C6 (mod)",     "Espressif",    2.10, "No",
        "Med",
        "若 SoC 改成 CV28M，自带 Wi-Fi 则此项可去除。"),
    (3,  "DDR2 SDRAM",                "M14D5121632A-2.5BG", "ESMT",        1.20,
        "EM68C32CWQG-25H",    "Etron",        1.30, "Yes",
        "MT47H32M16HR-25",    "Micron",       1.80, "Yes",
        "Low",
        "JEDEC 标准 BGA-90，3 家可二选一过料。"),
    (5,  "eMMC 8GB",                  "THGBMNG5D1LBAIL",    "Kioxia",      3.40,
        "EM68C32-8GB",        "Samsung",      3.80, "Yes",
        "MTFC8GAKAJCN-4M",    "Micron",       3.95, "Yes",
        "Low",
        "BGA-153 标准封装。"),
    (8,  "CMOS Sensor (Main)",        "OV13B10",            "OmniVision",  4.00,
        "IMX681",             "Sony",         6.00, "No",
        "SC500AI",            "SmartSens",    3.50, "No",
        "High",
        "切换需重新 Tuning ISP；建议主备各跑一支 EVT 模组。"),
    (13, "Sub Sensor (AON)",          "GC02M2",             "GalaxyCore",  1.20,
        "OV02B1B",            "OmniVision",   1.40, "No",
        "SC202CS",            "SmartSens",    1.10, "No",
        "Med",
        "AON 路像素要求低，国产替代非常充足。"),
    (15, "MEMS Mic L/R",              "SPH0641LU4H-1",      "Knowles",     1.60,
        "ICS-41351",          "TDK InvenSense", 1.30, "Pad-compat (review)",
        "T4040A0",            "Goertek",      0.95, "Yes (most pads)",
        "Med",
        "国产 Goertek 性价比突出，AOP 略低（118 vs 120 dB）。"),
    (20, "6-axis IMU",                "BMI270",             "Bosch",       1.10,
        "ICM-42688-P",        "TDK InvenSense", 1.40, "No",
        "QMI8658C",           "QST",          0.55, "No",
        "Med",
        "DMP 算法差异；如用 BMI270，要买 BSX 库授权。"),
    (22, "ToF",                       "VL53L1X",            "STMicro",     1.40,
        "TMF8801",            "ams OSRAM",    1.50, "No",
        "VL53L0X",            "STMicro",      0.90, "Pin-compat down-spec",
        "Low",
        "VL53L0X 距离 2m 起 → 仅 PoC 阶段可用。"),
    (25, "PMIC",                      "MAX77654AEWX+",      "Maxim/ADI",   1.85,
        "BQ25180+TPS62840+BQ27220 set", "TI", 2.10, "No",
        "RT9426A+RT9470 set", "Richtek",      1.65, "No",
        "High",
        "PMIC 一旦 freeze 极难替换。EVT 必须只押 1 颗。"),
    (27, "Battery 350mAh",            "ATL-353038-350",     "ATL",         2.40,
        "EVE-353038-350",     "EVE Energy",   2.10, "Yes (cell only)",
        "Lishen-353038-350",  "Lishen",       2.30, "Yes (cell only)",
        "Med",
        "PCM/连接器需各厂分别打样验证；UN38.3 须按选定厂送测。"),
    (31, "Pogo Connector",            "Mill-Max 850 series","Mill-Max",    2.40,
        "Xinyi XY-PG-6P-Au076","Xinyi (信音)", 1.20, "Footprint compat (要 review)",
        "Lianzhan custom",    "Lianzhan (连展)", 1.40, "Custom",
        "Med",
        "国产 1.2 vs 进口 2.4，DVT 起切换可省 BOM ¥8.4。"),
    (37, "Antenna chip",              "ANT-W63WS5-T",       "Yageo",       0.35,
        "Pulse W3011",        "Pulse",        0.55, "No",
        "FPC custom antenna", "Local",        0.45, "No",
        "Med",
        "FPC 天线性能更好但占空间。"),
    (44, "PCB Rigid-Flex",            "Compeq 6L HDI-RF",   "Compeq",      5.20,
        "KCE 6L HDI-RF",      "KCE",          4.80, "Yes (gerber同源)",
        "Aoshikang 6L HDI-RF","Aoshikang",    4.30, "Yes",
        "Low",
        "二阶 HDI + 1 层 Flex；3 家询价对比。"),
]

# ---------------------------------------------------------------------------
# DATA: Cost Curve (volume-based unit cost evolution)
# ---------------------------------------------------------------------------
COST_CURVE_HEADERS = [
    "Volume Tier", "Build Phase", "Total BOM USD",
    "SoC + WiFi", "Imaging (Main+AON)", "Battery+PMIC", "Mech+Pack", "Other",
    "FOB target USD", "MSRP target USD"
]
COST_CURVE_ROWS = [
    ("EVT (50pcs)",   "EVT", 78.00, 9.50, 17.50, 7.20, 12.00, 31.80, 0.00, 0.00),
    ("DVT (500pcs)",  "DVT", 52.00, 7.80, 14.20, 6.40, 8.50,  15.10, 0.00, 0.00),
    ("PVT (2K)",      "PVT", 38.00, 6.80, 12.10, 5.80, 6.50,  6.80,  0.00, 0.00),
    ("MP 10K",        "MP",  31.00, 6.30, 10.00, 5.45, 5.40,  3.85,  35.00, 149.00),
    ("MP 50K",        "MP",  27.50, 5.80, 9.20,  5.10, 4.80,  2.60,  31.50, 149.00),
    ("MP 200K",       "MP",  24.00, 5.20, 8.30,  4.80, 4.30,  1.40,  27.50, 149.00),
]

# ---------------------------------------------------------------------------
# DATA: NRE / Tooling capex (one-time)
# ---------------------------------------------------------------------------
NRE_HEADERS = ["Phase", "Item", "Vendor Type", "USD Estimate", "Notes"]
NRE_ROWS = [
    ("ID",  "Industrial design fees",       "ID house",       6000,  "2 rounds + final CMF"),
    ("ID",  "Hand model x10 (CNC+paint)",   "Model shop",     2500,  "for usability test"),
    ("MD",  "Mech design + DFM review",     "MD house",       8000,  "Includes drop sim"),
    ("EE",  "Schematic + Layout",           "EE house",       12000, "6L HDI Rigid-Flex"),
    ("EE",  "EVT PCB fab + SMT (50pcs)",    "Compeq + OEM",   18000, "Two SoC variants"),
    ("MD",  "EVT CNC housings (50sets)",    "CNC shop",       4500,  "Aluminum + PC+ABS"),
    ("DVT", "Soft tool (aluminum mold)",    "Mold maker",     45000, "Front+rear, ~5K shots"),
    ("DVT", "DVT samples 200pcs",           "OEM",            22000, "Small-line build"),
    ("DVT", "Reliability lab tests",        "SGS / TUV",      9000,  "Drop, HALT, IPX7, etc"),
    ("DVT", "Pre-compliance (FCC/CE/SRRC)", "Cert lab",       7500,  "Pre-cert before PVT"),
    ("PVT", "Steel mold (P20) production",  "Mold maker",     110000,"Front+rear, ~300K shots"),
    ("PVT", "FCT + ICT fixtures",           "Fixture house",  50000, "FCT+ICT for production"),
    ("PVT", "PVT build 2000pcs",            "OEM",            85000, "Production-line build"),
    ("MP",  "Cert formal (FCC/CE/SRRC)",    "Cert lab",       18000, "Formal certification"),
    ("MP",  "UN38.3 + IEC62133 + KC",       "Cert lab",       9500,  "Battery transport+safety"),
    ("MP",  "Tooling H13 (high-volume add-on)","Mold maker",  60000, "Optional after 200K"),
    ("MP",  "Inventory + safety stock",     "OEM",            45000, "Working capital reserve"),
]

# ---------------------------------------------------------------------------
# DATA: Risk register
# ---------------------------------------------------------------------------
RISK_HEADERS = ["#", "Risk", "Linked BOM", "Likelihood (1-5)", "Impact (1-5)", "Score", "Mitigation"]
RISK_ROWS = [
    (1, "SoC 缺货 / 价格波动",             "1, 2",      4, 5, "=D2*E2",
        "EVT 阶段 A/B 双方案并行；战略备货 6 个月 + 与代理商签 LTA"),
    (2, "Lens 模组良率 < 80%",              "9, 10",     3, 5, "=D3*E3",
        "AA 工艺锁定，二供 Q-Tech 同步打样；MTF 抽检 100%"),
    (3, "Battery 厂质量不一致",            "27",        3, 5, "=D4*E4",
        "ATL 主供 + EVE 备供；UN38.3 各厂分别送测"),
    (4, "PCB Rigid-Flex 良率",              "44",        3, 4, "=D5*E5",
        "EVT 起 3 家询价；折弯半径 / 铜厚提前与 PCB 厂联评"),
    (5, "PMIC freeze 后无法替换",           "25",        2, 5, "=D6*E6",
        "EVT 阶段务必单押；Layout 预留二供 Footprint 合并方案"),
    (6, "FCC/CE 一次过不了",                "37, 38",    3, 4, "=D7*E7",
        "EVT 末做 Pre-Compliance；π 型匹配电路保留"),
    (7, "防水 IP67 量产漏失",               "47-50, 19", 4, 4, "=D8*E8",
        "100% 气密测试 + 抽检 IPX7 浸水"),
    (8, "Pogo Pin 寿命 < 10k 次",            "31, 32",    2, 4, "=D9*E9",
        "进口 Mill-Max 主供，国产二供寿命单独验证"),
    (9, "ID 与 MD 公差冲突",                "47, 48",    3, 3, "=D10*E10",
        "MD 工程师从 ID 第 2 周介入做可行性评审"),
    (10,"模具改模 / 钢模废模",              "47, 48",    3, 5, "=D11*E11",
        "DVT 走软模，PVT 才上钢模；3D 打印验证 2 轮以上"),
]


# ============================================================================
# XLSX builder (no third-party deps)
# ============================================================================

def col_letter(idx):
    """1 -> A, 27 -> AA"""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def cell_xml(col, row, value, style_id=0, formula=None):
    ref = f"{col_letter(col)}{row}"
    if formula is not None:
        return f'<c r="{ref}" s="{style_id}"><f>{escape(formula)}</f></c>'
    if value is None or value == "":
        return f'<c r="{ref}" s="{style_id}"/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" s="{style_id}" t="n"><v>{value}</v></c>'
    # inline string
    text = escape(str(value))
    return f'<c r="{ref}" s="{style_id}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def build_sheet_xml(rows, col_widths=None, freeze_row=1, freeze_col=0, merges=None):
    """
    rows: list of lists; each inner list is [(value, style_id, formula_or_None), ...]
    """
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')

    # Frozen pane
    if freeze_row or freeze_col:
        x = freeze_col
        y = freeze_row
        top_left = f"{col_letter(x+1)}{y+1}"
        state = 'frozen'
        parts.append('<sheetViews><sheetView tabSelected="1" workbookViewId="0">')
        parts.append(f'<pane xSplit="{x}" ySplit="{y}" topLeftCell="{top_left}" activePane="bottomRight" state="{state}"/>')
        parts.append('</sheetView></sheetViews>')

    # Column widths
    if col_widths:
        parts.append('<cols>')
        for i, w in enumerate(col_widths, start=1):
            parts.append(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
        parts.append('</cols>')

    parts.append('<sheetData>')
    for r_idx, row in enumerate(rows, start=1):
        # row height for header (taller)
        ht = ' ht="32" customHeight="1"' if r_idx == 1 else ''
        parts.append(f'<row r="{r_idx}"{ht}>')
        for c_idx, cell in enumerate(row, start=1):
            value, style_id, formula = cell
            parts.append(cell_xml(c_idx, r_idx, value, style_id, formula))
        parts.append('</row>')
    parts.append('</sheetData>')

    if merges:
        parts.append(f'<mergeCells count="{len(merges)}">')
        for m in merges:
            parts.append(f'<mergeCell ref="{m}"/>')
        parts.append('</mergeCells>')

    parts.append('</worksheet>')
    return ''.join(parts)


# ---------------------------------------------------------------------------
# Stylesheet
# ---------------------------------------------------------------------------
# Style IDs we will use:
# 0 - default
# 1 - title (bold 14, fill blue, white, center, wrap)
# 2 - header (bold, fill dark, white, center, wrap, border)
# 3 - text body, border, wrap
# 4 - number USD 2dp, border
# 5 - integer, border, center
# 6 - tier T1 (red fill)
# 7 - tier T2 (yellow fill)
# 8 - tier T3 (green fill)
# 9 - risk High (red)
# 10 - risk Med (yellow)
# 11 - risk Low (green)
# 12 - sub-section header (bold, light fill)
# 13 - total row (bold, top border, USD)
# 14 - text body bold (label cells in cover/summary)
STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
    <numFmt numFmtId="165" formatCode="0"/>
  </numFmts>
  <fonts count="6">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><sz val="10"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF7F0000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="10">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border/>
    <border>
      <left style="thin"><color rgb="FFAAAAAA"/></left>
      <right style="thin"><color rgb="FFAAAAAA"/></right>
      <top style="thin"><color rgb="FFAAAAAA"/></top>
      <bottom style="thin"><color rgb="FFAAAAAA"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="15">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="165" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="4" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'''


def style_for_tier(t):
    return {"T1": 6, "T2": 7, "T3": 8}.get(t, 0)


def style_for_risk(r):
    return {"High": 9, "Med": 10, "Low": 11}.get(r, 0)


# ---------------------------------------------------------------------------
# Build each sheet
# ---------------------------------------------------------------------------

def build_cover():
    rows = []
    # Title row
    rows.append([("Looki-Class Wearable AI Capture Device  —  Hardware BOM v1.0", 1, None)] + [("", 1, None)] * 5)
    rows.append([("", 0, None)] * 6)
    info = [
        ("Project Code", "LOOKI-X1"),
        ("Document Owner", "Hardware Architect"),
        ("Date", str(date.today())),
        ("Build Phase", "EVT-ready, costed @ MP 10K MOQ"),
        ("Currency", "USD (10K MOQ landed cost)"),
        ("Target FOB", "$35.00 / unit @ 10K"),
        ("Target MSRP", "$149 / unit"),
        ("Target Weight", "<= 22 g"),
        ("Target IP Rating", "IP67 (diff. SKU) / IP54 (base)"),
        ("Battery", "Li-Po 350 mAh, UN38.3"),
    ]
    for k, v in info:
        rows.append([(k, 14, None), (v, 3, None)] + [("", 0, None)] * 4)

    rows.append([("", 0, None)] * 6)
    rows.append([("Sheet Index", 12, None)] + [("", 0, None)] * 5)
    sheets = [
        ("01_BOM_Main",        "完整主 BOM（60 项），按子系统分组、标 Tier/Risk"),
        ("02_Alternates",      "关键料替代矩阵（Alt-1 / Alt-2 + 切换难度）"),
        ("03_Cost_Curve",      "EVT→DVT→PVT→MP 量产规模成本下降曲线"),
        ("04_NRE_CapEx",       "一次性投入：模具/PCB/治具/认证 (USD)"),
        ("05_Risk_Register",   "TOP-10 硬件量产风险及 Mitigation"),
        ("06_Notes_AvoidPits", "BOM 使用约定 + 工程师必读的避坑要点"),
    ]
    for s, d in sheets:
        rows.append([(s, 14, None), (d, 3, None)] + [("", 0, None)] * 4)

    merges = ["A1:F1"]
    return build_sheet_xml(rows, col_widths=[24, 60, 12, 12, 12, 12], freeze_row=0, freeze_col=0, merges=merges)


def build_main_bom():
    rows = []
    # Header
    rows.append([(h, 2, None) for h in BOM_HEADERS])

    cur_section = None
    for r in BOM_ROWS:
        # Section divider
        section = r[1]
        if section != cur_section:
            rows.append([(section, 12, None)] + [("", 12, None)] * (len(BOM_HEADERS) - 1))
            cur_section = section

        # row index in sheet (1-based) of this data row will be len(rows)+1
        sheet_row = len(rows) + 1
        item_no, cat, func, mfr, mpn, pkg, spec, qty, unit, dist, lt, tier, risk = r

        cells = [
            (item_no, 5, None),
            (cat, 3, None),
            (func, 3, None),
            (mfr, 3, None),
            (mpn, 3, None),
            (pkg, 3, None),
            (spec, 3, None),
            (qty, 5, None),
            (unit, 4, None),
            # Ext cost = qty * unit, formula referencing this same row's cols H * I
            (None, 4, f"H{sheet_row}*I{sheet_row}"),
            (dist, 3, None),
            (lt, 5, None),
            (tier, style_for_tier(tier), None),
            (risk, style_for_risk(risk), None),
        ]
        rows.append(cells)

    # Total row
    total_row = len(rows) + 1
    rows.append([
        ("", 13, None),
        ("", 13, None),
        ("TOTAL BOM (USD per unit @ 10K MOQ)", 13, None),
        ("", 13, None), ("", 13, None), ("", 13, None), ("", 13, None),
        ("", 13, None),
        ("Subtotal:", 13, None),
        (None, 13, f"SUM(J2:J{total_row-1})"),
        ("", 13, None), ("", 13, None), ("", 13, None), ("", 13, None),
    ])

    col_widths = [5, 18, 32, 16, 26, 18, 42, 6, 12, 12, 22, 8, 7, 9]
    return build_sheet_xml(rows, col_widths=col_widths, freeze_row=1, freeze_col=2)


def build_alternates():
    rows = []
    rows.append([(h, 2, None) for h in ALT_HEADERS])
    for r in ALT_ROWS:
        rows.append([
            (r[0], 5, None),
            (r[1], 3, None),
            (r[2], 3, None), (r[3], 3, None), (r[4], 4, None),
            (r[5], 3, None), (r[6], 3, None), (r[7], 4, None), (r[8], 3, None),
            (r[9], 3, None), (r[10], 3, None), (r[11], 4, None), (r[12], 3, None),
            (r[13], style_for_risk({"High": "High", "Med": "Med", "Low": "Low"}.get(r[13], "Low")), None),
            (r[14], 3, None),
        ])
    col_widths = [6, 22, 22, 14, 11, 22, 14, 11, 18, 22, 14, 11, 18, 14, 50]
    return build_sheet_xml(rows, col_widths=col_widths, freeze_row=1, freeze_col=2)


def build_cost_curve():
    rows = []
    rows.append([(h, 2, None) for h in COST_CURVE_HEADERS])
    for r in COST_CURVE_ROWS:
        # Total = sum of components
        sheet_row = len(rows) + 1
        total_formula = f"D{sheet_row}+E{sheet_row}+F{sheet_row}+G{sheet_row}+H{sheet_row}"
        rows.append([
            (r[0], 3, None),
            (r[1], 5, None),
            (None, 4, total_formula),                # C: Total (formula)
            (r[3], 4, None),
            (r[4], 4, None),
            (r[5], 4, None),
            (r[6], 4, None),
            (r[7], 4, None),
            (r[8] if r[8] else "", 4, None),
            (r[9] if r[9] else "", 4, None),
        ])
    col_widths = [16, 10, 14, 12, 18, 14, 14, 10, 14, 14]
    return build_sheet_xml(rows, col_widths=col_widths, freeze_row=1, freeze_col=1)


def build_nre():
    rows = []
    rows.append([(h, 2, None) for h in NRE_HEADERS])
    total_row_count = len(NRE_ROWS)
    for r in NRE_ROWS:
        rows.append([
            (r[0], 5, None),
            (r[1], 3, None),
            (r[2], 3, None),
            (r[3], 4, None),
            (r[4], 3, None),
        ])
    last_data_row = len(rows)
    rows.append([
        ("", 13, None),
        ("TOTAL one-time CapEx (USD)", 13, None),
        ("", 13, None),
        (None, 13, f"SUM(D2:D{last_data_row})"),
        ("", 13, None),
    ])
    col_widths = [8, 38, 18, 16, 50]
    return build_sheet_xml(rows, col_widths=col_widths, freeze_row=1, freeze_col=0)


def build_risk():
    rows = []
    rows.append([(h, 2, None) for h in RISK_HEADERS])
    for r in RISK_ROWS:
        sheet_row = len(rows) + 1
        rows.append([
            (r[0], 5, None),
            (r[1], 3, None),
            (r[2], 3, None),
            (r[3], 5, None),
            (r[4], 5, None),
            (None, 4, f"D{sheet_row}*E{sheet_row}"),
            (r[6], 3, None),
        ])
    col_widths = [5, 36, 14, 14, 14, 10, 60]
    return build_sheet_xml(rows, col_widths=col_widths, freeze_row=1, freeze_col=1)


def build_notes():
    rows = []
    rows.append([("Notes / Avoid-Pits Checklist for using this BOM", 1, None)] + [("", 1, None)] * 1)
    rows.append([("", 0, None), ("", 0, None)])

    notes = [
        ("Convention", "所有单价均为 10K MOQ 含税到岸价 (USD)。EVT/DVT 阶段实际单价请上浮 30~80%。"),
        ("Convention", "Tier T1 = 战略料 (单点失效会卡死项目)；T2 = 标准件 (有 2+ 二供)；T3 = 通用件。"),
        ("Convention", "Risk = 量产端缺料、良率、价格波动综合评估。High 项必须周报跟踪。"),
        ("Convention", "Qty=0 的行表示该料目前不选用 (备选/可选)，但 PCB Layout 须保留 Footprint。"),
        ("Avoid-Pit",  "SoC 与 Sensor 的 Tuning 包是绑定的；切换主控会丢失 6-10 周 ISP 调试。EVT 双方案要等 Tuning 各自跑通再 PK。"),
        ("Avoid-Pit",  "PMIC 一旦 Layout freeze 极难替换，单押 1 颗，但 EVT 阶段验静态电流必须用 µA 级源表实测，不信 datasheet typ 值。"),
        ("Avoid-Pit",  "Lens AA 工艺单价不可省。普通 MTF 模组在 1.0µm sensor 上良率 < 70%。"),
        ("Avoid-Pit",  "Battery 任何换厂都要重过 UN38.3 (~$3000, 6 周)，不得在 PVT 后切换。"),
        ("Avoid-Pit",  "PCB Rigid-Flex 折弯半径 < 6×板厚会拉低良率，与 PCB 厂在 EE Layout 第一周就联评。"),
        ("Avoid-Pit",  "天线净空区 (Antenna Keep-Out) 在 ID 第一稿就要 hold 住，金属外壳必须开塑料窗。"),
        ("Avoid-Pit",  "EVT 末必跑 Pre-Compliance (FCC/CE/SRRC)，¥1.5-3 万一次比 DVT 后翻车便宜 10 倍。"),
        ("Avoid-Pit",  "DVT 走软模 (¥45K)，PVT 才上钢模 (¥110K)。先开钢模再改结构 = 报废一套模具。"),
        ("Avoid-Pit",  "Pogo Pin 寿命要按 10K 次实测，国产二供 (Xinyi/Lianzhan) 必须独立验证。"),
        ("Avoid-Pit",  "IP67 量产 100% 做气密 (差压法 < 0.5 kPa/min)；IPX7 浸水按 AQL 0.65 抽检。"),
        ("Avoid-Pit",  "色差 (CMF ΔE) 必须明文写进 PO，<= 1.5；否则不同模穴出色不一致量产时退货困难。"),
    ]
    for k, v in notes:
        rows.append([(k, 14, None), (v, 3, None)])

    col_widths = [16, 100]
    merges = ["A1:B1"]
    return build_sheet_xml(rows, col_widths=col_widths, freeze_row=0, freeze_col=0, merges=merges)


# ---------------------------------------------------------------------------
# Build workbook (relations, content_types, workbook.xml)
# ---------------------------------------------------------------------------

SHEETS = [
    ("00_Cover",            build_cover),
    ("01_BOM_Main",         build_main_bom),
    ("02_Alternates",       build_alternates),
    ("03_Cost_Curve",       build_cost_curve),
    ("04_NRE_CapEx",        build_nre),
    ("05_Risk_Register",    build_risk),
    ("06_Notes_AvoidPits",  build_notes),
]


def build_content_types(num_sheets):
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    parts.append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    parts.append('<Default Extension="xml" ContentType="application/xml"/>')
    parts.append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    parts.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    for i in range(1, num_sheets + 1):
        parts.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    parts.append('</Types>')
    return ''.join(parts)


def build_root_rels():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')


def build_workbook_xml(sheet_names):
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    parts.append('<sheets>')
    for i, name in enumerate(sheet_names, start=1):
        parts.append(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>')
    parts.append('</sheets>')
    parts.append('</workbook>')
    return ''.join(parts)


def build_workbook_rels(num_sheets):
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for i in range(1, num_sheets + 1):
        parts.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
    parts.append(f'<Relationship Id="rId{num_sheets+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    parts.append('</Relationships>')
    return ''.join(parts)


def main():
    sheet_names = [s[0] for s in SHEETS]
    sheet_xmls = [s[1]() for s in SHEETS]

    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), OUTFILE)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', build_content_types(len(SHEETS)))
        z.writestr('_rels/.rels', build_root_rels())
        z.writestr('xl/workbook.xml', build_workbook_xml(sheet_names))
        z.writestr('xl/_rels/workbook.xml.rels', build_workbook_rels(len(SHEETS)))
        z.writestr('xl/styles.xml', STYLES_XML)
        for i, xml in enumerate(sheet_xmls, start=1):
            z.writestr(f'xl/worksheets/sheet{i}.xml', xml)

    size_kb = os.path.getsize(out_path) / 1024
    print(f"Wrote: {out_path}  ({size_kb:.1f} KB, {len(SHEETS)} sheets)")


if __name__ == "__main__":
    main()
