# -*- coding: utf-8 -*-
"""
Insurance Calculator Module
保险计算器模块 — 主险计算 + 附加险计算
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QPushButton, QComboBox, QDoubleSpinBox, QSpinBox,
    QScrollArea, QFrame, QSlider, QFileDialog, QLineEdit,
    QListWidget, QListWidgetItem, QDateEdit, QTextEdit,
    QGraphicsDropShadowEffect, QMessageBox, QGroupBox,
    QSizePolicy, QAbstractItemView
)
from PyQt5.QtCore import Qt, pyqtSignal, QDate
from PyQt5.QtGui import QFont, QColor
import json
import os
import math
import re
from datetime import date

# 从主脚本导入设计系统
try:
    from Clause_Comparison_Assistant import AnthropicColors, AnthropicFonts, GlassCard
except ImportError:
    # 备用颜色定义
    class AnthropicColors:
        BG_PRIMARY = "#faf9f5"
        BG_CARD = "#f0eee6"
        BG_MINT = "#bcd1ca"
        BG_LAVENDER = "#cbcadb"
        BG_DARK = "#141413"
        ACCENT = "#d97757"
        ACCENT_DARK = "#c6613f"
        ACCENT_LIGHT = "#e8956f"
        TEXT_PRIMARY = "#141413"
        TEXT_SECONDARY = "#b0aea5"
        TEXT_TERTIARY = "#8a8880"
        TEXT_LIGHT = "#faf9f5"
        BORDER = "#d8d6ce"
        BORDER_DARK = "#c0beb6"

    class AnthropicFonts:
        TITLE_LARGE = ("Anthropic Sans", 28)
        TITLE = ("Anthropic Sans", 22)
        TITLE_SMALL = ("Anthropic Sans", 16)
        BODY = ("Anthropic Serif", 14)
        BODY_SMALL = ("Anthropic Serif", 12)

    class GlassCard(QFrame):
        def __init__(self, parent=None, variant="default"):
            super().__init__(parent)
            bg = {"mint": "#bcd1ca", "lavender": "#cbcadb"}.get(variant, "#f0eee6")
            self.setStyleSheet(f"QFrame {{ background: {bg}; border: 1px solid #d8d6ce; border-radius: 12px; }}")


# =============================================
# 数据常量
# =============================================

MC_PRODUCTS = {
    "employerLiability": {
        "productName": "雇主责任险",
        "productType": "liability",
        "amountUnit": "万元",
        "amountLabel": "每人限额",
        "premiumCap": 0.70,
        "versions": {
            "original": {
                "label": "雇主责任险费率",
                "baseRates": {
                    "fixed": {"class1": 0.0011, "class2": 0.0017, "class3": 0.0029},
                    "salary": {"class1": 0.0033, "class2": 0.0051, "class3": 0.0085}
                },
                "coefficients": [
                    {
                        "id": "perPersonLimit", "name": "每人赔偿限额调整系数", "applicableTo": ["fixed"],
                        "note": "未列明限额可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤10万元", "min": 1.2, "max": 1.3, "type": "range"},
                            {"parameter": "30万元", "value": 1.1, "type": "fixed"},
                            {"parameter": "50万元", "value": 1.0, "type": "fixed"},
                            {"parameter": "80万元", "value": 0.9, "type": "fixed"},
                            {"parameter": "≥100万元", "min": 0.8, "max": 0.85, "type": "range"}
                        ]
                    },
                    {
                        "id": "employeeCount", "name": "承保人数调整系数", "applicableTo": ["fixed"],
                        "note": "未列明人数可按线性插值法计算",
                        "rows": [
                            {"parameter": "＜100人", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "[100, 500)人", "min": 0.9, "max": 1.0, "minExclusive": True, "type": "range"},
                            {"parameter": "[500, 1000)人", "min": 0.8, "max": 0.9, "minExclusive": True, "type": "range"},
                            {"parameter": "≥1000人", "min": 0.7, "max": 0.8, "type": "range"}
                        ]
                    },
                    {
                        "id": "deathDisabilityMonths", "name": "死亡/伤残每人赔偿限额调整系数", "applicableTo": ["salary"],
                        "rows": [
                            {"parameter": "36/48个月", "value": 1.0, "type": "fixed"},
                            {"parameter": "48/60个月", "value": 1.25, "type": "fixed"},
                            {"parameter": "60/72个月", "value": 1.4, "type": "fixed"},
                            {"parameter": "72/84个月", "value": 1.5, "type": "fixed"}
                        ]
                    },
                    {
                        "id": "medicalLimit", "name": "医疗费用每人赔偿限额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "医疗费用每人赔偿限额÷每人赔偿限额；未列明比例可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤5%", "min": 0.9, "max": 0.95, "type": "range"},
                            {"parameter": "10%", "value": 1.0, "type": "fixed"},
                            {"parameter": "15%", "value": 1.05, "type": "fixed"},
                            {"parameter": "20%", "value": 1.1, "type": "fixed"},
                            {"parameter": "≥25%", "min": 1.15, "max": 1.3, "type": "range"}
                        ]
                    },
                    {
                        "id": "lostWorkLimit", "name": "误工费用每人赔偿限额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "误工费用每人赔偿限额÷每人赔偿限额；未列明比例可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤5%", "min": 0.9, "max": 0.95, "type": "range"},
                            {"parameter": "10%", "value": 1.0, "type": "fixed"},
                            {"parameter": "15%", "value": 1.05, "type": "fixed"},
                            {"parameter": "≥20%", "min": 1.1, "max": 1.2, "type": "range"}
                        ]
                    },
                    {
                        "id": "perAccidentRatio", "name": "每次事故赔偿限额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "每次事故赔偿限额÷每人赔偿限额；未列明比例可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤3倍", "min": 0.9, "max": 0.95, "type": "range"},
                            {"parameter": "5倍", "value": 1.0, "type": "fixed"},
                            {"parameter": "10倍", "value": 1.05, "type": "fixed"},
                            {"parameter": "≥15倍", "min": 1.1, "max": 1.2, "type": "range"}
                        ]
                    },
                    {
                        "id": "cumulativeRatio", "name": "累计赔偿限额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "累计赔偿限额÷每次事故赔偿限额；未列明比例可按线性插值法计算",
                        "rows": [
                            {"parameter": "1倍", "value": 0.95, "type": "fixed"},
                            {"parameter": "2倍", "value": 1.0, "type": "fixed"},
                            {"parameter": "3倍", "value": 1.05, "type": "fixed"},
                            {"parameter": "≥4倍", "min": 1.1, "max": 1.2, "type": "range"}
                        ]
                    },
                    {
                        "id": "deductible", "name": "免赔额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "每次事故医疗费用每人免赔额；未列明免赔额可按线性插值法计算",
                        "rows": [
                            {"parameter": "0元", "value": 1.0, "type": "fixed"},
                            {"parameter": "500元", "value": 0.97, "type": "fixed"},
                            {"parameter": "1000元", "value": 0.95, "type": "fixed"},
                            {"parameter": "≥2000元", "min": 0.8, "max": 0.9, "type": "range"}
                        ]
                    },
                    {
                        "id": "employeeCategory", "name": "雇员类别调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "管理人员", "min": 0.7, "max": 0.8, "type": "range"},
                            {"parameter": "后勤人员", "min": 0.9, "max": 1.0, "type": "range"},
                            {"parameter": "一线操作人员", "min": 1.0, "max": 2.0, "type": "range"}
                        ]
                    },
                    {
                        "id": "workInjuryInsurance", "name": "工伤保险情况调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "已投保工伤保险", "value": 1.0, "type": "fixed"},
                            {"parameter": "未投保工伤保险", "value": 1.2, "type": "fixed"}
                        ]
                    },
                    {
                        "id": "managementLevel", "name": "管理水平调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "制度完善，无明显缺陷", "min": 0.7, "max": 1.0, "type": "range"},
                            {"parameter": "较完善，存在个别缺陷", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "不完善或存在较多缺陷", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "lossRatio", "name": "赔付率调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "[0, 20%]", "min": 0.5, "max": 0.6, "type": "range"},
                            {"parameter": "(20%, 45%]", "min": 0.6, "max": 0.8, "minExclusive": True, "type": "range"},
                            {"parameter": "(45%, 70%]", "min": 0.8, "max": 1.0, "minExclusive": True, "type": "range"},
                            {"parameter": "(70%, 95%]", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "＞95%", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "hazardInspection", "name": "企业隐患排查整改调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "无隐患", "min": 0.7, "max": 1.0, "type": "range"},
                            {"parameter": "整改完成", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "存在重大隐患且未整改", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "historicalAccident", "name": "历史事故与损失情况调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "极少", "min": 0.5, "max": 0.7, "type": "range"},
                            {"parameter": "较少", "min": 0.7, "max": 1.0, "minExclusive": True, "type": "range"},
                            {"parameter": "一般", "min": 1.0, "max": 1.3, "minExclusive": True, "type": "range"},
                            {"parameter": "较多", "min": 1.3, "max": 1.5, "minExclusive": True, "type": "range"},
                            {"parameter": "很多", "min": 1.5, "max": 2.0, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "safetyTraining", "name": "员工安全教育培训调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "每年定期对员工进行安全教育和培训", "min": 0.7, "max": 1.0, "type": "range"},
                            {"parameter": "不定期对员工进行安全教育和培训", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "较少对员工进行安全教育和培训", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "safetyEquipment", "name": "安全设施和装备配置情况调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "安全设施和装备配置齐全", "min": 0.7, "max": 1.0, "type": "range"},
                            {"parameter": "安全设施和装备配置较齐全", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "安全设施和装备配置不齐全", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "renewal", "name": "续保调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "新保", "value": 1.0, "type": "fixed"},
                            {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                            {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                        ]
                    }
                ]
            },
            "v2026": {
                "label": "雇主责任险（2026版）费率",
                "baseRates": {
                    "fixed": {"class1": 0.0012, "class2": 0.0017, "class3": 0.0029},
                    "salary": {"class1": 0.0035, "class2": 0.0051, "class3": 0.0085}
                },
                "coefficients": [
                    {
                        "id": "perPersonLimit", "name": "每人赔偿限额调整系数", "applicableTo": ["fixed"],
                        "note": "每人赔偿限额按死亡/伤残赔偿限额高者取值；未列明限额可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤10万元", "min": 1.2, "max": 1.3, "type": "range"},
                            {"parameter": "30万元", "value": 1.1, "type": "fixed"},
                            {"parameter": "50万元", "value": 1.0, "type": "fixed"},
                            {"parameter": "80万元", "value": 0.9, "type": "fixed"},
                            {"parameter": "≥100万元", "min": 0.8, "max": 0.85, "type": "range"}
                        ]
                    },
                    {
                        "id": "employeeCount", "name": "承保人数调整系数", "applicableTo": ["fixed"],
                        "note": "未列明人数可按线性插值法计算",
                        "rows": [
                            {"parameter": "＜100人", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "[100, 500)人", "min": 0.9, "max": 1.0, "minExclusive": True, "type": "range"},
                            {"parameter": "[500, 1000)人", "min": 0.8, "max": 0.9, "minExclusive": True, "type": "range"},
                            {"parameter": "≥1000人", "min": 0.7, "max": 0.8, "type": "range"}
                        ]
                    },
                    {
                        "id": "deathDisabilityMonths", "name": "每人赔偿限额调整系数（工资月数）", "applicableTo": ["salary"],
                        "rows": [
                            {"parameter": "36/48个月", "value": 1.0, "type": "fixed"},
                            {"parameter": "48/60个月", "value": 1.25, "type": "fixed"},
                            {"parameter": "60/72个月", "value": 1.4, "type": "fixed"},
                            {"parameter": "72/84个月", "value": 1.5, "type": "fixed"}
                        ]
                    },
                    {
                        "id": "medicalLimit", "name": "每人医疗费用赔偿限额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "每人医疗费用赔偿限额÷每人赔偿限额；未列明比例可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤5%", "min": 0.9, "max": 0.95, "type": "range"},
                            {"parameter": "10%", "value": 1.0, "type": "fixed"},
                            {"parameter": "15%", "value": 1.05, "type": "fixed"},
                            {"parameter": "20%", "value": 1.1, "type": "fixed"},
                            {"parameter": "≥25%", "min": 1.15, "max": 1.3, "type": "range"}
                        ]
                    },
                    {
                        "id": "lostWorkDaily", "name": "每人每天误工费用金额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "基于雇员月平均工资；未列明金额可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤月工资÷40", "min": 0.9, "max": 0.95, "type": "range"},
                            {"parameter": "月工资÷30", "value": 1.0, "type": "fixed"},
                            {"parameter": "≥月工资÷20", "min": 1.05, "max": 1.1, "type": "range"}
                        ]
                    },
                    {
                        "id": "lostWorkDays", "name": "单次及累计赔偿误工费用天数调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "未列明天数可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤120天", "min": 0.95, "max": 0.97, "type": "range"},
                            {"parameter": "180天", "value": 1.0, "type": "fixed"},
                            {"parameter": "240天", "value": 1.03, "type": "fixed"},
                            {"parameter": "≥300天", "min": 1.06, "max": 1.1, "type": "range"}
                        ]
                    },
                    {
                        "id": "lostWorkLimitRatio", "name": "每人误工费用赔偿限额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "每人误工费用赔偿限额÷（每人每天误工费用金额×天数）；未列明比例可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤50%", "min": 0.95, "max": 0.96, "type": "range"},
                            {"parameter": "75%", "value": 0.98, "type": "fixed"},
                            {"parameter": "100%", "value": 1.0, "type": "fixed"}
                        ]
                    },
                    {
                        "id": "cumulativeRatio", "name": "累计赔偿限额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "累计赔偿限额÷每人赔偿限额；未列明比例可按线性插值法计算",
                        "rows": [
                            {"parameter": "≤5倍", "min": 0.9, "max": 0.95, "type": "range"},
                            {"parameter": "10倍", "value": 1.0, "type": "fixed"},
                            {"parameter": "20倍", "value": 1.05, "type": "fixed"},
                            {"parameter": "≥30倍", "min": 1.1, "max": 1.2, "type": "range"}
                        ]
                    },
                    {
                        "id": "deductibleRate", "name": "医疗费用免赔率调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "每次事故每人医疗费用免赔率；若同时约定免赔率与免赔额，以两者系数的低者取值",
                        "rows": [
                            {"parameter": "0", "value": 1.0, "type": "fixed"},
                            {"parameter": "10%", "value": 0.97, "type": "fixed"},
                            {"parameter": "20%", "value": 0.94, "type": "fixed"},
                            {"parameter": "30%", "value": 0.91, "type": "fixed"}
                        ]
                    },
                    {
                        "id": "deductibleAmount", "name": "医疗费用免赔额调整系数", "applicableTo": ["fixed", "salary"],
                        "note": "每次事故每人医疗费用免赔额；未列明免赔额可按线性插值法计算",
                        "rows": [
                            {"parameter": "0元", "value": 1.0, "type": "fixed"},
                            {"parameter": "500元", "value": 0.97, "type": "fixed"},
                            {"parameter": "1000元", "value": 0.94, "type": "fixed"},
                            {"parameter": "≥1500元", "min": 0.85, "max": 0.9, "type": "range"}
                        ]
                    },
                    {
                        "id": "employeeCategory", "name": "雇员类别调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "管理人员", "min": 0.7, "max": 0.8, "type": "range"},
                            {"parameter": "后勤人员", "min": 0.9, "max": 1.0, "type": "range"},
                            {"parameter": "一线操作人员", "min": 1.0, "max": 2.0, "type": "range"}
                        ]
                    },
                    {
                        "id": "managementLevel", "name": "管理水平调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "制度完善，无明显缺陷", "min": 0.7, "max": 1.0, "type": "range"},
                            {"parameter": "较完善，存在个别缺陷", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "不完善或存在较多缺陷", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "lossRatio", "name": "赔付率调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "[0, 20%]", "min": 0.5, "max": 0.6, "type": "range"},
                            {"parameter": "(20%, 45%]", "min": 0.6, "max": 0.8, "minExclusive": True, "type": "range"},
                            {"parameter": "(45%, 70%]", "min": 0.8, "max": 1.0, "minExclusive": True, "type": "range"},
                            {"parameter": "(70%, 95%]", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "＞95%", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "hazardInspection", "name": "企业隐患排查整改调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "无隐患", "min": 0.7, "max": 1.0, "type": "range"},
                            {"parameter": "整改完成", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "存在重大隐患且未整改", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "historicalAccident", "name": "历史事故与损失情况调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "极少", "min": 0.5, "max": 0.7, "type": "range"},
                            {"parameter": "较少", "min": 0.7, "max": 1.0, "minExclusive": True, "type": "range"},
                            {"parameter": "一般", "min": 1.0, "max": 1.3, "minExclusive": True, "type": "range"},
                            {"parameter": "较多", "min": 1.3, "max": 1.5, "minExclusive": True, "type": "range"},
                            {"parameter": "很多", "min": 1.5, "max": 2.0, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "safetyTraining", "name": "员工安全教育培训调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "每年定期对员工进行安全教育和培训", "min": 0.7, "max": 1.0, "type": "range"},
                            {"parameter": "不定期对员工进行安全教育和培训", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "较少对员工进行安全教育和培训", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "safetyEquipment", "name": "安全设施和装备配置情况调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "安全设施和装备配置齐全", "min": 0.7, "max": 1.0, "type": "range"},
                            {"parameter": "安全设施和装备配置较齐全", "min": 1.0, "max": 1.2, "minExclusive": True, "type": "range"},
                            {"parameter": "安全设施和装备配置不齐全", "min": 1.2, "max": 1.5, "minExclusive": True, "type": "range"}
                        ]
                    },
                    {
                        "id": "renewal", "name": "续保调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "新保", "value": 1.0, "type": "fixed"},
                            {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                            {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                        ]
                    },
                    {
                        "id": "workInjuryInsurance", "name": "工伤保险情况调整系数", "applicableTo": ["fixed", "salary"],
                        "rows": [
                            {"parameter": "已投保工伤保险", "value": 1.0, "type": "fixed"},
                            {"parameter": "未投保工伤保险", "value": 1.2, "type": "fixed"}
                        ]
                    }
                ]
            }
        }
    },
    "propertyAllRisk": {
        "productName": "财产一切险",
        "productType": "property",
        "amountUnit": "亿元",
        "amountLabel": "保险金额",
        "premiumCap": 0.70,
        "versions": {
            "original": {
                "label": "财产一切险费率",
                "baseRates": {"default": 0.0020},
                "coefficients": [
                    {"id": "propertyType", "name": "财产类别调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "房屋建筑", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "机器设备", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "存货", "min": 1.0, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "buildingStructure", "name": "建筑结构调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "钢结构/钢混结构", "min": 0.7, "max": 0.9, "type": "range"},
                         {"parameter": "砖混结构", "min": 0.9, "max": 1.1, "type": "range"},
                         {"parameter": "砖木/木结构", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "fireProtection", "name": "消防设施调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "完善（自动喷淋+报警）", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "较完善（灭火器+消火栓）", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "不完善", "min": 1.0, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "securityMeasures", "name": "安全措施调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "24小时监控+保安", "min": 0.7, "max": 0.9, "type": "range"},
                         {"parameter": "基本安保措施", "value": 1.0, "type": "fixed"},
                         {"parameter": "安保措施不足", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "location", "name": "地理位置调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "非自然灾害高发区", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "一般地区", "value": 1.0, "type": "fixed"},
                         {"parameter": "自然灾害高发区", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "deductibleProp", "name": "免赔额/免赔率调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "高免赔（≥10万或5%）", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "标准免赔", "value": 1.0, "type": "fixed"},
                         {"parameter": "低免赔/零免赔", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "lossHistory", "name": "历史赔付调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "近3年无赔付", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "赔付率＜50%", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "赔付率50%-80%", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "赔付率＞80%", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "insuredAmount", "name": "保险金额规模调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "＞10亿元", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "5-10亿元", "min": 0.85, "max": 0.95, "type": "range"},
                         {"parameter": "1-5亿元", "value": 1.0, "type": "fixed"},
                         {"parameter": "＜1亿元", "min": 1.05, "max": 1.2, "type": "range"}
                     ]},
                    {"id": "industryRisk", "name": "行业风险调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "低风险（办公/商业）", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "中风险（轻工制造）", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "较高风险（重工/化工）", "min": 1.0, "max": 1.3, "type": "range"},
                         {"parameter": "高风险（易燃易爆）", "min": 1.3, "max": 2.0, "type": "range"}
                     ]},
                    {"id": "managementLevelProp", "name": "风险管理水平调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "优秀", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "良好", "min": 0.85, "max": 1.0, "type": "range"},
                         {"parameter": "一般", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "较差", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "renewalProp", "name": "续保调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "新保", "value": 1.0, "type": "fixed"},
                         {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                         {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                     ]},
                    {"id": "coverageScope", "name": "承保范围调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "标准承保范围", "value": 1.0, "type": "fixed"},
                         {"parameter": "扩展暴风/洪水/地震", "min": 1.1, "max": 1.3, "type": "range"},
                         {"parameter": "限缩承保范围", "min": 0.8, "max": 0.95, "type": "range"}
                     ]},
                    {"id": "policyPeriod", "name": "保险期间调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "一年期", "value": 1.0, "type": "fixed"},
                         {"parameter": "长期（2-3年）", "min": 0.9, "max": 0.95, "type": "range"}
                     ]},
                    {"id": "coinsurance", "name": "共保/再保调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "有再保安排", "min": 0.9, "max": 1.0, "type": "range"},
                         {"parameter": "无再保安排", "value": 1.0, "type": "fixed"}
                     ]},
                    {"id": "specialAgreement", "name": "特别约定调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "标准条款", "value": 1.0, "type": "fixed"},
                         {"parameter": "含利于被保险人条款", "min": 1.05, "max": 1.2, "type": "range"},
                         {"parameter": "含利于保险人条款", "min": 0.85, "max": 0.95, "type": "range"}
                     ]},
                    {"id": "marketCompetition", "name": "市场竞争调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "竞争激烈", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般", "value": 1.0, "type": "fixed"},
                         {"parameter": "竞争较少", "min": 1.05, "max": 1.15, "type": "range"}
                     ]}
                ]
            }
        }
    },
    "propertyComprehensive": {
        "productName": "财产综合险",
        "productType": "property",
        "amountUnit": "亿元",
        "amountLabel": "保险金额",
        "premiumCap": 0.70,
        "versions": {
            "original": {
                "label": "财产综合险费率",
                "baseRates": {"default": 0.0017},
                "coefficients": [
                    {"id": "propertyType", "name": "财产类别调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "房屋建筑", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "机器设备", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "存货", "min": 1.0, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "buildingStructure", "name": "建筑结构调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "钢结构/钢混结构", "min": 0.7, "max": 0.9, "type": "range"},
                         {"parameter": "砖混结构", "min": 0.9, "max": 1.1, "type": "range"},
                         {"parameter": "砖木/木结构", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "fireProtection", "name": "消防设施调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "完善（自动喷淋+报警）", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "较完善（灭火器+消火栓）", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "不完善", "min": 1.0, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "securityMeasures", "name": "安全措施调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "24小时监控+保安", "min": 0.7, "max": 0.9, "type": "range"},
                         {"parameter": "基本安保措施", "value": 1.0, "type": "fixed"},
                         {"parameter": "安保措施不足", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "location", "name": "地理位置调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "非自然灾害高发区", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "一般地区", "value": 1.0, "type": "fixed"},
                         {"parameter": "自然灾害高发区", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "deductibleProp", "name": "免赔额/免赔率调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "高免赔（≥10万或5%）", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "标准免赔", "value": 1.0, "type": "fixed"},
                         {"parameter": "低免赔/零免赔", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "lossHistory", "name": "历史赔付调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "近3年无赔付", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "赔付率＜50%", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "赔付率50%-80%", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "赔付率＞80%", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "insuredAmount", "name": "保险金额规模调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "＞10亿元", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "5-10亿元", "min": 0.85, "max": 0.95, "type": "range"},
                         {"parameter": "1-5亿元", "value": 1.0, "type": "fixed"},
                         {"parameter": "＜1亿元", "min": 1.05, "max": 1.2, "type": "range"}
                     ]},
                    {"id": "industryRisk", "name": "行业风险调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "低风险（办公/商业）", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "中风险（轻工制造）", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "较高风险（重工/化工）", "min": 1.0, "max": 1.3, "type": "range"},
                         {"parameter": "高风险（易燃易爆）", "min": 1.3, "max": 2.0, "type": "range"}
                     ]},
                    {"id": "managementLevelProp", "name": "风险管理水平调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "优秀", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "良好", "min": 0.85, "max": 1.0, "type": "range"},
                         {"parameter": "一般", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "较差", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "renewalProp", "name": "续保调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "新保", "value": 1.0, "type": "fixed"},
                         {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                         {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                     ]},
                    {"id": "coverageScope", "name": "承保范围调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "标准承保范围", "value": 1.0, "type": "fixed"},
                         {"parameter": "扩展暴风/洪水/地震", "min": 1.1, "max": 1.3, "type": "range"},
                         {"parameter": "限缩承保范围", "min": 0.8, "max": 0.95, "type": "range"}
                     ]},
                    {"id": "policyPeriod", "name": "保险期间调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "一年期", "value": 1.0, "type": "fixed"},
                         {"parameter": "长期（2-3年）", "min": 0.9, "max": 0.95, "type": "range"}
                     ]},
                    {"id": "coinsurance", "name": "共保/再保调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "有再保安排", "min": 0.9, "max": 1.0, "type": "range"},
                         {"parameter": "无再保安排", "value": 1.0, "type": "fixed"}
                     ]},
                    {"id": "specialAgreement", "name": "特别约定调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "标准条款", "value": 1.0, "type": "fixed"},
                         {"parameter": "含利于被保险人条款", "min": 1.05, "max": 1.2, "type": "range"},
                         {"parameter": "含利于保险人条款", "min": 0.85, "max": 0.95, "type": "range"}
                     ]},
                    {"id": "marketCompetition", "name": "市场竞争调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "竞争激烈", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般", "value": 1.0, "type": "fixed"},
                         {"parameter": "竞争较少", "min": 1.05, "max": 1.15, "type": "range"}
                     ]}
                ]
            }
        }
    },
    "machineryBreakdown": {
        "productName": "机器损坏保险",
        "productType": "property",
        "amountUnit": "千万元",
        "amountLabel": "保险金额",
        "premiumCap": None,
        "versions": {
            "original": {
                "label": "机器损坏保险费率",
                "baseRates": {"default": 0.0020},
                "coefficients": [
                    {"id": "machineType", "name": "机器类别调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "通用机械设备", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "精密仪器/电子设备", "min": 1.0, "max": 1.3, "type": "range"},
                         {"parameter": "大型发电/变压设备", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "machineAge", "name": "设备年限调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "＜3年（新设备）", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "3-8年", "value": 1.0, "type": "fixed"},
                         {"parameter": "8-15年", "min": 1.1, "max": 1.3, "type": "range"},
                         {"parameter": "＞15年", "min": 1.3, "max": 1.6, "type": "range"}
                     ]},
                    {"id": "maintenanceLevel", "name": "维护保养水平调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "定期专业维保+完整记录", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "基本维保", "value": 1.0, "type": "fixed"},
                         {"parameter": "维保不足", "min": 1.1, "max": 1.4, "type": "range"}
                     ]},
                    {"id": "operatorQualification", "name": "操作人员资质调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "专业持证上岗", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "经过培训", "value": 1.0, "type": "fixed"},
                         {"parameter": "资质不明", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "environment", "name": "使用环境调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "室内恒温恒湿", "min": 0.7, "max": 0.9, "type": "range"},
                         {"parameter": "室内一般环境", "value": 1.0, "type": "fixed"},
                         {"parameter": "室外/恶劣环境", "min": 1.1, "max": 1.4, "type": "range"}
                     ]},
                    {"id": "deductibleMach", "name": "免赔额调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "高免赔（≥保额1%）", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "标准免赔", "value": 1.0, "type": "fixed"},
                         {"parameter": "低免赔/零免赔", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "lossHistoryMach", "name": "历史赔付调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "近3年无赔付", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "赔付率＜50%", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "赔付率50%-80%", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "赔付率＞80%", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "insuredAmountMach", "name": "保险金额规模调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "＞5000万元", "min": 0.75, "max": 0.9, "type": "range"},
                         {"parameter": "1000-5000万元", "value": 1.0, "type": "fixed"},
                         {"parameter": "＜1000万元", "min": 1.05, "max": 1.2, "type": "range"}
                     ]},
                    {"id": "sparePartsAvail", "name": "备品备件可获得性调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "国产常见型号", "min": 0.8, "max": 0.95, "type": "range"},
                         {"parameter": "进口但有国内代理", "value": 1.0, "type": "fixed"},
                         {"parameter": "特殊定制/停产型号", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "powerSupply", "name": "供电质量调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "双回路+UPS", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般供电", "value": 1.0, "type": "fixed"},
                         {"parameter": "供电不稳定", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "renewalMach", "name": "续保调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "新保", "value": 1.0, "type": "fixed"},
                         {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                         {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                     ]},
                    {"id": "manufacturer", "name": "制造商/品牌调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "知名品牌/可靠性高", "min": 0.85, "max": 0.95, "type": "range"},
                         {"parameter": "一般品牌", "value": 1.0, "type": "fixed"},
                         {"parameter": "非知名品牌", "min": 1.05, "max": 1.15, "type": "range"}
                     ]},
                    {"id": "marketMach", "name": "市场竞争调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "竞争激烈", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般", "value": 1.0, "type": "fixed"},
                         {"parameter": "竞争较少", "min": 1.05, "max": 1.15, "type": "range"}
                     ]}
                ]
            }
        }
    },
    "compositeHengli2025": {
        "productName": "突发和意外的直接物质损失一切险(恒力2025版)",
        "productType": "composite",
        "amountUnit": "元",
        "amountLabel": "保险金额",
        "premiumCap": None,
        "subRisks": ["materialDamage", "machineryBreakdown"],
        "subRiskLabels": {"materialDamage": "物质损失", "machineryBreakdown": "机器损坏"},
        "versions": {
            "hengli2025": {
                "label": "恒力2025版费率",
                "baseRates": {"materialDamage": 0.00089, "machineryBreakdown": 0.00362},
                "coefficients": [
                    {"id": "propertyTypeHL", "name": "财产类别调整系数", "applicableTo": ["materialDamage", "machineryBreakdown"],
                     "rows": [
                         {"parameter": "房屋建筑+机器设备+存货", "value": 1.0, "type": "fixed"},
                         {"parameter": "仅房屋建筑", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "仅机器设备", "min": 1.0, "max": 1.2, "type": "range"}
                     ]},
                    {"id": "buildingStructureHL", "name": "建筑结构调整系数", "applicableTo": ["materialDamage"],
                     "rows": [
                         {"parameter": "钢结构/钢混结构", "min": 0.7, "max": 0.9, "type": "range"},
                         {"parameter": "砖混结构", "min": 0.9, "max": 1.1, "type": "range"},
                         {"parameter": "砖木/木结构", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "fireProtectionHL", "name": "消防设施调整系数", "applicableTo": ["materialDamage"],
                     "rows": [
                         {"parameter": "完善", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "较完善", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "不完善", "min": 1.0, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "machineTypeHL", "name": "机器类别调整系数", "applicableTo": ["machineryBreakdown"],
                     "rows": [
                         {"parameter": "通用机械", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "精密仪器/电子设备", "min": 1.0, "max": 1.3, "type": "range"},
                         {"parameter": "大型特种设备", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "machineAgeHL", "name": "设备年限调整系数", "applicableTo": ["machineryBreakdown"],
                     "rows": [
                         {"parameter": "＜5年", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "5-10年", "value": 1.0, "type": "fixed"},
                         {"parameter": "＞10年", "min": 1.1, "max": 1.4, "type": "range"}
                     ]},
                    {"id": "maintenanceLevelHL", "name": "维护保养水平调整系数", "applicableTo": ["machineryBreakdown"],
                     "rows": [
                         {"parameter": "定期专业维保", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "基本维保", "value": 1.0, "type": "fixed"},
                         {"parameter": "维保不足", "min": 1.1, "max": 1.4, "type": "range"}
                     ]},
                    {"id": "deductibleHL", "name": "免赔额/免赔率调整系数", "applicableTo": ["materialDamage", "machineryBreakdown"],
                     "rows": [
                         {"parameter": "高免赔", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "标准免赔", "value": 1.0, "type": "fixed"},
                         {"parameter": "低免赔", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "lossHistoryHL", "name": "历史赔付调整系数", "applicableTo": ["materialDamage", "machineryBreakdown"],
                     "rows": [
                         {"parameter": "近3年无赔付", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "赔付率＜50%", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "赔付率50%-80%", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "赔付率＞80%", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "locationHL", "name": "地理位置调整系数", "applicableTo": ["materialDamage"],
                     "rows": [
                         {"parameter": "非自然灾害高发区", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "一般地区", "value": 1.0, "type": "fixed"},
                         {"parameter": "自然灾害高发区", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "industryRiskHL", "name": "行业风险调整系数", "applicableTo": ["materialDamage", "machineryBreakdown"],
                     "rows": [
                         {"parameter": "低风险", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "中风险", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "较高风险", "min": 1.0, "max": 1.3, "type": "range"},
                         {"parameter": "高风险", "min": 1.3, "max": 2.0, "type": "range"}
                     ]},
                    {"id": "managementHL", "name": "风险管理水平调整系数", "applicableTo": ["materialDamage", "machineryBreakdown"],
                     "rows": [
                         {"parameter": "优秀", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "良好", "min": 0.85, "max": 1.0, "type": "range"},
                         {"parameter": "一般", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "较差", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "renewalHL", "name": "续保调整系数", "applicableTo": ["materialDamage", "machineryBreakdown"],
                     "rows": [
                         {"parameter": "新保", "value": 1.0, "type": "fixed"},
                         {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                         {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                     ]},
                    {"id": "coverageScopeHL", "name": "承保范围调整系数", "applicableTo": ["materialDamage"],
                     "rows": [
                         {"parameter": "标准承保范围", "value": 1.0, "type": "fixed"},
                         {"parameter": "扩展自然灾害", "min": 1.1, "max": 1.3, "type": "range"},
                         {"parameter": "限缩承保范围", "min": 0.8, "max": 0.95, "type": "range"}
                     ]},
                    {"id": "marketHL", "name": "市场竞争调整系数", "applicableTo": ["materialDamage", "machineryBreakdown"],
                     "rows": [
                         {"parameter": "竞争激烈", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般", "value": 1.0, "type": "fixed"},
                         {"parameter": "竞争较少", "min": 1.05, "max": 1.15, "type": "range"}
                     ]}
                ]
            }
        }
    },
    "businessInterruption": {
        "productName": "营业中断保险",
        "productType": "interruption",
        "amountUnit": "元",
        "amountLabel": "毛利润损失保险金额",
        "premiumCap": None,
        "versions": {
            "original": {
                "label": "营业中断保险费率",
                "baseRates": {"default": 0.0011},
                "coefficients": [
                    {"id": "industryBI", "name": "行业类别调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "低风险行业（服务/零售）", "min": 0.7, "max": 0.9, "type": "range"},
                         {"parameter": "中风险行业（一般制造）", "value": 1.0, "type": "fixed"},
                         {"parameter": "高风险行业（化工/能源）", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "indemnityPeriod", "name": "赔偿期间调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "6个月", "min": 0.7, "max": 0.8, "type": "range"},
                         {"parameter": "12个月", "value": 1.0, "type": "fixed"},
                         {"parameter": "18个月", "min": 1.2, "max": 1.4, "type": "range"},
                         {"parameter": "24个月", "min": 1.4, "max": 1.6, "type": "range"}
                     ]},
                    {"id": "waitingPeriod", "name": "等待期调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "无等待期", "min": 1.1, "max": 1.2, "type": "range"},
                         {"parameter": "3天", "value": 1.0, "type": "fixed"},
                         {"parameter": "7天", "min": 0.9, "max": 0.95, "type": "range"},
                         {"parameter": "14天及以上", "min": 0.8, "max": 0.9, "type": "range"}
                     ]},
                    {"id": "supplyChainRisk", "name": "供应链依赖度调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "供应商分散/可替代", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般依赖", "value": 1.0, "type": "fixed"},
                         {"parameter": "高度依赖单一供应商", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "backupFacility", "name": "备用设施/预案调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "有完善备用设施和预案", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "有基本预案", "value": 1.0, "type": "fixed"},
                         {"parameter": "无备用设施和预案", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "lossHistoryBI", "name": "历史赔付调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "近3年无赔付", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "赔付率＜50%", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "赔付率＞50%", "min": 1.0, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "deductibleBI", "name": "免赔额调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "高免赔", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "标准免赔", "value": 1.0, "type": "fixed"},
                         {"parameter": "低免赔", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "renewalBI", "name": "续保调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "新保", "value": 1.0, "type": "fixed"},
                         {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                         {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                     ]},
                    {"id": "managementBI", "name": "风险管理水平调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "优秀", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "良好", "min": 0.85, "max": 1.0, "type": "range"},
                         {"parameter": "一般", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "较差", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "marketBI", "name": "市场竞争调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "竞争激烈", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般", "value": 1.0, "type": "fixed"},
                         {"parameter": "竞争较少", "min": 1.05, "max": 1.15, "type": "range"}
                     ]}
                ]
            }
        }
    },
    "cashComprehensive": {
        "productName": "现金综合保险",
        "productType": "property",
        "amountUnit": "元",
        "amountLabel": "保险金额",
        "premiumCap": None,
        "extraFields": ["dailyCashTurnover"],
        "versions": {
            "original": {
                "label": "现金综合保险费率",
                "baseRates": {"default": 0.00082},
                "coefficients": [
                    {"id": "cashStorageMethod", "name": "现金保管方式调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "银行级保险柜+双人管理", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "标准保险柜", "value": 1.0, "type": "fixed"},
                         {"parameter": "简易保管设施", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "transportSecurity", "name": "运送安全措施调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "专业押运公司", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "自行押运（双人+专车）", "value": 1.0, "type": "fixed"},
                         {"parameter": "简易运送", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "securitySystem", "name": "安防系统调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "完善（监控+报警+保安）", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "基本安防", "value": 1.0, "type": "fixed"},
                         {"parameter": "安防不足", "min": 1.1, "max": 1.4, "type": "range"}
                     ]},
                    {"id": "locationCash", "name": "经营地点调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "商业中心/治安良好区", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般商业区", "value": 1.0, "type": "fixed"},
                         {"parameter": "偏远/治安较差区", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "lossHistoryCash", "name": "历史赔付调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "近3年无赔付", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "赔付率＜50%", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "赔付率＞50%", "min": 1.0, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "deductibleCash", "name": "免赔额调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "高免赔", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "标准免赔", "value": 1.0, "type": "fixed"},
                         {"parameter": "低免赔", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "renewalCash", "name": "续保调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "新保", "value": 1.0, "type": "fixed"},
                         {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                         {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                     ]}
                ]
            }
        }
    },
    "jewelryComprehensive": {
        "productName": "珠宝商综合保险",
        "productType": "jewelry",
        "amountUnit": "元",
        "amountLabel": "保险金额",
        "premiumCap": None,
        "merchantTypes": {"wholesale": "批发商", "retail": "零售商"},
        "coverageCategories": {
            "storeDisplay": {"label": "店铺陈列", "baseRates": {"wholesale": 0.0035, "retail": 0.0040}},
            "safeStorage": {"label": "保险柜保管", "baseRates": {"wholesale": 0.0015, "retail": 0.0018}},
            "transit": {"label": "运输途中", "baseRates": {"wholesale": 0.0050, "retail": 0.0045}},
            "exhibition": {"label": "展览展示", "baseRates": {"wholesale": 0.0060, "retail": 0.0055}}
        },
        "versions": {
            "original": {
                "label": "珠宝商综合保险费率",
                "baseRates": {},
                "coefficients": [
                    {"id": "securityLevelJewel", "name": "安保等级调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "高等级（武装押运+银行级保险柜+24h监控）", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "中等级（保安+标准保险柜+监控）", "value": 1.0, "type": "fixed"},
                         {"parameter": "低等级（基本安保）", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "locationJewel", "name": "经营地点调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "高端商场/商业中心", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般商业街", "value": 1.0, "type": "fixed"},
                         {"parameter": "偏远地区/治安较差", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "goodsType", "name": "货物类别调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "黄金/白银", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "钻石/宝石", "min": 1.0, "max": 1.2, "type": "range"},
                         {"parameter": "高价值名表/艺术品", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "lossHistoryJewel", "name": "历史赔付调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "近3年无赔付", "min": 0.6, "max": 0.8, "type": "range"},
                         {"parameter": "赔付率＜50%", "min": 0.8, "max": 1.0, "type": "range"},
                         {"parameter": "赔付率＞50%", "min": 1.0, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "deductibleJewel", "name": "免赔额调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "高免赔（≥保额2%）", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "标准免赔", "value": 1.0, "type": "fixed"},
                         {"parameter": "低免赔", "min": 1.1, "max": 1.3, "type": "range"}
                     ]},
                    {"id": "transitMethod", "name": "运输方式调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "专业押运车", "min": 0.7, "max": 0.85, "type": "range"},
                         {"parameter": "普通车辆+押运人员", "value": 1.0, "type": "fixed"},
                         {"parameter": "快递/邮寄", "min": 1.2, "max": 1.5, "type": "range"}
                     ]},
                    {"id": "renewalJewel", "name": "续保调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "新保", "value": 1.0, "type": "fixed"},
                         {"parameter": "续保一年", "value": 0.95, "type": "fixed"},
                         {"parameter": "续保两年及以上", "min": 0.8, "max": 0.9, "type": "range"}
                     ]},
                    {"id": "marketJewel", "name": "市场竞争调整系数", "applicableTo": ["all"],
                     "rows": [
                         {"parameter": "竞争激烈", "min": 0.8, "max": 0.9, "type": "range"},
                         {"parameter": "一般", "value": 1.0, "type": "fixed"},
                         {"parameter": "竞争较少", "min": 1.05, "max": 1.15, "type": "range"}
                     ]}
                ]
            }
        }
    }
}

MC_DISABILITY_TABLES = {
    "table1": {
        "label": "附表1",
        "ratios": [
            {"level": "一级伤残", "pct": 100}, {"level": "二级伤残", "pct": 80},
            {"level": "三级伤残", "pct": 70}, {"level": "四级伤残", "pct": 60},
            {"level": "五级伤残", "pct": 50}, {"level": "六级伤残", "pct": 40},
            {"level": "七级伤残", "pct": 30}, {"level": "八级伤残", "pct": 20},
            {"level": "九级伤残", "pct": 10}, {"level": "十级伤残", "pct": 5}
        ]
    },
    "table2": {
        "label": "附表2",
        "ratios": [
            {"level": "一级伤残", "pct": 100}, {"level": "二级伤残", "pct": 80},
            {"level": "三级伤残", "pct": 65}, {"level": "四级伤残", "pct": 55},
            {"level": "五级伤残", "pct": 45}, {"level": "六级伤残", "pct": 25},
            {"level": "七级伤残", "pct": 15}, {"level": "八级伤残", "pct": 10},
            {"level": "九级伤残", "pct": 4}, {"level": "十级伤残", "pct": 1}
        ]
    },
    "table3": {
        "label": "附表3",
        "ratios": [
            {"level": "一级伤残", "pct": 100}, {"level": "二级伤残", "pct": 90},
            {"level": "三级伤残", "pct": 80}, {"level": "四级伤残", "pct": 70},
            {"level": "五级伤残", "pct": 60}, {"level": "六级伤残", "pct": 50},
            {"level": "七级伤残", "pct": 40}, {"level": "八级伤残", "pct": 30},
            {"level": "九级伤残", "pct": 20}, {"level": "十级伤残", "pct": 10}
        ]
    }
}

MC_DISABILITY_ADDON_OPTIONS = [
    {"group": "A", "label": "A组: 九级10% 十级5%", "p9": 10, "p10": 5, "coeff": {"table1": 1.000, "table2": 1.077, "table3": 0.924}},
    {"group": "A", "label": "A组: 九级8% 十级5%", "p9": 8, "p10": 5, "coeff": {"table1": 0.996, "table2": 1.073, "table3": 0.921}},
    {"group": "A", "label": "A组: 九级15% 十级5%", "p9": 15, "p10": 5, "coeff": {"table1": 1.009, "table2": 1.087, "table3": 0.933}},
    {"group": "A", "label": "A组: 九级10% 十级3%", "p9": 10, "p10": 3, "coeff": {"table1": 0.995, "table2": 1.072, "table3": 0.919}},
    {"group": "A", "label": "A组: 九级15% 十级3%", "p9": 15, "p10": 3, "coeff": {"table1": 1.004, "table2": 1.082, "table3": 0.928}},
    {"group": "B", "label": "B组: 九级4% 十级1%", "p9": 4, "p10": 1, "coeff": {"table1": 0.928, "table2": 1.000, "table3": 0.858}},
    {"group": "B", "label": "B组: 九级2% 十级1%", "p9": 2, "p10": 1, "coeff": {"table1": 0.924, "table2": 0.996, "table3": 0.854}},
    {"group": "B", "label": "B组: 九级3% 十级1%", "p9": 3, "p10": 1, "coeff": {"table1": 0.926, "table2": 0.998, "table3": 0.856}},
    {"group": "B", "label": "B组: 九级3% 十级2%", "p9": 3, "p10": 2, "coeff": {"table1": 0.929, "table2": 1.001, "table3": 0.858}},
    {"group": "B", "label": "B组: 九级4% 十级2%", "p9": 4, "p10": 2, "coeff": {"table1": 0.931, "table2": 1.003, "table3": 0.860}},
    {"group": "C", "label": "C组: 九级20% 十级10%", "p9": 20, "p10": 10, "coeff": {"table1": 1.082, "table2": 1.166, "table3": 1.000}},
    {"group": "C", "label": "C组: 九级15% 十级8%", "p9": 15, "p10": 8, "coeff": {"table1": 1.068, "table2": 1.151, "table3": 0.987}},
    {"group": "C", "label": "C组: 九级20% 十级8%", "p9": 20, "p10": 8, "coeff": {"table1": 1.077, "table2": 1.161, "table3": 0.995}},
    {"group": "C", "label": "C组: 九级15% 十级6%", "p9": 15, "p10": 6, "coeff": {"table1": 1.063, "table2": 1.145, "table3": 0.982}},
    {"group": "C", "label": "C组: 九级20% 十级6%", "p9": 20, "p10": 6, "coeff": {"table1": 1.072, "table2": 1.155, "table3": 0.991}}
]

MC_DISABILITY_GROUP_DESC = {
    "A": "二级80% 三级70% 四级60% 五级50% 六级40% 七级30% 八级20%",
    "B": "二级80% 三级65% 四级55% 五级45% 六级25% 七级15% 八级10%",
    "C": "二级90% 三级80% 四级70% 五级60% 六级50% 七级40% 八级30%"
}


# =============================================
# 工具函数
# =============================================

def fmt_currency(num):
    """格式化货币"""
    if num is None or math.isnan(num):
        return "¥0.00"
    return f"¥{abs(num):,.2f}"


def fmt_num(num, digits=4):
    """格式化数字"""
    return f"{float(num):.{digits}f}"


def is_leap_year(year):
    """判断闰年"""
    return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)


# =============================================
# 通用样式
# =============================================

def get_common_styles():
    """返回通用控件样式"""
    return f"""
        QLabel {{
            color: {AnthropicColors.TEXT_PRIMARY};
        }}
        QComboBox {{
            background: {AnthropicColors.BG_PRIMARY};
            border: 1px solid {AnthropicColors.BORDER};
            border-radius: 8px;
            padding: 8px 12px;
            color: {AnthropicColors.TEXT_PRIMARY};
            font-size: 13px;
            min-height: 20px;
        }}
        QComboBox:focus {{
            border-color: {AnthropicColors.ACCENT};
        }}
        QComboBox::drop-down {{
            border: none;
            width: 24px;
        }}
        QSpinBox, QDoubleSpinBox {{
            background: {AnthropicColors.BG_PRIMARY};
            border: 1px solid {AnthropicColors.BORDER};
            border-radius: 8px;
            padding: 8px 12px;
            color: {AnthropicColors.TEXT_PRIMARY};
            font-size: 13px;
            min-height: 20px;
        }}
        QSpinBox:focus, QDoubleSpinBox:focus {{
            border-color: {AnthropicColors.ACCENT};
        }}
        QLineEdit {{
            background: {AnthropicColors.BG_PRIMARY};
            border: 1px solid {AnthropicColors.BORDER};
            border-radius: 8px;
            padding: 8px 12px;
            color: {AnthropicColors.TEXT_PRIMARY};
            font-size: 13px;
        }}
        QLineEdit:focus {{
            border-color: {AnthropicColors.ACCENT};
        }}
        QPushButton {{
            background: {AnthropicColors.BG_CARD};
            border: 1px solid {AnthropicColors.BORDER};
            border-radius: 8px;
            padding: 8px 16px;
            color: {AnthropicColors.TEXT_PRIMARY};
            font-weight: 500;
            font-size: 13px;
        }}
        QPushButton:hover {{
            background: {AnthropicColors.BG_PRIMARY};
            border-color: {AnthropicColors.ACCENT};
        }}
        QTextEdit {{
            background: {AnthropicColors.BG_PRIMARY};
            border: 1px solid {AnthropicColors.BORDER};
            border-radius: 8px;
            padding: 8px;
            color: {AnthropicColors.TEXT_PRIMARY};
            font-size: 12px;
        }}
        QScrollArea {{
            border: none;
            background: transparent;
        }}
        QDateEdit {{
            background: {AnthropicColors.BG_PRIMARY};
            border: 1px solid {AnthropicColors.BORDER};
            border-radius: 8px;
            padding: 8px 12px;
            color: {AnthropicColors.TEXT_PRIMARY};
            font-size: 13px;
            min-height: 20px;
        }}
        QDateEdit:focus {{
            border-color: {AnthropicColors.ACCENT};
        }}
    """


def make_accent_button(text):
    """创建强调色按钮"""
    btn = QPushButton(text)
    btn.setCursor(Qt.PointingHandCursor)
    btn.setStyleSheet(f"""
        QPushButton {{
            background: {AnthropicColors.ACCENT};
            color: {AnthropicColors.TEXT_LIGHT};
            border: none;
            border-radius: 8px;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: 600;
        }}
        QPushButton:hover {{
            background: {AnthropicColors.ACCENT_DARK};
        }}
    """)
    return btn


def make_success_button(text):
    """创建成功色按钮"""
    btn = QPushButton(text)
    btn.setCursor(Qt.PointingHandCursor)
    btn.setStyleSheet(f"""
        QPushButton {{
            background: #10b981;
            color: white;
            border: none;
            border-radius: 8px;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: 600;
        }}
        QPushButton:hover {{
            background: #059669;
        }}
    """)
    return btn


# =============================================
# MainInsuranceTab — 主险计算器
# =============================================

class MainInsuranceTab(QWidget):
    """主险计算器 Tab"""
    premium_calculated = pyqtSignal(float, float)
    # 完整计算结果信号，传递给附加险使用
    full_result_calculated = pyqtSignal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_product = "employerLiability"
        self.selected_version = "original"
        self.current_plan = MC_PRODUCTS["employerLiability"]["versions"]["original"]
        self.coeff_selections = {}
        self.selected_disability_table = "none"
        self.selected_disability_option = -1
        self.result = None
        self._setup_ui()

    def _setup_ui(self):
        self.setStyleSheet(get_common_styles())
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(15, 10, 15, 10)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(scroll_widget)
        self.scroll_layout.setSpacing(12)
        scroll.setWidget(scroll_widget)
        main_layout.addWidget(scroll, 1)

        self._build_control_bar()
        self._build_params_section()
        self._build_disability_section()
        self._build_coeff_section()
        self._build_action_buttons()
        self._build_result_section()
        self._build_log_section()
        self.scroll_layout.addStretch()

    def _build_control_bar(self):
        card = GlassCard()
        layout = QHBoxLayout(card)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.addWidget(QLabel("险种:"))
        self.product_combo = QComboBox()
        for pid, pdata in MC_PRODUCTS.items():
            self.product_combo.addItem(pdata["productName"], pid)
        self.product_combo.currentIndexChanged.connect(self._on_product_change)
        layout.addWidget(self.product_combo)
        layout.addWidget(QLabel("版本:"))
        self.version_combo = QComboBox()
        self._refresh_version_combo()
        self.version_combo.currentIndexChanged.connect(self._on_version_change)
        layout.addWidget(self.version_combo)
        layout.addStretch()
        import_btn = QPushButton("📂 导入费率方案")
        import_btn.setCursor(Qt.PointingHandCursor)
        import_btn.clicked.connect(self._import_rate_plan)
        layout.addWidget(import_btn)
        self.scroll_layout.addWidget(card)

    def _refresh_version_combo(self, select_version=None):
        self.version_combo.blockSignals(True)
        self.version_combo.clear()
        product = MC_PRODUCTS.get(self.selected_product)
        if product:
            for vid, vdata in product["versions"].items():
                self.version_combo.addItem(vdata["label"], vid)
        if select_version:
            idx = self.version_combo.findData(select_version)
            if idx >= 0:
                self.version_combo.setCurrentIndex(idx)
        self.version_combo.blockSignals(False)
        self.selected_version = self.version_combo.currentData() or "original"
        self.current_plan = MC_PRODUCTS.get(self.selected_product, {}).get("versions", {}).get(self.selected_version, {})

    def _on_product_change(self):
        self.selected_product = self.product_combo.currentData()
        self._refresh_version_combo()
        self._on_version_change()
        self._log(f"切换险种: {MC_PRODUCTS[self.selected_product]['productName']}")

    def _on_version_change(self):
        self.selected_version = self.version_combo.currentData() or "original"
        product = MC_PRODUCTS.get(self.selected_product, {})
        self.current_plan = product.get("versions", {}).get(self.selected_version, {})
        self.coeff_selections = {}
        self.result = None
        self._clear_result()
        self._render_coefficients()
        if self.current_plan:
            self._log(f"切换版本: {self.current_plan.get('label', '')}")

    def _build_params_section(self):
        card = GlassCard()
        grid = QGridLayout(card)
        grid.setContentsMargins(16, 12, 16, 12)
        grid.setSpacing(10)

        grid.addWidget(QLabel("行业类别:"), 0, 0)
        self.industry_combo = QComboBox()
        self.industry_combo.addItem("一类行业", "class1")
        self.industry_combo.addItem("二类行业", "class2")
        self.industry_combo.addItem("三类行业", "class3")
        grid.addWidget(self.industry_combo, 0, 1)

        grid.addWidget(QLabel("计费方式:"), 0, 2)
        self.method_combo = QComboBox()
        self.method_combo.addItem("固定限额", "fixed")
        self.method_combo.addItem("工资总额", "salary")
        self.method_combo.currentIndexChanged.connect(self._on_method_change)
        grid.addWidget(self.method_combo, 0, 3)

        self.limit_label = QLabel("每人限额(万元):")
        grid.addWidget(self.limit_label, 1, 0)
        self.limit_spin = QDoubleSpinBox()
        self.limit_spin.setRange(1, 10000)
        self.limit_spin.setValue(50)
        self.limit_spin.setDecimals(2)
        self.limit_spin.setSuffix(" 万元")
        grid.addWidget(self.limit_spin, 1, 1)

        self.salary_label = QLabel("年度工资总额(元):")
        grid.addWidget(self.salary_label, 1, 2)
        self.salary_spin = QDoubleSpinBox()
        self.salary_spin.setRange(0, 999999999999)
        self.salary_spin.setValue(5000000)
        self.salary_spin.setDecimals(2)
        self.salary_spin.setSuffix(" 元")
        grid.addWidget(self.salary_spin, 1, 3)
        self.salary_label.hide()
        self.salary_spin.hide()

        grid.addWidget(QLabel("承保人数:"), 2, 0)
        self.count_spin = QSpinBox()
        self.count_spin.setRange(1, 999999)
        self.count_spin.setValue(100)
        self.count_spin.setSuffix(" 人")
        grid.addWidget(self.count_spin, 2, 1)

        grid.addWidget(QLabel("保险期间:"), 2, 2)
        self.term_combo = QComboBox()
        self.term_combo.addItem("年度", "annual")
        self.term_combo.addItem("短期", "short")
        self.term_combo.currentIndexChanged.connect(self._on_term_change)
        grid.addWidget(self.term_combo, 2, 3)

        self.days_label = QLabel("保险天数:")
        grid.addWidget(self.days_label, 3, 0)
        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 365)
        self.days_spin.setValue(180)
        self.days_spin.setSuffix(" 天")
        grid.addWidget(self.days_spin, 3, 1)
        self.days_label.hide()
        self.days_spin.hide()
        self.scroll_layout.addWidget(card)

    def _on_method_change(self):
        is_fixed = self.method_combo.currentData() == "fixed"
        self.limit_label.setVisible(is_fixed)
        self.limit_spin.setVisible(is_fixed)
        self.salary_label.setVisible(not is_fixed)
        self.salary_spin.setVisible(not is_fixed)
        self.coeff_selections = {}
        self._render_coefficients()
        self._log(f"切换计费方式: {'固定限额' if is_fixed else '工资总额'}")

    def _on_term_change(self):
        is_short = self.term_combo.currentData() == "short"
        self.days_label.setVisible(is_short)
        self.days_spin.setVisible(is_short)

    def _build_disability_section(self):
        card = GlassCard()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 12, 16, 12)
        row = QHBoxLayout()
        row.addWidget(QLabel("伤残赔偿比例附表:"))
        self.disability_combo = QComboBox()
        self.disability_combo.addItem("不使用", "none")
        for tid, tdata in MC_DISABILITY_TABLES.items():
            self.disability_combo.addItem(tdata["label"], tid)
        self.disability_combo.currentIndexChanged.connect(self._on_disability_table_change)
        row.addWidget(self.disability_combo)
        row.addStretch()
        layout.addLayout(row)

        self.disability_display = QLabel("")
        self.disability_display.setWordWrap(True)
        self.disability_display.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 12px;")
        self.disability_display.hide()
        layout.addWidget(self.disability_display)

        self.disability_options_area = QWidget()
        self.disability_options_layout = QVBoxLayout(self.disability_options_area)
        self.disability_options_layout.setContentsMargins(0, 0, 0, 0)
        self.disability_options_area.hide()
        layout.addWidget(self.disability_options_area)
        self.scroll_layout.addWidget(card)

    def _on_disability_table_change(self):
        self.selected_disability_table = self.disability_combo.currentData()
        self.selected_disability_option = -1
        if self.selected_disability_table == "none":
            self.disability_display.hide()
            self.disability_options_area.hide()
            self._render_coefficients()
            self._log("已关闭伤残赔偿比例附表")
            return
        tbl = MC_DISABILITY_TABLES[self.selected_disability_table]
        lines = [f"📋 {tbl['label']} 伤残赔偿比例:"]
        for r in tbl["ratios"]:
            lines.append(f"  {r['level']}: {r['pct']}%")
        self.disability_display.setText("\n".join(lines))
        self.disability_display.show()
        self._render_disability_options()
        self.disability_options_area.show()
        self._render_coefficients()
        self._log(f"选择伤残赔偿比例: {tbl['label']}")

    def _render_disability_options(self):
        while self.disability_options_layout.count():
            item = self.disability_options_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        if self.selected_disability_table == "none":
            return
        title = QLabel("附加伤残赔偿金赔偿限额比例调整系数")
        title.setStyleSheet(f"font-weight: 600; color: {AnthropicColors.ACCENT}; font-size: 14px;")
        self.disability_options_layout.addWidget(title)
        for group_name in ["A", "B", "C"]:
            group_label = QLabel(f"{group_name}组 · 一级100% {MC_DISABILITY_GROUP_DESC[group_name]}")
            group_label.setStyleSheet(f"font-size: 11px; color: {AnthropicColors.TEXT_SECONDARY}; margin-top: 8px;")
            self.disability_options_layout.addWidget(group_label)
            for idx, opt in enumerate(MC_DISABILITY_ADDON_OPTIONS):
                if opt["group"] != group_name:
                    continue
                coeff_val = opt["coeff"][self.selected_disability_table]
                is_selected = self.selected_disability_option == idx
                btn = QPushButton(f"九级{opt['p9']}% 十级{opt['p10']}%  →  系数 {fmt_num(coeff_val, 3)}")
                btn.setCursor(Qt.PointingHandCursor)
                bg = AnthropicColors.ACCENT if is_selected else AnthropicColors.BG_PRIMARY
                fg = AnthropicColors.TEXT_LIGHT if is_selected else AnthropicColors.TEXT_PRIMARY
                btn.setStyleSheet(f"""
                    QPushButton {{ background: {bg}; color: {fg}; border: 1px solid {AnthropicColors.BORDER};
                        border-radius: 6px; padding: 6px 12px; font-size: 12px; text-align: left; }}
                    QPushButton:hover {{ border-color: {AnthropicColors.ACCENT}; }}
                """)
                btn.clicked.connect(lambda checked, i=idx: self._select_disability_option(i))
                self.disability_options_layout.addWidget(btn)

    def _select_disability_option(self, idx):
        self.selected_disability_option = idx
        self._render_disability_options()
        self._render_coefficients()
        opt = MC_DISABILITY_ADDON_OPTIONS[idx]
        self._log(f"选择伤残方案: {opt['label']} → 系数 {fmt_num(opt['coeff'][self.selected_disability_table], 3)}")

    def _build_coeff_section(self):
        self.coeff_container = QWidget()
        self.coeff_layout = QVBoxLayout(self.coeff_container)
        self.coeff_layout.setContentsMargins(0, 0, 0, 0)
        self.coeff_layout.setSpacing(8)
        self.scroll_layout.addWidget(self.coeff_container)
        self._render_coefficients()

    def _render_coefficients(self):
        while self.coeff_layout.count():
            item = self.coeff_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        if not self.current_plan:
            return
        method = self.method_combo.currentData()
        applicable = [c for c in self.current_plan.get("coefficients", []) if method in c.get("applicableTo", [])]
        if not applicable and self.selected_disability_table == "none":
            lbl = QLabel("当前计费方式无可用系数表")
            lbl.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; padding: 16px;")
            self.coeff_layout.addWidget(lbl)
            return
        for coeff in applicable:
            card = self._create_coeff_card(coeff)
            self.coeff_layout.addWidget(card)

    def _create_coeff_card(self, coeff):
        card = GlassCard()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(6)
        sel = self.coeff_selections.get(coeff["id"])
        sel_value = sel["value"] if sel else None
        title_text = coeff["name"]
        title_text += f"  [{fmt_num(sel_value, 4)}]" if sel_value is not None else "  [未选]"
        title = QLabel(title_text)
        title.setStyleSheet(f"font-weight: 600; font-size: 13px; color: {AnthropicColors.TEXT_PRIMARY};")
        layout.addWidget(title)
        if coeff.get("note"):
            note = QLabel(f"注: {coeff['note']}")
            note.setWordWrap(True)
            note.setStyleSheet(f"font-size: 11px; color: {AnthropicColors.TEXT_TERTIARY};")
            layout.addWidget(note)
        for ri, row in enumerate(coeff["rows"]):
            is_selected = sel and sel.get("rowIndex") == ri
            if row["type"] == "range":
                bl = "(" if row.get("minExclusive") else "["
                br = ")" if row.get("maxExclusive") else "]"
                val_text = f"{bl}{fmt_num(row['min'], 2)}, {fmt_num(row['max'], 2)}{br}"
            else:
                val_text = fmt_num(row["value"], 2)
            btn = QPushButton(f"{row['parameter']}    {val_text}")
            btn.setCursor(Qt.PointingHandCursor)
            bg = AnthropicColors.ACCENT if is_selected else AnthropicColors.BG_PRIMARY
            fg = AnthropicColors.TEXT_LIGHT if is_selected else AnthropicColors.TEXT_PRIMARY
            btn.setStyleSheet(f"""
                QPushButton {{ background: {bg}; color: {fg}; border: 1px solid {AnthropicColors.BORDER};
                    border-radius: 6px; padding: 6px 12px; font-size: 12px; text-align: left; }}
                QPushButton:hover {{ border-color: {AnthropicColors.ACCENT}; }}
            """)
            click_value = row["min"] if row["type"] == "range" else row["value"]
            btn.clicked.connect(lambda checked, cid=coeff["id"], ridx=ri, cv=click_value: self._select_coeff_row(cid, ridx, cv))
            layout.addWidget(btn)
        if sel and sel.get("rowIndex") is not None:
            row = coeff["rows"][sel["rowIndex"]]
            if row["type"] == "range":
                slider_layout = QHBoxLayout()
                slider_layout.addWidget(QLabel("精确系数: "))
                slider_label = QLabel(fmt_num(sel["value"], 4))
                slider_label.setStyleSheet(f"font-weight: 600; color: {AnthropicColors.ACCENT};")
                slider_layout.addWidget(slider_label)
                slider = QSlider(Qt.Horizontal)
                slider.setMinimum(int(row["min"] * 10000))
                slider.setMaximum(int(row["max"] * 10000))
                slider.setValue(int(sel["value"] * 10000))
                slider.setSingleStep(100)
                cid = coeff["id"]
                slider.valueChanged.connect(lambda v, c=cid, lr=slider_label: self._on_slider_change(c, v, lr))
                slider_layout.addWidget(slider, 1)
                range_info = QLabel(f"[{fmt_num(row['min'], 2)}, {fmt_num(row['max'], 2)}]")
                range_info.setStyleSheet(f"font-size: 11px; color: {AnthropicColors.TEXT_TERTIARY};")
                slider_layout.addWidget(range_info)
                layout.addLayout(slider_layout)
        return card

    def _select_coeff_row(self, coeff_id, row_index, value):
        self.coeff_selections[coeff_id] = {"rowIndex": row_index, "value": value}
        self._render_coefficients()

    def _on_slider_change(self, coeff_id, int_value, label_widget):
        value = int_value / 10000.0
        if coeff_id in self.coeff_selections:
            self.coeff_selections[coeff_id]["value"] = value
        label_widget.setText(fmt_num(value, 4))

    def _build_action_buttons(self):
        row = QHBoxLayout()
        calc_btn = make_accent_button("🧮 计算主险保费")
        calc_btn.clicked.connect(self.calculate)
        row.addWidget(calc_btn)
        reset_btn = QPushButton("🔄 重置参数")
        reset_btn.setCursor(Qt.PointingHandCursor)
        reset_btn.clicked.connect(self._reset)
        row.addWidget(reset_btn)
        self.send_btn = make_success_button("📤 传入附加险计算")
        self.send_btn.clicked.connect(self._send_to_addon)
        self.send_btn.hide()
        row.addWidget(self.send_btn)
        row.addStretch()
        self.scroll_layout.addLayout(row)

    def _build_result_section(self):
        self.result_display = QTextEdit()
        self.result_display.setReadOnly(True)
        self.result_display.setMaximumHeight(300)
        self.result_display.hide()
        self.scroll_layout.addWidget(self.result_display)

    def _clear_result(self):
        self.result_display.clear()
        self.result_display.hide()
        self.send_btn.hide()

    def _build_log_section(self):
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setMaximumHeight(150)
        self.log_display.setStyleSheet(f"""
            QTextEdit {{ background: {AnthropicColors.BG_DARK}; color: {AnthropicColors.TEXT_LIGHT};
                border-radius: 8px; padding: 8px; font-size: 11px; font-family: monospace; }}
        """)
        self.scroll_layout.addWidget(self.log_display)

    def _log(self, msg, level="info"):
        from datetime import datetime
        time_str = datetime.now().strftime("%H:%M:%S")
        prefix = {"error": "❌", "warn": "⚠️", "success": "✅"}.get(level, "ℹ️")
        self.log_display.append(f"[{time_str}] {prefix} {msg}")

    def calculate(self):
        method = self.method_combo.currentData()
        industry_class = self.industry_combo.currentData()
        employee_count = self.count_spin.value()
        term_type = self.term_combo.currentData()
        days = self.days_spin.value() if term_type == "short" else 365
        if employee_count <= 0:
            self._log("计算失败: 承保人数无效", "error")
            return
        base_rates = self.current_plan.get("baseRates", {}).get(method, {})
        base_rate = base_rates.get(industry_class)
        if not base_rate:
            self._log(f"计算失败: 基准费率不存在 method={method} class={industry_class}", "error")
            return
        self._log("--- 开始计算 ---")
        self._log(f"版本: {self.current_plan.get('label', '')} | 计费: {'固定限额' if method == 'fixed' else '工资总额'} | 行业: {industry_class}")
        self._log(f"基准费率: {base_rate * 100:.4f}%")
        applicable = [c for c in self.current_plan.get("coefficients", []) if method in c.get("applicableTo", [])]
        coeff_product = 1.0
        coeff_details = []
        unselected_count = 0
        for coeff in applicable:
            sel = self.coeff_selections.get(coeff["id"])
            if sel:
                coeff_product *= sel["value"]
                coeff_details.append({"name": coeff["name"], "value": sel["value"]})
                self._log(f"  系数 [{coeff['name']}] = {fmt_num(sel['value'], 4)}")
            else:
                coeff_details.append({"name": coeff["name"], "value": 1.0, "unselected": True})
                unselected_count += 1
        if unselected_count > 0:
            self._log(f"  注意: {unselected_count} 个系数未选择，按基准 1.0 计算", "warn")
        self._log(f"系数乘积: {fmt_num(coeff_product, 6)}")
        adjusted_rate = base_rate * coeff_product
        is_capped = False
        if adjusted_rate > 0.70:
            self._log(f"调整后费率 {adjusted_rate * 100:.4f}% 超过70%封顶", "warn")
            adjusted_rate = 0.70
            is_capped = True
        self._log(f"调整后费率: {adjusted_rate * 100:.4f}%{'（封顶）' if is_capped else ''}")
        per_person_premium = 0.0
        total_premium = 0.0
        formula = ""
        if method == "fixed":
            limit_yuan = self.limit_spin.value() * 10000
            per_person_premium = limit_yuan * adjusted_rate
            if term_type == "short":
                per_person_premium *= (days / 365)
            total_premium = per_person_premium * employee_count
            formula = f"每人保费 = {fmt_currency(limit_yuan)} × {adjusted_rate * 100:.4f}%"
            if term_type == "short":
                formula += f" × ({days}/365)"
            formula += f" = {fmt_currency(per_person_premium)}"
            formula += f"\n主险保费 = {fmt_currency(per_person_premium)} × {employee_count}人 = {fmt_currency(total_premium)}"
        else:
            salary_yuan = self.salary_spin.value()
            total_premium = salary_yuan * adjusted_rate
            if term_type == "short":
                total_premium *= (days / 365)
            per_person_premium = total_premium / employee_count if employee_count > 0 else 0
            formula = f"年保费 = {fmt_currency(salary_yuan)} × {adjusted_rate * 100:.4f}%"
            if term_type == "short":
                formula += f" × ({days}/365)"
            formula += f" = {fmt_currency(total_premium)}"
            formula += f"\n每人均摊: {fmt_currency(total_premium)} / {employee_count}人 = {fmt_currency(per_person_premium)}"
        disability_coeff = 1.0
        disability_desc = ""
        if self.selected_disability_table != "none" and self.selected_disability_option >= 0:
            d_opt = MC_DISABILITY_ADDON_OPTIONS[self.selected_disability_option]
            disability_coeff = d_opt["coeff"][self.selected_disability_table]
            before_premium = total_premium
            total_premium *= disability_coeff
            per_person_premium *= disability_coeff
            tbl_label = MC_DISABILITY_TABLES[self.selected_disability_table]["label"]
            disability_desc = f"附加伤残赔偿比例({tbl_label} · {d_opt['label']})"
            formula += f"\n\n扩展伤残赔偿比例: {fmt_currency(before_premium)} × {fmt_num(disability_coeff, 3)} = {fmt_currency(total_premium)}"
            self._log(f"伤残赔偿比例调整: × {fmt_num(disability_coeff, 3)} ({disability_desc})")
        self._log(f"主险保费: {fmt_currency(total_premium)}", "success")
        self._log("--- 计算完成 ---", "success")
        self.result = {
            "version": self.current_plan.get("label", ""), "method": method, "industryClass": industry_class,
            "baseRate": base_rate, "coeffProduct": coeff_product, "disabilityCoeff": disability_coeff,
            "disabilityDesc": disability_desc, "adjustedRate": adjusted_rate, "isCapped": is_capped,
            "perPersonPremium": per_person_premium, "totalPremium": total_premium,
            "employeeCount": employee_count, "termType": term_type, "days": days,
            "formulaBreakdown": formula, "coeffDetails": coeff_details
        }
        self._render_result()
        self.send_btn.show()

    def _render_result(self):
        if not self.result:
            return
        r = self.result
        lines = [
            "═══════════════ 📊 计算结果 ═══════════════", "",
            f"  主险保费:   {fmt_currency(r['totalPremium'])}",
            f"  每人保费:   {fmt_currency(r['perPersonPremium'])}", "",
            "─────────── 公式分解 ───────────",
            r["formulaBreakdown"], "",
            "─────────── 参数明细 ───────────",
            f"  费率版本: {r['version']}",
            f"  计费方式: {'固定限额' if r['method'] == 'fixed' else '工资总额'}",
            f"  行业类别: {r['industryClass']}",
            f"  基准费率: {r['baseRate'] * 100:.4f}%",
            f"  系数乘积: {fmt_num(r['coeffProduct'], 6)}",
            f"  调整后费率: {r['adjustedRate'] * 100:.4f}%{' (封顶)' if r['isCapped'] else ''}",
            f"  承保人数: {r['employeeCount']}人",
            f"  保险期间: {'年度' if r['termType'] == 'annual' else str(r['days']) + '天'}",
        ]
        if r["disabilityCoeff"] != 1.0:
            lines.append(f"  伤残赔偿比例系数: {fmt_num(r['disabilityCoeff'], 3)}")
            lines.append(f"  {r['disabilityDesc']}")
        lines.extend(["", "─────────── 系数明细 ───────────"])
        for d in r["coeffDetails"]:
            suffix = "（未选，默认）" if d.get("unselected") else ""
            lines.append(f"  {d['name']}: {fmt_num(d['value'], 4)}{suffix}")
        self.result_display.setText("\n".join(lines))
        self.result_display.show()

    def _reset(self):
        self.coeff_selections = {}
        self.result = None
        self.industry_combo.setCurrentIndex(0)
        self.method_combo.setCurrentIndex(0)
        self.limit_spin.setValue(50)
        self.salary_spin.setValue(5000000)
        self.count_spin.setValue(100)
        self.term_combo.setCurrentIndex(0)
        self.days_spin.setValue(180)
        self.selected_disability_table = "none"
        self.selected_disability_option = -1
        self.disability_combo.setCurrentIndex(0)
        self._clear_result()
        self._render_coefficients()
        self._log("已重置参数和系数选择（险种/版本不变）")

    def _send_to_addon(self):
        if not self.result:
            return
        self.premium_calculated.emit(self.result["totalPremium"], self.result["perPersonPremium"])
        # 构建完整数据包传递给附加险
        full_data = dict(self.result)
        full_data["perPersonLimit"] = self.limit_spin.value() * 10000 if self.method_combo.currentData() == "fixed" else 0
        full_data["annualSalary"] = self.salary_spin.value() if self.method_combo.currentData() == "salary" else 0
        full_data["disabilityTable"] = self.selected_disability_table
        full_data["disabilityOption"] = self.selected_disability_option
        self.full_result_calculated.emit(full_data)
        self._log(f"已将主险保费 {fmt_currency(self.result['totalPremium'])}、每人保费 {fmt_currency(self.result['perPersonPremium'])} 传入附加险计算", "success")

    def _import_rate_plan(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "导入费率方案", "", "费率方案文件 (*.json *.docx);;JSON (*.json);;Word (*.docx)")
        if not file_path:
            return
        if file_path.endswith(".json"):
            self._import_json(file_path)
        elif file_path.endswith(".docx"):
            self._import_docx(file_path)

    def _import_json(self, file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._process_imported_data(data)
        except Exception as e:
            self._log(f"JSON 导入失败: {e}", "error")

    def _process_imported_data(self, data):
        if data.get("productName"):
            product_id = data.get("productId", f"custom_{id(data)}")
            product_name = data["productName"]
            versions = {}
            if isinstance(data.get("versions"), list):
                for idx, v in enumerate(data["versions"]):
                    vid = v.get("versionId", f"v{idx + 1}")
                    if not all(k in v for k in ("label", "baseRates", "coefficients")):
                        raise ValueError(f"版本 {vid} 缺少必要字段")
                    versions[vid] = {"label": v["label"], "baseRates": v["baseRates"], "coefficients": v["coefficients"]}
            else:
                raise ValueError("新格式 JSON 需包含 versions 数组")
        else:
            if not all(k in data for k in ("label", "baseRates", "coefficients")):
                raise ValueError("JSON 缺少必要字段: label, baseRates, coefficients")
            product_id = f"custom_{id(data)}"
            product_name = data["label"]
            versions = {"v1": {"label": data["label"], "baseRates": data["baseRates"], "coefficients": data["coefficients"]}}
        first_version = None
        if product_id in MC_PRODUCTS:
            existing = MC_PRODUCTS[product_id]
            for vid, vdata in versions.items():
                final_vid = vid if vid not in existing["versions"] else f"{vid}_{id(vdata)}"
                existing["versions"][final_vid] = vdata
                if not first_version:
                    first_version = final_vid
            self._log(f"向险种 [{existing['productName']}] 追加了 {len(versions)} 个新版本", "success")
        else:
            MC_PRODUCTS[product_id] = {"productName": product_name, "versions": versions}
            first_version = list(versions.keys())[0]
            self._log(f"导入新险种: {product_name}，包含 {len(versions)} 个版本", "success")
        self.product_combo.blockSignals(True)
        self.product_combo.clear()
        for pid, pdata in MC_PRODUCTS.items():
            self.product_combo.addItem(pdata["productName"], pid)
        idx = self.product_combo.findData(product_id)
        if idx >= 0:
            self.product_combo.setCurrentIndex(idx)
        self.product_combo.blockSignals(False)
        self.selected_product = product_id
        self._refresh_version_combo(first_version)
        self._on_version_change()

    def _import_docx(self, file_path):
        try:
            from docx import Document
        except ImportError:
            self._log("python-docx 未安装，请运行: pip install python-docx", "error")
            return
        try:
            doc = Document(file_path)
            text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
            parsed = self._parse_rate_plan_text(text)
            reply = QMessageBox.question(
                self, "Docx 导入确认",
                f"险种名称: {parsed['productName']}\n基准费率: {len(parsed['baseRates'].get('fixed', {}))} 个固定 + {len(parsed['baseRates'].get('salary', {}))} 个工资\n系数表: {len(parsed['coefficients'])} 个\n\n确认导入?",
                QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                import_data = {"productName": parsed["productName"], "productId": f"docx_{id(parsed)}",
                               "versions": [{"versionId": "v1", "label": parsed["label"], "baseRates": parsed["baseRates"], "coefficients": parsed["coefficients"]}]}
                self._process_imported_data(import_data)
        except Exception as e:
            self._log(f"Docx 解析失败: {e}", "error")

    def _parse_rate_plan_text(self, text):
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        product_name = "未知险种"
        for line in lines[:5]:
            if "费率" in line or "保险" in line:
                product_name = re.sub(r"费率方案|费率表|附件[:：]?\s*", "", line).strip()[:20]
                break
        base_rates = {"fixed": {}, "salary": {}}
        class_map = {"一": "class1", "二": "class2", "三": "class3", "1": "class1", "2": "class2", "3": "class3"}
        full_text = "\n".join(lines)
        rate_pattern = re.compile(r"[第]?([一二三1-3])[类].*?(\d+\.?\d*)\s*[%‰％]")
        fixed_section = re.search(r"固定[赔偿]*限额[\s\S]*?(?=工资|$)", full_text, re.IGNORECASE)
        if fixed_section:
            for m in rate_pattern.finditer(fixed_section.group()):
                cls = class_map.get(m.group(1))
                if cls:
                    val = float(m.group(2))
                    base_rates["fixed"][cls] = val / 1000 if "‰" in m.group() else val / 100
        salary_section = re.search(r"工资[收入]*[\s\S]*?(?=费率调整|调整系数|$)", full_text, re.IGNORECASE)
        if salary_section:
            for m in rate_pattern.finditer(salary_section.group()):
                cls = class_map.get(m.group(1))
                if cls:
                    val = float(m.group(2))
                    base_rates["salary"][cls] = val / 1000 if "‰" in m.group() else val / 100
        if not base_rates["fixed"] and not base_rates["salary"]:
            raise ValueError("未能从文本中提取到基准费率数据")
        if not base_rates["fixed"]:
            base_rates["fixed"] = dict(base_rates["salary"])
        if not base_rates["salary"]:
            base_rates["salary"] = dict(base_rates["fixed"])
        return {"productName": product_name, "label": f"{product_name}费率", "baseRates": base_rates, "coefficients": []}


# =============================================
# 附加险分类常量
# =============================================

ADDON_TYPES = {
    "modifier_coeff": {"label": "主险系数调整", "color": "#8b5cf6"},
    "sudden_death": {"label": "突发疾病身故", "color": "#ef4444"},
    "per_person_rate": {"label": "每人费率", "color": "#f59e0b"},
    "per_person_base": {"label": "每人定额", "color": "#06b6d4"},
    "disability_adjust": {"label": "伤残调整", "color": "#10b981"},
    "property_loss": {"label": "财产损失", "color": "#0d9488"},
    "formula_sum": {"label": "求和公式", "color": "#3b82f6"},
    "deduction": {"label": "减收", "color": "#ec4899"},
    "no_calc": {"label": "无需计算", "color": "#6b7280"},
    "simple_percentage": {"label": "百分比", "color": "#d97706"},
    "table_coefficient": {"label": "系数表", "color": "#7c3aed"},
    "regulatory": {"label": "规范类", "color": "#9ca3af"},
}

# 关键词映射：文件名关键词 → 附加险类型
ADDON_KEYWORD_MAP = [
    (["误工费"], "modifier_coeff"),
    (["突发疾病身故"], "sudden_death"),
    (["特定财产损失"], "property_loss"),
    (["工伤补充", "特定人员"], "per_person_rate"),
    (["药品服务", "药品费用"], "per_person_base"),
    (["劳务关系人员"], "disability_adjust"),
    (["雇主法律责任", "法律费用责任"], "formula_sum"),
    (["一次性伤残"], "formula_sum"),
    (["突发疾病除外", "猝死除外"], "deduction"),
    (["月申报", "员工自动承保", "每月申报"], "no_calc"),
]

# 伤残调整系数 (劳务关系人员)
DISABILITY_ADJUST_COEFFS = {
    "table1": 0.995,
    "table2": 1.072,
    "table3": 0.919,
}


# =============================================
# AddonInsuranceTab — 附加险计算器 (重构版)
# =============================================

class AddonInsuranceTab(QWidget):
    """附加险计算器 Tab — 支持11种附加险分类和专属计算逻辑"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.rate_data = None
        self.filtered_entries = []
        self.selected_entry = None
        self.coeff_selections = {}
        # 主险数据
        self.main_premium = 0.0
        self.per_person_premium = 0.0
        self.full_main_data = None  # 完整主险计算结果
        # 保费汇总
        self.premium_items = []
        self._setup_ui()

    def _setup_ui(self):
        self.setStyleSheet(get_common_styles())
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(10, 8, 10, 8)

        # 顶栏: 主险信息 + 导入按钮
        top_bar = GlassCard()
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(16, 10, 16, 10)

        top_layout.addWidget(QLabel("主险保费:"))
        self.main_premium_input = QDoubleSpinBox()
        self.main_premium_input.setRange(0, 999999999999)
        self.main_premium_input.setDecimals(2)
        self.main_premium_input.setSuffix(" 元")
        self.main_premium_input.valueChanged.connect(lambda v: setattr(self, 'main_premium', v))
        top_layout.addWidget(self.main_premium_input)

        top_layout.addWidget(QLabel("每人保费:"))
        self.per_person_input = QDoubleSpinBox()
        self.per_person_input.setRange(0, 999999999999)
        self.per_person_input.setDecimals(2)
        self.per_person_input.setSuffix(" 元")
        self.per_person_input.valueChanged.connect(lambda v: setattr(self, 'per_person_premium', v))
        top_layout.addWidget(self.per_person_input)

        # 主险数据状态指示
        self.main_data_status = QLabel("⚪ 未接收主险数据")
        self.main_data_status.setStyleSheet(f"font-size: 11px; color: {AnthropicColors.TEXT_TERTIARY};")
        top_layout.addWidget(self.main_data_status)

        top_layout.addStretch()

        folder_btn = QPushButton("📁 导入文件夹")
        folder_btn.setCursor(Qt.PointingHandCursor)
        folder_btn.clicked.connect(self._load_folder)
        top_layout.addWidget(folder_btn)

        json_btn = QPushButton("📂 导入JSON")
        json_btn.setCursor(Qt.PointingHandCursor)
        json_btn.clicked.connect(self._load_json)
        top_layout.addWidget(json_btn)

        inquiry_btn = QPushButton("📋 导入询价")
        inquiry_btn.setCursor(Qt.PointingHandCursor)
        inquiry_btn.clicked.connect(self._handle_inquiry_import)
        top_layout.addWidget(inquiry_btn)

        main_layout.addWidget(top_bar)

        # 三列布局
        content = QHBoxLayout()
        content.setSpacing(8)

        # 左列: 搜索 + 条款列表
        left_panel = QWidget()
        left_panel.setFixedWidth(340)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 搜索条款名称...")
        self.search_input.textChanged.connect(self._filter_entries)
        left_layout.addWidget(self.search_input)

        self.load_status = QLabel("未加载费率数据")
        self.load_status.setStyleSheet(f"font-size: 11px; color: {AnthropicColors.TEXT_TERTIARY};")
        left_layout.addWidget(self.load_status)

        self.clause_list = QListWidget()
        self.clause_list.setStyleSheet(f"""
            QListWidget {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                font-size: 12px;
            }}
            QListWidget::item {{
                padding: 8px 12px;
                border-bottom: 1px solid {AnthropicColors.BORDER};
            }}
            QListWidget::item:selected {{
                background: {AnthropicColors.ACCENT};
                color: {AnthropicColors.TEXT_LIGHT};
            }}
            QListWidget::item:hover:!selected {{
                background: {AnthropicColors.BG_CARD};
            }}
        """)
        self.clause_list.currentRowChanged.connect(self._on_clause_selected)
        left_layout.addWidget(self.clause_list, 1)

        self.inquiry_status = QLabel("")
        self.inquiry_status.setWordWrap(True)
        self.inquiry_status.setStyleSheet(f"font-size: 11px; color: {AnthropicColors.TEXT_TERTIARY};")
        left_layout.addWidget(self.inquiry_status)

        self.batch_calc_btn = make_accent_button("⚡ 一键批量计算")
        self.batch_calc_btn.clicked.connect(self._batch_calculate)
        self.batch_calc_btn.hide()
        left_layout.addWidget(self.batch_calc_btn)

        content.addWidget(left_panel)

        # 中列: 详情 + 计算
        mid_scroll = QScrollArea()
        mid_scroll.setWidgetResizable(True)
        mid_scroll.setFrameShape(QFrame.NoFrame)
        self.detail_widget = QWidget()
        self.detail_layout = QVBoxLayout(self.detail_widget)
        self.detail_layout.setContentsMargins(8, 0, 8, 0)
        self.detail_layout.setSpacing(8)
        mid_scroll.setWidget(self.detail_widget)

        self.detail_placeholder = QLabel("📊 请导入费率方案文件夹，然后从左侧选择条款")
        self.detail_placeholder.setAlignment(Qt.AlignCenter)
        self.detail_placeholder.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 16px; padding: 60px;")
        self.detail_layout.addWidget(self.detail_placeholder)
        self.detail_layout.addStretch()

        content.addWidget(mid_scroll, 1)

        # 右列: 保费汇总
        right_panel = QWidget()
        right_panel.setFixedWidth(280)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        summary_title = QLabel("💰 保费汇总")
        summary_title.setStyleSheet(f"font-weight: 600; font-size: 14px; color: {AnthropicColors.TEXT_PRIMARY};")
        right_layout.addWidget(summary_title)

        self.premium_list_area = QScrollArea()
        self.premium_list_area.setWidgetResizable(True)
        self.premium_list_area.setFrameShape(QFrame.NoFrame)
        self.premium_list_widget = QWidget()
        self.premium_list_layout = QVBoxLayout(self.premium_list_widget)
        self.premium_list_layout.setContentsMargins(0, 0, 0, 0)
        self.premium_list_layout.setSpacing(4)
        self.premium_list_area.setWidget(self.premium_list_widget)
        right_layout.addWidget(self.premium_list_area, 1)

        self.premium_empty_label = QLabel("计算附加险保费后\n将自动添加到此列表")
        self.premium_empty_label.setAlignment(Qt.AlignCenter)
        self.premium_empty_label.setStyleSheet(f"color: {AnthropicColors.TEXT_TERTIARY}; font-size: 12px; padding: 20px;")
        self.premium_list_layout.addWidget(self.premium_empty_label)

        self.addon_total_label = QLabel("附加险合计: ¥0.00")
        self.addon_total_label.setStyleSheet(f"font-weight: 600; font-size: 13px; color: {AnthropicColors.ACCENT}; padding: 8px;")
        right_layout.addWidget(self.addon_total_label)

        self.annual_total_label = QLabel("保单预估年保费: ¥0.00")
        self.annual_total_label.setStyleSheet(f"font-weight: 700; font-size: 15px; color: #10b981; padding: 8px; background: #ecfdf5; border-radius: 8px;")
        right_layout.addWidget(self.annual_total_label)

        # 短期保费计算
        short_card = GlassCard()
        short_layout = QVBoxLayout(short_card)
        short_layout.setContentsMargins(12, 10, 12, 10)
        short_layout.setSpacing(6)
        short_layout.addWidget(QLabel("📅 短期保费计算"))

        date_row1 = QHBoxLayout()
        date_row1.addWidget(QLabel("起保日:"))
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate())
        self.start_date.dateChanged.connect(self._calc_short_term)
        date_row1.addWidget(self.start_date)
        short_layout.addLayout(date_row1)

        date_row2 = QHBoxLayout()
        date_row2.addWidget(QLabel("终止日:"))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate().addYears(1))
        self.end_date.dateChanged.connect(self._calc_short_term)
        date_row2.addWidget(self.end_date)
        short_layout.addLayout(date_row2)

        self.short_term_result = QLabel("")
        self.short_term_result.setWordWrap(True)
        self.short_term_result.setStyleSheet(f"font-size: 12px; color: {AnthropicColors.TEXT_PRIMARY};")
        short_layout.addWidget(self.short_term_result)
        right_layout.addWidget(short_card)

        content.addWidget(right_panel)
        main_layout.addLayout(content, 1)

        # 底部日志
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setMaximumHeight(100)
        self.log_display.setStyleSheet(f"""
            QTextEdit {{ background: {AnthropicColors.BG_DARK}; color: {AnthropicColors.TEXT_LIGHT};
                border-radius: 8px; padding: 6px; font-size: 11px; font-family: monospace; }}
        """)
        main_layout.addWidget(self.log_display)

    # ---------- 信号接收 ----------
    def receive_main_premium(self, total, per_person):
        self.main_premium = total
        self.per_person_premium = per_person
        self.main_premium_input.setValue(total)
        self.per_person_input.setValue(per_person)
        self._log(f"收到主险保费: {fmt_currency(total)}，每人: {fmt_currency(per_person)}", "success")

    def receive_full_data(self, data):
        """接收主险完整计算数据"""
        self.full_main_data = data
        self.main_premium = data.get("totalPremium", 0)
        self.per_person_premium = data.get("perPersonPremium", 0)
        self.main_premium_input.setValue(self.main_premium)
        self.per_person_input.setValue(self.per_person_premium)
        method = data.get("method", "")
        base_rate = data.get("baseRate", 0)
        coeff_product = data.get("coeffProduct", 1)
        count = data.get("employeeCount", 0)
        industry = data.get("industryClass", "")
        limit_val = data.get("perPersonLimit", 0)
        salary_val = data.get("annualSalary", 0)
        dis_table = data.get("disabilityTable", "none")
        self.main_data_status.setText(
            f"🟢 {method}制 · {industry}类 · 基准{base_rate:.6f} · 系数积{coeff_product:.4f} · "
            f"{'限额' + fmt_currency(limit_val) if limit_val else '工资' + fmt_currency(salary_val)} · "
            f"{count}人 · 伤残{dis_table}"
        )
        self.main_data_status.setStyleSheet(f"font-size: 11px; color: #10b981;")
        self._log(f"收到主险完整数据: 基准费率={base_rate:.6f}, 系数积={coeff_product:.4f}, {count}人", "success")

    # ---------- 日志 ----------
    def _log(self, msg, level="info"):
        from datetime import datetime
        time_str = datetime.now().strftime("%H:%M:%S")
        prefix = {"error": "❌", "warn": "⚠️", "success": "✅"}.get(level, "ℹ️")
        self.log_display.append(f"[{time_str}] {prefix} {msg}")

    # ---------- 数据加载 ----------
    def _load_json(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "加载费率数据", "", "JSON (*.json)")
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data.get("entries"), list):
                raise ValueError("JSON 格式无效: 缺少 entries 数组")
            self.rate_data = data
            self.filtered_entries = list(data["entries"])
            self.load_status.setText(f"已加载 {len(data['entries'])} 条 ({os.path.basename(file_path)})")
            self._render_clause_list()
            self._log(f"加载成功: {len(data['entries'])} 个费率方案", "success")
        except Exception as e:
            self._log(f"JSON 加载失败: {e}", "error")

    def _load_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择费率方案文件夹")
        if not folder:
            return
        try:
            from docx import Document as DocxDocument
        except ImportError:
            self._log("python-docx 未安装，请运行: pip install python-docx", "error")
            return
        entries = []
        docx_files = [f for f in os.listdir(folder) if "费率方案" in f and f.endswith(".docx") and not f.startswith("~$")]
        if not docx_files:
            self._log("未找到费率方案 docx 文件", "warn")
            return
        self._log(f"发现 {len(docx_files)} 个费率方案文件，开始智能分类解析...")
        type_counts = {}
        for fname in sorted(docx_files):
            try:
                fpath = os.path.join(folder, fname)
                doc = DocxDocument(fpath)
                paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
                tables = []
                for tbl in doc.tables:
                    rows = []
                    for row in tbl.rows:
                        cells = [cell.text.strip() for cell in row.cells]
                        rows.append(cells)
                    if rows:
                        tables.append(rows)
                entry = self._classify_entry(fname, paragraphs, tables)
                if entry:
                    entries.append(entry)
                    rt = entry.get("rateType", "unknown")
                    type_counts[rt] = type_counts.get(rt, 0) + 1
            except Exception as e:
                self._log(f"解析失败: {fname} - {e}", "warn")
        self.rate_data = {"entries": entries}
        self.filtered_entries = list(entries)
        self.load_status.setText(f"已加载 {len(entries)} 条 (文件夹)")
        self._render_clause_list()
        type_summary = ", ".join(f"{ADDON_TYPES.get(k, {}).get('label', k)}:{v}" for k, v in sorted(type_counts.items()))
        self._log(f"解析完成: {len(entries)} 个费率方案 [{type_summary}]", "success")

    def _classify_entry(self, filename, paragraphs, tables):
        """智能分类附加险类型 — 基于文件名关键词 + 文本内容分析"""
        name = filename.replace(".docx", "").replace("中国太平洋财产保险股份有限公司", "")
        clause_name = name.replace("费率方案", "").strip()
        m = re.match(r"附加(.+?)(?:条款|保险)?$", clause_name)
        if m:
            clause_name = "附加" + m.group(1)

        entry = {
            "clauseName": clause_name,
            "fullName": filename.replace(".docx", ""),
            "industry": "雇主责任保险",
            "sourceFile": filename,
        }
        full_text = " ".join(paragraphs)
        substantive = [p for p in paragraphs
                       if "中国太平洋财产保险股份有限公司" not in p
                       and not (p.endswith("费率方案") and len(p) < 100)]

        # ===== Step 1: 关键词匹配确定类型 =====
        detected_type = None
        for keywords, addon_type in ADDON_KEYWORD_MAP:
            if any(kw in clause_name or kw in full_text for kw in keywords):
                detected_type = addon_type
                break

        # ===== Step 2: 规范类检测 =====
        reg_keywords = ["不涉及保险费的调整", "属于规范类", "不涉及费率", "不另收保险费"]
        is_regulatory = (not tables and not detected_type and substantive and
                         all(any(kw in p for kw in reg_keywords) or "保单最终保险费" in p or "工资总额" in p
                             for p in substantive))
        if is_regulatory:
            return {**entry, "rateType": "regulatory", "description": substantive[0] if substantive else ""}

        # ===== Step 3: 按检测类型构建 entry =====

        # 解析所有表格（通用）
        coeff_tables = self._parse_tables(tables, full_text)

        # 提取百分比
        percentages = []
        for p in paragraphs:
            for pct_m in re.finditer(r"([\d.]+)\s*[%％]", p):
                percentages.append({"value": float(pct_m.group(1)), "context": p})

        # --- modifier_coeff: 误工费 ---
        if detected_type == "modifier_coeff":
            return {**entry, "rateType": "modifier_coeff",
                    "coefficientTables": coeff_tables,
                    "description": substantive[0] if substantive else "",
                    "formula": "调整后主险保费 = 主险保费 × 免赔天数调整系数"}

        # --- sudden_death: 突发疾病身故 ---
        if detected_type == "sudden_death":
            base_pct = None
            for pi in percentages:
                if "基准保险费" in pi["context"] or "每人每次事故赔偿限额" in pi["context"]:
                    base_pct = pi["value"]
                    break
            if not base_pct:
                base_pct = 6.6  # 默认
            return {**entry, "rateType": "sudden_death",
                    "basePercent": base_pct,
                    "coefficientTables": coeff_tables,
                    "description": substantive[0] if substantive else "",
                    "formula": f"基准保费 = 每人限额 × {base_pct}% × 人数，再乘以系数调整"}

        # --- per_person_rate: 工伤补充/特定人员 ---
        if detected_type == "per_person_rate":
            rate_info = {}
            for pi in percentages:
                ctx = pi["context"]
                if "已购买工伤保险" in ctx or "有工伤" in ctx:
                    rate_info["with_injury_insurance"] = pi["value"]
                elif "未购买工伤保险" in ctx or "无工伤" in ctx:
                    rate_info["without_injury_insurance"] = pi["value"]
                elif not rate_info:
                    rate_info["default"] = pi["value"]
            return {**entry, "rateType": "per_person_rate",
                    "rateInfo": rate_info,
                    "coefficientTables": coeff_tables,
                    "description": substantive[0] if substantive else "",
                    "formula": "保费 = 每人保费 × 费率% × 人数 × 系数"}

        # --- property_loss: 员工个人特定财产损失 ---
        if detected_type == "property_loss":
            base_premium = 20  # 基本保险费=基准保险费=20元
            base_rate_pct = 1.5  # 基准费率1.5%
            for pi in percentages:
                base_rate_pct = pi["value"]
                break
            for p in paragraphs:
                amt_m = re.search(r"基[本准]保险费[=＝]?\s*(\d+)\s*元", p)
                if amt_m:
                    base_premium = int(amt_m.group(1))
                    break
            return {**entry, "rateType": "property_loss",
                    "basePremium": base_premium,
                    "baseRatePercent": base_rate_pct,
                    "coefficientTables": coeff_tables,
                    "description": substantive[0] if substantive else "",
                    "formula": f"保费 = ({base_premium}元 + 每人赔偿限额 × {base_rate_pct}%) × 系数积 × 承保人数"}

        # --- per_person_base: 药品服务 ---
        if detected_type == "per_person_base":
            base_amount = 300  # 默认
            for p in paragraphs:
                amt_m = re.search(r"(\d+)\s*元[/／每]人", p)
                if amt_m:
                    base_amount = int(amt_m.group(1))
                    break
            return {**entry, "rateType": "per_person_base",
                    "baseAmount": base_amount,
                    "coefficientTables": coeff_tables,
                    "description": substantive[0] if substantive else "",
                    "formula": f"保费 = {base_amount}元/人 × 系数 × 人数"}

        # --- disability_adjust: 劳务关系人员 ---
        if detected_type == "disability_adjust":
            return {**entry, "rateType": "disability_adjust",
                    "adjustCoeffs": dict(DISABILITY_ADJUST_COEFFS),
                    "description": substantive[0] if substantive else "",
                    "formula": "保费 = 每人保费 × 伤残调整系数 × 人数"}

        # --- formula_sum: 雇主法律责任/一次性伤残 ---
        if detected_type == "formula_sum":
            base_rate_factor = 1.0
            for pi in percentages:
                if "90" in str(pi["value"]):
                    base_rate_factor = 0.9
                elif "95" in str(pi["value"]):
                    base_rate_factor = 0.95
                elif "100" in str(pi["value"]):
                    base_rate_factor = 1.0
                elif "110" in str(pi["value"]):
                    base_rate_factor = 1.1
                elif "120" in str(pi["value"]):
                    base_rate_factor = 1.2
            if "90" in full_text and "一次性伤残" not in clause_name:
                base_rate_factor = 0.9
            elif "95" in full_text and "一次性伤残" not in clause_name:
                base_rate_factor = 0.95
            # 一次性伤残根据 ABCD 款
            if "一次性伤残" in clause_name:
                if "A款" in clause_name or "（A）" in clause_name:
                    base_rate_factor = 0.9
                elif "B款" in clause_name or "（B）" in clause_name:
                    base_rate_factor = 1.0
                elif "C款" in clause_name or "（C）" in clause_name:
                    base_rate_factor = 1.1
                elif "D款" in clause_name or "（D）" in clause_name:
                    base_rate_factor = 1.2
            return {**entry, "rateType": "formula_sum",
                    "baseRateFactor": base_rate_factor,
                    "coefficientTables": coeff_tables,
                    "description": substantive[0] if substantive else "",
                    "formula": f"保费 = Σ(每人限额 × 主险基准费率 × {base_rate_factor} × 人数 × 系数积)"}

        # --- deduction: 突发疾病除外 ---
        if detected_type == "deduction":
            deduct_pct = 5.0  # 默认5%
            for pi in percentages:
                deduct_pct = pi["value"]
                break
            return {**entry, "rateType": "deduction",
                    "deductPercent": deduct_pct,
                    "description": substantive[0] if substantive else "",
                    "formula": f"减收 = 主险保费 × {deduct_pct}%"}

        # --- no_calc: 月申报/员工自动承保 ---
        if detected_type == "no_calc":
            return {**entry, "rateType": "no_calc",
                    "description": substantive[0] if substantive else full_text[:300],
                    "formula": "本条款有计费说明但无需单独计算附加保费"}

        # ===== Step 4: 未匹配关键词，按内容分析 =====
        if coeff_tables:
            base_premium = {"description": "未找到基准保险费描述"}
            for p in paragraphs:
                pct_m = re.search(r"([\d.]+)\s*[%％]", p)
                if pct_m and ("基准保险费" in p or "主险保险费的" in p):
                    base_premium = {"description": p, "percentage": float(pct_m.group(1))}
                    break
                mult_m = re.search(r"主险保险费的\s*([\d.]+)\s*倍", p)
                if mult_m:
                    base_premium = {"description": p, "multiplier": float(mult_m.group(1))}
                    break
            formula = "保险费 = 基准保险费 × 各项费率调整系数的乘积"
            for p in paragraphs:
                if "保险费" in p and ("×" in p or "＝" in p or "乘积" in p):
                    formula = p
                    break
            return {**entry, "rateType": "table_coefficient", "basePremium": base_premium,
                    "coefficientTables": coeff_tables, "formula": formula,
                    "description": substantive[0] if substantive else ""}

        # 简单百分比
        for p in paragraphs:
            pct_m = re.search(r"([\d.]+)\s*[%％]", p)
            if pct_m:
                return {**entry, "rateType": "simple_percentage",
                        "percentage": float(pct_m.group(1)), "description": p}
            mult_m = re.search(r"主险保险费的\s*([\d.]+)\s*倍", p)
            if mult_m:
                mult = float(mult_m.group(1))
                return {**entry, "rateType": "simple_percentage",
                        "percentage": mult * 100, "multiplier": mult, "description": p}

        return {**entry, "rateType": "regulatory", "description": full_text[:200]}

    def _parse_tables(self, tables, full_text=""):
        """通用表格解析"""
        coeff_tables = []
        for raw_table in tables:
            if len(raw_table) < 2:
                continue
            header = raw_table[0]
            rows = []
            for i in range(1, len(raw_table)):
                if len(raw_table[i]) < 2:
                    continue
                param, coeff = raw_table[i][0], raw_table[i][1]
                if not param or not coeff:
                    continue
                rows.append({"parameter": param, "coefficient": coeff,
                             "parsedValue": self._parse_coeff_value(coeff)})
            if rows:
                coeff_tables.append({
                    "name": header[0] if header else "调整系数",
                    "headerRow": header,
                    "supportsInterpolation": "线性插值" in full_text or "插值" in full_text,
                    "rows": rows,
                })
        return coeff_tables

    def _parse_coeff_value(self, text):
        text = text.strip().replace("，", ",").replace("（", "(").replace("）", ")")
        range_m = re.match(r"[\[(]?\s*([\d.]+)\s*[,]\s*([\d.]+)\s*[\])]?", text)
        if range_m:
            return {"type": "range", "min": float(range_m.group(1)), "max": float(range_m.group(2)), "display": text}
        num_m = re.match(r"^([\d.]+)$", text)
        if num_m:
            return {"type": "fixed", "value": float(num_m.group(1)), "display": text}
        return {"type": "text", "display": text}

    # ---------- 搜索/筛选 ----------
    def _filter_entries(self, keyword=""):
        if not self.rate_data:
            return
        query = keyword.strip().lower()
        self.filtered_entries = [e for e in self.rate_data["entries"]
                                 if not query or query in e.get("clauseName", "").lower()
                                 or query in e.get("fullName", "").lower()]
        self._render_clause_list()

    def _render_clause_list(self):
        self.clause_list.clear()
        for entry in self.filtered_entries:
            rt = entry.get("rateType", "")
            type_info = ADDON_TYPES.get(rt, {"label": rt, "color": "#6b7280"})
            item = QListWidgetItem(f"{entry['clauseName']}  [{type_info['label']}]")
            item.setData(Qt.UserRole, entry)
            self.clause_list.addItem(item)

    def _on_clause_selected(self, row):
        if row < 0 or row >= len(self.filtered_entries):
            return
        self.selected_entry = self.filtered_entries[row]
        self.coeff_selections = {}
        self._render_detail()
        rt = self.selected_entry.get("rateType", "")
        type_label = ADDON_TYPES.get(rt, {}).get("label", rt)
        self._log(f"选中: {self.selected_entry['clauseName']} [{type_label}]")

    # ---------- 详情渲染 ----------
    def _render_detail(self):
        while self.detail_layout.count():
            item = self.detail_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        entry = self.selected_entry
        if not entry:
            placeholder = QLabel("📊 请导入费率方案文件夹，然后从左侧选择条款")
            placeholder.setAlignment(Qt.AlignCenter)
            placeholder.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 16px; padding: 60px;")
            self.detail_layout.addWidget(placeholder)
            self.detail_layout.addStretch()
            return

        rate_type = entry.get("rateType", "")
        type_info = ADDON_TYPES.get(rate_type, {"label": rate_type, "color": "#6b7280"})

        # 条款名称 + 类型标签
        header_row = QHBoxLayout()
        name_label = QLabel(entry["clauseName"])
        name_label.setStyleSheet(f"font-weight: 700; font-size: 16px; color: {AnthropicColors.TEXT_PRIMARY};")
        header_row.addWidget(name_label)
        type_badge = QLabel(type_info["label"])
        type_badge.setStyleSheet(f"background: {type_info['color']}; color: white; padding: 2px 10px; "
                                 f"border-radius: 10px; font-size: 11px; font-weight: 600;")
        type_badge.setFixedHeight(22)
        header_row.addWidget(type_badge)
        header_row.addStretch()
        header_widget = QWidget()
        header_widget.setLayout(header_row)
        self.detail_layout.addWidget(header_widget)

        # 描述
        if entry.get("description"):
            desc = QLabel(entry["description"])
            desc.setWordWrap(True)
            desc.setStyleSheet(f"font-size: 12px; color: {AnthropicColors.TEXT_SECONDARY};")
            self.detail_layout.addWidget(desc)

        # 公式说明
        if entry.get("formula"):
            formula_label = QLabel(f"📐 {entry['formula']}")
            formula_label.setWordWrap(True)
            formula_label.setStyleSheet(f"padding: 10px; background: #eff6ff; border-radius: 8px; font-size: 12px; color: #1e40af;")
            self.detail_layout.addWidget(formula_label)

        # ===== 按类型渲染专属输入区 =====

        if rate_type == "regulatory":
            reg_label = QLabel("📋 规范类条款\n本条款不涉及保险费的调整")
            reg_label.setAlignment(Qt.AlignCenter)
            reg_label.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 14px; padding: 30px;")
            self.detail_layout.addWidget(reg_label)
            self.detail_layout.addStretch()
            return

        if rate_type == "no_calc":
            no_label = QLabel("📋 本条款有计费规则说明，但无需单独计算附加保费\n其影响已包含在主险或其他条款中")
            no_label.setAlignment(Qt.AlignCenter)
            no_label.setWordWrap(True)
            no_label.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 13px; padding: 30px;")
            self.detail_layout.addWidget(no_label)
            self._add_manual_section()
            self.detail_layout.addStretch()
            return

        if rate_type == "simple_percentage":
            pct = entry.get("percentage", 0)
            mult = entry.get("multiplier")
            if mult:
                info = QLabel(f"费率: 主险保费 × {mult}")
            else:
                info = QLabel(f"费率: 主险保费 × {pct}%")
            info.setStyleSheet(f"padding: 12px; background: #eff6ff; border-radius: 8px; font-size: 13px;")
            self.detail_layout.addWidget(info)

        elif rate_type == "modifier_coeff":
            hint = QLabel("⚡ 优先级最高：此条款调整主险保费系数\n选择免赔天数对应的调整系数后计算")
            hint.setWordWrap(True)
            hint.setStyleSheet(f"padding: 10px; background: #faf5ff; border: 1px solid #c084fc; border-radius: 8px; font-size: 12px; color: #7c3aed;")
            self.detail_layout.addWidget(hint)
            for ti, table in enumerate(entry.get("coefficientTables", [])):
                self._render_addon_coeff_table(table, ti)

        elif rate_type == "sudden_death":
            base_pct = entry.get("basePercent", 6.6)
            hint = QLabel(f"基准费率: 每人限额 × {base_pct}%\n"
                          "附加险限额 < 主险限额 → 减收\n"
                          "附加险限额 > 主险限额 → 加收")
            hint.setWordWrap(True)
            hint.setStyleSheet(f"padding: 10px; background: #fef2f2; border: 1px solid #fca5a5; border-radius: 8px; font-size: 12px;")
            self.detail_layout.addWidget(hint)
            # 限额输入
            grid = QGridLayout()
            grid.addWidget(QLabel("附加险每人限额(万元):"), 0, 0)
            self.addon_limit_input = QDoubleSpinBox()
            self.addon_limit_input.setRange(0, 9999)
            self.addon_limit_input.setDecimals(1)
            self.addon_limit_input.setSuffix(" 万元")
            grid.addWidget(self.addon_limit_input, 0, 1)
            grid.addWidget(QLabel("主险每人限额(万元):"), 1, 0)
            self.main_limit_display = QDoubleSpinBox()
            self.main_limit_display.setRange(0, 9999)
            self.main_limit_display.setDecimals(1)
            self.main_limit_display.setSuffix(" 万元")
            if self.full_main_data:
                limit_wan = self.full_main_data.get("perPersonLimit", 0) / 10000
                self.main_limit_display.setValue(limit_wan)
            grid.addWidget(self.main_limit_display, 1, 1)
            grid_w = QWidget()
            grid_w.setLayout(grid)
            self.detail_layout.addWidget(grid_w)
            for ti, table in enumerate(entry.get("coefficientTables", [])):
                self._render_addon_coeff_table(table, ti)

        elif rate_type == "per_person_rate":
            rate_info = entry.get("rateInfo", {})
            hint_text = "费率说明:\n"
            if rate_info.get("with_injury_insurance"):
                hint_text += f"  已购买工伤保险: {rate_info['with_injury_insurance']}%\n"
            if rate_info.get("without_injury_insurance"):
                hint_text += f"  未购买工伤保险: {rate_info['without_injury_insurance']}%\n"
            if rate_info.get("default"):
                hint_text += f"  默认费率: {rate_info['default']}%\n"
            hint = QLabel(hint_text.strip())
            hint.setWordWrap(True)
            hint.setStyleSheet(f"padding: 10px; background: #fffbeb; border: 1px solid #fbbf24; border-radius: 8px; font-size: 12px;")
            self.detail_layout.addWidget(hint)
            # 工伤保险状态选择
            self.injury_insurance_combo = QComboBox()
            self.injury_insurance_combo.addItem("已购买工伤保险", "with")
            self.injury_insurance_combo.addItem("未购买工伤保险", "without")
            self.detail_layout.addWidget(self.injury_insurance_combo)
            # 独立人数输入（不一定全员）
            default_count = self.full_main_data.get("employeeCount", 1) if self.full_main_data else 1
            count_row = QHBoxLayout()
            count_row.addWidget(QLabel("适用人数:"))
            self.addon_count_input = QSpinBox()
            self.addon_count_input.setRange(1, 999999)
            self.addon_count_input.setValue(default_count)
            self.addon_count_input.setSuffix("")
            count_row.addWidget(self.addon_count_input)
            count_row.addStretch()
            count_w = QWidget()
            count_w.setLayout(count_row)
            self.detail_layout.addWidget(count_w)
            for ti, table in enumerate(entry.get("coefficientTables", [])):
                self._render_addon_coeff_table(table, ti)

        elif rate_type == "property_loss":
            bp = entry.get("basePremium", 20)
            br = entry.get("baseRatePercent", 1.5)
            hint = QLabel(f"公式: 保费 = (基本保险费 + 每次事故每人赔偿限额 × 基准费率) × 系数积 × 承保人数\n基本保险费: {bp}元，基准费率: {br}%")
            hint.setWordWrap(True)
            hint.setStyleSheet(f"padding: 10px; background: #f0fdfa; border: 1px solid #5eead4; border-radius: 8px; font-size: 12px;")
            self.detail_layout.addWidget(hint)
            grid = QGridLayout()
            grid.addWidget(QLabel("每次事故每人赔偿限额(万元):"), 0, 0)
            self.property_limit_input = QDoubleSpinBox()
            self.property_limit_input.setRange(0, 9999)
            self.property_limit_input.setDecimals(1)
            self.property_limit_input.setSuffix(" 万元")
            grid.addWidget(self.property_limit_input, 0, 1)
            default_count = self.full_main_data.get("employeeCount", 1) if self.full_main_data else 1
            grid.addWidget(QLabel("适用人数:"), 1, 0)
            self.addon_count_input = QSpinBox()
            self.addon_count_input.setRange(1, 999999)
            self.addon_count_input.setValue(default_count)
            self.addon_count_input.setSuffix("")
            grid.addWidget(self.addon_count_input, 1, 1)
            grid_w = QWidget()
            grid_w.setLayout(grid)
            self.detail_layout.addWidget(grid_w)
            for ti, table in enumerate(entry.get("coefficientTables", [])):
                self._render_addon_coeff_table(table, ti)

        elif rate_type == "per_person_base":
            base_amt = entry.get("baseAmount", 300)
            hint = QLabel(f"基准: {base_amt}元/人，乘以调整系数后 × 人数")
            hint.setStyleSheet(f"padding: 10px; background: #ecfeff; border: 1px solid #67e8f9; border-radius: 8px; font-size: 12px;")
            self.detail_layout.addWidget(hint)
            for ti, table in enumerate(entry.get("coefficientTables", [])):
                self._render_addon_coeff_table(table, ti)

        elif rate_type == "disability_adjust":
            coeffs = entry.get("adjustCoeffs", DISABILITY_ADJUST_COEFFS)
            hint = QLabel(f"根据主险选择的伤残赔偿附表调整:\n"
                          f"  附表1: ×{coeffs.get('table1', 0.995)}\n"
                          f"  附表2: ×{coeffs.get('table2', 1.072)}\n"
                          f"  附表3: ×{coeffs.get('table3', 0.919)}")
            hint.setStyleSheet(f"padding: 10px; background: #ecfdf5; border: 1px solid #6ee7b7; border-radius: 8px; font-size: 12px;")
            self.detail_layout.addWidget(hint)
            if self.full_main_data:
                dis_table = self.full_main_data.get("disabilityTable", "none")
                if dis_table != "none":
                    auto_label = QLabel(f"🔗 已自动识别主险伤残附表: {dis_table}")
                    auto_label.setStyleSheet(f"color: #10b981; font-size: 12px; font-weight: 600;")
                    self.detail_layout.addWidget(auto_label)
            # 独立人数输入（不一定全员）
            default_count = self.full_main_data.get("employeeCount", 1) if self.full_main_data else 1
            count_row = QHBoxLayout()
            count_row.addWidget(QLabel("适用人数:"))
            self.addon_count_input = QSpinBox()
            self.addon_count_input.setRange(1, 999999)
            self.addon_count_input.setValue(default_count)
            self.addon_count_input.setSuffix("")
            count_row.addWidget(self.addon_count_input)
            count_row.addStretch()
            count_w = QWidget()
            count_w.setLayout(count_row)
            self.detail_layout.addWidget(count_w)

        elif rate_type == "formula_sum":
            factor = entry.get("baseRateFactor", 1.0)
            hint = QLabel(f"求和公式: Σ(每人限额 × 主险基准费率 × {factor} × 人数 × 系数积)\n"
                          f"基准费率来源: 主险基准费率 × {factor}")
            hint.setWordWrap(True)
            hint.setStyleSheet(f"padding: 10px; background: #eff6ff; border: 1px solid #93c5fd; border-radius: 8px; font-size: 12px;")
            self.detail_layout.addWidget(hint)
            for ti, table in enumerate(entry.get("coefficientTables", [])):
                self._render_addon_coeff_table(table, ti)

        elif rate_type == "deduction":
            deduct_pct = entry.get("deductPercent", 5.0)
            hint = QLabel(f"减收: 主险保费 × {deduct_pct}%\n将从主险保费中扣减此金额")
            hint.setStyleSheet(f"padding: 10px; background: #fdf2f8; border: 1px solid #f9a8d4; border-radius: 8px; font-size: 12px;")
            self.detail_layout.addWidget(hint)

        elif rate_type == "table_coefficient":
            if entry.get("basePremium"):
                bp = entry["basePremium"]
                bp_label = QLabel(f"基准保险费: {bp.get('description', '')}")
                bp_label.setWordWrap(True)
                bp_label.setStyleSheet(f"padding: 10px; background: #eff6ff; border-radius: 8px; font-size: 12px;")
                self.detail_layout.addWidget(bp_label)
            for ti, table in enumerate(entry.get("coefficientTables", [])):
                self._render_addon_coeff_table(table, ti)

        # 计算按钮（非 regulatory / no_calc）
        if rate_type not in ("regulatory", "no_calc"):
            calc_btn = make_accent_button("🧮 计算附加险保费")
            calc_btn.clicked.connect(self._calculate)
            self.detail_layout.addWidget(calc_btn)

        # 结果区
        self.addon_result_label = QLabel("")
        self.addon_result_label.setWordWrap(True)
        self.addon_result_label.setStyleSheet(f"font-size: 13px; padding: 10px;")
        self.addon_result_label.hide()
        self.detail_layout.addWidget(self.addon_result_label)

        # 核保经验计费
        self._add_manual_section()
        self.detail_layout.addStretch()

    def _add_manual_section(self):
        """添加核保经验计费区域"""
        manual_card = GlassCard()
        manual_layout = QVBoxLayout(manual_card)
        manual_layout.setContentsMargins(12, 10, 12, 10)
        manual_layout.addWidget(QLabel("✏️ 核保经验计费"))
        manual_hint = QLabel("手动输入附加险保费（覆盖公式计算结果）")
        manual_hint.setWordWrap(True)
        manual_hint.setStyleSheet(f"font-size: 11px; color: {AnthropicColors.TEXT_TERTIARY};")
        manual_layout.addWidget(manual_hint)
        manual_row = QHBoxLayout()
        self.manual_input = QDoubleSpinBox()
        self.manual_input.setRange(-999999999, 999999999)
        self.manual_input.setDecimals(2)
        self.manual_input.setSuffix(" 元")
        manual_row.addWidget(self.manual_input, 1)
        manual_btn = QPushButton("确认计入")
        manual_btn.setCursor(Qt.PointingHandCursor)
        manual_btn.setStyleSheet(f"QPushButton {{ background: #f59e0b; color: white; border: none; border-radius: 6px; padding: 8px 14px; }}")
        manual_btn.clicked.connect(self._add_manual_premium)
        manual_row.addWidget(manual_btn)
        manual_layout.addLayout(manual_row)
        self.detail_layout.addWidget(manual_card)

    def _render_addon_coeff_table(self, table, table_idx):
        card = GlassCard()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)
        title = QLabel(table.get("name", "调整系数"))
        title.setStyleSheet(f"font-weight: 600; font-size: 13px;")
        layout.addWidget(title)
        if table.get("supportsInterpolation"):
            interp = QLabel("支持线性插值")
            interp.setStyleSheet(f"font-size: 11px; color: #3b82f6;")
            layout.addWidget(interp)
        for ri, row in enumerate(table.get("rows", [])):
            sel = self.coeff_selections.get(table_idx)
            is_selected = sel and sel.get("rowIdx") == ri
            btn = QPushButton(f"{row['parameter']}    {row['coefficient']}")
            btn.setCursor(Qt.PointingHandCursor)
            bg = AnthropicColors.ACCENT if is_selected else AnthropicColors.BG_PRIMARY
            fg = AnthropicColors.TEXT_LIGHT if is_selected else AnthropicColors.TEXT_PRIMARY
            btn.setStyleSheet(f"""
                QPushButton {{ background: {bg}; color: {fg}; border: 1px solid {AnthropicColors.BORDER};
                    border-radius: 6px; padding: 5px 10px; font-size: 12px; text-align: left; }}
                QPushButton:hover {{ border-color: {AnthropicColors.ACCENT}; }}
            """)
            btn.clicked.connect(lambda checked, ti=table_idx, r=ri: self._select_addon_coeff_row(ti, r))
            layout.addWidget(btn)
        sel = self.coeff_selections.get(table_idx)
        if sel and sel.get("parsedValue", {}).get("type") == "range":
            pv = sel["parsedValue"]
            current_val = sel.get("value", pv["min"])
            slider_layout = QHBoxLayout()
            slider_label = QLabel(f"{current_val:.4f}")
            slider_label.setStyleSheet(f"font-weight: 600; color: {AnthropicColors.ACCENT};")
            slider_layout.addWidget(QLabel("精确系数:"))
            slider_layout.addWidget(slider_label)
            slider = QSlider(Qt.Horizontal)
            slider.setMinimum(int(pv["min"] * 1000))
            slider.setMaximum(int(pv["max"] * 1000))
            slider.setValue(int(current_val * 1000))
            ti_ref = table_idx
            slider.valueChanged.connect(lambda v, ti=ti_ref, lbl=slider_label: self._on_addon_slider_change(ti, v, lbl))
            slider_layout.addWidget(slider, 1)
            layout.addLayout(slider_layout)
        self.detail_layout.addWidget(card)

    def _select_addon_coeff_row(self, table_idx, row_idx):
        entry = self.selected_entry
        if not entry or not entry.get("coefficientTables"):
            return
        table = entry["coefficientTables"][table_idx]
        row = table["rows"][row_idx]
        pv = row["parsedValue"]
        value = pv.get("value", pv.get("min", 1.0)) if pv["type"] != "text" else 1.0
        self.coeff_selections[table_idx] = {
            "rowIdx": row_idx, "value": value, "parsedValue": pv,
            "parameter": row["parameter"], "coefficient": row["coefficient"],
        }
        self._render_detail()

    def _on_addon_slider_change(self, table_idx, int_value, label_widget):
        value = int_value / 1000.0
        if table_idx in self.coeff_selections:
            self.coeff_selections[table_idx]["value"] = value
        label_widget.setText(f"{value:.4f}")

    def _get_coeff_product(self, entry):
        """计算所有已选系数表的系数乘积"""
        product = 1.0
        details = []
        for ti in range(len(entry.get("coefficientTables", []))):
            sel = self.coeff_selections.get(ti)
            if not sel:
                table_name = entry["coefficientTables"][ti]["name"]
                raise ValueError(f"请选择「{table_name}」的系数值")
            product *= sel["value"]
            details.append({"table": entry["coefficientTables"][ti]["name"],
                            "parameter": sel["parameter"], "value": sel["value"]})
        return product, details

    # ---------- 计算引擎 (重构版) ----------
    def _calculate(self):
        entry = self.selected_entry
        if not entry:
            self._log("请先选择费率方案", "warn")
            return
        self.main_premium = self.main_premium_input.value()
        rate_type = entry.get("rateType", "")
        if rate_type in ("regulatory", "no_calc"):
            return
        if self.main_premium <= 0 and rate_type not in ("per_person_base", "property_loss"):
            self._log("请输入有效的主险保费", "warn")
            return
        try:
            calc_method = {
                "simple_percentage": self._calc_simple,
                "modifier_coeff": self._calc_modifier_coeff,
                "sudden_death": self._calc_sudden_death,
                "per_person_rate": self._calc_per_person_rate,
                "per_person_base": self._calc_per_person_base,
                "disability_adjust": self._calc_disability_adjust,
                "property_loss": self._calc_property_loss,
                "formula_sum": self._calc_formula_sum,
                "deduction": self._calc_deduction,
                "table_coefficient": self._calc_table,
            }.get(rate_type)
            if not calc_method:
                self._log(f"未知计算类型: {rate_type}", "error")
                return
            result = calc_method(entry)
        except Exception as e:
            self._log(f"计算错误: {e}", "error")
            return
        premium_text = fmt_currency(result["premium"])
        if result["premium"] < 0:
            premium_text = f"-{fmt_currency(abs(result['premium']))}"
        self.addon_result_label.setText(f"✅ {result['formulaDisplay']}\n保费: {premium_text}")
        bg_color = "#fef2f2" if result["premium"] < 0 else "#ecfdf5"
        fg_color = "#991b1b" if result["premium"] < 0 else "#065f46"
        self.addon_result_label.setStyleSheet(f"font-size: 13px; padding: 12px; background: {bg_color}; border-radius: 8px; color: {fg_color};")
        self.addon_result_label.show()
        self._add_premium_item(entry["clauseName"], result["premium"], result["formulaDisplay"])
        self._log(f"计算完成 [{ADDON_TYPES.get(rate_type, {}).get('label', '')}]: {premium_text}", "success")

    def _calc_simple(self, entry):
        pct = entry.get("percentage", 0)
        mult = entry.get("multiplier")
        if mult:
            rate = mult
            premium = self.main_premium * rate
            formula_str = f"{fmt_currency(self.main_premium)} × {mult} = {fmt_currency(premium)}"
        else:
            rate = pct / 100
            premium = self.main_premium * rate
            formula_str = f"{fmt_currency(self.main_premium)} × {pct}% = {fmt_currency(premium)}"
        return {"type": "simple_percentage", "premium": premium, "formulaDisplay": formula_str}

    def _calc_modifier_coeff(self, entry):
        """误工费: 调整主险保费系数"""
        product, details = self._get_coeff_product(entry)
        adjusted = self.main_premium * product
        diff = adjusted - self.main_premium
        coeff_str = " × ".join(f"{d['value']:.4f}" for d in details)
        formula_str = (f"调整后主险保费 = {fmt_currency(self.main_premium)} × {coeff_str} = {fmt_currency(adjusted)}\n"
                       f"差额: {'+' if diff >= 0 else ''}{fmt_currency(diff)}")
        return {"type": "modifier_coeff", "premium": diff, "formulaDisplay": formula_str}

    def _calc_sudden_death(self, entry):
        """突发疾病身故: 6.6% × (限额差异)"""
        base_pct = entry.get("basePercent", 6.6) / 100
        addon_limit = getattr(self, 'addon_limit_input', None)
        main_limit_w = getattr(self, 'main_limit_display', None)
        addon_val = (addon_limit.value() * 10000) if addon_limit else 0
        main_val = (main_limit_w.value() * 10000) if main_limit_w else 0
        if not main_val:
            if self.full_main_data:
                main_val = self.full_main_data.get("perPersonLimit", 0)
            if not main_val:
                raise ValueError("请输入主险每人限额")
        if not addon_val:
            raise ValueError("请输入附加险每人限额")
        count = self.full_main_data.get("employeeCount", 1) if self.full_main_data else 1
        # 系数
        product = 1.0
        coeff_str = ""
        if entry.get("coefficientTables"):
            product, details = self._get_coeff_product(entry)
            coeff_str = " × " + " × ".join(f"{d['value']:.4f}" for d in details)
        if addon_val < main_val:
            # 减收
            ratio = 1 - addon_val / main_val
            premium = -(main_val * base_pct * count * product * ratio)
            formula_str = (f"减收: {fmt_currency(main_val)} × {entry.get('basePercent', 6.6)}% × {count}人{coeff_str} × "
                           f"(1 - {addon_val}/{main_val}) = {fmt_currency(abs(premium))}")
        elif addon_val > main_val:
            ratio = addon_val / main_val - 1
            premium = main_val * base_pct * count * product * ratio
            formula_str = (f"加收: {fmt_currency(main_val)} × {entry.get('basePercent', 6.6)}% × {count}人{coeff_str} × "
                           f"({addon_val}/{main_val} - 1) = {fmt_currency(premium)}")
        else:
            premium = 0
            formula_str = "附加险限额 = 主险限额，不调整"
        return {"type": "sudden_death", "premium": premium, "formulaDisplay": formula_str}

    def _calc_per_person_rate(self, entry):
        """工伤补充/特定人员: 每人保费 × 费率% × 人数"""
        rate_info = entry.get("rateInfo", {})
        combo = getattr(self, 'injury_insurance_combo', None)
        if combo and combo.currentData() == "with":
            pct = rate_info.get("with_injury_insurance", rate_info.get("default", 0))
        else:
            pct = rate_info.get("without_injury_insurance", rate_info.get("default", 0))
        rate = pct / 100
        count_input = getattr(self, 'addon_count_input', None)
        count = count_input.value() if count_input else (self.full_main_data.get("employeeCount", 1) if self.full_main_data else 1)
        product = 1.0
        coeff_str = ""
        if entry.get("coefficientTables"):
            product, details = self._get_coeff_product(entry)
            coeff_str = " × " + " × ".join(f"{d['value']:.4f}" for d in details)
        premium = self.per_person_premium * rate * count * product
        status = "已购买" if (combo and combo.currentData() == "with") else "未购买"
        formula_str = (f"{fmt_currency(self.per_person_premium)} × {pct}%({status}工伤) × {count}人{coeff_str} = "
                       f"{fmt_currency(premium)}")
        return {"type": "per_person_rate", "premium": premium, "formulaDisplay": formula_str}

    def _calc_per_person_base(self, entry):
        """药品服务: 定额/人 × 系数 × 人数"""
        base_amt = entry.get("baseAmount", 300)
        count = self.full_main_data.get("employeeCount", 1) if self.full_main_data else 1
        product = 1.0
        coeff_str = ""
        if entry.get("coefficientTables"):
            product, details = self._get_coeff_product(entry)
            coeff_str = " × " + " × ".join(f"{d['value']:.4f}" for d in details)
        premium = base_amt * product * count
        formula_str = f"{base_amt}元/人 × {count}人{coeff_str} = {fmt_currency(premium)}"
        return {"type": "per_person_base", "premium": premium, "formulaDisplay": formula_str}

    def _calc_disability_adjust(self, entry):
        """劳务关系人员: 每人保费 × 伤残调整系数 × 人数"""
        coeffs = entry.get("adjustCoeffs", DISABILITY_ADJUST_COEFFS)
        dis_table = "none"
        if self.full_main_data:
            dis_table = self.full_main_data.get("disabilityTable", "none")
        if dis_table == "none" or dis_table not in coeffs:
            raise ValueError("需要主险选择伤残赔偿附表（请先在主险计算Tab选择伤残附表并传入）")
        coeff = coeffs[dis_table]
        count_input = getattr(self, 'addon_count_input', None)
        count = count_input.value() if count_input else (self.full_main_data.get("employeeCount", 1) if self.full_main_data else 1)
        premium = self.per_person_premium * coeff * count
        formula_str = (f"{fmt_currency(self.per_person_premium)} × {coeff}({dis_table}) × {count}人 = "
                       f"{fmt_currency(premium)}")
        return {"type": "disability_adjust", "premium": premium, "formulaDisplay": formula_str}

    def _calc_property_loss(self, entry):
        """员工个人特定财产损失: (基本保险费 + 限额×费率%) × 系数积 × 承保人数"""
        base_premium = entry.get("basePremium", 20)
        base_rate_pct = entry.get("baseRatePercent", 1.5)
        limit_input = getattr(self, 'property_limit_input', None)
        limit_val = limit_input.value() if limit_input else 0
        count_input = getattr(self, 'addon_count_input', None)
        count = count_input.value() if count_input else (self.full_main_data.get("employeeCount", 1) if self.full_main_data else 1)
        if limit_val <= 0:
            raise ValueError("请输入每次事故每人赔偿限额")
        limit_yuan = limit_val * 10000
        rate_part = limit_yuan * base_rate_pct / 100
        per_person = base_premium + rate_part
        product = 1.0
        coeff_str = ""
        if entry.get("coefficientTables"):
            product, details = self._get_coeff_product(entry)
            coeff_str = " × " + " × ".join(f"{d['value']:.4f}" for d in details)
        premium = per_person * product * count
        formula_str = (f"({base_premium}元 + {limit_val}万×{base_rate_pct}%) × 系数{coeff_str.lstrip(' × ')} × {count}人 = "
                       f"({base_premium} + {fmt_currency(rate_part)}) × {product:.4f} × {count} = "
                       f"{fmt_currency(premium)}")
        return {"type": "property_loss", "premium": premium, "formulaDisplay": formula_str}

    def _calc_formula_sum(self, entry):
        """雇主法律责任/一次性伤残: Σ公式"""
        factor = entry.get("baseRateFactor", 1.0)
        if not self.full_main_data:
            raise ValueError("需要主险完整数据（请先在主险计算Tab计算并传入附加险）")
        base_rate = self.full_main_data.get("baseRate", 0)
        per_limit = self.full_main_data.get("perPersonLimit", 0)
        count = self.full_main_data.get("employeeCount", 1)
        main_coeff = self.full_main_data.get("coeffProduct", 1.0)
        if not base_rate:
            raise ValueError("主险基准费率为0，无法计算")
        if not per_limit:
            raise ValueError("主险每人限额为0（仅固定限额计费方式支持此类附加险）")
        # 附加险自身的系数
        addon_product = 1.0
        coeff_str = ""
        if entry.get("coefficientTables"):
            addon_product, details = self._get_coeff_product(entry)
            coeff_str = " × " + " × ".join(f"{d['value']:.4f}" for d in details)
        adjusted_rate = base_rate * factor
        premium = per_limit * adjusted_rate * count * main_coeff * addon_product
        formula_str = (f"{fmt_currency(per_limit)} × {base_rate:.6f} × {factor} × {count}人 × "
                       f"主险系数{main_coeff:.4f}{coeff_str} = {fmt_currency(premium)}")
        return {"type": "formula_sum", "premium": premium, "formulaDisplay": formula_str}

    def _calc_deduction(self, entry):
        """突发疾病除外: 减收主险保费的X%"""
        deduct_pct = entry.get("deductPercent", 5.0)
        premium = -(self.main_premium * deduct_pct / 100)
        formula_str = f"减收: {fmt_currency(self.main_premium)} × {deduct_pct}% = {fmt_currency(abs(premium))}"
        return {"type": "deduction", "premium": premium, "formulaDisplay": formula_str}

    def _calc_table(self, entry):
        """通用系数表计算"""
        base_premium = self.main_premium
        bp = entry.get("basePremium", {})
        if bp.get("multiplier"):
            base_premium = self.main_premium * bp["multiplier"]
        elif bp.get("percentage"):
            base_premium = self.main_premium * (bp["percentage"] / 100)
        product, details = self._get_coeff_product(entry)
        premium = base_premium * product
        base_str = (f"{fmt_currency(self.main_premium)} × {bp['multiplier']}" if bp.get("multiplier")
                    else f"{fmt_currency(self.main_premium)} × {bp['percentage']}%" if bp.get("percentage")
                    else fmt_currency(self.main_premium))
        coeff_str = " × ".join(f"{c['value']:.4f}" for c in details)
        return {"type": "table_coefficient", "premium": premium,
                "formulaDisplay": f"基准 {base_str} = {fmt_currency(base_premium)} × 系数 ({coeff_str}) = {fmt_currency(premium)}"}

    # ---------- 保费汇总管理 ----------
    def _add_premium_item(self, clause_name, premium, formula):
        existing_idx = next((i for i, item in enumerate(self.premium_items) if item["clauseName"] == clause_name), -1)
        new_item = {"id": id(formula), "clauseName": clause_name, "premium": premium, "formula": formula}
        if existing_idx >= 0:
            self.premium_items[existing_idx] = new_item
            self._log(f"已更新: {clause_name} 的保费")
        else:
            self.premium_items.append(new_item)
        self._render_premium_summary()

    def _remove_premium_item(self, item_id):
        self.premium_items = [item for item in self.premium_items if item["id"] != item_id]
        self._render_premium_summary()
        self._log("已移除一项附加险保费")

    def _render_premium_summary(self):
        while self.premium_list_layout.count():
            item = self.premium_list_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.premium_items:
            empty = QLabel("计算附加险保费后\n将自动添加到此列表")
            empty.setAlignment(Qt.AlignCenter)
            empty.setStyleSheet(f"color: {AnthropicColors.TEXT_TERTIARY}; font-size: 12px; padding: 20px;")
            self.premium_list_layout.addWidget(empty)
            self.addon_total_label.setText("附加险合计: ¥0.00")
            self.annual_total_label.setText("保单预估年保费: ¥0.00")
            return

        addon_total = 0.0
        for item in self.premium_items:
            addon_total += item["premium"]
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(8, 4, 8, 4)
            info_layout = QVBoxLayout()
            name_lbl = QLabel(item["clauseName"])
            name_lbl.setStyleSheet(f"font-size: 12px; font-weight: 500;")
            info_layout.addWidget(name_lbl)
            amount_color = "#ef4444" if item["premium"] < 0 else AnthropicColors.ACCENT
            amount_text = f"-{fmt_currency(abs(item['premium']))}" if item["premium"] < 0 else fmt_currency(item["premium"])
            amount_lbl = QLabel(amount_text)
            amount_lbl.setStyleSheet(f"font-size: 11px; color: {amount_color};")
            info_layout.addWidget(amount_lbl)
            row_layout.addLayout(info_layout, 1)
            del_btn = QPushButton("×")
            del_btn.setFixedSize(24, 24)
            del_btn.setCursor(Qt.PointingHandCursor)
            del_btn.setStyleSheet(f"QPushButton {{ background: transparent; color: {AnthropicColors.TEXT_TERTIARY}; border: none; font-size: 16px; }} QPushButton:hover {{ color: #ef4444; }}")
            item_id = item["id"]
            del_btn.clicked.connect(lambda checked, iid=item_id: self._remove_premium_item(iid))
            row_layout.addWidget(del_btn)
            row_widget.setStyleSheet(f"background: {AnthropicColors.BG_CARD}; border-radius: 6px;")
            self.premium_list_layout.addWidget(row_widget)

        addon_total_text = f"-{fmt_currency(abs(addon_total))}" if addon_total < 0 else fmt_currency(addon_total)
        self.addon_total_label.setText(f"附加险合计: {addon_total_text}")
        main_val = self.main_premium_input.value()
        annual_total = main_val + addon_total
        annual_text = f"-{fmt_currency(abs(annual_total))}" if annual_total < 0 else fmt_currency(annual_total)
        self.annual_total_label.setText(f"保单预估年保费: {annual_text}")
        self._calc_short_term()

    # ---------- 短期保费计算 ----------
    def _calc_short_term(self):
        start = self.start_date.date()
        end = self.end_date.date()
        if end <= start:
            self.short_term_result.setText("终止日须晚于起保日")
            self.short_term_result.setStyleSheet(f"color: #ef4444; font-size: 11px;")
            return
        insurance_days = start.daysTo(end)
        start_year = start.year()
        year_days = 366 if is_leap_year(start_year) else 365
        main_val = self.main_premium_input.value()
        addon_total = sum(item["premium"] for item in self.premium_items)
        annual_total = main_val + addon_total
        short_premium = annual_total / year_days * insurance_days
        leap_text = f"（闰年 {year_days}天）" if year_days == 366 else f"（平年 {year_days}天）"
        short_text = f"-{fmt_currency(abs(short_premium))}" if short_premium < 0 else fmt_currency(short_premium)
        self.short_term_result.setText(
            f"保险天数: {insurance_days} 天 · {start_year}年{leap_text}\n"
            f"{fmt_currency(abs(annual_total))} ÷ {year_days} × {insurance_days}\n"
            f"短期保费: {short_text}")
        self.short_term_result.setStyleSheet(f"font-size: 12px; color: {AnthropicColors.ACCENT}; font-weight: 600;")

    # ---------- 核保经验计费 ----------
    def _add_manual_premium(self):
        if not self.selected_entry:
            self._log("请先选择条款", "warn")
            return
        manual_val = self.manual_input.value()
        self._add_premium_item(self.selected_entry["clauseName"], manual_val, f"核保经验计费: {fmt_currency(manual_val)}")
        self._log(f"核保经验计费: {self.selected_entry['clauseName']} → {fmt_currency(manual_val)}", "success")

    # ---------- 询价导入 ----------
    def _handle_inquiry_import(self):
        if not self.rate_data or not self.rate_data.get("entries"):
            self._log("请先加载费率方案数据", "warn")
            return
        file_path, _ = QFileDialog.getOpenFileName(self, "导入询价文件", "", "询价文件 (*.xlsx *.docx);;Excel (*.xlsx);;Word (*.docx)")
        if not file_path:
            return
        if file_path.endswith(".xlsx"):
            self._parse_inquiry_excel(file_path)
        elif file_path.endswith(".docx"):
            self._parse_inquiry_docx(file_path)

    def _parse_inquiry_excel(self, file_path):
        try:
            from openpyxl import load_workbook
        except ImportError:
            self._log("openpyxl 未安装，请运行: pip install openpyxl", "error")
            return
        try:
            wb = load_workbook(file_path, read_only=True)
            ws = wb.active
            clause_names = []
            for row in ws.iter_rows():
                if len(row) > 5 and row[5].value:
                    val = str(row[5].value).strip()
                    if len(val) > 2 and val not in ("附加条款", "条款名称") and "附加" in val:
                        clause_names.append(val)
            if not clause_names:
                for row in ws.iter_rows():
                    for cell in row:
                        val = str(cell.value or "").strip()
                        if len(val) > 4 and "附加" in val and val not in clause_names:
                            clause_names.append(val)
            wb.close()
            self._match_inquiry_clauses(clause_names, os.path.basename(file_path))
        except Exception as e:
            self._log(f"Excel 解析失败: {e}", "error")

    def _parse_inquiry_docx(self, file_path):
        try:
            from docx import Document
        except ImportError:
            self._log("python-docx 未安装", "error")
            return
        try:
            doc = Document(file_path)
            clause_names = []
            for p in doc.paragraphs:
                for run in p.runs:
                    if run.font.color and run.font.color.rgb:
                        rgb = str(run.font.color.rgb)
                        if rgb.startswith("0000") or rgb.lower() in ("0000ff", "0000cd", "0000ee"):
                            text = run.text.strip()
                            if len(text) > 2:
                                clause_names.append(text)
            if not clause_names:
                for p in doc.paragraphs:
                    text = p.text.strip()
                    if "附加" in text and 4 < len(text) < 60:
                        clause_names.append(text)
            self._match_inquiry_clauses(clause_names, os.path.basename(file_path))
        except Exception as e:
            self._log(f"Docx 解析失败: {e}", "error")

    def _match_clause_name(self, imported_name, entries):
        normalized = imported_name.replace(" ", "").replace("（", "(").replace("）", ")")
        for e in entries:
            entry_norm = e["clauseName"].replace(" ", "").replace("（", "(").replace("）", ")")
            if entry_norm == normalized:
                return e
        for e in entries:
            entry_norm = e["clauseName"].replace(" ", "").replace("（", "(").replace("）", ")")
            if normalized in entry_norm or entry_norm in normalized:
                return e
        core = re.sub(r"^附加", "", normalized)
        core = re.sub(r"条款$|扩展$", "", core)
        if len(core) < 3:
            return None
        for e in entries:
            entry_core = re.sub(r"^附加", "", e["clauseName"].replace(" ", ""))
            entry_core = re.sub(r"条款$|扩展$", "", entry_core)
            if core in entry_core or entry_core in core:
                return e
        return None

    def _match_inquiry_clauses(self, clause_names, file_name):
        if not clause_names:
            self.inquiry_status.setText("未识别到条款名称")
            return
        entries = self.rate_data["entries"]
        matched = []
        unmatched = []
        seen = set()
        for name in clause_names:
            entry = self._match_clause_name(name, entries)
            if entry and entry["sourceFile"] not in seen:
                matched.append({"importedName": name, "entry": entry})
                seen.add(entry["sourceFile"])
            elif not entry:
                unmatched.append(name)
        self.inquiry_status.setText(f"{file_name} → 识别 {len(clause_names)} 条，匹配 {len(matched)} 条")
        if unmatched:
            self._log(f"未匹配: {', '.join(unmatched[:5])}", "warn")
        if matched:
            self.batch_calc_btn.show()
            self.batch_calc_btn.setText(f"⚡ 一键计算全部（{len(matched)} 条）")
            self._batch_matched = matched
        else:
            self.batch_calc_btn.hide()
        self._log(f"导入 {file_name}：识别 {len(clause_names)} 条，匹配 {len(matched)} 条")

    def _batch_calculate(self):
        matched = getattr(self, '_batch_matched', [])
        if not matched:
            self._log("无匹配条款可计算", "warn")
            return
        self.main_premium = self.main_premium_input.value()
        if self.main_premium <= 0:
            self._log("请先输入主险保费", "warn")
            return
        calc_count = 0
        skip_count = 0
        for item in matched:
            entry = item["entry"]
            rt = entry.get("rateType", "")
            if rt in ("regulatory", "no_calc"):
                skip_count += 1
                continue
            if rt == "simple_percentage":
                result = self._calc_simple(entry)
                self._add_premium_item(entry["clauseName"], result["premium"], result["formulaDisplay"])
                calc_count += 1
            elif rt == "deduction":
                result = self._calc_deduction(entry)
                self._add_premium_item(entry["clauseName"], result["premium"], result["formulaDisplay"])
                calc_count += 1
            else:
                self._add_premium_item(entry["clauseName"], 0,
                                       f"需手动计算 [{ADDON_TYPES.get(rt, {}).get('label', rt)}]")
                skip_count += 1
        self._log(f"批量计算完成: {calc_count} 条已计算, {skip_count} 条需手动处理", "success")
