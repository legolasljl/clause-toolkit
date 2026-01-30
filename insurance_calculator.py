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
        calc_btn = make_accent_button("🧮 计算保费")
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
            formula += f"\n总保费 = {fmt_currency(per_person_premium)} × {employee_count}人 = {fmt_currency(total_premium)}"
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
        self._log(f"总保费: {fmt_currency(total_premium)}", "success")
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
            f"  总保费:     {fmt_currency(r['totalPremium'])}",
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
