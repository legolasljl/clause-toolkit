# -*- coding: utf-8 -*-
"""
Clause Comparison Assistant V18.0 (Multi-Tab Toolkit Edition)
智能条款工具箱
- [性能] 预处理索引加速匹配 5-10x
- [算法] 编辑距离容错 + 混合相似度
- [重构] 多级匹配策略拆分
- [功能] 批量处理多文件
- [健壮] 完善异常处理和日志
- [配置] 外部化JSON配置
- [新增] 用户自定义映射管理（单条/批量）
- [新增] 导出时使用库内条款名
- [v17.0] 中文分词支持 (jieba)
- [v17.0] 扩展英中映射表 (200+条款)
- [v17.0] 中英混合条款智能分离匹配
- [v17.0] TF-IDF向量快速候选筛选
- [v17.0] 动态权重调整
- [v17.0] 扩展语义别名和关键词库
- [v17.1] 多结果匹配（每条客户条款返回最多3条匹配供选择）
- [v17.1] 除外条款智能过滤（除非客户明确包含"除外"）
- [v17.1] 条款查询功能（仅查询标题，支持模糊搜索）
- [v17.1] 用户映射优先（有映射时只返回映射的那一条）
- [V18.0] Tab页面布局（条款提取/条款比对/条款输出）
- [V18.0] 条款提取功能（支持docx/pdf，文件夹智能分类）
- [V18.0] 条款输出功能（Excel比对报告转Word文档）
- [V18.0] 文件夹分类ZIP打包导出
- [V18.0] UI优化：紧凑型统计面板
- [V18.0] Tab标签显示优化（加宽+字体调整）
- [V18.0] 分类预览框样式修复（字体渲染）
- [V18.0] .doc自动转换为.docx功能（macOS textutil/LibreOffice）
- [V18.0] 统计栏水平对齐优化（分隔符布局）
- [V18.0] 文件列表字体颜色修复（高对比度）
- [V18.0] Excel导出Anthropic风格美化
- [V18.0] 条款输出Tab完整功能实现
- [V18.0] 支持从条款提取结果或Excel文件加载数据
- [V18.0] 三种输出模式：按条款逐个/按分类合并/全部合并
- [V18.0] Word样式自定义：标题字号/正文字号/行距/注册号显示
- [V18.0] 条款预览列表支持多选/全选
- [V18.0] 智能Excel列识别（自动匹配条款名称/注册号/内容列）
- [V18.0] Word文档Anthropic配色方案

Author: Dachi Yijin
Date: 2025-12-23
Updated: 2026-01-22 (V18.0 Multi-Tab Toolkit Edition)
"""

import sys
import os
import re
import difflib
import traceback
import logging
import subprocess
import platform
from typing import List, Dict, Tuple, Optional, Set, Any

# Windows 控制台编码修复（仅当有控制台时）
if platform.system() == 'Windows':
    import io
    # GUI 应用（--windowed）没有控制台，stdout/stderr 可能是 None
    if sys.stdout is not None and hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if sys.stderr is not None and hasattr(sys.stderr, 'buffer'):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
from dataclasses import dataclass, field
from enum import Enum
from collections import defaultdict
from functools import lru_cache
from pathlib import Path
from datetime import datetime
import pandas as pd
from docx import Document

# ASCII Art Logo
APP_LOGO = """
╔═══════════════════════════════════════════════════════════════════╗
║          ██████╗  ██████╗██╗   ██╗     ██╗██╗███╗   ██╗           ║
║          ██╔══██╗██╔════╝╚██╗ ██╔╝     ██║██║████╗  ██║           ║
║          ██║  ██║██║      ╚████╔╝      ██║██║██╔██╗ ██║           ║
║          ██║  ██║██║       ╚██╔╝  ██   ██║██║██║╚██╗██║           ║
║          ██████╔╝╚██████╗   ██║   ╚█████╔╝██║██║ ╚████║           ║
║          ╚═════╝  ╚═════╝   ╚═╝    ╚════╝ ╚═╝╚═╝  ╚═══╝           ║
║                    🚀 智能条款比对工具 🚀                         ║
║                     Author: Dachi_Yijin                           ║
╚═══════════════════════════════════════════════════════════════════╝
"""
# 打印Logo
print(APP_LOGO)

# ==========================================
# 中文分词支持
# ==========================================
try:
    import jieba
    jieba.setLogLevel(logging.WARNING)  # 减少jieba日志输出
    HAS_JIEBA = True
except ImportError:
    HAS_JIEBA = False

# ==========================================
# TF-IDF向量匹配支持
# ==========================================
try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    import numpy as np
    HAS_SKLEARN = True
except ImportError:
    HAS_SKLEARN = False

# ==========================================
# PDF解析支持
# ==========================================
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False
    try:
        import PyPDF2
        HAS_PYPDF2 = True
    except ImportError:
        HAS_PYPDF2 = False

# ==========================================
# ZIP打包支持
# ==========================================
import zipfile
import shutil

# ==========================================
# 日志配置
# ==========================================
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"Clause_Comparison_Assistant_{datetime.now():%Y%m%d}.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ==========================================
# 导入配置管理器
# ==========================================
try:
    from clause_config_manager import get_config, ClauseConfigManager
    HAS_CONFIG_MANAGER = True
except ImportError:
    HAS_CONFIG_MANAGER = False
    logger.warning("未找到 clause_config_manager，使用内置配置")

# 导入映射管理器
try:
    from clause_mapping_manager import ClauseMappingManager, get_mapping_manager
    from clause_mapping_dialog import ClauseMappingDialog
    HAS_MAPPING_MANAGER = True
except ImportError:
    HAS_MAPPING_MANAGER = False
    logger.warning("未找到 clause_mapping_manager，映射管理功能不可用")

# ==========================================
# macOS PyQt5 Plugin Fix
# ==========================================
try:
    import PyQt5
    plugin_path = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins')
    os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path
except ImportError:
    pass

try:
    from deep_translator import GoogleTranslator
    HAS_TRANSLATOR = True
except ImportError:
    HAS_TRANSLATOR = False

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QProgressBar, QTextEdit,
    QFileDialog, QMessageBox, QFrame, QGraphicsDropShadowEffect,
    QDialog, QFormLayout, QListWidget, QListWidgetItem, QCheckBox,
    QTabWidget, QSpinBox, QDoubleSpinBox, QGroupBox, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QFont, QColor, QDesktopServices, QTextCursor

# ==========================================
# macOS 打包防闪退
# ==========================================
class NullWriter:
    def write(self, text): pass
    def flush(self): pass

if getattr(sys, 'frozen', False):
    sys.stdout = NullWriter()
    sys.stderr = NullWriter()

def global_exception_handler(exctype, value, tb):
    error_msg = "".join(traceback.format_exception(exctype, value, tb))
    logger.error(f"未捕获异常: {error_msg}")
    try:
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setText("程序发生意外错误")
        msg_box.setInformativeText(str(value))
        msg_box.setDetailedText(error_msg)
        msg_box.exec_()
    except Exception as e:
        logger.error(f"无法显示错误对话框: {e}")

sys.excepthook = global_exception_handler


# ==========================================
# 安全工具函数
# ==========================================
def validate_file_path(file_path: str, allowed_extensions: list = None) -> bool:
    """
    验证文件路径安全性，防止路径遍历攻击

    Args:
        file_path: 要验证的文件路径
        allowed_extensions: 允许的文件扩展名列表 (如 ['.docx', '.xlsx'])

    Returns:
        True 如果路径安全，False 否则
    """
    if not file_path:
        return False

    # 转换为绝对路径并规范化
    try:
        abs_path = os.path.abspath(os.path.normpath(file_path))
    except (TypeError, ValueError):
        return False

    # 检查路径遍历攻击 (.. 序列)
    if '..' in file_path:
        logger.warning(f"检测到路径遍历尝试: {file_path}")
        return False

    # 检查是否访问敏感系统目录
    sensitive_dirs = ['/etc', '/usr', '/bin', '/sbin', '/var', '/root', '/System', '/Library']
    for sensitive in sensitive_dirs:
        if abs_path.startswith(sensitive):
            logger.warning(f"检测到敏感目录访问: {file_path}")
            return False

    # 检查文件扩展名
    if allowed_extensions:
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in [e.lower() for e in allowed_extensions]:
            return False

    return True


def sanitize_error_message(error: Exception) -> str:
    """
    清理错误信息，移除敏感路径和系统信息

    Args:
        error: 异常对象

    Returns:
        清理后的错误信息
    """
    error_str = str(error)

    # 移除完整文件路径，只保留文件名
    import re
    # 匹配类Unix路径
    error_str = re.sub(r'/(?:Users|home)/[^/]+/[^\s\'"]+', '<路径已隐藏>', error_str)
    # 匹配Windows路径
    error_str = re.sub(r'[A-Z]:\\[^\s\'"]+', '<路径已隐藏>', error_str)

    return error_str


# ==========================================
# Anthropic UI 设计系统
# ==========================================
class AnthropicColors:
    """Anthropic 官方色彩系统"""
    # 背景色
    BG_PRIMARY = "#faf9f5"      # 主背景/奶油白
    BG_CARD = "#f0eee6"         # 卡片背景/浅米色
    BG_MINT = "#bcd1ca"         # 特殊卡片/薄荷绿
    BG_LAVENDER = "#cbcadb"     # 特殊卡片/淡紫色
    BG_DARK = "#141413"         # 深色区域

    # 强调色
    ACCENT = "#d97757"          # 主强调色/陶土色
    ACCENT_DARK = "#c6613f"     # 次强调色/深赭红
    ACCENT_HOVER = "#e8956f"    # 悬停色

    # 文字色
    TEXT_PRIMARY = "#141413"    # 主要文字
    TEXT_SECONDARY = "#b0aea5"  # 次要文字（仅用于装饰性文字）
    TEXT_MUTED = "#6b6a65"      # 中等对比度文字（用于按钮/标签）
    TEXT_LIGHT = "#faf9f5"      # 深色背景上的文字

    # 状态色
    SUCCESS = "#5a9a7a"         # 成功/绿色
    WARNING = "#d9a557"         # 警告/金色
    ERROR = "#c75050"           # 错误/红色
    INFO = "#5a7a9a"            # 信息/蓝灰

    # 边框色
    BORDER = "#e5e3db"          # 浅边框
    BORDER_DARK = "#d0cec6"     # 深边框


class AnthropicFonts:
    """Anthropic 字体配置"""
    # 标题字体
    TITLE_LARGE = ("Anthropic Sans", 28)
    TITLE = ("Anthropic Sans", 22)
    TITLE_SMALL = ("Anthropic Sans", 16)

    # 正文字体
    BODY = ("Anthropic Serif", 14)
    BODY_SMALL = ("Anthropic Serif", 12)

    # UI 元素
    BUTTON = ("Anthropic Sans", 14)
    LABEL = ("Anthropic Sans", 13)

    # 代码字体
    CODE = ("JetBrains Mono", 12)

    # 中文回退（Anthropic 字体不含中文）
    CN_FALLBACK = "PingFang SC"


# ==========================================
# 常量定义
# ==========================================
class ExcelColumns:
    """Excel列名常量 - v17.1支持多结果匹配"""
    SEQ = '序号'
    CLIENT_ORIG = '客户条款(原)'
    CLIENT_TRANS = '客户条款(译)'
    CLIENT_CONTENT = '客户原始内容'

    # 多结果匹配列 (v17.1)
    # 匹配1
    MATCH1_NAME = '匹配1_条款名称'
    MATCH1_REG = '匹配1_注册号'
    MATCH1_CONTENT = '匹配1_条款内容'
    MATCH1_SCORE = '匹配1_匹配度'
    MATCH1_LEVEL = '匹配1_匹配级别'
    # 匹配2
    MATCH2_NAME = '匹配2_条款名称'
    MATCH2_REG = '匹配2_注册号'
    MATCH2_CONTENT = '匹配2_条款内容'
    MATCH2_SCORE = '匹配2_匹配度'
    MATCH2_LEVEL = '匹配2_匹配级别'
    # 匹配3
    MATCH3_NAME = '匹配3_条款名称'
    MATCH3_REG = '匹配3_注册号'
    MATCH3_CONTENT = '匹配3_条款内容'
    MATCH3_SCORE = '匹配3_匹配度'
    MATCH3_LEVEL = '匹配3_匹配级别'

    # 保留旧列名以兼容（主匹配结果）
    MATCHED_NAME = '匹配条款库名称'
    REG_NO = '产品注册号'
    MATCHED_CONTENT = '匹配条款库内容'
    SCORE = '综合匹配度'
    MATCH_LEVEL = '匹配级别'
    DIFF_ANALYSIS = '保障差异提示'
    TITLE_SCORE = '标题相似度'
    CONTENT_SCORE = '内容相似度'

    # 列索引（1-based, 需根据新格式调整）
    SCORE_COL_IDX = 8  # 匹配1_匹配度
    LEVEL_COL_IDX = 9  # 匹配1_匹配级别


# ==========================================
# 数据结构
# ==========================================
class MatchLevel(Enum):
    """匹配级别"""
    EXACT = "精确匹配"
    SEMANTIC = "语义匹配"
    KEYWORD = "关键词匹配"
    FUZZY = "模糊匹配"
    NONE = "无匹配"

@dataclass
class MatchThresholds:
    """匹配阈值"""
    exact_min: float = 0.98
    semantic_min: float = 0.85
    keyword_min: float = 0.60
    fuzzy_min: float = 0.40
    accept_min: float = 0.15

@dataclass
class ClauseItem:
    """条款项"""
    title: str
    content: str
    original_title: str = ""

@dataclass
class MatchResult:
    """匹配结果"""
    matched_name: str = ""
    matched_content: str = ""
    matched_reg: str = ""
    score: float = 0.0
    title_score: float = 0.0
    content_score: float = 0.0
    match_level: MatchLevel = MatchLevel.NONE
    diff_analysis: str = ""

@dataclass
class LibraryIndex:
    """条款库索引结构"""
    by_name_norm: Dict[str, int] = field(default_factory=dict)
    by_keyword: Dict[str, List[int]] = field(default_factory=lambda: defaultdict(list))
    cleaned_cache: Dict[int, Dict[str, str]] = field(default_factory=dict)
    data: List[Dict] = field(default_factory=list)


# ==========================================
# 内置默认配置（当配置管理器不可用时）
# ==========================================
class DefaultConfig:
    """默认配置 - v17.0 扩展版"""

    # ========================================
    # 英中条款映射表 (基于TOC.docx扩展，200+条目)
    # ========================================
    CLIENT_EN_CN_MAP = {
        # === 通用条款 ===
        "30 days notice of cancellation clause": "30天注销保单通知条款",
        "30 days notice of cancellation": "30天注销保单通知条款",
        "60 days non-renewal notice clause": "60天不续保通知条款",
        "60 days non-renewal notice": "60天不续保通知条款",
        "72 hours clause": "72小时条款",
        "72 hours": "72小时条款",
        "time adjustment": "72小时条款",
        "50/50 clause": "50/50条款",
        "85% co-insurance": "85％扩展条款",
        "co-insurance extension clause": "85％扩展条款",

        # === 索赔与控制 ===
        "claims control clause": "理赔控制条款",
        "claims control": "理赔控制条款",
        "joint-insured clause": "共同被保险人条款",
        "joint insured clause": "共同被保险人条款",
        "joint insured": "共同被保险人条款",
        "jurisdiction clause": "司法管辖权条款",
        "jurisdiction": "司法管辖权条款",
        "loss adjusters clause": "指定公估人条款",
        "loss adjusters": "指定公估人条款",
        "nomination of loss adjusters clause": "指定公估人条款",
        "loss adjuster clause": "指定公估人条款",
        "loss notification clause": "损失通知条款",
        "loss notification": "损失通知条款",

        # === 控制与取消 ===
        "no control clause": "不受控制条款",
        "no control": "不受控制条款",
        "non-cancellation clause": "不可注销保单条款",
        "non cancellation clause": "不可注销保单条款",
        "non-invalidation clause": "不使失效条款",
        "non invalidation clause": "不使失效条款",
        "non invalidation": "不使失效条款",

        # === 付款与费用 ===
        "payment on account clause": "预付赔款条款",
        "payment on account": "预付赔款条款",
        "premium installment clause": "分期付费条款",
        "premium instalment clause": "分期付费条款",
        "premium adjustment clause": "保费调整条款",
        "premium adjustment": "保费调整条款",
        "professional fees clause": "专业费用扩展条款",
        "professional fees": "专业费用扩展条款",
        "professional fee extension clause": "专业费用扩展条款",

        # === 代位与利益 ===
        "severability of interest clause": "利益可分性条款",
        "severability of interest": "利益可分性条款",
        "waiver of subrogation clause": "放弃代位追偿扩展条款",
        "waiver of subrogation extension clause": "放弃代位追偿扩展条款",
        "waiver of subrogation": "放弃代位追偿扩展条款",

        # === 自然灾害 ===
        "earthquake extension clause": "地震扩展条款",
        "earthquake extension": "地震扩展条款",
        "earthquake and tsunami": "地震扩展条款",
        "earthquake": "地震扩展条款",
        "flood extension clause": "洪水扩展条款",
        "flood extension": "洪水扩展条款",
        "flood and inundation exclusion clause": "洪水除外条款",
        "flood exclusion": "洪水除外条款",
        "flood prevention warranty clause": "防洪保证条款",
        "typhoon and hurricane extension clause": "台风、飓风扩展条款",
        "typhoon and hurricane": "台风、飓风扩展条款",
        "typhoon extension": "台风、飓风扩展条款",
        "hurricane extension": "台风、飓风扩展条款",
        "tornado extension clause": "龙卷风扩展条款",
        "tornado exclusion clause": "龙卷风除外条款",
        "hailstone extension clause": "冰雹扩展条款",
        "hailstone extension": "冰雹扩展条款",
        "snowstorm and icicle extension clause": "暴雪、冰凌扩展条款",
        "snowstorm extension": "暴雪、冰凌扩展条款",
        "storm and tempest extension clause": "暴风雨扩展条款",
        "storm and tempest exclusion clause": "暴风雨除外条款",
        "sandstorm extension clause": "沙尘暴扩展条款",
        "sandstorm exclusion clause": "沙尘暴除外条款",
        "lightning extension clause": "雷电扩展条款",
        "lightning extension": "雷电扩展条款",

        # === 地质灾害 ===
        "accidental subsidence of ground extension clause": "地面突然下陷下沉扩展条款",
        "subsidence extension": "地面突然下陷下沉扩展条款",
        "landslip & subsidence clause": "地崩及地陷条款",
        "landslip and subsidence": "地崩及地陷条款",
        "mud-rock flow, avalanche and sudden landslip extension clause": "泥石流、崩塌、突发性滑坡扩展条款",
        "mudslide extension": "泥石流、崩塌、突发性滑坡扩展条款",

        # === 盗窃与安全 ===
        "theft, burglary and robbery extension clause": "盗窃、抢劫扩展条款",
        "theft and robbery": "盗窃、抢劫扩展条款",
        "theft extension": "盗窃、抢劫扩展条款",
        "burglary insurance clause": "盗窃险条款",
        "burglary extension": "盗窃险条款",
        "malicious damage extension clause": "恶意破坏扩展条款",
        "malicious damage": "恶意破坏扩展条款",

        # === 罢工与暴乱 ===
        "strike riot and civil commotion extension clause": "罢工、暴乱及民众骚乱扩展条款",
        "strike, riot and civil commotion": "罢工、暴乱及民众骚乱扩展条款",
        "strike riot civil commotion": "罢工、暴乱及民众骚乱扩展条款",
        "srcc": "罢工、暴乱及民众骚乱扩展条款",
        "terrorism extension clause": "恐怖活动扩展条款",
        "terrorism extension": "恐怖活动扩展条款",
        "act of terrorism extension clause": "恐怖活动扩展条款",

        # === 价值与金额 ===
        "reinstatement value clause": "重置价值条款",
        "reinstatement value": "重置价值条款",
        "agreed value insurance clause": "定值保险条款",
        "agreed value": "定值保险条款",
        "automatic reinstatement of sum insured clause": "自动恢复保险金额条款",
        "automatic reinstatement of sum insured": "自动恢复保险金额条款",
        "automatic reinstatement": "自动恢复保险金额条款",
        "escalation extension clause": "自动升值扩展条款",
        "escalation extension": "自动升值扩展条款",

        # === 费用扩展 ===
        "removal of debris clause": "清理残骸费用扩展条款",
        "removal of debris": "清理残骸费用扩展条款",
        "debris removal expenses extension clause": "清理残骸费用扩展条款",
        "debris removal": "清理残骸费用扩展条款",
        "fire fighting cost extension clause": "灭火费用扩展条款",
        "fire fighting cost": "灭火费用扩展条款",
        "fire brigade charges extension clause": "消防队灭火费用扩展条款",
        "air freight fee extension clause": "空运费扩展条款",
        "air freight extension": "空运费扩展条款",
        "airfreight clause": "空运费扩展条款",
        "extra charges extension clause": "特别费用扩展条款",
        "extra charges clause": "特别费用扩展条款",
        "extra charges": "特别费用扩展条款",

        # === 财产与建筑 ===
        "all other contents extension clause": "其他物品扩展条款",
        "alteration of building clause": "建筑物变动扩展条款",
        "building alterations clause": "建筑物改变条款",
        "capital additions extension clause": "增加资产扩展条款",
        "capital additions": "增加资产扩展条款",
        "contract price extension clause": "合同价格扩展条款",
        "designation of property clause": "财物种别条款",
        "foundation exclusion clause": "地基除外条款",
        "simple building exclusion": "简易建筑除外条款",
        "property in the open or simple building extension clause": "露天存放及简易建筑内财产扩展条款",
        "property in the open": "露天存放及简易建筑内财产扩展条款",
        "off premises property clause": "场所外财产条款",
        "outside ancillary devices of building extension clause": "建筑物外部附属设施扩展条款",

        # === 设备与机械 ===
        "boilers and pressure vessels extension clause": "锅炉、压力容器扩展条款",
        "boiler explosion": "锅炉爆炸责任条款",
        "breakage of glass extension clause": "玻璃破碎扩展条款",
        "breakage of glass clause": "玻璃破碎扩展条款",
        "glass breakage": "玻璃破碎扩展条款",
        "bursting of water tank or water pipe extension clause": "水箱、水管爆裂扩展条款",
        "water damage": "水箱、水管爆裂扩展条款",
        "hoisting and transport machinery extension clause": "起重、运输机械扩展条款",
        "locomotive extension clause": "铁路机车车辆扩展条款",
        "refrigerating plants extension clause": "冷库扩展条款",
        "sprinkler leakage damage extension clause": "自动喷淋系统水损扩展条款",
        "sprinkler leakage": "自动喷淋系统水损扩展条款",
        "portable devices on premises extension clause": "便携式设备扩展条款",

        # === 运输与移动 ===
        "inland transit extension clause": "内陆运输扩展条款",
        "inland transit clause": "内陆运输扩展条款",
        "inland transit": "内陆运输扩展条款",
        "transit clause": "运输条款",
        "temporary removal extension clause": "临时移动扩展条款",
        "temporary removal": "临时移动扩展条款",
        "temporary removal between factories extension clause": "厂区间临时移动扩展条款",
        "loaded property extension clause": "车辆装载物扩展条款",

        # === 责任与赔偿 ===
        "public authority extension clause": "公共当局扩展条款",
        "public authority": "公共当局扩展条款",
        "civil authorities clause": "公共当局扩展条款",
        "civil authorities": "公共当局扩展条款",
        "error and omissions clause": "错误和遗漏条款",
        "errors and omissions clause": "错误和遗漏条款",
        "errors and omissions": "错误和遗漏条款",
        "breach of conditions clause": "违反条件条款",
        "breach of conditions": "违反条件条款",
        "cross liability clause": "交叉责任条款",
        "cross liability": "交叉责任条款",
        "contractual liability clause": "契约责任扩展条款",
        "contractual liability": "契约责任扩展条款",

        # === 其他扩展 ===
        "automatic cover clause": "自动承保条款",
        "automatic cover": "自动承保条款",
        "average relief clause": "分摊豁免条款",
        "brand & trademark clause": "商标条款",
        "brand and trademark": "商标条款",
        "cost of duplication extension clause": "复制费用扩展条款",
        "documents clause": "索赔单据条款",
        "emergency rescue clause": "紧急抢险条款",
        "falling of flying objects extension clause": "飞行物体及其他空中运行物体坠落扩展条款",
        "fire prevention facilities warranty clause": "消防保证条款",
        "impact damage extension clause": "碰撞扩展条款",
        "impact damage exclusion clause": "碰撞除外条款",
        "inhibition clause": "阻止条款",
        "it clarification clause": "数据损失澄清条款",
        "legal requirements warranty": "遵守法律规定保证条款",
        "loss payee clause": "赔款接受人条款",
        "mortgage clause": "抵押权条款",
        "mortgagee clause": "抵押条款",
        "non occupying landlord clause": "非占用者业主条款",
        "oil or gas pipeline damage extension clause": "油气管道损坏扩展条款",
        "out-sourcing processing extension clause": "委托加工扩展条款",
        "pair & set clause": "成对或成套设备条款",
        "personal effects of employees extension clause": "雇员个人物品扩展条款",
        "smoke damage extension clause": "烟熏扩展条款",
        "spontaneous combustion extension clause": "自燃扩展条款",
        "spontaneous combustion exclusion clause": "自燃除外条款",
        "stock declaration and adjustment clause": "仓储财产申报条款",
        "storage warranty": "存放保证条款",
        "supply failure extension clause": "供应中断扩展条款",
        "supply suspension extension clause": "供应中断扩展条款",
        "temporary protection extension clause": "临时保护措施扩展条款",
        "undamaged building extra charges extension clause": "建筑物未受损部分额外费用扩展条款",
        "workmen clause": "装修工人条款",
        "assignment clause": "权益转让条款",

        # === 工程险条款 ===
        "camps and stores clause": "工棚、库房特别条款",
        "cement storage warranty": "水泥存储保证条款",
        "construction material clause": "建筑材料特别条款",
        "construction plant, equipment and machinery clause": "施工用机具特别条款",
        "construction machinery clause": "建筑、安装施工机具、设备扩展条款",
        "erection machinery clause": "建筑、安装施工机具、设备扩展条款",
        "contract works taken over or put into service clause": "工程完工部分扩展条款",
        "cost for decontamination clause": "清除污染费用扩展条款",
        "customs duties clause": "海关关税条款",
        "dams and water reservoirs clause": "大坝、水库工程除外特别条款",
        "defective design, materials and workmanship": "设计错误、原材料缺陷及工艺不善条款",
        "defects liability period clause": "扩展责任保证期扩展条款",
        "extended maintenance clause": "扩展责任保证期扩展条款",
        "designer's risk clause": "设计师风险扩展条款",
        "drilling work for water wells clause": "钻井工程特别条款",
        "employer's property extension": "雇主财产财产扩展条款",
        "escalation clause": "10％增值条款",
        "existing structures and surrounding property clause": "原有建筑物及周围财产扩展条款",
        "extinguishing expenses clause": "灭火费用条款",
        "fire-fighting facilities clause": "防火设施特别条款",
        "free issue materials clause": "免费提供物料扩展条款",
        "guarantee period clause": "保证期特别扩展条款",
        "hoisting, cranes and unregistered vehicles liability clause": "起重机、未登记车辆责任扩展条款",
        "hydrocarbon processing industries clause": "碳氢化合物生产业特别条款",
        "indemnity to principals clause": "业主保障条款",
        "laying pipelines, ducts and cables clause": "铺设管道、电缆特别条款",
        "laying water supply and sewer pipes clause": "铺设供水、污水管特别条款",
        "leak search costs when laying pipelines clause": "埋管查漏费用特别条款",
        "maintenance & inspection clause": "检查维护条款",
        "maintenance visits clause": "有限责任保证期扩展条款",
        "marine cargo insurance clause": "运输险、工程险责任分摊条款",
        "marine work special condition": "海工特别条款",
        "non-negligent indemnity": "非疏忽过错赔偿条款",
        "nuclear fuel elements clause": "核燃料组件条款",
        "obstruction & nuisance clause": "阻碍或妨害条款",
        "offsite storage clause": "工地外储存物特别条款",
        "plans and documents clause": "工程图纸、文件特别条款",
        "principal's property clause": "业主财产扩展条款",
        "quarterly declaration clause": "季度申报条款",
        "reactor pressure vessel with internals clause": "压力反应堆特别扩展条款",
        "removal of debris from landslides clause": "清除滑坡土石方特别除外条款",
        "run off clause": "保单延续条款",
        "safety precautions clause": "安全防范条款",
        "structures in earthquake zones clause": "地震地区建筑物特别条款",
        "testing & commissioning clause": "试车条款",
        "time adjustment clause": "时间调整特别条款",
        "time schedule clause": "建筑、安装工程时间进度特别条款",
        "tunnels and galleries clause": "隧道工程特别除外条款",
        "underground cables, pipes and other facilities clause": "地下电缆、管道及设施特别条款",
        "underground service clause": "地下服务设施条款",
        "underground works clause": "地下工程条款",
        "unexploded bombs clause": "地下炸弹特别条款",
        "used machinery clause": "旧设备除外条款",
        "vibration, removal or weakening of support clause": "震动、移动或减弱支撑扩展条款",
        "sue & labor clause": "诉讼及劳务费用特别条款",
        "manufacturer's risks clause": "制造商风险扩展条款",
        "piling, foundation and retaining wall construction work": "打桩及挡土墙除外条款",
        "burning & welding clause": "烧焊条款",

        # === 责任险条款 ===
        "accidental pollution clause": "意外污染条款",
        "advertising signs and decorations liability clause": "广告及装饰装置责任条款",
        "car park liability clause": "停车场责任条款",
        "car park service clause": "泊车服务条款",
        "contractors contingent liability clause": "承包人意外责任条款",
        "defective sanitary installation clause": "有缺陷的卫生装置条款",
        "delivery goods extension": "运输货物扩展条款",
        "elevator and escalator clause": "电梯责任条款",
        "lifts, elevators and escalators liability clause": "电梯、升降机责任扩展条款",
        "employees temporary working overseas": "海外公干条款",
        "employees temporarily working oversea clause": "员工公（劳）务出国条款",
        "exhibition and sales demonstration": "展览和销售演示条款",
        "fire & explosion extension clause": "火灾和爆炸责任条款",
        "fire brigade and water damage clause": "灭火及水损责任条款",
        "first aid liability clause": "急救责任条款",
        "first aid treatment clause": "急救费用条款",
        "food and drink clause": "食品、饮料责任条款",
        "goods and services clause": "提供物品及服务条款",
        "guest's property clause": "客人财产责任条款",
        "hire and non-owned automobiles liability clause": "租用及非拥有机动车辆责任条款",
        "hoists, cranes and unregistered vehicles liability clause": "起重机及起重设备责任条款",
        "indemnity to landlord clause": "房东保障条款",
        "independent contractors liability clause": "独立承建商责任条款",
        "laundry liability clause": "洗衣房责任条款",
        "loading and unloading of vehicles clause": "车辆装卸责任条款",
        "maintenance, repair and decoration of the premises clause": "修改、修理及保养责任条款",
        "motor contingent liability clause": "租用汽车责任条款",
        "personal injury liability clause": "人身侵害责任条款",
        "social and welfare club clause": "联谊及康乐活动责任附加条款",
        "swimming pool liability clause": "游泳池责任条款",
        "tenant's liability clause": "出租人责任条款",
        "third party liability of directors and executives clause": "董事及高级管理人员个人第三者责任条款",
        "catering facilities clause": "膳食条款",
        "extraordinary weather condition clause": "反常天气条款",
        "social activities clause": "社会活动条款",

        # === 产品责任险条款 ===
        "absolute asbestos exclusion": "石棉除外条款",
        "absolute pollution exclusion": "污染除外条款",
        "allergy exclusion clause": "过敏除外条款",
        "batch clause": "同一批次产品条款",
        "circuit board & battery exclusion clause": "电路板、电池除外条款",
        "claim made basis clause": "以索赔提出为基础条款",
        "defense cost within the limit of indemnity": "抗辩费用条款",
        "designated vendor liability": "指定经销商责任条款",
        "efficacy exclusion clause": "功效除外条款",
        "electromagnetic radiation exclusion": "电磁辐射、无线电波除外条款",
        "employees bodily injury exclusion": "雇员人身伤害除外条款",
        "genetically modified organisms exclusion": "转基因体除外条款",
        "gmo exclusion": "转基因体除外条款",
        "lead exclusion": "铅物质除外条款",
        "nuclear energy liability exclusion": "核能责任除外条款",
        "occurrence basis clause": "以发生为基础条款",
        "product-completed operation": "完工操作风险条款",
        "punitive damage exclusion": "惩罚性赔偿除外条款",
        "exemplary damage exclusion": "惩罚性赔偿除外条款",
        "silica exclusion": "硅除外条款",
        "us canada domiciled operations exclusion clause": "美加地区操作除外条款",
        "vendor broad form liability": "列明经销商扩展条款",
        "war and terrorism exclusion": "战争及恐怖主义除外条款",

        # === 营业中断险条款 ===
        "accumulated stocks clause": "货物累积条款",
        "bomb scare extension": "炸弹恐吓条款",
        "denial of access": "通道堵塞条款",
        "departmental clause": "部门条款",
        "inclusion of all turnover": "包括全部营业额条款",
        "infectious disease murder and closure clause": "谋杀等条款",
        "public utilities extension": "公共事业设备扩展条款",
        "reinstatement of sum insured clause": "恢复保险金额条款",
        "uninsured standing charges clause": "未保险的维持费用条款",
        "waiver of excess clause": "免赔额豁免条款",
        "loss of book debts clause": "遗失欠款帐册条款",

        # === 机损险条款 ===
        "overhaul of electric motors": "电动马达检修条款",
        "overhaul of steam, water and gas turbines": "蒸气、水、气体涡轮机及涡轮发电机条款",

        # === 通用简写 ===
        "interpretation & headings": "通译和标题条款",
        "year 2000 problem exclusion clause": "财产险2000年问题除外责任条款",
    }

    # ========================================
    # 语义别名映射表 (扩展版)
    # ========================================
    SEMANTIC_ALIAS_MAP = {
        # === 污染相关 ===
        "污染保险": "意外污染责任",
        "污染责任": "意外污染责任",
        "污染条款": "意外污染条款",
        "环境污染": "意外污染责任",

        # === 财产存放 ===
        "露天财产": "露天存放及简易建筑内财产",
        "露天物品": "露天存放及简易建筑内财产",
        "简易建筑": "露天存放及简易建筑内财产",
        "临时建筑": "露天存放及简易建筑内财产",

        # === 施救费用 ===
        "损害防止": "阻止损失",
        "施救费用": "阻止损失",
        "救援费用": "阻止损失",
        "抢险费用": "紧急抢险",

        # === 地质灾害 ===
        "崩塌沉降": "地面突然下陷下沉",
        "地面下陷": "地面突然下陷下沉",
        "地陷": "地面突然下陷下沉",
        "地面沉降": "地面突然下陷下沉",
        "山体滑坡": "泥石流、崩塌、突发性滑坡",
        "滑坡": "泥石流、崩塌、突发性滑坡",
        "泥石流": "泥石流、崩塌、突发性滑坡",

        # === 盗窃相关 ===
        "盗窃险": "盗窃、抢劫扩展",
        "盗抢险": "盗窃、抢劫扩展",
        "抢劫险": "盗窃、抢劫扩展",
        "入室盗窃": "盗窃、抢劫扩展",

        # === 自然灾害 ===
        "地震海啸": "地震扩展",
        "震动": "地震扩展",
        "台风": "台风、飓风扩展",
        "飓风": "台风、飓风扩展",
        "暴风": "暴风雨扩展",
        "暴雨": "暴风雨扩展",
        "水灾": "洪水扩展",
        "水淹": "洪水扩展",
        "内涝": "洪水扩展",
        "雷击": "雷电扩展",
        "雷电": "雷电扩展",
        "冰雹": "冰雹扩展",
        "雪灾": "暴雪、冰凌扩展",
        "冰凌": "暴雪、冰凌扩展",

        # === 机械设备 ===
        "锅炉爆炸": "锅炉、压力容器扩展",
        "压力容器": "锅炉、压力容器扩展",
        "玻璃破损": "玻璃破碎扩展",
        "玻璃险": "玻璃破碎扩展",
        "水管爆裂": "水箱、水管爆裂扩展",
        "水管破裂": "水箱、水管爆裂扩展",
        "喷淋系统": "自动喷淋系统水损扩展",
        "消防喷淋": "自动喷淋系统水损扩展",

        # === 责任相关 ===
        "公共责任": "公众责任",
        "第三者责任": "公众责任",
        "雇主责任": "雇员责任",
        "工伤责任": "雇员责任",
        "产品责任": "产品责任",
        "职业责任": "专业责任",

        # === 费用相关 ===
        "残骸清理": "清理残骸费用",
        "清除残骸": "清理残骸费用",
        "灭火费用": "灭火费用扩展",
        "消防费用": "灭火费用扩展",
        "空运费": "空运费扩展",
        "加急运费": "空运费扩展",
        "专业费": "专业费用扩展",
        "公估费": "专业费用扩展",

        # === 价值相关 ===
        "重置价值": "重置价值条款",
        "重建价值": "重置价值条款",
        "恢复保额": "自动恢复保险金额",
        "自动恢复": "自动恢复保险金额",

        # === 罢工暴乱 ===
        "罢工": "罢工、暴乱及民众骚乱",
        "暴乱": "罢工、暴乱及民众骚乱",
        "民众骚乱": "罢工、暴乱及民众骚乱",
        "骚乱": "罢工、暴乱及民众骚乱",
        "恐怖活动": "恐怖活动扩展",
        "恐怖袭击": "恐怖活动扩展",

        # === 运输相关 ===
        "内陆运输": "内陆运输扩展",
        "陆上运输": "内陆运输扩展",
        "临时移动": "临时移动扩展",
        "厂区运输": "厂区间临时移动扩展",
    }

    # ========================================
    # 关键词映射表 (扩展版)
    # ========================================
    KEYWORD_MAP = {
        # === 自然灾害 ===
        "地震": ["地震", "震动", "earthquake", "seismic"],
        "海啸": ["海啸", "tsunami"],
        "洪水": ["洪水", "水灾", "水淹", "内涝", "flood", "inundation"],
        "台风": ["台风", "飓风", "typhoon", "hurricane", "cyclone"],
        "龙卷风": ["龙卷风", "tornado", "twister"],
        "暴风雨": ["暴风", "暴雨", "storm", "tempest"],
        "雷电": ["雷电", "雷击", "闪电", "lightning"],
        "冰雹": ["冰雹", "hail", "hailstone"],
        "暴雪": ["暴雪", "雪灾", "冰凌", "snowstorm", "icicle"],
        "沙尘暴": ["沙尘暴", "sandstorm", "dust storm"],

        # === 地质灾害 ===
        "滑坡": ["滑坡", "崩塌", "泥石流", "landslip", "landslide", "mudslide", "avalanche"],
        "地陷": ["地陷", "下陷", "沉降", "subsidence", "sinkhole"],

        # === 盗窃相关 ===
        "盗窃": ["盗窃", "盗抢", "抢劫", "入室", "burglary", "theft", "robbery"],
        "恶意破坏": ["恶意", "蓄意", "malicious", "vandalism"],

        # === 罢工暴乱 ===
        "罢工": ["罢工", "strike"],
        "暴乱": ["暴乱", "暴动", "riot"],
        "骚乱": ["骚乱", "民众骚乱", "civil commotion"],
        "恐怖": ["恐怖", "terrorism", "terrorist"],

        # === 污染相关 ===
        "污染": ["污染", "意外污染", "环境污染", "pollution", "contamination"],

        # === 设备相关 ===
        "锅炉": ["锅炉", "boiler"],
        "压力容器": ["压力容器", "pressure vessel"],
        "玻璃": ["玻璃", "glass"],
        "水管": ["水管", "水箱", "水损", "water pipe", "water tank"],
        "喷淋": ["喷淋", "消防喷淋", "sprinkler"],
        "电梯": ["电梯", "升降机", "扶梯", "elevator", "escalator", "lift"],
        "起重机": ["起重机", "起重", "吊车", "crane", "hoist"],

        # === 火灾相关 ===
        "火灾": ["火灾", "火险", "fire"],
        "自燃": ["自燃", "spontaneous combustion"],
        "爆炸": ["爆炸", "explosion"],
        "烟熏": ["烟熏", "smoke"],

        # === 价值相关 ===
        "重置": ["重置", "重建", "reinstatement", "replacement"],
        "定值": ["定值", "约定价值", "agreed value"],
        "恢复保额": ["恢复保额", "恢复保险金额", "reinstatement of sum"],
        "升值": ["升值", "增值", "escalation"],

        # === 费用相关 ===
        "残骸": ["残骸", "清理残骸", "debris", "removal of debris"],
        "灭火": ["灭火", "消防", "fire fighting", "fire brigade"],
        "空运费": ["空运费", "空运", "air freight", "airfreight"],
        "专业费用": ["专业费用", "公估", "professional fee"],
        "施救": ["施救", "救援", "抢险", "sue and labor", "sue & labor"],

        # === 责任相关 ===
        "公众责任": ["公众责任", "第三者", "public liability", "third party"],
        "产品责任": ["产品责任", "product liability"],
        "雇主责任": ["雇主责任", "雇员责任", "employer", "employee liability"],
        "交叉责任": ["交叉责任", "cross liability"],
        "契约责任": ["契约责任", "合同责任", "contractual liability"],

        # === 运输相关 ===
        "运输": ["运输", "transit", "transport"],
        "内陆运输": ["内陆运输", "inland transit"],
        "临时移动": ["临时移动", "temporary removal"],

        # === 工程相关 ===
        "工程": ["工程", "construction", "erection"],
        "试车": ["试车", "testing", "commissioning"],
        "保证期": ["保证期", "维护期", "maintenance", "defects liability"],
        "隧道": ["隧道", "tunnel"],
        "打桩": ["打桩", "桩基", "piling"],

        # === 其他 ===
        "72小时": ["72小时", "时间调整", "72 hours", "time adjustment"],
        "代位追偿": ["代位追偿", "代位", "subrogation"],
        "共同被保险人": ["共同被保险人", "joint insured"],
        "免赔额": ["免赔额", "免赔", "deductible", "excess"],
    }

    PENALTY_KEYWORDS = ["打孔盗气"]

    NOISE_WORDS = [
        "企业财产保险", "附加", "扩展", "条款", "险",
        "（A款）", "（B款）", "(A款)", "(B款)",
        "2025版", "2024版", "2023版", "2022版", "版",
        "clause", "extension", "cover", "insurance",
        "特别", "责任", "保险", "费用",
    ]

    # ========================================
    # 语义聚类（用于更智能的匹配）
    # ========================================
    SEMANTIC_CLUSTERS = {
        "地震类": ["地震", "震动", "地震海啸", "地震扩展", "earthquake"],
        "水灾类": ["洪水", "水灾", "暴雨", "水淹", "内涝", "flood", "inundation"],
        "盗窃类": ["盗窃", "盗抢", "抢劫", "入室盗窃", "burglary", "theft", "robbery"],
        "施救类": ["施救费用", "损害防止", "阻止损失", "救援费用", "sue and labor"],
        "台风类": ["台风", "飓风", "热带风暴", "typhoon", "hurricane"],
        "火灾类": ["火灾", "火险", "燃烧", "fire"],
        "罢工类": ["罢工", "暴乱", "骚乱", "民众骚乱", "strike", "riot", "civil commotion"],
        "责任类": ["责任", "赔偿", "liability", "indemnity"],
    }


# ==========================================
# 编辑距离算法
# ==========================================
@lru_cache(maxsize=10000)
def levenshtein_distance(s1: str, s2: str) -> int:
    """计算编辑距离（带缓存）"""
    if len(s1) < len(s2):
        return levenshtein_distance(s2, s1)

    if len(s2) == 0:
        return len(s1)

    previous_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        current_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = previous_row[j + 1] + 1
            deletions = current_row[j] + 1
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row

    return previous_row[-1]


def levenshtein_ratio(s1: str, s2: str) -> float:
    """计算编辑距离相似度"""
    if not s1 or not s2:
        return 0.0

    # 长度差异过大直接返回低分
    len_diff = abs(len(s1) - len(s2))
    max_len = max(len(s1), len(s2))
    if len_diff > max_len * 0.6:
        return 0.0

    distance = levenshtein_distance(s1, s2)
    return 1 - (distance / max_len)


# ==========================================
# 核心匹配逻辑（重构版）
# ==========================================
class ClauseMatcherLogic:
    """条款匹配核心逻辑 - 优化版"""

    # 条款库中的常见样板内容（这些内容不影响匹配度计算）
    BOILERPLATE_PHRASES = [
        "本条款所述费用在本保险单明细表所列保险金额之外另行赔付。",
        "本附加险根据保险单的约定收取保险费。",
        "本保险单所载其他条件不变。",
        "本附加条款与主条款内容相悖之处，以本附加条款为准；未尽之处，以主条款为准。",
        "限额由保险双方约定并在保险单中载明。",
        "本条款未尽事宜，以主保险合同的条款为准。",
        "本附加险条款与主险条款相抵触之处，以本附加险条款为准。",
        "本保险合同所载其他条款、条件和除外责任不变。",
        "本附加险保费按主险保费的一定比例收取。",
        "本条款中任何未定义的词语或术语具有主保险合同中规定的含义。",
    ]

    def __init__(self):
        """初始化匹配器"""
        # 加载配置
        if HAS_CONFIG_MANAGER:
            self.config = get_config()
            self._use_external_config = True
        else:
            self.config = None
            self._use_external_config = False

        self.thresholds = MatchThresholds()
        self._index: Optional[LibraryIndex] = None

        # v17.0: TF-IDF向量索引
        self._tfidf_vectorizer = None
        self._tfidf_vectors = None
        self._tfidf_names = []

        logger.info(f"匹配器初始化完成，外部配置: {self._use_external_config}")
        logger.info(f"jieba分词: {HAS_JIEBA}, sklearn(TF-IDF): {HAS_SKLEARN}")

    @classmethod
    def remove_boilerplate(cls, content: str) -> str:
        """
        从内容中移除样板文字，用于更准确的相似度计算
        """
        if not content:
            return ""
        result = content
        for phrase in cls.BOILERPLATE_PHRASES:
            result = result.replace(phrase, "")
        # 移除多余的空白和换行
        result = re.sub(r'\s+', ' ', result).strip()
        return result

    # ========================================
    # 配置访问方法
    # ========================================

    def _get_client_mapping(self, term: str) -> Optional[str]:
        """获取英中映射"""
        if self._use_external_config:
            return self.config.get_client_mapping(term)
        return DefaultConfig.CLIENT_EN_CN_MAP.get(term.lower())

    def _get_semantic_alias(self, text: str) -> Optional[str]:
        """获取语义别名"""
        alias_map = (self.config.semantic_alias_map if self._use_external_config
                     else DefaultConfig.SEMANTIC_ALIAS_MAP)
        for alias, target in alias_map.items():
            if alias in text:
                return target
        return None

    def _get_keywords(self, text: str) -> Set[str]:
        """提取关键词"""
        keywords = set()
        text_lower = text.lower()
        keyword_map = (self.config.keyword_extract_map if self._use_external_config
                       else DefaultConfig.KEYWORD_MAP)
        for core, variants in keyword_map.items():
            for v in variants:
                if v.lower() in text_lower:
                    keywords.add(core)
                    break
        return keywords

    def _is_penalty_keyword(self, text: str) -> bool:
        """检查惩罚关键词"""
        penalty_list = (self.config.penalty_keywords if self._use_external_config
                        else DefaultConfig.PENALTY_KEYWORDS)
        return any(kw in text for kw in penalty_list)

    def _get_noise_words(self) -> List[str]:
        """获取噪音词列表"""
        return (self.config.noise_words if self._use_external_config
                else DefaultConfig.NOISE_WORDS)

    # ========================================
    # 文本处理方法
    # ========================================

    @staticmethod
    def normalize_text(text: str) -> str:
        """标准化文本"""
        if not isinstance(text, str):
            return ""
        text = text.lower().strip()
        text = re.sub(r"['\"\'\'\"\"\(\)（）\[\]【】]", '', text)
        text = re.sub(r'\s+', ' ', text)
        return text

    def clean_title(self, text: str) -> str:
        """清理标题"""
        if not isinstance(text, str):
            return ""
        text = re.sub(r'[\(（].*?[\)）]', '', text)
        for w in self._get_noise_words():
            text = text.replace(w, "").replace(w.lower(), "")
        text = re.sub(r'[0-9\s]+', '', text)
        return text.strip()

    @classmethod
    def clean_content(cls, text: str) -> str:
        """
        清理内容，用于相似度计算
        1. 移除样板文字
        2. 移除括号内容
        3. 移除空白和数字
        """
        if not isinstance(text, str):
            return ""
        # 先移除样板文字
        text = cls.remove_boilerplate(text)
        # 移除括号及其内容
        text = re.sub(r'[\(（].*?[\)）]', '', text)
        # 移除空白
        text = re.sub(r'\s+', '', text)
        # 移除数字
        text = re.sub(r'[0-9]+', '', text)
        return text

    # ========================================
    # 除外条款检测 (v17.1)
    # ========================================

    @staticmethod
    def is_exclusion_clause(lib_name: str) -> bool:
        """
        v17.1: 判断条款库中的条款是否为除外条款
        例如: "企业财产保险附加洪水除外条款（A款）"
        """
        if not lib_name:
            return False
        return '除外' in lib_name

    @staticmethod
    def client_wants_exclusion(client_title: str) -> bool:
        """
        v17.1: 判断客户条款名称是否明确包含"除外"
        只有当客户条款明确包含"除外"时，才应匹配除外类条款
        """
        if not client_title:
            return False
        return '除外' in client_title

    @staticmethod
    def extract_extra_info(text: str) -> str:
        """提取括号内额外信息"""
        if not isinstance(text, str):
            return ""
        matches = re.findall(r'([\(（].*?[\)）])', text)
        return " ".join(matches) if matches else ""

    @staticmethod
    def is_english(text: str) -> bool:
        """判断是否为英文"""
        if not isinstance(text, str) or len(text) <= 3:
            return False
        zh_count = len(re.findall(r'[\u4e00-\u9fa5]', text))
        return zh_count < len(text) * 0.15

    @staticmethod
    def is_bilingual(text: str) -> bool:
        """
        v17.0: 判断是否为中英混合文本
        如: "Earthquake Extension Clause 地震扩展条款"
        """
        if not isinstance(text, str) or len(text) < 5:
            return False
        zh_count = len(re.findall(r'[\u4e00-\u9fa5]', text))
        en_count = len(re.findall(r'[a-zA-Z]', text))
        total = len(text)
        # 中英文各占15%以上视为双语
        return (zh_count / total >= 0.15) and (en_count / total >= 0.15)

    @staticmethod
    def split_bilingual(text: str) -> Tuple[str, str]:
        """
        v17.0: 分离中英文部分
        返回: (中文部分, 英文部分)
        """
        if not isinstance(text, str):
            return "", ""
        # 提取中文（包括中文标点）
        cn_chars = re.findall(r'[\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef]+', text)
        cn_part = ''.join(cn_chars)
        # 提取英文单词
        en_words = re.findall(r'[a-zA-Z]+(?:\s+[a-zA-Z]+)*', text)
        en_part = ' '.join(en_words)
        return cn_part.strip(), en_part.strip().lower()

    @staticmethod
    def tokenize_chinese(text: str) -> List[str]:
        """
        v17.0: 中文分词
        使用jieba分词，如果不可用则使用字符级别分割
        """
        if not text:
            return []
        if HAS_JIEBA:
            # 使用jieba分词，过滤单字符和标点
            words = list(jieba.cut(text))
            return [w for w in words if len(w) > 1 and re.search(r'[\u4e00-\u9fa5a-zA-Z]', w)]
        else:
            # 降级：使用2-gram字符分割
            chars = re.findall(r'[\u4e00-\u9fa5]', text)
            if len(chars) < 2:
                return chars
            return [chars[i] + chars[i+1] for i in range(len(chars) - 1)]

    # ========================================
    # 相似度计算（v17.0 增强版）
    # ========================================

    @staticmethod
    def calculate_similarity(text1: str, text2: str) -> float:
        """
        混合相似度计算：
        - SequenceMatcher（序列匹配）
        - Levenshtein（编辑距离）
        取较高值
        """
        if not text1 or not text2:
            return 0.0

        # 序列匹配
        seq_ratio = difflib.SequenceMatcher(None, text1, text2).ratio()

        # 编辑距离（仅对较短文本使用，避免性能问题）
        if len(text1) <= 100 and len(text2) <= 100:
            lev_ratio = levenshtein_ratio(text1, text2)
            return max(seq_ratio, lev_ratio)

        return seq_ratio

    @classmethod
    def calculate_similarity_chinese(cls, text1: str, text2: str) -> float:
        """
        v17.0: 中文增强相似度计算
        结合词级别Jaccard相似度和字符级别相似度
        """
        if not text1 or not text2:
            return 0.0

        # 1. 字符级别相似度（基础）
        char_sim = cls.calculate_similarity(text1, text2)

        # 2. 词级别Jaccard相似度（如果jieba可用）
        if HAS_JIEBA:
            words1 = set(cls.tokenize_chinese(text1))
            words2 = set(cls.tokenize_chinese(text2))

            if words1 and words2:
                intersection = words1 & words2
                union = words1 | words2
                jaccard_sim = len(intersection) / len(union) if union else 0

                # 加权组合：词级别权重更高
                combined_sim = 0.6 * jaccard_sim + 0.4 * char_sim
                return max(combined_sim, char_sim)  # 取较高值

        return char_sim

    def calculate_bilingual_similarity(self, text1: str, text2: str) -> float:
        """
        v17.0: 双语条款相似度计算
        分别计算中英文部分的相似度，然后加权组合
        """
        if not text1 or not text2:
            return 0.0

        # 检查是否为双语文本
        is_bi1 = self.is_bilingual(text1)
        is_bi2 = self.is_bilingual(text2)

        if not is_bi1 and not is_bi2:
            # 都不是双语，使用标准相似度
            return self.calculate_similarity_chinese(text1, text2)

        # 分离中英文
        cn1, en1 = self.split_bilingual(text1)
        cn2, en2 = self.split_bilingual(text2)

        scores = []

        # 计算中文部分相似度
        if cn1 and cn2:
            cn_sim = self.calculate_similarity_chinese(cn1, cn2)
            scores.append(('cn', cn_sim, len(cn1) + len(cn2)))

        # 计算英文部分相似度
        if en1 and en2:
            en_sim = self.calculate_similarity(en1, en2)
            scores.append(('en', en_sim, len(en1) + len(en2)))

        # 尝试英文映射匹配
        if en1:
            mapped = self._get_client_mapping(en1)
            if mapped and cn2:
                map_sim = self.calculate_similarity_chinese(mapped, cn2)
                if map_sim > 0.7:
                    scores.append(('map', map_sim, 100))  # 高权重

        if not scores:
            return self.calculate_similarity(text1, text2)

        # 按长度加权平均
        total_weight = sum(s[2] for s in scores)
        weighted_sim = sum(s[1] * s[2] for s in scores) / total_weight if total_weight > 0 else 0

        # 取加权结果和最高单项的较大值
        max_sim = max(s[1] for s in scores)
        return max(weighted_sim, max_sim * 0.9)

    # ========================================
    # TF-IDF 向量匹配 (v17.0)
    # ========================================

    def build_tfidf_index(self, lib_data: List[Dict]) -> None:
        """
        v17.0: 构建TF-IDF向量索引，用于快速候选筛选
        """
        if not HAS_SKLEARN:
            logger.warning("sklearn不可用，跳过TF-IDF索引构建")
            return

        names = []
        for lib in lib_data:
            name = str(lib.get('条款名称', ''))
            if name.strip():
                # 对中文进行分词处理
                if HAS_JIEBA:
                    tokens = ' '.join(self.tokenize_chinese(name))
                    names.append(tokens if tokens else name)
                else:
                    names.append(name)

        if not names:
            return

        try:
            # 使用字符n-gram，适合中文
            self._tfidf_vectorizer = TfidfVectorizer(
                analyzer='char',
                ngram_range=(2, 4),
                max_features=5000,
                min_df=1
            )
            self._tfidf_vectors = self._tfidf_vectorizer.fit_transform(names)
            self._tfidf_names = names
            logger.info(f"TF-IDF索引构建完成，向量维度: {self._tfidf_vectors.shape}")
        except Exception as e:
            logger.warning(f"TF-IDF索引构建失败: {e}")
            self._tfidf_vectorizer = None
            self._tfidf_vectors = None

    def find_tfidf_candidates(self, query: str, top_k: int = 10) -> List[Tuple[int, float]]:
        """
        v17.0: 使用TF-IDF快速找到候选条款
        返回: [(索引, 相似度分数), ...]
        """
        if not HAS_SKLEARN or self._tfidf_vectorizer is None or self._tfidf_vectors is None:
            return []

        try:
            # 对查询进行同样的预处理
            if HAS_JIEBA:
                query_tokens = ' '.join(self.tokenize_chinese(query))
                query_text = query_tokens if query_tokens else query
            else:
                query_text = query

            query_vec = self._tfidf_vectorizer.transform([query_text])
            similarities = cosine_similarity(query_vec, self._tfidf_vectors).flatten()

            # 获取top_k个最相似的索引
            top_indices = np.argsort(similarities)[-top_k:][::-1]
            results = [(int(idx), float(similarities[idx])) for idx in top_indices if similarities[idx] > 0.1]

            return results
        except Exception as e:
            logger.debug(f"TF-IDF候选查找失败: {e}")
            return []

    # ========================================
    # 动态权重计算 (v17.0)
    # ========================================

    def calculate_dynamic_weight(self, title: str, content: str) -> Tuple[float, float]:
        """
        v17.0: 根据条款特征动态调整标题/内容权重
        返回: (标题权重, 内容权重)
        """
        # 默认权重
        title_weight = 0.7
        content_weight = 0.3

        title_len = len(title) if title else 0
        content_len = len(content) if content else 0

        # 情况1: 标题很短且内容丰富 -> 增加内容权重
        if title_len < 10 and content_len > 100:
            title_weight = 0.4
            content_weight = 0.6

        # 情况2: 标题很长（可能包含详细描述）-> 增加标题权重
        elif title_len > 30:
            title_weight = 0.8
            content_weight = 0.2

        # 情况3: 无内容 -> 全部使用标题
        elif content_len < 10:
            title_weight = 1.0
            content_weight = 0.0

        # 情况4: 标题包含特定关键词（表示具体条款类型）-> 增加标题权重
        specific_keywords = ['扩展条款', '除外条款', '特别条款', '附加险', 'extension', 'exclusion']
        if any(kw in title.lower() for kw in specific_keywords):
            title_weight = min(title_weight + 0.1, 0.9)
            content_weight = 1.0 - title_weight

        return title_weight, content_weight

    # ========================================
    # 索引构建（性能优化核心）
    # ========================================

    def build_index(self, lib_data: List[Dict]) -> LibraryIndex:
        """
        预构建条款库索引，加速匹配
        时间复杂度从 O(n*m) 降至 O(n + m)
        """
        logger.info(f"开始构建索引，条款数: {len(lib_data)}")

        index = LibraryIndex(data=lib_data)

        for i, lib in enumerate(lib_data):
            name = str(lib.get('条款名称', ''))
            if not name.strip():
                continue

            # 预计算清理结果（避免重复计算）
            name_norm = self.normalize_text(name)
            name_clean = self.clean_title(name)

            index.cleaned_cache[i] = {
                'norm': name_norm,
                'clean': name_clean,
                'original': name,
            }

            # 名称索引（精确匹配用）
            index.by_name_norm[name_norm] = i
            index.by_name_norm[name_clean] = i

            # 关键词倒排索引
            keywords = self._get_keywords(name)
            for kw in keywords:
                index.by_keyword[kw].append(i)

        logger.info(f"索引构建完成: {len(index.by_name_norm)} 名称, {len(index.by_keyword)} 关键词")
        self._index = index

        # v17.0: 构建TF-IDF索引
        self.build_tfidf_index(lib_data)

        return index

    @staticmethod
    def _fullwidth_to_halfwidth(text: str) -> str:
        """全角字符转半角"""
        result = []
        for char in text:
            code = ord(char)
            if code == 0x3000:  # 全角空格
                result.append(' ')
            elif 0xFF01 <= code <= 0xFF5E:  # 全角字符范围
                result.append(chr(code - 0xFEE0))
            else:
                result.append(char)
        return ''.join(result)

    def find_library_entry_by_name(self, target_name: str, index: LibraryIndex) -> Optional[Dict]:
        """
        根据名称在条款库中查找条目
        支持全角/半角字符模糊匹配
        """
        if not target_name:
            return None

        # 标准化目标名称
        target_norm = self._fullwidth_to_halfwidth(target_name.lower().strip())
        target_clean = re.sub(r'[^\u4e00-\u9fa5a-z0-9%]', '', target_norm)

        best_match_idx = -1
        best_score = 0.0

        for i, cached in index.cleaned_cache.items():
            lib_name = cached['original']
            lib_norm = self._fullwidth_to_halfwidth(lib_name.lower().strip())
            lib_clean = re.sub(r'[^\u4e00-\u9fa5a-z0-9%]', '', lib_norm)

            # 精确匹配（标准化后）
            if target_clean == lib_clean:
                return index.data[i]

            # 包含匹配
            if target_clean in lib_clean or lib_clean in target_clean:
                score = len(target_clean) / max(len(lib_clean), 1)
                if score > best_score:
                    best_score = score
                    best_match_idx = i

            # 相似度匹配
            sim = self.calculate_similarity(target_clean, lib_clean)
            if sim > best_score and sim > 0.8:
                best_score = sim
                best_match_idx = i

        if best_match_idx >= 0:
            return index.data[best_match_idx]

        return None

    @staticmethod
    def clean_reg_number(reg: str) -> str:
        """清理注册号，移除前缀"""
        if not reg:
            return ""
        # 移除 "产品注册号：" 等前缀
        reg = re.sub(r'^(产品)?注册号[：:]\s*', '', str(reg).strip())
        return reg

    # ========================================
    # 多级匹配策略（拆分重构）
    # ========================================

    def _try_exact_match(self, title_norm: str, title_clean: str,
                         index: LibraryIndex, original_title: str = "") -> Optional[Tuple[int, float]]:
        """
        级别1: 精确匹配 (v17.0 增强)
        - 标准化名称匹配
        - 英中映射表匹配
        - 双语条款分离匹配
        """
        # 标准化名称精确匹配
        if title_norm in index.by_name_norm:
            return index.by_name_norm[title_norm], 1.0

        # 清理后名称精确匹配
        if title_clean in index.by_name_norm:
            return index.by_name_norm[title_clean], self.thresholds.exact_min

        # v17.0: 英中映射表匹配
        if original_title:
            # 提取英文部分尝试映射
            _, en_part = self.split_bilingual(original_title)
            if en_part:
                mapped_cn = self._get_client_mapping(en_part)
                if mapped_cn:
                    mapped_norm = self.normalize_text(mapped_cn)
                    mapped_clean = self.clean_title(mapped_cn)
                    if mapped_norm in index.by_name_norm:
                        return index.by_name_norm[mapped_norm], 0.95
                    if mapped_clean in index.by_name_norm:
                        return index.by_name_norm[mapped_clean], 0.93
                    # 部分匹配
                    for i, cached in index.cleaned_cache.items():
                        if mapped_cn in cached['original'] or cached['original'] in mapped_cn:
                            return i, 0.90

        return None

    def _try_semantic_match(self, title: str, index: LibraryIndex) -> Optional[Tuple[int, float]]:
        """级别2: 语义别名匹配"""
        semantic_target = self._get_semantic_alias(title)
        if not semantic_target:
            return None

        # 在索引中查找目标
        for i, cached in index.cleaned_cache.items():
            if semantic_target in cached['original']:
                return i, self.thresholds.semantic_min

        return None

    def _try_keyword_match(self, title: str, index: LibraryIndex) -> Optional[Tuple[int, float]]:
        """级别3: 关键词匹配"""
        c_keywords = self._get_keywords(title)
        if not c_keywords:
            return None

        # 统计候选项得分
        candidate_scores: Dict[int, float] = defaultdict(float)

        for kw in c_keywords:
            if kw in index.by_keyword:
                for idx in index.by_keyword[kw]:
                    candidate_scores[idx] += 1

        if not candidate_scores:
            return None

        # 找最高分候选
        best_idx = max(candidate_scores, key=candidate_scores.get)
        best_count = candidate_scores[best_idx]

        # 计算关键词匹配度
        l_keywords = self._get_keywords(index.cleaned_cache[best_idx]['original'])
        if l_keywords:
            keyword_ratio = best_count / max(len(c_keywords), len(l_keywords))
            if keyword_ratio >= 0.5:
                score = self.thresholds.keyword_min + keyword_ratio * 0.2
                return best_idx, score

        return None

    def _try_fuzzy_match(self, title_clean: str, content: str,
                         index: LibraryIndex, is_title_only: bool,
                         original_title: str = "", max_results: int = 1) -> Any:
        """
        级别4: 模糊匹配 (v17.1 增强版)
        - 使用TF-IDF快速候选筛选
        - 使用中文分词增强相似度
        - 支持双语匹配
        - 动态权重调整
        - v17.1: 支持返回多条匹配结果
        - v17.1: 除外条款过滤

        Args:
            max_results: 返回结果数量，1为单个结果(兼容旧接口)，>1为多个结果列表

        Returns:
            当max_results=1时: Tuple[int, float, float, float] - (idx, score, title_sim, content_sim)
            当max_results>1时: List[Tuple[int, float, float, float]] - 候选列表
        """
        # v17.1: 检查客户是否需要除外类条款
        wants_exclusion = self.client_wants_exclusion(original_title)

        candidates = []

        # v17.0: 计算动态权重
        title_weight, content_weight = self.calculate_dynamic_weight(title_clean, content)

        # v17.0: 使用TF-IDF快速筛选候选（如果可用）
        candidate_indices = set()
        tfidf_candidates = self.find_tfidf_candidates(original_title or title_clean, top_k=20)
        if tfidf_candidates:
            candidate_indices = {idx for idx, _ in tfidf_candidates}
            # 同时也检查所有条款（以防TF-IDF遗漏）
            # 但优先处理TF-IDF候选
        else:
            candidate_indices = set(index.cleaned_cache.keys())

        # 如果TF-IDF候选较少，添加所有条款确保覆盖
        if len(candidate_indices) < 10:
            candidate_indices = set(index.cleaned_cache.keys())

        for i in candidate_indices:
            if i not in index.cleaned_cache:
                continue
            cached = index.cleaned_cache[i]
            l_name_clean = cached['clean']
            l_name_original = cached['original']

            # v17.1: 除外条款过滤 - 除非客户明确需要除外条款，否则跳过库内的除外条款
            if not wants_exclusion and self.is_exclusion_clause(l_name_original):
                continue

            # v17.0: 使用增强相似度计算
            # 先检查是否为双语匹配
            if original_title and (self.is_bilingual(original_title) or self.is_bilingual(l_name_original)):
                title_sim = self.calculate_bilingual_similarity(original_title, l_name_original)
            else:
                # 使用中文增强相似度
                title_sim = self.calculate_similarity_chinese(title_clean, l_name_clean)

            # 内容相似度
            content_sim = 0.0
            if not is_title_only and content.strip():
                c_content_clean = self.clean_content(content)
                l_content = str(index.data[i].get('条款内容', ''))
                l_content_clean = self.clean_content(l_content)
                if c_content_clean and l_content_clean:
                    # v17.0: 对内容也使用中文增强相似度
                    content_sim = self.calculate_similarity_chinese(c_content_clean, l_content_clean)

            # v17.0: 使用动态权重加权得分
            if is_title_only or not content.strip():
                score = title_sim
            else:
                score = title_weight * title_sim + content_weight * content_sim

            # 惩罚项
            if self._is_penalty_keyword(cached['original']) and not self._is_penalty_keyword(title_clean):
                score -= 0.5

            if score > self.thresholds.accept_min:
                candidates.append((i, score, title_sim, content_sim))

        # 按分数降序排序
        candidates.sort(key=lambda x: x[1], reverse=True)

        # v17.1: 根据max_results返回不同格式
        if max_results == 1:
            # 兼容旧接口
            if candidates:
                return candidates[0]
            return (-1, 0.0, 0.0, 0.0)
        else:
            # 返回多条结果
            return candidates[:max_results]

    def match_clause(self, clause: ClauseItem, index: LibraryIndex,
                     is_title_only: bool) -> MatchResult:
        """
        主匹配入口 - 多级策略
        优先级: 精确 > 语义 > 关键词 > 模糊
        """
        result = MatchResult()
        title = clause.title
        content = clause.content

        title_clean = self.clean_title(title)
        title_norm = self.normalize_text(title)

        matched_idx = -1
        match_level = MatchLevel.NONE
        score = 0.0
        title_score = 0.0
        content_score = 0.0

        # === 级别1: 精确匹配 (v17.0: 传递原始标题用于英中映射) ===
        exact_result = self._try_exact_match(title_norm, title_clean, index, original_title=title)
        if exact_result:
            matched_idx, score = exact_result
            match_level = MatchLevel.EXACT
            title_score = score

        # === 级别2: 语义匹配 ===
        if matched_idx < 0:
            semantic_result = self._try_semantic_match(title, index)
            if semantic_result:
                matched_idx, score = semantic_result
                match_level = MatchLevel.SEMANTIC
                title_score = score

        # === 级别3: 关键词匹配 ===
        if matched_idx < 0:
            keyword_result = self._try_keyword_match(title, index)
            if keyword_result:
                matched_idx, score = keyword_result
                match_level = MatchLevel.KEYWORD
                title_score = score

        # === 级别4: 模糊匹配 (v17.0 增强) ===
        if matched_idx < 0:
            fuzzy_idx, fuzzy_score, t_sim, c_sim = self._try_fuzzy_match(
                title_clean, content, index, is_title_only,
                original_title=title  # v17.0: 传递原始标题用于双语匹配
            )
            if fuzzy_score > self.thresholds.accept_min:
                matched_idx = fuzzy_idx
                score = fuzzy_score
                match_level = MatchLevel.FUZZY
                title_score = t_sim
                content_score = c_sim

        # 构建结果
        if matched_idx >= 0 and score > self.thresholds.accept_min:
            lib = index.data[matched_idx]
            base_name = lib.get('条款名称', '')
            extra_params = self.extract_extra_info(clause.original_title or clause.title)

            result.matched_name = f"{base_name} {extra_params}".strip() if extra_params else base_name
            result.matched_content = lib.get('条款内容', '')
            result.matched_reg = lib.get('产品注册号', lib.get('注册号', ''))
            result.score = max(0, score)
            result.title_score = title_score
            result.content_score = content_score
            result.match_level = match_level

            # 差异分析（低分时）
            if score < 0.6:
                result.diff_analysis = self.analyze_difference(content, result.matched_content)

        return result

    def match_clause_multiple(self, clause: ClauseItem, index: LibraryIndex,
                               is_title_only: bool, max_results: int = 3) -> List[MatchResult]:
        """
        v17.1: 多结果匹配入口
        返回最多max_results条匹配结果供用户选择

        Args:
            clause: 待匹配的条款
            index: 条款库索引
            is_title_only: 是否仅匹配标题
            max_results: 最多返回结果数，默认3条

        Returns:
            List[MatchResult]: 匹配结果列表，按分数降序排列
        """
        title = clause.title
        content = clause.content
        original_title = clause.original_title or title

        title_clean = self.clean_title(title)

        results = []
        seen_names = set()

        # 获取多条模糊匹配候选
        fuzzy_candidates = self._try_fuzzy_match(
            title_clean, content, index, is_title_only,
            original_title=original_title,
            max_results=max_results + 5  # 多获取一些以便去重
        )

        # fuzzy_candidates是列表: [(idx, score, title_sim, content_sim), ...]
        if isinstance(fuzzy_candidates, tuple):
            # 单结果模式返回的tuple
            if fuzzy_candidates[0] >= 0:
                fuzzy_candidates = [fuzzy_candidates]
            else:
                fuzzy_candidates = []

        for idx, score, title_sim, content_sim in fuzzy_candidates:
            if len(results) >= max_results:
                break

            if score <= self.thresholds.accept_min:
                continue

            lib = index.data[idx]
            base_name = lib.get('条款名称', '')

            # 去重
            if base_name in seen_names:
                continue
            seen_names.add(base_name)

            extra_params = self.extract_extra_info(original_title)

            result = MatchResult(
                matched_name=f"{base_name} {extra_params}".strip() if extra_params else base_name,
                matched_content=lib.get('条款内容', ''),
                matched_reg=lib.get('产品注册号', lib.get('注册号', '')),
                score=max(0, score),
                title_score=title_sim,
                content_score=content_sim,
                match_level=MatchLevel.FUZZY,
                diff_analysis=""
            )

            # 差异分析（低分时）
            if score < 0.6 and content:
                result.diff_analysis = self.analyze_difference(content, result.matched_content)

            results.append(result)

        # 如果没有任何匹配，返回空匹配结果
        if not results:
            results.append(MatchResult())

        return results

    def search_library_titles(self, query: str, index: LibraryIndex,
                               max_results: int = 5) -> List[Dict]:
        """
        v17.1: 条款查询功能 - 仅查询条款标题
        用于快速模糊查询条款库中的条款

        Args:
            query: 查询字符串（条款名称或关键词）
            index: 条款库索引
            max_results: 最多返回结果数，默认5条

        Returns:
            List[Dict]: 查询结果列表，每项包含:
                - name: 条款名称
                - content: 条款内容
                - reg: 注册号
                - score: 匹配分数
                - matchType: 匹配类型
        """
        if not query or not index.data:
            return []

        # 检查查询是否需要除外条款
        wants_exclusion = self.client_wants_exclusion(query)

        query_lower = query.lower()
        query_clean = self.clean_title(query)
        query_norm = self.normalize_text(query)

        candidates = []

        for i, cached in index.cleaned_cache.items():
            lib_name = cached['original']
            lib_name_clean = cached['clean']
            lib_name_norm = cached['norm']
            lib_name_lower = lib_name.lower()

            # 除外条款过滤
            if not wants_exclusion and self.is_exclusion_clause(lib_name):
                continue

            match_type = ""
            score = 0.0

            # 精确匹配
            if query_norm == lib_name_norm or query_clean == lib_name_clean:
                match_type = "exact"
                score = 1.0
            # 包含匹配
            elif query_lower in lib_name_lower or lib_name_lower in query_lower:
                match_type = "contain"
                score = 0.9
            elif query_clean in lib_name_clean or lib_name_clean in query_clean:
                match_type = "contain"
                score = 0.85
            else:
                # 模糊相似度匹配
                title_sim = self.calculate_similarity_chinese(query_clean, lib_name_clean)
                if title_sim > 0.3:  # 较低的阈值以便显示更多可能的匹配
                    match_type = "fuzzy"
                    score = title_sim

            if score > 0:
                lib = index.data[i]
                candidates.append({
                    'name': lib_name,
                    'content': lib.get('条款内容', ''),
                    'reg': lib.get('产品注册号', lib.get('注册号', '')),
                    'score': score,
                    'matchType': match_type
                })

        # 按分数降序排序
        candidates.sort(key=lambda x: x['score'], reverse=True)

        # 去重并限制结果数量
        results = []
        seen_names = set()
        for c in candidates:
            if len(results) >= max_results:
                break
            if c['name'] not in seen_names:
                seen_names.add(c['name'])
                results.append(c)

        return results

    # ========================================
    # 翻译和差异分析
    # ========================================

    def translate_title(self, title: str) -> Tuple[str, bool]:
        """翻译英文标题"""
        if not self.is_english(title):
            return title, False

        title_norm = self.normalize_text(title)

        # 1. 查询映射
        mapped = self._get_client_mapping(title_norm)
        if mapped:
            return mapped, True

        # 2. 部分匹配
        client_map = (self.config.client_en_cn_map if self._use_external_config
                      else DefaultConfig.CLIENT_EN_CN_MAP)
        for eng, chn in client_map.items():
            if eng in title_norm or title_norm in eng:
                return chn, True

        # 3. 在线翻译
        if HAS_TRANSLATOR:
            try:
                translated = GoogleTranslator(source='auto', target='zh-CN').translate(title)
                logger.debug(f"在线翻译: {title} -> {translated}")
                return translated, True
            except ConnectionError as e:
                logger.warning(f"翻译服务连接失败: {e}")
            except TimeoutError as e:
                logger.warning(f"翻译服务超时: {e}")
            except Exception as e:
                logger.error(f"翻译失败: {type(e).__name__}: {e}")

        return title, False

    @staticmethod
    def analyze_difference(c_content: str, l_content: str) -> str:
        """分析保障差异"""
        c_text, l_text = str(c_content), str(l_content)
        if not c_text.strip():
            return ""

        analysis = []
        keywords = {
            "限额": ["Limit", "限额", "最高", "limit"],
            "免赔": ["Deductible", "Excess", "免赔", "deductible"],
            "除外": ["Exclusion", "除外", "不负责", "exclusion"],
            "观察期": ["Waiting Period", "观察期", "等待期"],
            "赔偿期": ["Indemnity Period", "赔偿期间"],
        }

        for key, words in keywords.items():
            c_has = any(w.lower() in c_text.lower() for w in words)
            l_has = any(w.lower() in l_text.lower() for w in words)
            if c_has and not l_has:
                analysis.append(f"⚠️ 客户提及[{key}]但库内未提及")
            elif not c_has and l_has:
                analysis.append(f"ℹ️ 库内包含[{key}]但客户未提及")

        return " | ".join(analysis)

    # ========================================
    # 文档解析
    # ========================================

    @staticmethod
    def is_likely_title(text: str) -> bool:
        """
        判断是否像标题（严格模式）
        只有明确符合标题特征的才返回True
        v17.1: 增强过滤规则
        """
        if not text or len(text) < 3:
            return False

        # 太长的不是标题
        if len(text) > 100:
            return False

        # 以句号等结尾的通常是内容
        if text.endswith(('。', '；', '.', ';', '，', ',')):
            return False

        # ===== v17.1: 优先检查是否为标题（"条款"关键词最优先）=====

        # 包含"条款"关键词，但排除以"本条款"、"本扩展条款"、"本附加条款"开头的内容句
        # 这个检查必须在 descriptive_keywords 之前！否则"恢复保险金额条款"会被错误排除
        if '条款' in text:
            if text.startswith(('本条款', '本扩展条款', '本附加条款')):
                return False
            return True

        # ===== 排除明确不是标题的内容 =====

        # 1. 排除包含金额的内容（如 "RMB50万元"、"CNY5000万元"、"人民币100万"）
        money_pattern = r'(RMB|CNY|人民币|美元|USD|EUR|HKD|港币)?\s*\d+[\d,\.]*\s*(万元|元|万|亿|千元)'
        if re.search(money_pattern, text, re.IGNORECASE):
            return False

        # 2. 排除包含"赔偿限额"、"保险金额"等描述性文字的内容
        # 注意：如果包含"条款"关键词，上面已经返回True，不会到达这里
        descriptive_keywords = ['赔偿限额', '保险金额', '责任限额', '每次事故', '累计赔偿',
                                '免赔额', '自负额', '保险费', '费率', '保险期间']
        if any(kw in text for kw in descriptive_keywords):
            return False

        # 3. 排除特定的内容句（完整匹配或开头匹配）
        excluded_exact = [
            '本扩展条款受下列条件限制',
            '特约扩展责任',
        ]
        if text in excluded_exact:
            return False

        # ===== 其他标题模式检查 =====

        # 带数字编号的条款标题（如 "35、码头吊机、铁路车辆第三者责任险"）
        # 支持格式：1、xxx, 1.xxx, 1）xxx, (1) xxx, 一、xxx 等
        numbered_title_pattern = r'^(\d+|[一二三四五六七八九十]+)[、\.．）\)]'
        if re.match(numbered_title_pattern, text):
            # 但如果后面是描述性内容则排除
            title_part = re.sub(numbered_title_pattern, '', text).strip()
            if title_part and len(title_part) > 3 and not title_part.endswith(('。', '；', '，')):
                # 检查是否包含"险"、"条款"等标志性词汇
                if any(kw in title_part for kw in ['险', '条款', '责任', '扩展', '附加']):
                    return True

        # 附加保险条款，以 "(XXXX版)" 结尾（无"条款"字样）
        # 如：平安产险企业财产保险附加提前60天通知解除保单保险（2025版）
        if '附加' in text and '保险' in text and re.search(r'[（(]\d{4}版?[）)]$', text):
            return True

        # ===== 明确是内容的模式（不是标题）=====
        content_start_patterns = [
            # 条款内容常见开头
            r'^经双方同意',
            r'^兹经双方同意',
            r'^兹经保险',
            r'^兹经合同',
            r'^发生.*损失',
            r'^如果.*保险',
            r'^本保单',
            r'^本保险',
            r'^本条款',
            r'^本款项',
            r'^本公司',
            r'^本扩展条款',  # v17.1
            r'^本附加条款',  # v17.1
            r'^保险人',
            r'^被保险人',
            r'^投保人',
            r'^对于',
            r'^若',
            r'^但',
            r'^在保',
            r'^上述',
            r'^该',
            r'^其中',  # v17.1
            r'^此',
            r'^当',
            r'^财产险',
            r'^除',
            r'^凡',
            r'^任何',
            # 列表项（子条目，不是新条款）
            r'^[\(（]\s*[一二三四五六七八九十]+\s*[\)）]',  # (一)、（二）
            r'^[一二三四五六七八九十]+[、\.．]',  # 一、二、
            r'^\d+[、\.．\s](?![\.．\s]*[^\d].*条款)',  # 1、2、但不匹配 "1. xxx条款"
            r'^[\(（]\s*\d+\s*[\)）]',  # (1)、（2）
        ]

        for pattern in content_start_patterns:
            if re.match(pattern, text):
                return False

        # ===== 其他标题模式（已通过内容排除检查）=====
        # 英文条款关键词
        if re.search(r'\b(Clause|Extension|Coverage|Insurance)\b', text, re.IGNORECASE):
            return True

        # 全大写英文（可能是英文条款名）
        if text.isupper() and len(text) > 5 and re.search(r'[A-Z]{3,}', text):
            return True

        # 默认不是标题（保守策略）
        return False

    def parse_docx(self, doc_path: str) -> Tuple[List[ClauseItem], bool]:
        """解析Word文档 - 基于标题识别的智能分割"""
        logger.info(f"解析文档: {doc_path}")

        try:
            doc = Document(doc_path)
        except Exception as e:
            logger.error(f"文档打开失败: {e}")
            raise ValueError(f"无法打开文档: {e}")

        # 1. 读取普通段落
        all_lines = [p.text.strip() for p in doc.paragraphs]

        # 2. 读取表格中的内容
        table_lines = []
        for table in doc.tables:
            for row in table.rows:
                row_text = ' '.join(cell.text.strip() for cell in row.cells if cell.text.strip())
                if row_text:
                    table_lines.append(row_text)

        # 如果表格有内容且段落基本为空，优先使用表格内容
        non_empty_paragraphs = [l for l in all_lines if l]
        if table_lines and len(non_empty_paragraphs) < len(table_lines):
            logger.info(f"检测到表格内容: {len(table_lines)} 行，优先使用表格")
            all_lines = table_lines
        elif table_lines:
            logger.info(f"合并段落({len(non_empty_paragraphs)})和表格({len(table_lines)})内容")
            all_lines.extend(table_lines)

        # 过滤空行
        non_empty_lines = [l for l in all_lines if l]
        logger.info(f"非空行数: {len(non_empty_lines)}")

        # 3. 基于标题识别进行分割（不再依赖空行）
        clauses = []
        current_title = None
        current_content = []

        for line in non_empty_lines:
            if self.is_likely_title(line):
                # 保存前一个条款
                if current_title is not None:
                    clauses.append(ClauseItem(
                        title=current_title,
                        content="\n".join(current_content),
                        original_title=current_title
                    ))
                # 开始新条款
                current_title = line
                current_content = []
            else:
                # 内容行
                if current_title is not None:
                    current_content.append(line)
                else:
                    # 没有标题的内容，作为独立条款
                    clauses.append(ClauseItem(
                        title=line,
                        content="",
                        original_title=line
                    ))

        # 保存最后一个条款
        if current_title is not None:
            clauses.append(ClauseItem(
                title=current_title,
                content="\n".join(current_content),
                original_title=current_title
            ))

        is_title_only = all(not c.content for c in clauses)
        logger.info(f"解析完成: {len(clauses)} 条款, 纯标题模式: {is_title_only}")

        return clauses, is_title_only


# ==========================================
# 条款库加载器
# ==========================================
class LibraryLoader:
    """条款库加载器 - 支持自动列名识别和多Sheet选择"""

    @staticmethod
    def get_sheet_names(excel_path: str) -> List[str]:
        """
        获取Excel文件中所有Sheet名称
        """
        try:
            xl = pd.ExcelFile(excel_path)
            return xl.sheet_names
        except Exception as e:
            logger.warning(f"读取Sheet列表失败: {e}")
            return []

    @staticmethod
    def load_excel(excel_path: str, header_row: int = None, sheet_name: str = None) -> List[Dict]:
        """
        加载Excel条款库
        自动识别列名和表头行

        Args:
            excel_path: Excel文件路径
            header_row: 表头行索引（自动检测时为None）
            sheet_name: Sheet名称（None时使用第一个Sheet）
        """
        logger.info(f"加载条款库: {excel_path}, Sheet: {sheet_name or '默认'}")

        try:
            # 自动检测表头行
            read_params = {'header': None, 'nrows': 5}
            if sheet_name:
                read_params['sheet_name'] = sheet_name

            if header_row is None:
                # 先读取前几行检测表头
                df_test = pd.read_excel(excel_path, **read_params)
                header_row = 0  # 默认第0行
                for i in range(min(3, len(df_test))):
                    row_values = [str(v).lower() if pd.notna(v) else '' for v in df_test.iloc[i]]
                    # 检查是否包含表头关键词
                    if any('条款' in v or 'name' in v or '名称' in v for v in row_values):
                        header_row = i
                        break
                logger.info(f"自动检测表头行: {header_row}")

            read_params = {'header': header_row}
            if sheet_name:
                read_params['sheet_name'] = sheet_name
            df = pd.read_excel(excel_path, **read_params)
        except FileNotFoundError:
            raise ValueError(f"文件不存在: {excel_path}")
        except Exception as e:
            raise ValueError(f"Excel读取失败: {e}")

        df.columns = [str(c).strip() for c in df.columns]

        # 自动识别列名
        name_col = None
        content_col = None
        reg_col = None

        for col in df.columns:
            col_lower = col.lower()
            if name_col is None and ('条款名称' in col or '名称' in col or 'name' in col_lower):
                name_col = col
            elif content_col is None and ('条款内容' in col or '内容' in col or 'content' in col_lower):
                content_col = col
            elif reg_col is None and ('注册号' in col or '产品' in col or 'reg' in col_lower):
                reg_col = col

        # 回退到位置
        if not name_col and len(df.columns) > 0:
            name_col = df.columns[0]
        if not content_col and len(df.columns) > 2:
            content_col = df.columns[2]
        if not reg_col and len(df.columns) > 1:
            reg_col = df.columns[1]

        logger.info(f"列名识别: 名称={name_col}, 内容={content_col}, 注册号={reg_col}")

        # 构建数据
        lib_data = []
        for _, row in df.iterrows():
            name = str(row.get(name_col, '')) if pd.notna(row.get(name_col)) else ''
            if not name.strip():
                continue

            lib_data.append({
                '条款名称': name,
                '条款内容': str(row.get(content_col, '')) if content_col and pd.notna(row.get(content_col)) else '',
                '产品注册号': str(row.get(reg_col, '')) if reg_col and pd.notna(row.get(reg_col)) else '',
            })

        logger.info(f"加载完成: {len(lib_data)} 条有效记录")
        return lib_data


# ==========================================
# Excel样式器
# ==========================================
class ExcelStyler:
    """Excel样式应用器"""

    FILLS = {
        'green': PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        'yellow': PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        'red': PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        'blue': PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid"),
        'header': PatternFill(start_color="667eea", end_color="667eea", fill_type="solid"),
    }

    BORDER = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    # v17.1: 新格式列宽（支持3组匹配结果）
    # A=序号, B=客户条款(原), C=客户条款(译), D=客户原始内容
    # E-I=匹配1, J-N=匹配2, O-S=匹配3
    WIDTHS = {
        'A': 6, 'B': 35, 'C': 30, 'D': 45,
        # 匹配1
        'E': 40, 'F': 25, 'G': 50, 'H': 10, 'I': 12,
        # 匹配2
        'J': 40, 'K': 25, 'L': 50, 'M': 10, 'N': 12,
        # 匹配3
        'O': 40, 'P': 25, 'Q': 50, 'R': 10, 'S': 12,
    }

    @classmethod
    def apply_styles(cls, output_path: str):
        """应用Excel样式"""
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active

        # 表头
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = cls.FILLS['header']
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = cls.BORDER

        # 列宽
        for col, width in cls.WIDTHS.items():
            ws.column_dimensions[col].width = width

        # 数据行
        # v17.1: 新格式匹配度和匹配级别列索引
        # 匹配1: H(8)=匹配度, I(9)=级别
        # 匹配2: M(13)=匹配度, N(14)=级别
        # 匹配3: R(18)=匹配度, S(19)=级别
        score_cols = {8, 13, 18}  # 匹配度列索引
        level_cols = {9, 14, 19}  # 匹配级别列索引

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                cell.border = cls.BORDER

                # v17.1: 匹配度着色（支持3组）
                if cell.col_idx in score_cols:
                    try:
                        val = float(cell.value) if cell.value else 0
                        if val >= 0.8:
                            cell.fill = cls.FILLS['green']
                        elif val >= 0.5:
                            cell.fill = cls.FILLS['yellow']
                        elif val > 0:
                            cell.fill = cls.FILLS['red']
                    except (ValueError, TypeError):
                        pass

                # v17.1: 匹配级别着色（支持3组）
                if cell.col_idx in level_cols:
                    val = str(cell.value) if cell.value else ""
                    if "精确" in val:
                        cell.fill = cls.FILLS['green']
                    elif "语义" in val:
                        cell.fill = cls.FILLS['blue']
                    elif "关键词" in val:
                        cell.fill = cls.FILLS['yellow']

        # 冻结首行
        ws.freeze_panes = 'A2'

        wb.save(output_path)
        logger.info(f"Excel样式已应用: {output_path}")


# ==========================================
# 工作线程
# ==========================================
class MatchWorker(QThread):
    """单文件匹配工作线程"""
    log_signal = pyqtSignal(str, str)
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, doc_path: str, excel_path: str, output_path: str, sheet_name: str = None):
        super().__init__()
        self.doc_path = doc_path
        self.excel_path = excel_path
        self.output_path = output_path
        self.sheet_name = sheet_name  # 指定的Sheet名称

    def run(self):
        try:
            logic = ClauseMatcherLogic()

            # 状态信息
            self.log_signal.emit(f"📊 配置: 外部={logic._use_external_config}, 翻译={HAS_TRANSLATOR}", "info")

            # 解析文档
            self.log_signal.emit("⏳ 正在解析文档...", "info")
            clauses, is_title_only = logic.parse_docx(self.doc_path)
            mode_str = "纯标题模式" if is_title_only else "完整内容模式"
            self.log_signal.emit(f"📖 [{mode_str}] 提取到 {len(clauses)} 条", "success")

            # 加载条款库
            sheet_info = f" [{self.sheet_name}]" if self.sheet_name else ""
            self.log_signal.emit(f"📚 加载条款库{sheet_info}...", "info")
            lib_data = LibraryLoader.load_excel(self.excel_path, sheet_name=self.sheet_name)
            self.log_signal.emit(f"✓ 条款库 {len(lib_data)} 条", "success")

            # 构建索引
            self.log_signal.emit("🔧 构建索引...", "info")
            index = logic.build_index(lib_data)
            self.log_signal.emit(f"✓ 索引完成", "success")

            # 开始匹配 (v17.1 多结果匹配)
            self.log_signal.emit("🧠 开始智能匹配（v17.1 多结果模式）...", "info")
            results = []
            stats = {'exact': 0, 'semantic': 0, 'keyword': 0, 'fuzzy': 0, 'none': 0}

            for idx, clause in enumerate(clauses, 1):
                self.progress_signal.emit(idx, len(clauses))

                # 翻译
                original_title = clause.title
                translated_title, was_translated = logic.translate_title(clause.title)
                if was_translated:
                    clause.title = translated_title
                    clause.original_title = original_title

                # 检查用户自定义映射
                user_library_name = None
                if HAS_MAPPING_MANAGER:
                    mapping_mgr = get_mapping_manager()
                    # 按原标题或翻译后标题查找
                    user_library_name = mapping_mgr.get_library_name(original_title)
                    if not user_library_name and was_translated:
                        user_library_name = mapping_mgr.get_library_name(translated_title)

                # v17.1: 根据是否有用户映射决定匹配策略
                match_results = []
                if user_library_name:
                    # 有用户映射，只返回映射的那一条
                    lib_entry = logic.find_library_entry_by_name(user_library_name, index)
                    if lib_entry:
                        mapped_result = MatchResult(
                            matched_name=lib_entry.get('条款名称', user_library_name),
                            matched_reg=logic.clean_reg_number(lib_entry.get('产品注册号', lib_entry.get('注册号', ''))),
                            matched_content=lib_entry.get('条款内容', ''),
                            score=1.0,
                            match_level=MatchLevel.EXACT,
                            diff_analysis="用户自定义映射",
                            title_score=1.0,
                            content_score=0.0,
                        )
                    else:
                        mapped_result = MatchResult(
                            matched_name=user_library_name,
                            matched_reg="",
                            matched_content="",
                            score=1.0,
                            match_level=MatchLevel.EXACT,
                            diff_analysis="用户自定义映射（未在库中找到）",
                            title_score=1.0,
                            content_score=0.0,
                        )
                    match_results = [mapped_result]
                else:
                    # 无用户映射，使用多结果匹配（最多3条）
                    match_results = logic.match_clause_multiple(clause, index, is_title_only, max_results=3)

                # 统计使用第一个匹配结果
                primary_match = match_results[0] if match_results else MatchResult()
                if primary_match.match_level == MatchLevel.EXACT:
                    stats['exact'] += 1
                elif primary_match.match_level == MatchLevel.SEMANTIC:
                    stats['semantic'] += 1
                elif primary_match.match_level == MatchLevel.KEYWORD:
                    stats['keyword'] += 1
                elif primary_match.match_level == MatchLevel.FUZZY:
                    stats['fuzzy'] += 1
                else:
                    stats['none'] += 1

                # v17.1: 构建多结果行
                row = {
                    ExcelColumns.SEQ: idx,
                    ExcelColumns.CLIENT_ORIG: original_title,
                    ExcelColumns.CLIENT_TRANS: translated_title if was_translated else "",
                    ExcelColumns.CLIENT_CONTENT: clause.content[:500] if clause.content else "",
                }

                # 填充最多3条匹配结果
                for match_num in range(1, 4):
                    if match_num <= len(match_results):
                        mr = match_results[match_num - 1]
                        row[f'匹配{match_num}_条款名称'] = mr.matched_name or ""
                        row[f'匹配{match_num}_注册号'] = logic.clean_reg_number(mr.matched_reg)
                        row[f'匹配{match_num}_条款内容'] = mr.matched_content[:500] if mr.matched_content else ""
                        row[f'匹配{match_num}_匹配度'] = round(mr.score, 3)
                        row[f'匹配{match_num}_匹配级别'] = mr.match_level.value
                    else:
                        row[f'匹配{match_num}_条款名称'] = ""
                        row[f'匹配{match_num}_注册号'] = ""
                        row[f'匹配{match_num}_条款内容'] = ""
                        row[f'匹配{match_num}_匹配度'] = ""
                        row[f'匹配{match_num}_匹配级别'] = ""

                results.append(row)

            # 保存结果
            df_res = pd.DataFrame(results)
            df_res.to_excel(self.output_path, index=False)
            ExcelStyler.apply_styles(self.output_path)

            # 输出统计
            self.log_signal.emit(f"📊 匹配统计:", "info")
            self.log_signal.emit(f"   精确匹配: {stats['exact']}", "success")
            self.log_signal.emit(f"   语义匹配: {stats['semantic']}", "success")
            self.log_signal.emit(f"   关键词匹配: {stats['keyword']}", "info")
            self.log_signal.emit(f"   模糊匹配: {stats['fuzzy']}", "warning")
            self.log_signal.emit(f"   无匹配: {stats['none']}", "error")

            self.log_signal.emit(f"🎉 完成！", "success")
            self.log_signal.emit(f"💡 提示: 报告中每个客户条款最多显示3条匹配结果供您选择", "info")
            self.finished_signal.emit(True, self.output_path)

        except Exception as e:
            logger.exception("匹配过程出错")
            self.log_signal.emit(f"❌ 错误: {str(e)}", "error")
            self.finished_signal.emit(False, str(e))


class BatchMatchWorker(QThread):
    """批量匹配工作线程"""
    log_signal = pyqtSignal(str, str)
    progress_signal = pyqtSignal(int, int)
    batch_progress_signal = pyqtSignal(int, int, str)  # 当前文件, 总数, 文件名
    finished_signal = pyqtSignal(bool, str, int, int)  # 成功, 消息, 成功数, 总数

    def __init__(self, doc_paths: List[str], excel_path: str, output_dir: str, sheet_name: str = None):
        super().__init__()
        self.doc_paths = doc_paths
        self.excel_path = excel_path
        self.output_dir = output_dir
        self.sheet_name = sheet_name  # 指定的Sheet名称

    def run(self):
        try:
            logic = ClauseMatcherLogic()

            # 加载条款库（只需一次）
            sheet_info = f" [{self.sheet_name}]" if self.sheet_name else ""
            self.log_signal.emit(f"📚 加载条款库{sheet_info}...", "info")
            lib_data = LibraryLoader.load_excel(self.excel_path, sheet_name=self.sheet_name)
            self.log_signal.emit(f"✓ 条款库 {len(lib_data)} 条", "success")

            # 构建索引（只需一次）
            self.log_signal.emit("🔧 构建索引...", "info")
            index = logic.build_index(lib_data)

            success_count = 0
            total = len(self.doc_paths)

            for file_idx, doc_path in enumerate(self.doc_paths, 1):
                file_name = Path(doc_path).name
                self.batch_progress_signal.emit(file_idx, total, file_name)
                self.log_signal.emit(f"\n📄 [{file_idx}/{total}] {file_name}", "info")

                try:
                    # 解析文档
                    clauses, is_title_only = logic.parse_docx(doc_path)
                    self.log_signal.emit(f"   提取 {len(clauses)} 条款", "info")

                    # 匹配 (v17.1 多结果匹配)
                    results = []
                    mapping_mgr = get_mapping_manager() if HAS_MAPPING_MANAGER else None

                    for idx, clause in enumerate(clauses, 1):
                        original_title = clause.title
                        translated_title, was_translated = logic.translate_title(clause.title)
                        if was_translated:
                            clause.title = translated_title
                            clause.original_title = original_title

                        # 检查用户自定义映射
                        user_library_name = None
                        if mapping_mgr:
                            user_library_name = mapping_mgr.get_library_name(original_title)
                            if not user_library_name and was_translated:
                                user_library_name = mapping_mgr.get_library_name(translated_title)

                        # v17.1: 根据是否有用户映射决定匹配策略
                        match_results = []
                        if user_library_name:
                            # 有用户映射，只返回映射的那一条
                            lib_entry = logic.find_library_entry_by_name(user_library_name, index)
                            if lib_entry:
                                mapped_result = MatchResult(
                                    matched_name=lib_entry.get('条款名称', user_library_name),
                                    matched_reg=logic.clean_reg_number(lib_entry.get('产品注册号', lib_entry.get('注册号', ''))),
                                    matched_content=lib_entry.get('条款内容', ''),
                                    score=1.0,
                                    match_level=MatchLevel.EXACT,
                                    diff_analysis="用户自定义映射",
                                    title_score=1.0,
                                    content_score=0.0,
                                )
                            else:
                                mapped_result = MatchResult(
                                    matched_name=user_library_name,
                                    matched_reg="",
                                    matched_content="",
                                    score=1.0,
                                    match_level=MatchLevel.EXACT,
                                    diff_analysis="用户自定义映射（未在库中找到）",
                                    title_score=1.0,
                                    content_score=0.0,
                                )
                            match_results = [mapped_result]
                        else:
                            # 无用户映射，使用多结果匹配（最多3条）
                            match_results = logic.match_clause_multiple(clause, index, is_title_only, max_results=3)

                        # v17.1: 构建多结果行
                        row = {
                            ExcelColumns.SEQ: idx,
                            ExcelColumns.CLIENT_ORIG: original_title,
                            ExcelColumns.CLIENT_TRANS: translated_title if was_translated else "",
                            ExcelColumns.CLIENT_CONTENT: clause.content[:500] if clause.content else "",
                        }

                        # 填充最多3条匹配结果
                        for match_num in range(1, 4):
                            if match_num <= len(match_results):
                                mr = match_results[match_num - 1]
                                row[f'匹配{match_num}_条款名称'] = mr.matched_name or ""
                                row[f'匹配{match_num}_注册号'] = logic.clean_reg_number(mr.matched_reg)
                                row[f'匹配{match_num}_条款内容'] = mr.matched_content[:500] if mr.matched_content else ""
                                row[f'匹配{match_num}_匹配度'] = round(mr.score, 3)
                                row[f'匹配{match_num}_匹配级别'] = mr.match_level.value
                            else:
                                row[f'匹配{match_num}_条款名称'] = ""
                                row[f'匹配{match_num}_注册号'] = ""
                                row[f'匹配{match_num}_条款内容'] = ""
                                row[f'匹配{match_num}_匹配度'] = ""
                                row[f'匹配{match_num}_匹配级别'] = ""

                        results.append(row)

                    # 保存
                    output_name = f"报告_{Path(doc_path).stem}.xlsx"
                    output_path = Path(self.output_dir) / output_name
                    df_res = pd.DataFrame(results)
                    df_res.to_excel(output_path, index=False)
                    ExcelStyler.apply_styles(str(output_path))

                    self.log_signal.emit(f"   ✓ 已保存: {output_name}", "success")
                    success_count += 1

                except Exception as e:
                    self.log_signal.emit(f"   ✗ 失败: {e}", "error")

            self.log_signal.emit(f"\n🎉 批量处理完成: {success_count}/{total}", "success")
            self.finished_signal.emit(True, self.output_dir, success_count, total)

        except Exception as e:
            logger.exception("批量处理出错")
            self.log_signal.emit(f"❌ 错误: {str(e)}", "error")
            self.finished_signal.emit(False, str(e), 0, 0)


# ==========================================
# UI组件 - Anthropic 风格
# ==========================================
class AnthropicCard(QFrame):
    """Anthropic 风格卡片组件"""
    def __init__(self, parent=None, variant="default"):
        super().__init__(parent)
        # 根据变体选择背景色
        if variant == "mint":
            bg_color = AnthropicColors.BG_MINT
        elif variant == "lavender":
            bg_color = AnthropicColors.BG_LAVENDER
        else:
            bg_color = AnthropicColors.BG_CARD

        self.setStyleSheet(f"""
            AnthropicCard {{
                background: {bg_color};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 12px;
            }}
        """)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 25))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)


# 保留旧名称以兼容
GlassCard = AnthropicCard


class AddMappingDialog(QDialog):
    """添加映射对话框 - Anthropic 风格"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("添加术语映射")
        self.setMinimumWidth(400)
        self.setStyleSheet(f"""
            QDialog {{ background: {AnthropicColors.BG_PRIMARY}; }}
            QLabel {{ color: {AnthropicColors.TEXT_PRIMARY}; font-size: 14px; }}
            QLineEdit {{
                background: {AnthropicColors.BG_CARD};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px; padding: 10px; color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QLineEdit:focus {{ border-color: {AnthropicColors.ACCENT}; }}
            QPushButton {{
                background: {AnthropicColors.BG_DARK}; color: {AnthropicColors.TEXT_LIGHT}; border: none;
                border-radius: 8px; padding: 10px 20px; font-weight: bold;
            }}
            QPushButton:hover {{ background: {AnthropicColors.ACCENT}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        form = QFormLayout()
        self.eng_input = QLineEdit()
        self.eng_input.setPlaceholderText("例如: reinstatement value")
        form.addRow("英文术语:", self.eng_input)

        self.chn_input = QLineEdit()
        self.chn_input.setPlaceholderText("例如: 重置价值条款")
        form.addRow("中文翻译:", self.chn_input)
        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet(f"background: transparent; color: {AnthropicColors.TEXT_PRIMARY}; border: 1px solid {AnthropicColors.BORDER};")
        cancel_btn.clicked.connect(self.reject)
        save_btn = QPushButton("保存")
        save_btn.clicked.connect(self.accept)
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def get_mapping(self) -> Tuple[str, str]:
        return self.eng_input.text().strip(), self.chn_input.text().strip()


# ==========================================
# 打赏对话框
# ==========================================
class DonateDialog(QDialog):
    """支持作者对话框 - 微信和支付宝双二维码"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('💝 支持作者')
        self.setFixedSize(520, 520)
        self._setup_ui()

    def _get_qr_image_path(self, name: str) -> str:
        """获取二维码图片路径（支持PyInstaller打包）"""
        possible_paths = []

        # PyInstaller 打包后的路径
        if getattr(sys, 'frozen', False):
            # 运行在打包环境中
            bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
            possible_paths.append(os.path.join(bundle_dir, name))
            # macOS .app 的 Resources 目录
            possible_paths.append(os.path.join(bundle_dir, '..', 'Resources', name))

        # 常规开发路径
        possible_paths.extend([
            os.path.join(os.path.dirname(os.path.abspath(__file__)), name),
            os.path.join(os.getcwd(), name),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Resources', name),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', name),
        ])

        for path in possible_paths:
            if os.path.exists(path):
                return path
        return ""

    def _setup_ui(self):
        self.setStyleSheet(f"""
            QDialog {{
                background: {AnthropicColors.BG_PRIMARY};
            }}
            QLabel {{ color: {AnthropicColors.TEXT_PRIMARY}; }}
            QPushButton {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT}; border: none; border-radius: 12px;
                padding: 12px 30px; font-weight: bold; font-size: 14px;
            }}
            QPushButton:hover {{
                background: {AnthropicColors.ACCENT};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(30, 25, 30, 25)

        # 标题区域 - 带动画效果
        title = QLabel('✨ 感谢您的支持！✨')
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(f'''
            font-size: 22px; font-weight: bold;
            color: {AnthropicColors.ACCENT};
            padding: 5px;
        ''')
        layout.addWidget(title)

        desc = QLabel('如果这个工具对您有帮助，欢迎请作者喝杯咖啡 ☕')
        desc.setAlignment(Qt.AlignCenter)
        desc.setStyleSheet(f'color: {AnthropicColors.TEXT_SECONDARY}; font-size: 13px;')
        layout.addWidget(desc)

        # 打赏区域
        donate_container = QHBoxLayout()
        donate_container.setSpacing(25)

        # 微信支付
        wechat_widget = QWidget()
        wechat_layout = QVBoxLayout(wechat_widget)
        wechat_layout.setAlignment(Qt.AlignCenter)
        wechat_layout.setSpacing(8)

        wechat_label = QLabel('💚 微信支付')
        wechat_label.setAlignment(Qt.AlignCenter)
        wechat_label.setStyleSheet('font-weight: bold; font-size: 14px; color: #07C160;')
        wechat_layout.addWidget(wechat_label)

        wechat_qr_label = QLabel()
        wechat_qr_label.setFixedSize(160, 160)
        wechat_qr_label.setAlignment(Qt.AlignCenter)

        from PyQt5.QtGui import QPixmap
        wx_path = self._get_qr_image_path('wx.jpg')
        if wx_path:
            pixmap = QPixmap(wx_path)
            if not pixmap.isNull():
                wechat_qr_label.setPixmap(pixmap.scaled(154, 154, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                wechat_qr_label.setStyleSheet('''
                    background-color: white; border-radius: 12px;
                    border: 3px solid #07C160; padding: 3px;
                ''')
            else:
                wechat_qr_label.setText('💚\n微信扫码')
                wechat_qr_label.setStyleSheet(f'''
                    font-size: 20px; background-color: {AnthropicColors.BG_CARD}; border-radius: 12px;
                    border: 3px solid #07C160; color: #07C160;
                ''')
        else:
            wechat_qr_label.setText('💚\n微信扫码')
            wechat_qr_label.setStyleSheet('''
                font-size: 20px; background-color: rgba(255,255,255,0.1); border-radius: 12px;
                border: 3px solid #07C160; color: #07C160;
            ''')
        wechat_layout.addWidget(wechat_qr_label, alignment=Qt.AlignCenter)

        wechat_hint = QLabel('微信扫一扫')
        wechat_hint.setAlignment(Qt.AlignCenter)
        wechat_hint.setStyleSheet('font-size: 12px; color: #07C160;')
        wechat_layout.addWidget(wechat_hint)
        donate_container.addWidget(wechat_widget)

        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.VLine)
        separator.setStyleSheet(f'background-color: {AnthropicColors.BORDER};')
        donate_container.addWidget(separator)

        # 支付宝
        alipay_widget = QWidget()
        alipay_layout = QVBoxLayout(alipay_widget)
        alipay_layout.setAlignment(Qt.AlignCenter)
        alipay_layout.setSpacing(8)

        alipay_label = QLabel('💙 支付宝')
        alipay_label.setAlignment(Qt.AlignCenter)
        alipay_label.setStyleSheet('font-weight: bold; font-size: 14px; color: #1677FF;')
        alipay_layout.addWidget(alipay_label)

        alipay_qr_label = QLabel()
        alipay_qr_label.setFixedSize(160, 160)
        alipay_qr_label.setAlignment(Qt.AlignCenter)

        zfb_path = self._get_qr_image_path('zfb.jpg')
        if zfb_path:
            pixmap = QPixmap(zfb_path)
            if not pixmap.isNull():
                alipay_qr_label.setPixmap(pixmap.scaled(154, 154, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                alipay_qr_label.setStyleSheet('''
                    background-color: white; border-radius: 12px;
                    border: 3px solid #1677FF; padding: 3px;
                ''')
            else:
                alipay_qr_label.setText('💙\n支付宝扫码')
                alipay_qr_label.setStyleSheet(f'''
                    font-size: 20px; background-color: {AnthropicColors.BG_CARD}; border-radius: 12px;
                    border: 3px solid #1677FF; color: #1677FF;
                ''')
        else:
            alipay_qr_label.setText('💙\n支付宝扫码')
            alipay_qr_label.setStyleSheet('''
                font-size: 20px; background-color: rgba(255,255,255,0.1); border-radius: 12px;
                border: 3px solid #1677FF; color: #1677FF;
            ''')
        alipay_layout.addWidget(alipay_qr_label, alignment=Qt.AlignCenter)

        alipay_hint = QLabel('支付宝扫一扫')
        alipay_hint.setAlignment(Qt.AlignCenter)
        alipay_hint.setStyleSheet('font-size: 12px; color: #1677FF;')
        alipay_layout.addWidget(alipay_hint)
        donate_container.addWidget(alipay_widget)

        layout.addLayout(donate_container)

        # 感谢语
        thanks_label = QLabel('「大鑽戒基金會」へのご支援、誠にありがとうございます💎')
        thanks_label.setAlignment(Qt.AlignCenter)
        thanks_label.setStyleSheet(f'''
            font-size: 14px; font-weight: 500;
            color: {AnthropicColors.TEXT_PRIMARY}; padding: 15px 0 5px 0;
        ''')
        layout.addWidget(thanks_label)

        # 作者信息
        author_info = QLabel('Author: Dachi Yijin  |  智能条款比对工具')
        author_info.setAlignment(Qt.AlignCenter)
        author_info.setStyleSheet(f'color: {AnthropicColors.TEXT_SECONDARY}; font-size: 11px;')
        layout.addWidget(author_info)

        # 关闭按钮
        close_btn = QPushButton('关闭')
        close_btn.setFixedWidth(140)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn, alignment=Qt.AlignCenter)


class BatchSelectDialog(QDialog):
    """批量文件选择对话框 - Anthropic 风格"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("批量处理")
        self.setMinimumSize(500, 400)
        self.setStyleSheet(f"""
            QDialog {{ background: {AnthropicColors.BG_PRIMARY}; }}
            QLabel {{ color: {AnthropicColors.TEXT_PRIMARY}; font-size: 14px; }}
            QListWidget {{
                background: {AnthropicColors.BG_CARD};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px; color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QListWidget::item {{ padding: 10px 12px; border-radius: 6px; margin-bottom: 2px; }}
            QListWidget::item:hover {{ background: rgba(217, 119, 87, 0.08); }}
            QListWidget::item:selected {{ background: {AnthropicColors.BG_MINT}; color: {AnthropicColors.TEXT_PRIMARY}; }}
            QPushButton {{
                background: transparent;
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px; padding: 10px; color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QPushButton:hover {{ background: {AnthropicColors.BG_CARD}; border-color: {AnthropicColors.ACCENT}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        layout.addWidget(QLabel("选择要批量处理的 Word 文件:"))

        self.file_list = QListWidget()
        layout.addWidget(self.file_list)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("➕ 添加文件")
        add_btn.clicked.connect(self._add_files)
        clear_btn = QPushButton("🗑️ 清空")
        clear_btn.clicked.connect(self.file_list.clear)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(clear_btn)
        layout.addLayout(btn_row)

        action_row = QHBoxLayout()
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        start_btn = QPushButton("开始批量处理")
        start_btn.setStyleSheet(f"background: {AnthropicColors.BG_DARK}; color: {AnthropicColors.TEXT_LIGHT};")
        start_btn.clicked.connect(self.accept)
        action_row.addWidget(cancel_btn)
        action_row.addWidget(start_btn)
        layout.addLayout(action_row)

        self.selected_files: List[str] = []

    def _add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择Word文件", "", "Word Files (*.docx)")
        for f in files:
            if f not in self.selected_files:
                self.selected_files.append(f)
                self.file_list.addItem(Path(f).name)

    def get_files(self) -> List[str]:
        return self.selected_files


class ClauseQueryDialog(QDialog):
    """v17.1: 条款查询对话框 - 仅查询条款标题"""
    def __init__(self, parent=None, library_index=None, logic=None, mapping_mgr=None):
        super().__init__(parent)
        self.setWindowTitle("🔍 条款智能查询")
        self.setMinimumSize(600, 500)
        self.library_index = library_index
        self.logic = logic or ClauseMatcherLogic()
        self.mapping_mgr = mapping_mgr
        self._setup_ui()

    def _setup_ui(self):
        self.setStyleSheet(f"""
            QDialog {{ background: {AnthropicColors.BG_PRIMARY}; }}
            QLabel {{ color: {AnthropicColors.TEXT_PRIMARY}; font-size: 14px; }}
            QLineEdit, QTextEdit {{
                background: {AnthropicColors.BG_CARD};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px; padding: 10px; color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QLineEdit:focus, QTextEdit:focus {{ border-color: {AnthropicColors.ACCENT}; }}
            QPushButton {{
                background: transparent;
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px; padding: 10px 20px; color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QPushButton:hover {{ background: {AnthropicColors.BG_CARD}; border-color: {AnthropicColors.ACCENT}; }}
            QListWidget {{
                background: {AnthropicColors.BG_CARD};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px; color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QListWidget::item {{ padding: 10px 12px; border-radius: 6px; margin-bottom: 2px; }}
            QListWidget::item:hover {{ background: rgba(217, 119, 87, 0.08); }}
            QListWidget::item:selected {{ background: {AnthropicColors.BG_MINT}; color: {AnthropicColors.TEXT_PRIMARY}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 说明
        hint = QLabel("输入条款名称或关键词，系统将自动匹配最相近的条款库条款（仅匹配标题）")
        hint.setStyleSheet(f"color: {AnthropicColors.TEXT_MUTED}; font-size: 12px;")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        # 输入行
        input_row = QHBoxLayout()
        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText("例如: 自动升值 或 REINSTATEMENT VALUE...")
        self.query_input.returnPressed.connect(self._do_search)
        self.search_btn = QPushButton("🔍 搜索")
        self.search_btn.setStyleSheet(f"background: {AnthropicColors.ACCENT}; color: {AnthropicColors.TEXT_LIGHT};")
        self.search_btn.clicked.connect(self._do_search)
        input_row.addWidget(self.query_input, 4)
        input_row.addWidget(self.search_btn, 1)
        layout.addLayout(input_row)

        # 结果列表
        layout.addWidget(QLabel("查询结果（最多5条）:"))
        self.result_list = QListWidget()
        self.result_list.itemDoubleClicked.connect(self._show_detail)
        layout.addWidget(self.result_list, 1)

        # 详情区
        layout.addWidget(QLabel("选中条款详情:"))
        self.detail_text = QTextEdit()
        self.detail_text.setReadOnly(True)
        self.detail_text.setMaximumHeight(120)
        layout.addWidget(self.detail_text)

        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)

        # 存储结果数据
        self._search_results = []

    def _do_search(self):
        """执行查询"""
        query = self.query_input.text().strip()
        if not query:
            return

        self.result_list.clear()
        self.detail_text.clear()
        self._search_results = []

        if not self.library_index or not self.library_index.data:
            self.result_list.addItem("⚠️ 请先选择条款库文件")
            return

        # 检查是否有用户映射
        if self.mapping_mgr:
            mapped_name = self.mapping_mgr.get_library_name(query)
            if mapped_name:
                # 有映射，直接返回映射的那一条
                lib_entry = self.logic.find_library_entry_by_name(mapped_name, self.library_index)
                if lib_entry:
                    self._search_results = [{
                        'name': lib_entry.get('条款名称', mapped_name),
                        'content': lib_entry.get('条款内容', ''),
                        'reg': lib_entry.get('产品注册号', lib_entry.get('注册号', '')),
                        'score': 1.0,
                        'matchType': 'mapping'
                    }]
                else:
                    self._search_results = [{
                        'name': mapped_name,
                        'content': '(用户映射条款，未在库中找到)',
                        'reg': '',
                        'score': 1.0,
                        'matchType': 'mapping'
                    }]
                self._display_results()
                return

        # 使用search_library_titles进行查询
        results = self.logic.search_library_titles(query, self.library_index, max_results=5)
        self._search_results = results
        self._display_results()

    def _display_results(self):
        """显示查询结果"""
        if not self._search_results:
            self.result_list.addItem("未找到匹配的条款")
            return

        for i, r in enumerate(self._search_results):
            match_type = r.get('matchType', '')
            if match_type == 'mapping':
                type_str = "[用户映射]"
            elif match_type == 'exact':
                type_str = "[精确匹配]"
            elif match_type == 'contain':
                type_str = "[包含匹配]"
            else:
                type_str = f"[模糊 {r.get('score', 0):.2f}]"

            self.result_list.addItem(f"{i+1}. {type_str} {r.get('name', '')}")

        # 自动选择第一条
        if self.result_list.count() > 0:
            self.result_list.setCurrentRow(0)
            self._show_detail(self.result_list.item(0))

    def _show_detail(self, item):
        """显示条款详情"""
        row = self.result_list.row(item)
        if 0 <= row < len(self._search_results):
            r = self._search_results[row]
            detail = f"【条款名称】{r.get('name', '')}\n\n"
            detail += f"【产品注册号】{r.get('reg', '无')}\n\n"
            detail += f"【条款内容】\n{r.get('content', '无内容')[:500]}..."
            self.detail_text.setText(detail)


# ==========================================
# 条款提取Tab - V18.0新增
# ==========================================
class ClauseExtractorTab(QWidget):
    """条款提取Tab - 支持文件夹分类和条款提取"""

    # 文件分类信号
    extraction_log = pyqtSignal(str, str)  # message, level

    def __init__(self, parent=None):
        super().__init__(parent)
        self.selected_files = []
        self.classified_files = {'fujia': [], 'feilv': [], 'zhu': []}
        self.doc_files = []  # .doc文件列表（需要转换）
        self.extracted_data = []
        self.categories = set()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 15, 20, 15)

        # 紧凑型统计面板（水平对齐）- 初始隐藏，有数据时显示
        self.stats_frame = QFrame()
        self.stats_frame.setFixedHeight(45)
        self.stats_frame.setStyleSheet(f"""
            QFrame {{
                background: {AnthropicColors.BG_CARD};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
            }}
        """)
        self.stats_frame.setVisible(False)  # 初始隐藏
        stats_layout = QHBoxLayout(self.stats_frame)
        stats_layout.setContentsMargins(20, 0, 20, 0)
        stats_layout.setSpacing(0)

        # 使用固定宽度的统计项确保对齐
        stat_style = "font-size: 13px; font-family: 'Söhne', 'SF Pro Display', -apple-system, sans-serif;"

        self.stat_total_label = QLabel("待处理: 0")
        self.stat_total_label.setFixedWidth(120)
        self.stat_total_label.setAlignment(Qt.AlignCenter)
        self.stat_total_label.setStyleSheet(f"color: {AnthropicColors.ACCENT}; {stat_style} font-weight: 600;")

        sep1 = QLabel("|")
        sep1.setFixedWidth(20)
        sep1.setAlignment(Qt.AlignCenter)
        sep1.setStyleSheet(f"color: {AnthropicColors.BORDER}; font-size: 14px;")

        self.stat_extracted_label = QLabel("已提取: 0")
        self.stat_extracted_label.setFixedWidth(100)
        self.stat_extracted_label.setAlignment(Qt.AlignCenter)
        self.stat_extracted_label.setStyleSheet(f"color: {AnthropicColors.SUCCESS}; {stat_style} font-weight: 600;")

        sep2 = QLabel("|")
        sep2.setFixedWidth(20)
        sep2.setAlignment(Qt.AlignCenter)
        sep2.setStyleSheet(f"color: {AnthropicColors.BORDER}; font-size: 14px;")

        self.stat_categories_label = QLabel("分类数: 0")
        self.stat_categories_label.setFixedWidth(100)
        self.stat_categories_label.setAlignment(Qt.AlignCenter)
        self.stat_categories_label.setStyleSheet(f"color: {AnthropicColors.INFO}; {stat_style} font-weight: 600;")

        sep3 = QLabel("|")
        sep3.setFixedWidth(20)
        sep3.setAlignment(Qt.AlignCenter)
        sep3.setStyleSheet(f"color: {AnthropicColors.BORDER}; font-size: 14px;")

        self.stat_skipped_label = QLabel("已跳过: 0")
        self.stat_skipped_label.setFixedWidth(100)
        self.stat_skipped_label.setAlignment(Qt.AlignCenter)
        self.stat_skipped_label.setStyleSheet(f"color: {AnthropicColors.WARNING}; {stat_style} font-weight: 600;")

        stats_layout.addStretch()
        stats_layout.addWidget(self.stat_total_label)
        stats_layout.addWidget(sep1)
        stats_layout.addWidget(self.stat_extracted_label)
        stats_layout.addWidget(sep2)
        stats_layout.addWidget(self.stat_categories_label)
        stats_layout.addWidget(sep3)
        stats_layout.addWidget(self.stat_skipped_label)
        stats_layout.addStretch()

        layout.addWidget(self.stats_frame)

        # 文件选择卡片
        file_card = GlassCard()
        file_card_layout = QVBoxLayout(file_card)
        file_card_layout.setSpacing(12)

        card_title = QLabel("📂 选择条款文件或文件夹")
        card_title.setStyleSheet(f"color: {AnthropicColors.ACCENT}; font-weight: 600; font-size: 14px;")
        file_card_layout.addWidget(card_title)

        # 模式切换按钮
        mode_layout = QHBoxLayout()
        self.mode_files_btn = QPushButton("📄 选择文件")
        self.mode_folder_btn = QPushButton("📁 选择文件夹")

        mode_btn_style = f"""
            QPushButton {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                padding: 10px 20px;
                color: {AnthropicColors.TEXT_MUTED};
                font-weight: 500;
            }}
            QPushButton:hover {{
                border-color: {AnthropicColors.ACCENT};
                color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QPushButton:checked {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT};
                border-color: {AnthropicColors.BG_DARK};
            }}
        """
        self.mode_files_btn.setStyleSheet(mode_btn_style)
        self.mode_folder_btn.setStyleSheet(mode_btn_style)
        self.mode_files_btn.setCheckable(True)
        self.mode_folder_btn.setCheckable(True)
        self.mode_files_btn.setChecked(True)
        self.mode_files_btn.clicked.connect(lambda: self._switch_mode('files'))
        self.mode_folder_btn.clicked.connect(lambda: self._switch_mode('folder'))

        mode_layout.addWidget(self.mode_files_btn)
        mode_layout.addWidget(self.mode_folder_btn)
        mode_layout.addStretch()
        file_card_layout.addLayout(mode_layout)

        # 文件选择区域
        self.file_select_btn = QPushButton("点击选择文件 (.docx / .pdf)")
        self.file_select_btn.setMinimumHeight(80)
        self.file_select_btn.setCursor(Qt.PointingHandCursor)
        self.file_select_btn.setStyleSheet(f"""
            QPushButton {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 2px dashed {AnthropicColors.BORDER};
                border-radius: 12px;
                color: {AnthropicColors.TEXT_MUTED};
                font-size: 14px;
            }}
            QPushButton:hover {{
                border-color: {AnthropicColors.ACCENT};
                color: {AnthropicColors.TEXT_PRIMARY};
                background: rgba(217, 119, 87, 0.05);
            }}
        """)
        self.file_select_btn.clicked.connect(self._select_files)
        file_card_layout.addWidget(self.file_select_btn)

        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(180)
        self.file_list.setStyleSheet(f"""
            QListWidget {{
                background: {AnthropicColors.BG_CARD};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                padding: 10px;
                font-family: 'Söhne Mono', 'SF Mono', 'Menlo', monospace;
                font-size: 12px;
                color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QListWidget::item {{
                padding: 8px 12px;
                border-radius: 6px;
                color: {AnthropicColors.TEXT_PRIMARY};
                margin-bottom: 2px;
            }}
            QListWidget::item:hover {{
                background: rgba(217, 119, 87, 0.1);
            }}
            QListWidget::item:selected {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT};
            }}
        """)
        self.file_list.setVisible(False)
        file_card_layout.addWidget(self.file_list)

        # 分类预览区域（文件夹模式）
        self.classify_preview = QWidget()
        classify_layout = QHBoxLayout(self.classify_preview)
        classify_layout.setContentsMargins(0, 10, 0, 0)

        self.preview_fujia = self._create_classify_box("📗 附加条款", "0", "#d97757")
        self.preview_feilv = self._create_classify_box("📘 费率表", "0", "#6a9bcc")
        self.preview_zhu = self._create_classify_box("📙 主条款", "0", "#788c5d")

        classify_layout.addWidget(self.preview_fujia)
        classify_layout.addWidget(self.preview_feilv)
        classify_layout.addWidget(self.preview_zhu)
        self.classify_preview.setVisible(False)
        file_card_layout.addWidget(self.classify_preview)

        layout.addWidget(file_card)

        # 操作按钮行
        btn_layout = QHBoxLayout()

        self.extract_btn = QPushButton("🚀 开始提取")
        self.extract_btn.setMinimumHeight(48)
        self.extract_btn.setCursor(Qt.PointingHandCursor)
        self.extract_btn.setEnabled(False)
        self.extract_btn.setStyleSheet(f"""
            QPushButton {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT};
                border: none;
                border-radius: 8px;
                font-size: 15px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background: #2a2a28;
            }}
            QPushButton:disabled {{
                background: {AnthropicColors.BORDER};
                color: {AnthropicColors.TEXT_SECONDARY};
            }}
        """)
        self.extract_btn.clicked.connect(self._start_extraction)

        self.download_zip_btn = QPushButton("📦 下载分类ZIP")
        self.download_zip_btn.setMinimumHeight(48)
        self.download_zip_btn.setCursor(Qt.PointingHandCursor)
        self.download_zip_btn.setVisible(False)
        self.download_zip_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {AnthropicColors.TEXT_PRIMARY};
                border: 1px solid {AnthropicColors.BG_DARK};
                border-radius: 8px;
                font-size: 14px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT};
            }}
        """)
        self.download_zip_btn.clicked.connect(self._download_classified_zip)

        self.download_excel_btn = QPushButton("📊 下载Excel报告")
        self.download_excel_btn.setMinimumHeight(48)
        self.download_excel_btn.setCursor(Qt.PointingHandCursor)
        self.download_excel_btn.setVisible(False)
        self.download_excel_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {AnthropicColors.SUCCESS};
                border: 1px solid {AnthropicColors.SUCCESS};
                border-radius: 8px;
                font-size: 14px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                background: {AnthropicColors.SUCCESS};
                color: white;
            }}
        """)
        self.download_excel_btn.clicked.connect(self._download_excel_report)

        self.clear_btn = QPushButton("🗑 清空")
        self.clear_btn.setMinimumHeight(48)
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {AnthropicColors.TEXT_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                font-size: 14px;
            }}
            QPushButton:hover {{
                border-color: {AnthropicColors.ERROR};
                color: {AnthropicColors.ERROR};
            }}
        """)
        self.clear_btn.clicked.connect(self._clear_all)

        btn_layout.addWidget(self.extract_btn, 3)
        btn_layout.addWidget(self.download_zip_btn, 1)
        btn_layout.addWidget(self.download_excel_btn, 1)
        btn_layout.addWidget(self.clear_btn, 1)
        layout.addLayout(btn_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{ background: {AnthropicColors.BORDER}; border-radius: 2px; }}
            QProgressBar::chunk {{
                background: {AnthropicColors.ACCENT};
                border-radius: 2px;
            }}
        """)
        layout.addWidget(self.progress_bar)

        # 日志区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(f"""
            QTextEdit {{
                background: {AnthropicColors.BG_DARK};
                border: none;
                border-radius: 8px;
                color: {AnthropicColors.TEXT_LIGHT};
                padding: 15px;
                font-family: 'SF Mono', 'Consolas', monospace;
                font-size: 12px;
            }}
        """)
        layout.addWidget(self.log_text, 1)

        # 初始日志
        self._log("📊 条款提取工具已就绪", "info")
        self._log("   支持格式: .docx / .pdf", "info")
        self._log("   文件夹模式可自动分类：附加条款、费率表、主条款", "info")

    def _create_classify_box(self, title: str, count: str, color: str) -> QFrame:
        """创建分类预览框"""
        frame = QFrame()
        frame.setMinimumWidth(160)
        frame.setMinimumHeight(90)
        frame.setStyleSheet(f"""
            QFrame {{
                background: {AnthropicColors.BG_CARD};
                border: 2px solid {color};
                border-radius: 12px;
            }}
        """)
        layout = QVBoxLayout(frame)
        layout.setSpacing(8)
        layout.setContentsMargins(20, 15, 20, 15)

        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet(f"""
            QLabel {{
                color: {AnthropicColors.TEXT_PRIMARY};
                background: transparent;
                border: none;
                font-size: 14px;
                font-weight: 500;
            }}
        """)

        count_label = QLabel(count)
        count_label.setAlignment(Qt.AlignCenter)
        count_label.setStyleSheet(f"""
            QLabel {{
                color: {color};
                background: transparent;
                border: none;
                font-size: 24px;
                font-weight: bold;
                font-family: 'Menlo', 'Monaco', 'Consolas', monospace;
            }}
        """)
        count_label.setObjectName("count")

        layout.addWidget(title_label)
        layout.addWidget(count_label)
        return frame

    def _switch_mode(self, mode: str):
        """切换文件/文件夹模式"""
        self.mode_files_btn.setChecked(mode == 'files')
        self.mode_folder_btn.setChecked(mode == 'folder')

        if mode == 'files':
            self.file_select_btn.setText("点击选择文件 (.docx / .pdf)")
        else:
            self.file_select_btn.setText("点击选择文件夹")

        self._clear_all()
        self._log(f"📋 切换到{'文件模式' if mode == 'files' else '文件夹模式'}", "info")

    def _select_files(self):
        """选择文件或文件夹"""
        if self.mode_files_btn.isChecked():
            files, _ = QFileDialog.getOpenFileNames(
                self, "选择条款文件", "",
                "文档文件 (*.docx *.pdf);;Word文档 (*.docx);;PDF文档 (*.pdf)"
            )
            if files:
                self._handle_files(files)
        else:
            folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
            if folder:
                self._handle_folder(folder)

    def _handle_files(self, file_paths: list):
        """处理选择的文件"""
        self.selected_files = []
        self.file_list.clear()

        for fp in file_paths:
            fname = os.path.basename(fp)
            ext = fname.split('.')[-1].lower()

            # 检查格式
            if ext not in ['docx', 'pdf']:
                self._log(f"⚠️ 跳过不支持的格式: {fname}", "warning")
                continue

            # 检查是否包含"附加"
            if '附加' not in fname:
                self._log(f"⚠️ 跳过不含「附加」的文件: {fname}", "warning")
                continue

            # 排除费率文件
            if '费率' in fname:
                self._log(f"⚠️ 跳过费率文件: {fname}", "warning")
                continue

            self.selected_files.append(fp)
            self.file_list.addItem(f"📄 {fname}")

        if self.selected_files:
            self.file_list.setVisible(True)
            self.extract_btn.setEnabled(True)
            self._update_stats()
            self._log(f"✓ 已选择 {len(self.selected_files)} 个文件", "success")

    def _handle_folder(self, folder_path: str):
        """处理文件夹 - 自动分类（支持多层子目录穿透）"""
        self.classified_files = {'fujia': [], 'feilv': [], 'zhu': []}
        self.selected_files = []
        self.doc_files = []  # 需要转换的.doc文件
        self.file_list.clear()

        # 使用os.walk递归遍历所有子目录
        for root, dirs, files in os.walk(folder_path):
            # 跳过隐藏目录
            dirs[:] = [d for d in dirs if not d.startswith('.')]

            for fname in files:
                if fname.startswith('.') or fname.startswith('~'):
                    continue

                ext = fname.split('.')[-1].lower()
                if ext not in ['doc', 'docx', 'pdf']:
                    continue

                full_path = os.path.join(root, fname)
                category = self._classify_file(fname)
                self.classified_files[category].append(full_path)

                # 记录.doc文件（需要手动转换）
                if ext == 'doc':
                    self.doc_files.append(fname)

                # 只有附加条款的 docx/pdf 才能提取
                if category == 'fujia' and ext in ['docx', 'pdf']:
                    self.selected_files.append(full_path)

        # 更新分类预览
        self.classify_preview.setVisible(True)
        self.preview_fujia.findChild(QLabel, "count").setText(str(len(self.classified_files['fujia'])))
        self.preview_feilv.findChild(QLabel, "count").setText(str(len(self.classified_files['feilv'])))
        self.preview_zhu.findChild(QLabel, "count").setText(str(len(self.classified_files['zhu'])))

        # 显示文件列表
        self.file_list.clear()
        category_icons = {'fujia': '📗', 'feilv': '📘', 'zhu': '📙'}
        for cat in ['fujia', 'feilv', 'zhu']:
            for fp in self.classified_files[cat]:
                fname = os.path.basename(fp)
                ext = fname.split('.')[-1].lower()
                # 标记.doc文件
                suffix = " ⚠️" if ext == 'doc' else ""
                self.file_list.addItem(f"{category_icons[cat]} {fname}{suffix}")
        self.file_list.setVisible(True)

        total = sum(len(v) for v in self.classified_files.values())
        self._log(f"📁 文件夹加载完成，共 {total} 个文件", "info")
        self._log(f"   📗 附加条款: {len(self.classified_files['fujia'])} 个", "info")
        self._log(f"   📘 费率表: {len(self.classified_files['feilv'])} 个", "info")
        self._log(f"   📙 主条款: {len(self.classified_files['zhu'])} 个", "info")

        # 警告.doc文件 - 弹出对话框
        if self.doc_files:
            self._log(f"⚠️ 发现 {len(self.doc_files)} 个 .doc 文件需要先转换为 .docx 格式:", "warning")
            for df in self.doc_files[:5]:
                self._log(f"   • {df}", "warning")
            if len(self.doc_files) > 5:
                self._log(f"   ... 还有 {len(self.doc_files) - 5} 个文件", "warning")
            self._log("💡 请使用 Microsoft Word 或 LibreOffice 打开后另存为 .docx 格式", "info")

            # 显示警告对话框
            self._show_doc_warning_dialog()

        # 启用提取按钮 - 只要有可提取的文件就启用
        if self.selected_files:
            self.extract_btn.setEnabled(True)
            self._log(f"✓ 将提取 {len(self.selected_files)} 个附加条款(.docx/.pdf)", "success")
        else:
            # 如果没有可提取文件但有附加条款的.doc文件，也提示
            fujia_doc_count = sum(1 for f in self.classified_files['fujia'] if f.endswith('.doc'))
            if fujia_doc_count > 0:
                self._log(f"⚠️ 有 {fujia_doc_count} 个附加条款为.doc格式，转换后即可提取", "warning")
                self.extract_btn.setEnabled(False)
            else:
                self._log("ℹ️ 未找到可提取的附加条款文件", "info")
                self.extract_btn.setEnabled(False)

        # 显示ZIP下载按钮
        if total > 0:
            self.download_zip_btn.setVisible(True)

        self._update_stats()

    def _classify_file(self, filename: str) -> str:
        """文件分类"""
        if '费率表' in filename or '费率方案' in filename:
            return 'feilv'

        # 匹配"附加xxx保险"或"附加xxx条款"
        fujia_pattern = r'附加.{1,20}(保险|条款)'
        if re.search(fujia_pattern, filename):
            return 'fujia'

        return 'zhu'

    def _show_doc_warning_dialog(self):
        """显示.doc文件警告对话框 - 使用自定义Dialog确保按钮可见"""
        doc_count = len(self.doc_files)
        fujia_doc_count = sum(1 for f in self.classified_files.get('fujia', []) if f.endswith('.doc'))

        # 创建自定义对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("发现旧版Word文档")
        dialog.setMinimumWidth(500)
        dialog.setStyleSheet(f"background: {AnthropicColors.BG_PRIMARY};")

        layout = QVBoxLayout(dialog)
        layout.setSpacing(15)
        layout.setContentsMargins(25, 25, 25, 20)

        # 标题
        title = QLabel(f"发现 {doc_count} 个 .doc 格式文件")
        title.setStyleSheet(f"color: {AnthropicColors.TEXT_PRIMARY}; font-size: 16px; font-weight: bold;")
        layout.addWidget(title)

        subtitle = QLabel("是否自动转换为 .docx 格式？")
        subtitle.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 14px;")
        layout.addWidget(subtitle)

        # 文件列表
        file_list = QLabel()
        detail_text = "检测到以下 .doc 文件:\n"
        for df in self.doc_files[:8]:
            detail_text += f"  • {df}\n"
        if len(self.doc_files) > 8:
            detail_text += f"  ... 还有 {len(self.doc_files) - 8} 个文件"
        file_list.setText(detail_text)
        file_list.setStyleSheet(f"color: {AnthropicColors.TEXT_PRIMARY}; font-size: 12px; padding: 10px; background: {AnthropicColors.BG_CARD}; border-radius: 8px;")
        file_list.setWordWrap(True)
        layout.addWidget(file_list)

        # 提示信息
        if fujia_doc_count > 0:
            hint = QLabel(f"💡 其中 {fujia_doc_count} 个附加条款转换后可立即提取")
            hint.setStyleSheet(f"color: {AnthropicColors.ACCENT}; font-size: 13px; font-weight: 500;")
            layout.addWidget(hint)

        layout.addSpacing(10)

        # 按钮行
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        skip_btn = QPushButton("跳过")
        skip_btn.setMinimumSize(100, 44)
        skip_btn.setCursor(Qt.PointingHandCursor)
        skip_btn.setStyleSheet(f"""
            QPushButton {{
                background: {AnthropicColors.BG_CARD};
                color: {AnthropicColors.TEXT_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                font-size: 14px;
                font-weight: 500;
                padding: 10px 25px;
            }}
            QPushButton:hover {{
                background: {AnthropicColors.BG_PRIMARY};
                border-color: {AnthropicColors.TEXT_SECONDARY};
            }}
        """)
        skip_btn.clicked.connect(dialog.reject)

        convert_btn = QPushButton("✓ 自动转换")
        convert_btn.setMinimumSize(120, 44)
        convert_btn.setCursor(Qt.PointingHandCursor)
        convert_btn.setStyleSheet(f"""
            QPushButton {{
                background: {AnthropicColors.ACCENT};
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                font-weight: 600;
                padding: 10px 25px;
            }}
            QPushButton:hover {{
                background: #c96747;
            }}
        """)
        convert_btn.clicked.connect(dialog.accept)

        btn_layout.addWidget(skip_btn)
        btn_layout.addWidget(convert_btn)
        layout.addLayout(btn_layout)

        # 显示对话框
        if dialog.exec_() == QDialog.Accepted:
            self._convert_doc_files()

    def _convert_doc_files(self):
        """批量转换.doc文件为.docx格式"""
        import subprocess
        import platform

        self._log(f"🔄 开始转换 {len(self.doc_files)} 个 .doc 文件...", "info")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        converted = 0
        failed = 0
        converted_paths = []

        for i, doc_name in enumerate(self.doc_files):
            progress = int((i + 1) / len(self.doc_files) * 100)
            self.progress_bar.setValue(progress)
            QApplication.processEvents()

            # 查找完整路径
            doc_path = None
            for cat in ['fujia', 'feilv', 'zhu']:
                for fp in self.classified_files[cat]:
                    if os.path.basename(fp) == doc_name:
                        doc_path = fp
                        break
                if doc_path:
                    break

            if not doc_path:
                self._log(f"  ✗ 未找到文件路径: {doc_name}", "error")
                failed += 1
                continue

            docx_path = doc_path.rsplit('.', 1)[0] + '.docx'

            try:
                if platform.system() == 'Darwin':
                    # macOS: 使用 textutil 或 soffice
                    result = subprocess.run(
                        ['textutil', '-convert', 'docx', doc_path, '-output', docx_path],
                        capture_output=True, text=True, timeout=60
                    )
                    if result.returncode != 0:
                        # 尝试使用 LibreOffice
                        soffice_paths = [
                            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                            '/usr/local/bin/soffice'
                        ]
                        for soffice in soffice_paths:
                            if os.path.exists(soffice):
                                output_dir = os.path.dirname(doc_path)
                                result = subprocess.run(
                                    [soffice, '--headless', '--convert-to', 'docx', '--outdir', output_dir, doc_path],
                                    capture_output=True, text=True, timeout=120
                                )
                                if result.returncode == 0:
                                    break
                else:
                    # Windows/Linux: 使用 LibreOffice
                    soffice = 'soffice' if platform.system() == 'Linux' else 'soffice.exe'
                    output_dir = os.path.dirname(doc_path)
                    result = subprocess.run(
                        [soffice, '--headless', '--convert-to', 'docx', '--outdir', output_dir, doc_path],
                        capture_output=True, text=True, timeout=120
                    )

                if os.path.exists(docx_path):
                    converted += 1
                    converted_paths.append(docx_path)
                    self._log(f"  ✓ {doc_name} → .docx", "success")

                    # 更新分类列表
                    for cat in ['fujia', 'feilv', 'zhu']:
                        if doc_path in self.classified_files[cat]:
                            self.classified_files[cat].remove(doc_path)
                            self.classified_files[cat].append(docx_path)
                            # 如果是附加条款，添加到待提取列表
                            if cat == 'fujia':
                                self.selected_files.append(docx_path)
                            break
                else:
                    self._log(f"  ✗ 转换失败: {doc_name}", "error")
                    failed += 1

            except subprocess.TimeoutExpired:
                self._log(f"  ✗ 转换超时: {doc_name}", "error")
                failed += 1
            except Exception as e:
                self._log(f"  ✗ 转换错误: {doc_name} - {str(e)}", "error")
                failed += 1

        self.progress_bar.setValue(100)
        self._log(f"🎉 转换完成! 成功: {converted}, 失败: {failed}", "success" if failed == 0 else "warning")

        # 更新UI
        if converted > 0:
            self._refresh_file_list()
            self._update_stats()
            if self.selected_files:
                self.extract_btn.setEnabled(True)
                self._log(f"✓ 现在可以提取 {len(self.selected_files)} 个附加条款", "success")

    def _refresh_file_list(self):
        """刷新文件列表显示"""
        self.file_list.clear()
        category_icons = {'fujia': '📗', 'feilv': '📘', 'zhu': '📙'}
        for cat in ['fujia', 'feilv', 'zhu']:
            for fp in self.classified_files[cat]:
                fname = os.path.basename(fp)
                ext = fname.split('.')[-1].lower()
                suffix = " ⚠️" if ext == 'doc' else ""
                self.file_list.addItem(f"{category_icons[cat]} {fname}{suffix}")

        # 更新分类预览
        self.preview_fujia.findChild(QLabel, "count").setText(str(len(self.classified_files['fujia'])))
        self.preview_feilv.findChild(QLabel, "count").setText(str(len(self.classified_files['feilv'])))
        self.preview_zhu.findChild(QLabel, "count").setText(str(len(self.classified_files['zhu'])))

    def _start_extraction(self):
        """开始提取条款"""
        if not self.selected_files:
            self._log("⚠️ 请先选择文件", "warning")
            return

        self.extracted_data = []
        self.categories = set()
        self.extract_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self._log(f"🚀 开始处理 {len(self.selected_files)} 个文件...", "info")

        for i, fp in enumerate(self.selected_files):
            fname = os.path.basename(fp)
            progress = int((i + 1) / len(self.selected_files) * 100)
            self.progress_bar.setValue(progress)
            QApplication.processEvents()

            try:
                results = self._extract_clause(fp)
                for result in results:
                    self.extracted_data.append(result)
                    self.categories.add(result['Category'])
                    if result.get('Error'):
                        self._log(f"✗ {result['ClauseName']}: {result['Error']}", "error")
                    else:
                        self._log(f"✓ {result['ClauseName']} → {result['Category']}", "success")
            except Exception as e:
                self._log(f"✗ {fname}: {sanitize_error_message(e)}", "error")

        self.progress_bar.setValue(100)
        self._update_stats()

        success_count = len([d for d in self.extracted_data if not d.get('Error')])
        self._log(f"🎉 处理完成! 新增: {success_count} 条，共 {len(self.categories)} 个分类", "success")

        self.extract_btn.setEnabled(True)
        if self.extracted_data:
            self.download_excel_btn.setVisible(True)

    def _extract_clause(self, file_path: str) -> list:
        """提取单个文件的条款"""
        fname = os.path.basename(file_path)
        clause_name = os.path.splitext(fname)[0]
        today = datetime.now().strftime('%Y-%m-%d')
        ext = fname.split('.')[-1].lower()

        result = {
            'FileName': fname,
            'ClauseName': clause_name,
            'RegistrationNo': '',
            'Content': '',
            'Category': self._get_category(fname, clause_name),
            'AddDate': today,
            'Error': ''
        }

        try:
            if ext == 'pdf':
                paragraphs = self._parse_pdf(file_path)
            else:
                paragraphs = self._parse_docx(file_path)

            if not paragraphs:
                result['Error'] = '文档内容为空'
                return [result]

            # 提取注册号
            for i, para in enumerate(paragraphs[:8]):
                if '注册号' in para or re.search(r'[A-Z]\d{10,}', para):
                    match = re.search(r'[（\(]([^）\)]+)[）\)]', para)
                    if match:
                        result['RegistrationNo'] = match.group(1)
                    else:
                        result['RegistrationNo'] = re.sub(r'(产品)?注册号[:：]?', '', para).strip()
                    break

            # 提取正文（跳过标题和注册号部分）
            content_lines = []
            start_idx = 3 if len(paragraphs) >= 4 else 0
            for para in paragraphs[start_idx:]:
                clean = para.strip()
                if clean and clean != clause_name and not self._is_noise_line(clean):
                    content_lines.append(clean)

            result['Content'] = '\n'.join(content_lines)
            return [result]

        except Exception as e:
            result['Error'] = f'解析出错: {str(e)}'
            return [result]

    def _parse_docx(self, file_path: str) -> list:
        """解析Word文档"""
        doc = Document(file_path)
        paragraphs = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                paragraphs.append(text)
        return paragraphs

    def _parse_pdf(self, file_path: str) -> list:
        """解析PDF文档"""
        paragraphs = []

        if HAS_PDFPLUMBER:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        paragraphs.extend([l.strip() for l in lines if l.strip()])
        elif HAS_PYPDF2:
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        paragraphs.extend([l.strip() for l in lines if l.strip()])
        else:
            raise ImportError("未安装PDF解析库 (pdfplumber 或 PyPDF2)")

        return paragraphs

    def _is_noise_line(self, text: str) -> bool:
        """判断是否为噪声行"""
        noise_patterns = [
            r'^第?\s*\d+\s*页',
            r'^Page\s*\d+',
            r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}$',
            r'^www\.',
            r'^http',
        ]
        for pattern in noise_patterns:
            if re.match(pattern, text, re.IGNORECASE):
                return True
        return False

    def _get_category(self, filename: str, title: str) -> str:
        """获取条款分类"""
        text = title or filename
        if '附加' in text:
            parts = text.split('附加')
            prefix = parts[0].replace('条款', '').strip()
            if prefix and len(prefix) > 2:
                return prefix + '附加条款'
        return '通用附加条款'

    def _download_classified_zip(self):
        """下载分类后的ZIP文件"""
        if not any(self.classified_files.values()):
            self._log("⚠️ 没有可下载的文件", "warning")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存分类ZIP",
            f"条款分类_{datetime.now():%Y%m%d_%H%M}.zip",
            "ZIP文件 (*.zip)"
        )
        if not save_path:
            return

        self._log("📦 正在生成分类ZIP文件...", "info")

        try:
            with zipfile.ZipFile(save_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                folder_names = {'fujia': '附加条款', 'feilv': '费率表', 'zhu': '主条款'}
                for cat, files in self.classified_files.items():
                    folder_name = folder_names[cat]
                    for fp in files:
                        fname = os.path.basename(fp)
                        zf.write(fp, f"{folder_name}/{fname}")

            total = sum(len(v) for v in self.classified_files.values())
            self._log(f"✅ 分类ZIP已保存: {os.path.basename(save_path)}", "success")
            self._log(f"   包含 {len(self.classified_files['fujia'])} 附加条款 + {len(self.classified_files['feilv'])} 费率表 + {len(self.classified_files['zhu'])} 主条款", "info")
        except Exception as e:
            self._log(f"❌ ZIP生成失败: {sanitize_error_message(e)}", "error")

    def _download_excel_report(self):
        """下载Excel报告 - Anthropic风格"""
        if not self.extracted_data:
            self._log("⚠️ 没有可导出的数据", "warning")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存Excel报告",
            f"附加条款提取_{datetime.now():%Y%m%d_%H%M}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if not save_path:
            return

        try:
            # 按分类分组
            grouped = defaultdict(list)
            for item in self.extracted_data:
                cat = item.get('Category', '其他附加条款')
                grouped[cat].append(item)

            # 创建工作簿
            wb = openpyxl.Workbook()
            wb.remove(wb.active)

            # Anthropic 风格颜色
            header_fill = PatternFill(start_color="141413", end_color="141413", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True, size=11)
            accent_fill = PatternFill(start_color="FAF9F5", end_color="FAF9F5", fill_type="solid")
            success_font = Font(color="5A9A7A")
            error_font = Font(color="C75050")
            border_style = Border(
                left=Side(style='thin', color='E0DED5'),
                right=Side(style='thin', color='E0DED5'),
                top=Side(style='thin', color='E0DED5'),
                bottom=Side(style='thin', color='E0DED5')
            )

            for sheet_name, items in grouped.items():
                safe_name = sheet_name[:30].replace('/', ' ').replace('\\', ' ')
                ws = wb.create_sheet(title=safe_name)

                # 表头
                headers = ['条款名称', '注册号', '条款内容', '原文件名', '添加日期', '状态']
                ws.append(headers)

                # 表头样式
                for col in range(1, 7):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                # 数据
                for row_idx, item in enumerate(items, start=2):
                    ws.append([
                        item['ClauseName'],
                        item['RegistrationNo'],
                        item['Content'][:30000] if item['Content'] else '',
                        item['FileName'],
                        item['AddDate'],
                        f"失败: {item['Error']}" if item.get('Error') else '成功'
                    ])

                    # 数据行样式
                    for col in range(1, 7):
                        cell = ws.cell(row=row_idx, column=col)
                        cell.border = border_style
                        cell.alignment = Alignment(vertical='center', wrap_text=(col == 3))
                        if row_idx % 2 == 0:
                            cell.fill = accent_fill

                    # 状态列颜色
                    status_cell = ws.cell(row=row_idx, column=6)
                    if item.get('Error'):
                        status_cell.font = error_font
                    else:
                        status_cell.font = success_font

                # 列宽
                ws.column_dimensions['A'].width = 40
                ws.column_dimensions['B'].width = 25
                ws.column_dimensions['C'].width = 100
                ws.column_dimensions['D'].width = 45
                ws.column_dimensions['E'].width = 12
                ws.column_dimensions['F'].width = 12

                # 冻结首行
                ws.freeze_panes = 'A2'

            wb.save(save_path)
            self._log(f"✅ Excel报告已保存: {os.path.basename(save_path)}", "success")

            # 打开文件所在目录（跨平台）
            if platform.system() == 'Windows':
                os.startfile(os.path.dirname(save_path))
            elif platform.system() == 'Darwin':
                subprocess.run(['open', '-R', save_path], check=False)
            else:
                subprocess.run(['xdg-open', os.path.dirname(save_path)], check=False)

        except Exception as e:
            self._log(f"❌ Excel导出失败: {sanitize_error_message(e)}", "error")

    def _clear_all(self):
        """清空所有"""
        self.selected_files = []
        self.classified_files = {'fujia': [], 'feilv': [], 'zhu': []}
        self.doc_files = []
        self.extracted_data = []
        self.categories = set()

        self.file_list.clear()
        self.file_list.setVisible(False)
        self.classify_preview.setVisible(False)
        self.download_zip_btn.setVisible(False)
        self.download_excel_btn.setVisible(False)
        self.extract_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)

        self._update_stats()
        self._log("🗑 已清空所有文件", "info")

    def _update_stats(self):
        """更新统计 - 有数据时显示统计面板"""
        total = len(self.selected_files)
        extracted = len(self.extracted_data)
        categories = len(self.categories)
        skipped = len([d for d in self.extracted_data if d.get('Error')])

        self.stat_total_label.setText(f"待处理: {total}")
        self.stat_extracted_label.setText(f"已提取: {extracted}")
        self.stat_categories_label.setText(f"分类数: {categories}")
        self.stat_skipped_label.setText(f"已跳过: {skipped}")

        # 有任何数据时显示统计面板，否则隐藏
        has_data = total > 0 or extracted > 0
        self.stats_frame.setVisible(has_data)

    def _log(self, message: str, level: str = "info"):
        """添加日志"""
        colors = {
            'info': '#e0e0e0',
            'success': '#7ec9a0',
            'warning': '#e5c07b',
            'error': '#e06c75'
        }
        color = colors.get(level, '#e0e0e0')
        self.log_text.append(f'<span style="color: {color}">{message}</span>')


# ==========================================
# 条款输出Tab - V18.0 完整功能
# ==========================================
class ClauseOutputTab(QWidget):
    """条款输出Tab - Excel/提取结果转Word文档"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = parent
        self.report_data = []  # 存储读取的条款数据
        self.selected_clauses = []  # 用户选中的条款
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 15, 20, 15)

        # 紧凑型标题栏
        header = QHBoxLayout()
        title = QLabel("📝 条款输出")
        title.setStyleSheet(f"""
            color: {AnthropicColors.TEXT_PRIMARY};
            font-size: 18px;
            font-weight: bold;
            font-family: 'Söhne', 'SF Pro Display', -apple-system, sans-serif;
        """)
        header.addWidget(title)
        header.addStretch()

        # 输出模式选择
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["按条款逐个输出", "按分类合并输出", "全部合并为一个文档"])
        self.mode_combo.setStyleSheet(f"""
            QComboBox {{
                background: {AnthropicColors.BG_CARD};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 6px;
                padding: 8px 12px;
                color: {AnthropicColors.TEXT_PRIMARY};
                min-width: 160px;
            }}
            QComboBox:hover {{ border-color: {AnthropicColors.ACCENT}; }}
            QComboBox::drop-down {{
                border: none;
                width: 20px;
            }}
            QComboBox QAbstractItemView {{
                background: {AnthropicColors.BG_PRIMARY};
                color: {AnthropicColors.TEXT_PRIMARY};
                selection-background-color: {AnthropicColors.ACCENT};
                selection-color: white;
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 6px;
                padding: 5px;
            }}
        """)
        mode_label = QLabel("输出模式:")
        mode_label.setStyleSheet(f"color: {AnthropicColors.TEXT_PRIMARY}; font-weight: 500;")
        header.addWidget(mode_label)
        header.addWidget(self.mode_combo)
        layout.addLayout(header)

        # 数据源选择卡片
        source_card = GlassCard()
        source_layout = QVBoxLayout(source_card)
        source_layout.setSpacing(12)

        source_title = QLabel("📊 选择数据源")
        source_title.setStyleSheet(f"color: {AnthropicColors.ACCENT}; font-weight: 600; font-size: 14px;")
        source_layout.addWidget(source_title)

        # 数据源按钮行
        source_btn_layout = QHBoxLayout()

        self.from_extract_btn = QPushButton("📄 从条款提取获取")
        self.from_extract_btn.setCursor(Qt.PointingHandCursor)
        self.from_extract_btn.setStyleSheet(f"""
            QPushButton {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                padding: 12px 20px;
                color: {AnthropicColors.TEXT_PRIMARY};
                font-weight: 500;
            }}
            QPushButton:hover {{
                border-color: {AnthropicColors.ACCENT};
                background: rgba(217, 119, 87, 0.08);
            }}
        """)
        self.from_extract_btn.clicked.connect(self._load_from_extractor)

        self.from_file_btn = QPushButton("📁 从Excel文件加载")
        self.from_file_btn.setCursor(Qt.PointingHandCursor)
        self.from_file_btn.setStyleSheet(self.from_extract_btn.styleSheet())
        self.from_file_btn.clicked.connect(self._load_from_excel)

        source_btn_layout.addWidget(self.from_extract_btn)
        source_btn_layout.addWidget(self.from_file_btn)
        source_btn_layout.addStretch()
        source_layout.addLayout(source_btn_layout)

        # 文件路径显示
        self.source_label = QLabel("未选择数据源")
        self.source_label.setStyleSheet(f"color: {AnthropicColors.TEXT_MUTED}; font-size: 12px; padding: 5px 0;")
        source_layout.addWidget(self.source_label)

        layout.addWidget(source_card)

        # 条款预览列表
        preview_card = GlassCard()
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setSpacing(10)

        preview_header = QHBoxLayout()
        preview_title = QLabel("📋 条款预览")
        preview_title.setStyleSheet(f"color: {AnthropicColors.ACCENT}; font-weight: 600; font-size: 14px;")
        preview_header.addWidget(preview_title)

        self.clause_count_label = QLabel("共 0 条")
        self.clause_count_label.setStyleSheet(f"color: {AnthropicColors.TEXT_MUTED}; font-size: 12px;")
        preview_header.addWidget(self.clause_count_label)
        preview_header.addStretch()

        # 全选/取消按钮
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                color: {AnthropicColors.ACCENT};
                font-size: 12px;
                padding: 4px 8px;
            }}
            QPushButton:hover {{ text-decoration: underline; }}
        """)
        self.select_all_btn.clicked.connect(self._toggle_select_all)
        preview_header.addWidget(self.select_all_btn)

        preview_layout.addLayout(preview_header)

        # 条款列表
        self.clause_list = QListWidget()
        self.clause_list.setMinimumHeight(200)
        self.clause_list.setStyleSheet(f"""
            QListWidget {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                padding: 8px;
                font-size: 13px;
                color: {AnthropicColors.TEXT_PRIMARY};
            }}
            QListWidget::item {{
                padding: 10px 12px;
                border-radius: 6px;
                margin-bottom: 2px;
            }}
            QListWidget::item:hover {{
                background: rgba(217, 119, 87, 0.08);
            }}
            QListWidget::item:selected {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT};
            }}
        """)
        preview_layout.addWidget(self.clause_list)

        layout.addWidget(preview_card)

        # Word样式设置卡片
        style_card = GlassCard()
        style_layout = QVBoxLayout(style_card)
        style_layout.setSpacing(10)

        style_title = QLabel("🎨 Word样式设置")
        style_title.setStyleSheet(f"color: {AnthropicColors.ACCENT}; font-weight: 600; font-size: 14px;")
        style_layout.addWidget(style_title)

        style_grid = QHBoxLayout()

        # 标签通用样式
        label_style = f"color: {AnthropicColors.TEXT_PRIMARY}; font-size: 13px; font-weight: 500;"
        spin_style = f"""
            QSpinBox, QDoubleSpinBox {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 6px;
                padding: 8px;
                color: {AnthropicColors.TEXT_PRIMARY};
                font-size: 13px;
            }}
            QSpinBox:focus, QDoubleSpinBox:focus {{
                border-color: {AnthropicColors.ACCENT};
            }}
        """

        # 标题字号
        title_size_layout = QVBoxLayout()
        title_label = QLabel("标题字号")
        title_label.setStyleSheet(label_style)
        title_size_layout.addWidget(title_label)
        self.title_size_spin = QSpinBox()
        self.title_size_spin.setRange(12, 36)
        self.title_size_spin.setValue(16)
        self.title_size_spin.setStyleSheet(spin_style)
        title_size_layout.addWidget(self.title_size_spin)
        style_grid.addLayout(title_size_layout)

        # 正文字号
        body_size_layout = QVBoxLayout()
        body_label = QLabel("正文字号")
        body_label.setStyleSheet(label_style)
        body_size_layout.addWidget(body_label)
        self.body_size_spin = QSpinBox()
        self.body_size_spin.setRange(9, 18)
        self.body_size_spin.setValue(12)
        self.body_size_spin.setStyleSheet(spin_style)
        body_size_layout.addWidget(self.body_size_spin)
        style_grid.addLayout(body_size_layout)

        # 行距
        line_spacing_layout = QVBoxLayout()
        spacing_label = QLabel("行距")
        spacing_label.setStyleSheet(label_style)
        line_spacing_layout.addWidget(spacing_label)
        self.line_spacing_spin = QDoubleSpinBox()
        self.line_spacing_spin.setRange(1.0, 3.0)
        self.line_spacing_spin.setValue(1.5)
        self.line_spacing_spin.setSingleStep(0.25)
        self.line_spacing_spin.setStyleSheet(spin_style)
        line_spacing_layout.addWidget(self.line_spacing_spin)
        style_grid.addLayout(line_spacing_layout)

        # 包含注册号
        include_reg_layout = QVBoxLayout()
        reg_label = QLabel("包含注册号")
        reg_label.setStyleSheet(label_style)
        include_reg_layout.addWidget(reg_label)
        self.include_reg_check = QCheckBox("显示")
        self.include_reg_check.setChecked(True)
        self.include_reg_check.setStyleSheet(f"color: {AnthropicColors.TEXT_PRIMARY}; font-size: 13px;")
        include_reg_layout.addWidget(self.include_reg_check)
        style_grid.addLayout(include_reg_layout)

        style_grid.addStretch()
        style_layout.addLayout(style_grid)

        layout.addWidget(style_card)

        # 操作按钮行
        btn_layout = QHBoxLayout()

        self.generate_btn = QPushButton("📄 生成Word文档")
        self.generate_btn.setMinimumHeight(48)
        self.generate_btn.setCursor(Qt.PointingHandCursor)
        self.generate_btn.setEnabled(False)
        self.generate_btn.setStyleSheet(f"""
            QPushButton {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT};
                border: none;
                border-radius: 8px;
                font-size: 15px;
                font-weight: 600;
            }}
            QPushButton:hover {{ background: #2a2a28; }}
            QPushButton:disabled {{
                background: {AnthropicColors.BORDER};
                color: {AnthropicColors.TEXT_SECONDARY};
            }}
        """)
        self.generate_btn.clicked.connect(self._generate_word)

        self.preview_btn = QPushButton("👁 预览")
        self.preview_btn.setMinimumHeight(48)
        self.preview_btn.setCursor(Qt.PointingHandCursor)
        self.preview_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {AnthropicColors.TEXT_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                font-size: 14px;
                padding: 0 25px;
            }}
            QPushButton:hover {{
                border-color: {AnthropicColors.ACCENT};
                background: rgba(217, 119, 87, 0.05);
            }}
        """)
        self.preview_btn.clicked.connect(self._preview_output)

        btn_layout.addWidget(self.generate_btn, 3)
        btn_layout.addWidget(self.preview_btn, 1)
        layout.addLayout(btn_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{ background: {AnthropicColors.BORDER}; border-radius: 2px; }}
            QProgressBar::chunk {{ background: {AnthropicColors.ACCENT}; border-radius: 2px; }}
        """)
        layout.addWidget(self.progress_bar)

        # 日志区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet(f"""
            QTextEdit {{
                background: {AnthropicColors.BG_DARK};
                border: none;
                border-radius: 8px;
                color: {AnthropicColors.TEXT_LIGHT};
                padding: 12px;
                font-family: 'Söhne Mono', 'SF Mono', 'Menlo', monospace;
                font-size: 12px;
            }}
        """)
        layout.addWidget(self.log_text)

        self._log("📝 条款输出工具已就绪", "info")
        self._log("   支持从条款提取结果或Excel文件加载数据", "info")

    def _load_from_extractor(self):
        """从条款提取Tab获取数据"""
        if not self.parent_window:
            self._log("❌ 无法获取父窗口引用", "error")
            return

        try:
            extractor_tab = self.parent_window.extractor_tab
            if not extractor_tab.extracted_data:
                self._log("⚠️ 条款提取Tab中没有已提取的数据", "warning")
                self._log("   请先在「条款提取」页面提取条款", "info")
                return

            self.report_data = []
            for item in extractor_tab.extracted_data:
                if not item.get('Error'):
                    self.report_data.append({
                        'name': item.get('ClauseName', ''),
                        'regNo': item.get('RegistrationNo', ''),
                        'content': item.get('Content', ''),
                        'category': item.get('Category', '其他'),
                        'filename': item.get('FileName', '')
                    })

            if self.report_data:
                self._update_clause_list()
                self.source_label.setText(f"✓ 已从条款提取加载 {len(self.report_data)} 条数据")
                self._log(f"✓ 从条款提取Tab加载了 {len(self.report_data)} 条条款", "success")
                self.generate_btn.setEnabled(True)
            else:
                self._log("⚠️ 没有成功提取的条款数据", "warning")

        except Exception as e:
            self._log(f"❌ 加载失败: {sanitize_error_message(e)}", "error")

    def _load_from_excel(self):
        """从Excel文件加载数据"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "",
            "Excel文件 (*.xlsx);;所有文件 (*.*)"
        )
        if not file_path:
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(10)

        try:
            self._log(f"📖 读取文件: {os.path.basename(file_path)}", "info")

            wb = openpyxl.load_workbook(file_path, read_only=True)
            self.report_data = []

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                headers = [cell.value for cell in ws[1]] if ws[1] else []

                # 智能识别列
                col_map = self._detect_columns(headers)

                if not col_map.get('name'):
                    continue

                self.progress_bar.setValue(30)

                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row or not any(row):
                        continue

                    name = self._safe_get(row, col_map.get('name'))
                    if not name:
                        continue

                    self.report_data.append({
                        'name': name,
                        'regNo': self._safe_get(row, col_map.get('regNo')),
                        'content': self._safe_get(row, col_map.get('content')),
                        'category': sheet_name if sheet_name != 'Sheet' else '条款',
                        'filename': self._safe_get(row, col_map.get('filename'))
                    })

            wb.close()
            self.progress_bar.setValue(80)

            # 去重
            seen = set()
            unique_data = []
            for item in self.report_data:
                if item['name'] not in seen:
                    seen.add(item['name'])
                    unique_data.append(item)
            self.report_data = unique_data

            if self.report_data:
                self._update_clause_list()
                self.source_label.setText(f"✓ {os.path.basename(file_path)} ({len(self.report_data)} 条)")
                self._log(f"✓ 加载了 {len(self.report_data)} 条不重复条款", "success")
                self.generate_btn.setEnabled(True)
            else:
                self._log("⚠️ 文件中未找到有效条款数据", "warning")

        except Exception as e:
            self._log(f"❌ 读取Excel失败: {sanitize_error_message(e)}", "error")
        finally:
            self.progress_bar.setVisible(False)

    def _detect_columns(self, headers: list) -> dict:
        """智能识别Excel列 - 优先匹配「匹配1_」前缀的列（E/F/G）"""
        col_map = {}

        # 第一优先级：查找「匹配1_」前缀的列
        for i, h in enumerate(headers):
            if not h:
                continue
            h_str = str(h)
            if '匹配1_条款名称' in h_str or h_str == '匹配1_条款名称':
                col_map['name'] = i
            elif '匹配1_注册号' in h_str or '匹配1_产品注册号' in h_str:
                col_map['regNo'] = i
            elif '匹配1_条款内容' in h_str:
                col_map['content'] = i

        # 如果找到了匹配1_列，记录日志
        if col_map.get('name'):
            self._log(f"✓ 识别到匹配列: E={col_map.get('name')}, F={col_map.get('regNo')}, G={col_map.get('content')}", "success")
            return col_map

        # 第二优先级：直接使用E/F/G列（索引4/5/6）
        col_map['name'] = 4      # E列 = 匹配1_条款名称
        col_map['regNo'] = 5     # F列 = 匹配1_注册号
        col_map['content'] = 6   # G列 = 匹配1_条款内容
        self._log("ℹ️ 使用默认列: E=条款名称, F=注册号, G=条款内容", "info")

        return col_map

    def _safe_get(self, row: tuple, index: int) -> str:
        """安全获取行数据"""
        if index is None or index >= len(row):
            return ''
        return str(row[index] or '').strip()

    def _update_clause_list(self):
        """更新条款列表显示"""
        self.clause_list.clear()

        for item in self.report_data:
            list_item = QListWidgetItem()
            list_item.setCheckState(Qt.Checked)

            # 显示格式：条款名称 (分类)
            display_text = item['name']
            if item.get('category'):
                display_text += f"  [{item['category']}]"
            if item.get('regNo'):
                display_text += f"  ({item['regNo'][:20]}...)" if len(item.get('regNo', '')) > 20 else f"  ({item['regNo']})"

            list_item.setText(display_text)
            list_item.setData(Qt.UserRole, item)
            self.clause_list.addItem(list_item)

        self.clause_count_label.setText(f"共 {len(self.report_data)} 条")

    def _toggle_select_all(self):
        """切换全选/取消"""
        # 检查当前状态
        all_checked = all(
            self.clause_list.item(i).checkState() == Qt.Checked
            for i in range(self.clause_list.count())
        )

        new_state = Qt.Unchecked if all_checked else Qt.Checked
        for i in range(self.clause_list.count()):
            self.clause_list.item(i).setCheckState(new_state)

        self.select_all_btn.setText("取消全选" if not all_checked else "全选")

    def _get_selected_clauses(self) -> list:
        """获取选中的条款"""
        selected = []
        for i in range(self.clause_list.count()):
            item = self.clause_list.item(i)
            if item.checkState() == Qt.Checked:
                selected.append(item.data(Qt.UserRole))
        return selected

    def _preview_output(self):
        """预览输出"""
        selected = self._get_selected_clauses()
        if not selected:
            self._log("⚠️ 请至少选择一条条款", "warning")
            return

        preview_text = f"将输出 {len(selected)} 条条款:\n\n"
        for i, clause in enumerate(selected[:10], 1):
            preview_text += f"{i}. {clause['name']}\n"
        if len(selected) > 10:
            preview_text += f"... 还有 {len(selected) - 10} 条\n"

        preview_text += f"\n输出模式: {self.mode_combo.currentText()}"
        preview_text += f"\n标题字号: {self.title_size_spin.value()}pt"
        preview_text += f"\n正文字号: {self.body_size_spin.value()}pt"

        QMessageBox.information(self, "输出预览", preview_text)

    def _generate_word(self):
        """生成Word文档"""
        selected = self._get_selected_clauses()
        if not selected:
            self._log("⚠️ 请至少选择一条条款", "warning")
            return

        output_mode = self.mode_combo.currentIndex()

        if output_mode == 0:
            # 按条款逐个输出 - 选择输出目录
            output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
            if not output_dir:
                return
            self._generate_individual_docs(selected, output_dir)

        elif output_mode == 1:
            # 按分类合并输出 - 选择输出目录
            output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
            if not output_dir:
                return
            self._generate_category_docs(selected, output_dir)

        else:
            # 全部合并为一个文档
            save_path, _ = QFileDialog.getSaveFileName(
                self, "保存Word文档",
                f"条款汇总_{datetime.now():%Y%m%d_%H%M}.docx",
                "Word文档 (*.docx)"
            )
            if not save_path:
                return
            self._generate_combined_doc(selected, save_path)

    def _generate_individual_docs(self, clauses: list, output_dir: str):
        """按条款逐个生成Word文档"""
        self.progress_bar.setVisible(True)
        self._log(f"📄 开始生成 {len(clauses)} 个独立文档...", "info")

        success_count = 0
        for i, clause in enumerate(clauses):
            progress = int((i + 1) / len(clauses) * 100)
            self.progress_bar.setValue(progress)
            QApplication.processEvents()

            try:
                # 清理文件名
                safe_name = re.sub(r'[\\/*?:"<>|]', '_', clause['name'])[:50]
                file_path = os.path.join(output_dir, f"{safe_name}.docx")

                doc = self._create_clause_document(clause)
                doc.save(file_path)
                success_count += 1

            except Exception as e:
                self._log(f"  ✗ {clause['name']}: {str(e)}", "error")

        self.progress_bar.setValue(100)
        self._log(f"✅ 完成! 成功生成 {success_count}/{len(clauses)} 个文档", "success")
        self._log(f"   输出目录: {output_dir}", "info")
        self.progress_bar.setVisible(False)

        # 打开输出目录（跨平台）
        if platform.system() == 'Windows':
            os.startfile(output_dir)
        elif platform.system() == 'Darwin':
            subprocess.run(['open', output_dir], check=False)
        else:
            subprocess.run(['xdg-open', output_dir], check=False)

    def _generate_category_docs(self, clauses: list, output_dir: str):
        """按分类生成Word文档"""
        self.progress_bar.setVisible(True)

        # 按分类分组
        categorized = defaultdict(list)
        for clause in clauses:
            cat = clause.get('category', '其他') or '其他'
            categorized[cat].append(clause)

        self._log(f"📄 按 {len(categorized)} 个分类生成文档...", "info")

        total = len(categorized)
        for i, (category, cat_clauses) in enumerate(categorized.items()):
            progress = int((i + 1) / total * 100)
            self.progress_bar.setValue(progress)
            QApplication.processEvents()

            try:
                safe_cat = re.sub(r'[\\/*?:"<>|]', '_', category)[:30]
                file_path = os.path.join(output_dir, f"{safe_cat}_条款汇总.docx")

                doc = self._create_category_document(category, cat_clauses)
                doc.save(file_path)
                self._log(f"  ✓ {category}: {len(cat_clauses)} 条条款", "success")

            except Exception as e:
                self._log(f"  ✗ {category}: {str(e)}", "error")

        self.progress_bar.setValue(100)
        self._log(f"✅ 完成! 输出目录: {output_dir}", "success")
        self.progress_bar.setVisible(False)

        # 打开输出目录（跨平台）
        if platform.system() == 'Windows':
            os.startfile(output_dir)
        elif platform.system() == 'Darwin':
            subprocess.run(['open', output_dir], check=False)
        else:
            subprocess.run(['xdg-open', output_dir], check=False)

    def _set_run_font(self, run, size_pt: int, bold: bool = False, color_rgb=None):
        """设置run的字体：宋体(中文) + Times New Roman(英文)"""
        from docx.shared import Pt, RGBColor
        from docx.oxml.ns import qn

        run.font.size = Pt(size_pt)
        run.font.name = 'Times New Roman'  # 英文字体
        run.bold = bold

        # 设置中文字体为宋体
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

        if color_rgb:
            run.font.color.rgb = color_rgb

    def _generate_combined_doc(self, clauses: list, save_path: str):
        """生成合并的Word文档 - 使用宋体+Times New Roman"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(20)

        try:
            self._log(f"📄 生成合并文档，共 {len(clauses)} 条条款...", "info")

            doc = Document()

            from docx.shared import Pt, RGBColor
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            from docx.oxml.ns import qn

            # 文档标题
            title_para = doc.add_paragraph()
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_run = title_para.add_run('条款汇总清单')
            self._set_run_font(title_run, self.title_size_spin.value() + 4, bold=True)

            # 生成日期
            date_para = doc.add_paragraph()
            date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            date_run = date_para.add_run(f"生成日期: {datetime.now():%Y年%m月%d日}")
            self._set_run_font(date_run, 10, color_rgb=RGBColor(128, 128, 128))

            doc.add_paragraph()

            self.progress_bar.setValue(40)

            # 按分类组织
            categorized = defaultdict(list)
            for clause in clauses:
                cat = clause.get('category', '其他') or '其他'
                categorized[cat].append(clause)

            clause_num = 1
            for category, cat_clauses in categorized.items():
                # 分类标题
                cat_para = doc.add_paragraph()
                cat_run = cat_para.add_run(f"【{category}】")
                self._set_run_font(cat_run, self.title_size_spin.value(), bold=True, color_rgb=RGBColor(217, 119, 87))

                for clause in cat_clauses:
                    # 条款名称
                    name_para = doc.add_paragraph()
                    name_run = name_para.add_run(f"{clause_num}. {clause['name']}")
                    self._set_run_font(name_run, self.title_size_spin.value(), bold=True)

                    # 注册号
                    if self.include_reg_check.isChecked() and clause.get('regNo'):
                        reg_para = doc.add_paragraph()
                        reg_run = reg_para.add_run(f"注册号: {clause['regNo']}")
                        self._set_run_font(reg_run, self.body_size_spin.value(), color_rgb=RGBColor(100, 100, 100))

                    # 条款内容
                    if clause.get('content'):
                        for para_text in clause['content'].split('\n'):
                            para_text = para_text.strip()
                            if para_text:
                                content_para = doc.add_paragraph()
                                content_run = content_para.add_run(para_text)
                                self._set_run_font(content_run, self.body_size_spin.value())
                                content_para.paragraph_format.line_spacing = self.line_spacing_spin.value()

                    doc.add_paragraph()
                    clause_num += 1

            self.progress_bar.setValue(80)

            doc.save(save_path)

            self.progress_bar.setValue(100)
            self._log(f"✅ Word文档已生成: {os.path.basename(save_path)}", "success")
            self._log(f"   共导出 {len(clauses)} 条条款，{len(categorized)} 个分类", "info")

            # 打开生成的文档（跨平台）
            if platform.system() == 'Windows':
                os.startfile(save_path)
            elif platform.system() == 'Darwin':
                subprocess.run(['open', save_path], check=False)
            else:
                subprocess.run(['xdg-open', save_path], check=False)

        except Exception as e:
            self._log(f"❌ 生成失败: {sanitize_error_message(e)}", "error")
            logger.error(f"生成Word文档失败: {e}")  # 完整错误记录到日志
        finally:
            self.progress_bar.setVisible(False)

    def _create_clause_document(self, clause: dict) -> Document:
        """创建单个条款的Word文档 - 使用宋体+Times New Roman"""
        from docx.shared import RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = Document()

        # 条款名称
        title_para = doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run(clause['name'])
        self._set_run_font(title_run, self.title_size_spin.value(), bold=True)

        # 注册号
        if self.include_reg_check.isChecked() and clause.get('regNo'):
            reg_para = doc.add_paragraph()
            reg_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            reg_run = reg_para.add_run(clause['regNo'])
            self._set_run_font(reg_run, self.body_size_spin.value(), color_rgb=RGBColor(100, 100, 100))

        doc.add_paragraph()

        # 条款内容
        if clause.get('content'):
            for line in clause['content'].split('\n'):
                line = line.strip()
                if line:
                    para = doc.add_paragraph()
                    run = para.add_run(line)
                    self._set_run_font(run, self.body_size_spin.value())
                    para.paragraph_format.line_spacing = self.line_spacing_spin.value()

        return doc

    def _create_category_document(self, category: str, clauses: list) -> Document:
        """创建分类条款文档 - 使用宋体+Times New Roman"""
        from docx.shared import RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = Document()

        # 分类标题
        title_para = doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run(f"【{category}】条款汇总")
        self._set_run_font(title_run, self.title_size_spin.value() + 2, bold=True)

        date_para = doc.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_run = date_para.add_run(f"共 {len(clauses)} 条 · {datetime.now():%Y-%m-%d}")
        self._set_run_font(date_run, 10, color_rgb=RGBColor(128, 128, 128))

        doc.add_paragraph()

        for i, clause in enumerate(clauses, 1):
            # 条款名称
            name_para = doc.add_paragraph()
            name_run = name_para.add_run(f"{i}. {clause['name']}")
            self._set_run_font(name_run, self.title_size_spin.value(), bold=True)

            if self.include_reg_check.isChecked() and clause.get('regNo'):
                reg_para = doc.add_paragraph()
                reg_run = reg_para.add_run(f"注册号: {clause['regNo']}")
                self._set_run_font(reg_run, self.body_size_spin.value(), color_rgb=RGBColor(100, 100, 100))

            if clause.get('content'):
                for line in clause['content'].split('\n'):
                    line = line.strip()
                    if line:
                        para = doc.add_paragraph()
                        run = para.add_run(line)
                        self._set_run_font(run, self.body_size_spin.value())
                        para.paragraph_format.line_spacing = self.line_spacing_spin.value()

            doc.add_paragraph()

        return doc

    def _log(self, message: str, level: str = "info"):
        """添加日志"""
        colors = {
            'info': '#e0e0e0',
            'success': '#7ec9a0',
            'warning': '#e5c07b',
            'error': '#e06c75'
        }
        color = colors.get(level, '#e0e0e0')
        self.log_text.append(f'<span style="color: {color}">{message}</span>')


class ClauseComparisonAssistant(QMainWindow):
    """主界面 - Anthropic 风格 - V18.0 Tab版"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("智能条款工具箱 V18.0")
        self.setMinimumSize(1000, 900)
        # Anthropic 风格：温暖的奶油白背景
        self.setStyleSheet(f"""
            QMainWindow {{
                background: {AnthropicColors.BG_PRIMARY};
            }}
        """)

        if HAS_CONFIG_MANAGER:
            self._config = get_config()
        else:
            self._config = None

        # 初始化映射管理器
        if HAS_MAPPING_MANAGER:
            self._mapping_manager = get_mapping_manager()
            self._mapping_manager.load()
            # 应用用户映射到配置
            if self._config:
                self._mapping_manager.apply_to_config(self._config)
        else:
            self._mapping_manager = None

        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(12)
        layout.setContentsMargins(30, 20, 30, 20)

        # 标题行 - Anthropic 风格
        header_layout = QHBoxLayout()

        title = QLabel("🔧 智能条款工具箱")
        title.setStyleSheet(f"color: {AnthropicColors.TEXT_PRIMARY}; font-size: 26px; font-weight: bold;")
        header_layout.addWidget(title)

        header_layout.addStretch()

        # 版本信息
        subtitle = QLabel("V18.0 · 条款提取 · 条款比对 · 条款输出")
        subtitle.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 12px;")
        header_layout.addWidget(subtitle)

        # 支持作者按钮 - Anthropic 强调色风格
        self.donate_btn = QPushButton("💝 支持作者")
        self.donate_btn.setCursor(Qt.PointingHandCursor)
        self.donate_btn.setStyleSheet(f"""
            QPushButton {{
                background: {AnthropicColors.ACCENT};
                color: {AnthropicColors.TEXT_LIGHT};
                border: none;
                border-radius: 15px;
                padding: 6px 16px;
                font-size: 12px;
                font-weight: 500;
                margin-left: 15px;
            }}
            QPushButton:hover {{
                background: {AnthropicColors.ACCENT_DARK};
            }}
        """)
        self.donate_btn.clicked.connect(self._show_donate_dialog)

        # 为支持作者按钮添加柔和阴影
        donate_shadow = QGraphicsDropShadowEffect()
        donate_shadow.setBlurRadius(12)
        donate_shadow.setColor(QColor(217, 119, 87, 80))  # Anthropic ACCENT 色
        donate_shadow.setOffset(0, 2)
        self.donate_btn.setGraphicsEffect(donate_shadow)

        # 呼吸动画定时器
        self._donate_glow_step = 0
        self._donate_timer = QTimer(self)
        self._donate_timer.timeout.connect(self._animate_donate_button)
        self._donate_timer.start(50)  # 50ms间隔

        header_layout.addWidget(self.donate_btn)
        layout.addLayout(header_layout)

        # ==========================================
        # 主Tab区域 - Anthropic风格
        # ==========================================
        self.main_tabs = QTabWidget()
        self.main_tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: none;
                background: {AnthropicColors.BG_PRIMARY};
            }}
            QTabBar::tab {{
                background: {AnthropicColors.BG_CARD};
                color: {AnthropicColors.TEXT_SECONDARY};
                border: none;
                padding: 14px 40px;
                margin-right: 8px;
                border-radius: 8px 8px 0 0;
                font-size: 14px;
                font-weight: 600;
                min-width: 140px;
            }}
            QTabBar::tab:selected {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT};
            }}
            QTabBar::tab:hover:!selected {{
                background: {AnthropicColors.BG_CARD};
                color: {AnthropicColors.TEXT_PRIMARY};
            }}
        """)

        # Tab 1: 条款提取
        self.extractor_tab = ClauseExtractorTab(self)
        self.main_tabs.addTab(self.extractor_tab, "📄 条款提取")

        # Tab 2: 条款比对
        self.comparison_tab = self._create_comparison_tab()
        self.main_tabs.addTab(self.comparison_tab, "🔍 条款比对")

        # Tab 3: 条款输出
        self.output_tab = ClauseOutputTab(self)
        self.main_tabs.addTab(self.output_tab, "📝 条款输出")

        layout.addWidget(self.main_tabs, 1)

        # 版本信息
        version = QLabel("V18.0 Multi-Tab Edition · Made with ❤️ by Dachi Yijin")
        version.setAlignment(Qt.AlignCenter)
        version.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 11px;")
        layout.addWidget(version)

    def _create_comparison_tab(self) -> QWidget:
        """创建条款比对Tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 15, 20, 15)

        # 配置统计
        if self._config:
            stats = self._config.get_stats()
            user_mappings = self._mapping_manager.get_mapping_count() if self._mapping_manager else 0
            stats_text = f"📊 {stats['client_mappings']} 映射 | {user_mappings} 自定义 | {stats['semantic_aliases']} 别名"
        else:
            stats_text = "📊 使用内置配置"
        self.stats_label = QLabel(stats_text)
        self.stats_label.setAlignment(Qt.AlignCenter)
        self.stats_label.setStyleSheet(f"color: {AnthropicColors.TEXT_SECONDARY}; font-size: 11px;")
        layout.addWidget(self.stats_label)

        # 输入卡片 - Anthropic 风格
        card = GlassCard()
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(18)
        card_layout.setContentsMargins(30, 30, 30, 30)

        # Anthropic 风格的输入框样式
        style = f"""
            QLabel {{ color: {AnthropicColors.TEXT_PRIMARY}; font-weight: 500; }}
            QLineEdit {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px; padding: 12px 15px;
                color: {AnthropicColors.TEXT_PRIMARY}; font-size: 14px;
            }}
            QLineEdit:focus {{ border-color: {AnthropicColors.ACCENT}; }}
        """
        card.setStyleSheet(card.styleSheet() + style)

        btn_style = f"""
            QPushButton {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px; padding: 12px 18px;
                color: {AnthropicColors.TEXT_PRIMARY}; font-weight: 500;
            }}
            QPushButton:hover {{ background: {AnthropicColors.BG_CARD}; border-color: {AnthropicColors.ACCENT}; }}
        """

        self.doc_input = self._create_file_row(card_layout, "📂 客户文档",
            "Word 条款清单 (.docx)", "Word Files (*.docx)", btn_style)
        self.lib_input = self._create_file_row(card_layout, "📚 条款库",
            "Excel 条款库 (.xlsx)", "Excel Files (*.xlsx)", btn_style)

        # 添加Sheet选择行
        sheet_row = QHBoxLayout()
        sheet_label = QLabel("📋 险种Sheet")
        sheet_label.setFixedWidth(90)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumHeight(40)
        self.sheet_combo.setStyleSheet(f"""
            QComboBox {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 8px;
                padding: 8px 12px;
                color: {AnthropicColors.TEXT_PRIMARY};
                font-size: 13px;
            }}
            QComboBox:hover {{ border-color: {AnthropicColors.ACCENT}; }}
            QComboBox::drop-down {{
                border: none;
                width: 30px;
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid {AnthropicColors.ACCENT};
            }}
            QComboBox QAbstractItemView {{
                background: {AnthropicColors.BG_PRIMARY};
                border: 1px solid {AnthropicColors.ACCENT};
                selection-background-color: {AnthropicColors.ACCENT};
                color: {AnthropicColors.TEXT_PRIMARY};
            }}
        """)
        self.sheet_combo.addItem("自动选择第一个Sheet")
        self.sheet_combo.setToolTip("选择条款库中的险种Sheet（如财产险、责任险等）")
        # 当条款库文件改变时更新Sheet列表
        self.lib_input.textChanged.connect(self._update_sheet_list)
        sheet_row.addWidget(sheet_label)
        sheet_row.addWidget(self.sheet_combo, 1)
        card_layout.addLayout(sheet_row)

        line = QFrame()
        line.setFixedHeight(1)
        line.setStyleSheet(f"background: {AnthropicColors.BORDER};")
        card_layout.addWidget(line)

        row3 = QHBoxLayout()
        label3 = QLabel("💾 保存路径")
        label3.setFixedWidth(90)
        self.out_input = QLineEdit()
        self.out_input.setPlaceholderText("报告保存位置...")
        btn3 = QPushButton("选择")
        btn3.setCursor(Qt.PointingHandCursor)
        btn3.setStyleSheet(btn_style)
        btn3.clicked.connect(self._browse_save)
        row3.addWidget(label3)
        row3.addWidget(self.out_input, 1)
        row3.addWidget(btn3)
        card_layout.addLayout(row3)

        layout.addWidget(card)

        # 按钮行
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)

        self.start_btn = QPushButton("🚀 开始比对")
        self.start_btn.setCursor(Qt.PointingHandCursor)
        self.start_btn.setMinimumHeight(52)
        self.start_btn.setStyleSheet(f"""
            QPushButton {{
                background: {AnthropicColors.BG_DARK};
                color: {AnthropicColors.TEXT_LIGHT}; font-size: 16px; font-weight: bold;
                border-radius: 12px; border: none;
            }}
            QPushButton:hover {{
                background: {AnthropicColors.ACCENT};
            }}
            QPushButton:disabled {{ background: {AnthropicColors.BORDER_DARK}; color: {AnthropicColors.TEXT_MUTED}; }}
        """)
        self.start_btn.clicked.connect(self._start_process)

        self.batch_btn = QPushButton("📦 批量处理")
        self.batch_btn.setCursor(Qt.PointingHandCursor)
        self.batch_btn.setMinimumHeight(52)
        self.batch_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: {AnthropicColors.TEXT_PRIMARY};
                font-size: 14px; font-weight: 500;
                border-radius: 8px; border: 1px solid {AnthropicColors.BG_DARK};
            }}
            QPushButton:hover {{ background: {AnthropicColors.BG_DARK}; color: {AnthropicColors.TEXT_LIGHT}; }}
        """)
        self.batch_btn.clicked.connect(self._show_batch_dialog)

        self.add_btn = QPushButton("🔧 映射设置")
        self.add_btn.setCursor(Qt.PointingHandCursor)
        self.add_btn.setMinimumHeight(52)
        self.add_btn.setStyleSheet(self.batch_btn.styleSheet())
        self.add_btn.clicked.connect(self._show_add_mapping_dialog)

        # v17.1: 条款查询按钮
        self.query_btn = QPushButton("🔍 条款查询")
        self.query_btn.setCursor(Qt.PointingHandCursor)
        self.query_btn.setMinimumHeight(52)
        self.query_btn.setStyleSheet(self.batch_btn.styleSheet())
        self.query_btn.clicked.connect(self._show_query_dialog)

        self.open_btn = QPushButton("📂 打开目录")
        self.open_btn.setCursor(Qt.PointingHandCursor)
        self.open_btn.setMinimumHeight(52)
        self.open_btn.setEnabled(False)
        self.open_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: {AnthropicColors.TEXT_SECONDARY};
                font-size: 14px; font-weight: 500;
                border-radius: 8px; border: 1px solid {AnthropicColors.BORDER};
            }}
            QPushButton:hover {{ border-color: {AnthropicColors.ACCENT}; color: {AnthropicColors.ACCENT}; }}
            QPushButton:disabled {{ color: {AnthropicColors.BORDER}; border-color: {AnthropicColors.BORDER}; }}
        """)
        self.open_btn.clicked.connect(self._open_output_folder)

        btn_layout.addWidget(self.start_btn, 3)
        btn_layout.addWidget(self.batch_btn, 1)
        btn_layout.addWidget(self.add_btn, 1)
        btn_layout.addWidget(self.query_btn, 1)  # v17.1: 条款查询
        btn_layout.addWidget(self.open_btn, 1)
        layout.addLayout(btn_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{ background: {AnthropicColors.BORDER}; border-radius: 2px; }}
            QProgressBar::chunk {{
                background: {AnthropicColors.ACCENT};
                border-radius: 2px;
            }}
        """)
        layout.addWidget(self.progress_bar)

        # 日志
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(f"""
            QTextEdit {{
                background: {AnthropicColors.BG_CARD};
                border: 1px solid {AnthropicColors.BORDER};
                border-radius: 12px; color: {AnthropicColors.TEXT_PRIMARY};
                padding: 15px;
                font-family: 'JetBrains Mono', 'SF Mono', 'Menlo', monospace;
                font-size: 12px;
            }}
        """)
        layout.addWidget(self.log_text, 1)

        return tab

    def _create_file_row(self, layout, label_text: str, placeholder: str,
                         filter_str: str, btn_style: str) -> QLineEdit:
        row = QHBoxLayout()
        label = QLabel(label_text)
        label.setFixedWidth(90)
        line_edit = QLineEdit()
        line_edit.setPlaceholderText(placeholder)
        btn = QPushButton("浏览")
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(btn_style)
        btn.clicked.connect(lambda: self._browse_file(line_edit, filter_str))
        row.addWidget(label)
        row.addWidget(line_edit, 1)
        row.addWidget(btn)
        layout.addLayout(row)
        return line_edit

    def _browse_file(self, line_edit: QLineEdit, filter_str: str):
        f, _ = QFileDialog.getOpenFileName(self, "选择文件", "", filter_str)
        if f:
            line_edit.setText(f)
            if line_edit == self.doc_input and not self.out_input.text():
                self.out_input.setText(os.path.join(os.path.dirname(f), "条款比对报告.xlsx"))

    def _browse_save(self):
        f, _ = QFileDialog.getSaveFileName(self, "保存结果", "条款比对报告.xlsx", "Excel Files (*.xlsx)")
        if f:
            self.out_input.setText(f)

    def _animate_donate_button(self):
        """支持作者按钮的呼吸发光动画"""
        import math
        self._donate_glow_step = (self._donate_glow_step + 1) % 120

        # 使用正弦函数创建平滑的呼吸效果
        glow_intensity = int(80 + 70 * math.sin(self._donate_glow_step * math.pi / 60))
        blur_radius = int(12 + 8 * math.sin(self._donate_glow_step * math.pi / 60))

        effect = self.donate_btn.graphicsEffect()
        if effect and isinstance(effect, QGraphicsDropShadowEffect):
            effect.setBlurRadius(blur_radius)
            effect.setColor(QColor(217, 119, 87, glow_intensity))  # Anthropic accent color

    def _show_donate_dialog(self):
        """显示支持作者对话框"""
        dialog = DonateDialog(self)
        dialog.exec_()

    def _show_add_mapping_dialog(self):
        """打开条款映射管理对话框"""
        if HAS_MAPPING_MANAGER:
            # 获取当前条款库中的条款名称列表（用于下拉提示）
            library_clauses = self._get_library_clauses()

            dialog = ClauseMappingDialog(self, library_clauses=library_clauses)
            dialog.mappings_changed.connect(self._on_mappings_changed)
            dialog.exec_()
        elif self._config:
            # 兼容旧版：使用简单的添加对话框
            dialog = AddMappingDialog(self)
            if dialog.exec_() == QDialog.Accepted:
                eng, chn = dialog.get_mapping()
                if eng and chn:
                    self._config.add_client_mapping(eng, chn)
                    self._config.save()
                    self._append_log(f"✓ 已添加映射: '{eng}' -> '{chn}'", "success")
        else:
            QMessageBox.warning(self, "提示", "映射管理功能不可用")

    def _show_query_dialog(self):
        """v17.1: 打开条款查询对话框"""
        library_path = self.lib_input.text().strip()
        if not library_path or not os.path.exists(library_path):
            QMessageBox.warning(self, "提示", "请先选择条款库文件！")
            return

        try:
            # 加载条款库并构建索引
            logic = ClauseMatcherLogic()
            sheet_name = self._get_selected_sheet()
            lib_data = LibraryLoader.load_excel(library_path, sheet_name=sheet_name)
            library_index = logic.build_index(lib_data)

            # 获取映射管理器
            mapping_mgr = get_mapping_manager() if HAS_MAPPING_MANAGER else None

            # 打开查询对话框
            dialog = ClauseQueryDialog(
                parent=self,
                library_index=library_index,
                logic=logic,
                mapping_mgr=mapping_mgr
            )
            dialog.exec_()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"加载条款库失败: {sanitize_error_message(e)}")

    def _get_library_clauses(self) -> List[str]:
        """从当前条款库获取条款名称列表"""
        library_path = self.lib_input.text().strip()
        if not library_path or not os.path.exists(library_path):
            return []

        try:
            clauses = []
            wb = openpyxl.load_workbook(library_path, read_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(max_col=1, values_only=True):
                    if row[0] and isinstance(row[0], str):
                        name = row[0].strip()
                        if name and len(name) > 3 and name not in clauses:
                            clauses.append(name)
            wb.close()
            return clauses[:500]  # 限制数量防止内存问题
        except Exception as e:
            logger.warning(f"读取条款库失败: {e}")
            return []

    def _on_mappings_changed(self):
        """映射变更回调：更新配置"""
        if HAS_MAPPING_MANAGER and self._config:
            mapping_manager = get_mapping_manager()
            count = mapping_manager.apply_to_config(self._config)
            self._append_log(f"✓ 已应用 {count} 条用户映射", "success")

            # 更新统计显示
            stats = self._config.get_stats()
            user_mappings = mapping_manager.get_mapping_count()
            self.stats_label.setText(f"📊 {stats['client_mappings']} 映射 | {user_mappings} 自定义 | {stats['semantic_aliases']} 别名")

    def _show_batch_dialog(self):
        if not self.lib_input.text():
            QMessageBox.warning(self, "提示", "请先选择条款库")
            return

        dialog = BatchSelectDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            files = dialog.get_files()
            if not files:
                return

            output_dir = QFileDialog.getExistingDirectory(self, "选择输出目录")
            if not output_dir:
                return

            self._start_batch_process(files, output_dir)

    def _append_log(self, msg: str, level: str):
        colors = {
            "info": AnthropicColors.TEXT_SECONDARY,
            "success": AnthropicColors.SUCCESS,
            "error": AnthropicColors.ERROR,
            "warning": AnthropicColors.WARNING
        }
        self.log_text.append(f'<span style="color:{colors.get(level, AnthropicColors.TEXT_PRIMARY)}">{msg}</span>')
        self.log_text.moveCursor(QTextCursor.End)

    def _start_process(self):
        doc = self.doc_input.text().strip()
        excel = self.lib_input.text().strip()
        out = self.out_input.text().strip()

        if not all([doc, excel, out]):
            QMessageBox.warning(self, "提示", "请完善所有文件路径！")
            return

        self._set_ui_state(False)
        self.log_text.clear()

        # 获取选择的Sheet名称
        sheet_name = self._get_selected_sheet()

        self.worker = MatchWorker(doc, excel, out, sheet_name)
        self.worker.log_signal.connect(self._append_log)
        self.worker.progress_signal.connect(lambda c, t: self.progress_bar.setValue(int(c/t*100)))
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.start()

    def _start_batch_process(self, files: List[str], output_dir: str):
        self._set_ui_state(False)
        self.log_text.clear()

        # 获取选择的Sheet名称
        sheet_name = self._get_selected_sheet()

        self.batch_worker = BatchMatchWorker(files, self.lib_input.text(), output_dir, sheet_name)
        self.batch_worker.log_signal.connect(self._append_log)
        self.batch_worker.batch_progress_signal.connect(
            lambda c, t, n: self.progress_bar.setValue(int(c/t*100))
        )
        self.batch_worker.finished_signal.connect(self._on_batch_finished)
        self.batch_worker.start()

    def _get_selected_sheet(self) -> Optional[str]:
        """获取选择的Sheet名称"""
        if self.sheet_combo.currentIndex() == 0:  # "自动选择第一个Sheet"
            return None
        return self.sheet_combo.currentText()

    def _update_sheet_list(self, excel_path: str):
        """当条款库文件改变时更新Sheet列表"""
        self.sheet_combo.clear()
        self.sheet_combo.addItem("自动选择第一个Sheet")

        if not excel_path or not os.path.exists(excel_path):
            return

        try:
            sheets = LibraryLoader.get_sheet_names(excel_path)
            if sheets:
                for sheet in sheets:
                    self.sheet_combo.addItem(sheet)
                # 如果只有一个Sheet，保持默认选择
                if len(sheets) > 1:
                    self._append_log(f"📋 检测到 {len(sheets)} 个Sheet: {', '.join(sheets)}", "info")
        except Exception as e:
            logger.warning(f"读取Sheet列表失败: {e}")

    def _set_ui_state(self, enabled: bool):
        self.start_btn.setEnabled(enabled)
        self.batch_btn.setEnabled(enabled)
        self.start_btn.setText("🚀 开始比对" if enabled else "⏳ 处理中...")
        self.progress_bar.setVisible(not enabled)
        if not enabled:
            self.progress_bar.setValue(0)

    def _on_finished(self, success: bool, msg: str):
        self._set_ui_state(True)
        if success:
            self.open_btn.setEnabled(True)
            self.open_btn.setStyleSheet("""
                QPushButton {
                    background: transparent; color: #2ecc71;
                    font-size: 14px; font-weight: 500;
                    border-radius: 26px; border: 2px solid #2ecc71;
                }
                QPushButton:hover { background: #2ecc71; color: white; }
            """)
            QMessageBox.information(self, "完成", f"比对完成！\n{msg}")

    def _on_batch_finished(self, success: bool, msg: str, ok_count: int, total: int):
        self._set_ui_state(True)
        if success:
            self.open_btn.setEnabled(True)
            QMessageBox.information(self, "完成", f"批量处理完成！\n成功: {ok_count}/{total}\n输出目录: {msg}")

    def _open_output_folder(self):
        path = self.out_input.text().strip()
        if path and os.path.exists(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))


def main():
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setFont(QFont("PingFang SC", 13))

    window = ClauseComparisonAssistant()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
