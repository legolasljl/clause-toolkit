# -*- coding: utf-8 -*-
"""
设置页面UI模块
提供客户类型配置的可视化编辑界面
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QListWidget,
    QPushButton, QLabel, QLineEdit, QTextEdit, QTableWidget,
    QTableWidgetItem, QTabWidget, QComboBox, QCheckBox,
    QMessageBox, QFileDialog, QHeaderView, QSplitter,
    QFormLayout, QSpinBox, QDialog, QDialogButtonBox,
    QListWidgetItem, QAbstractItemView, QDoubleSpinBox,
    QScrollArea, QFrame
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont


class SettingsPage(QWidget):
    """设置页面"""

    config_changed = pyqtSignal()

    def __init__(self, config_manager):
        super().__init__()
        self.config_manager = config_manager
        self.current_customer_type = None
        self._init_ui()
        self._load_data()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        # 顶部工具栏
        toolbar = QHBoxLayout()
        self.btn_import = QPushButton("📥 导入配置")
        self.btn_export = QPushButton("📤 导出配置")
        self.btn_reset = QPushButton("🔄 重置默认")
        self.btn_import.clicked.connect(self._import_config)
        self.btn_export.clicked.connect(self._export_config)
        self.btn_reset.clicked.connect(self._reset_config)
        toolbar.addWidget(self.btn_import)
        toolbar.addWidget(self.btn_export)
        toolbar.addWidget(self.btn_reset)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        # 主区域：左侧客户类型列表 + 右侧编辑区
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左侧：客户类型列表
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 10, 0)

        lbl_types = QLabel("👤 客户类型")
        lbl_types.setStyleSheet("font-weight: bold; font-size: 14px; color: #e0e0ff;")
        left_layout.addWidget(lbl_types)

        self.list_customer_types = QListWidget()
        self.list_customer_types.currentRowChanged.connect(self._on_customer_type_selected)
        left_layout.addWidget(self.list_customer_types)

        btn_row = QHBoxLayout()
        self.btn_add_type = QPushButton("➕ 新增")
        self.btn_del_type = QPushButton("🗑️ 删除")
        self.btn_add_type.clicked.connect(self._add_customer_type)
        self.btn_del_type.clicked.connect(self._delete_customer_type)
        btn_row.addWidget(self.btn_add_type)
        btn_row.addWidget(self.btn_del_type)
        left_layout.addLayout(btn_row)

        splitter.addWidget(left_panel)

        # 右侧：编辑区（使用 Tab）
        right_panel = QTabWidget()

        # Tab 1: 基本信息
        self.tab_basic = self._create_basic_tab()
        right_panel.addTab(self.tab_basic, "📋 基本信息")

        # Tab 2: 汇总表表头
        self.tab_headers = self._create_headers_tab()
        right_panel.addTab(self.tab_headers, "📊 表头列定义")

        # Tab 3: 占位符
        self.tab_placeholders = self._create_placeholders_tab()
        right_panel.addTab(self.tab_placeholders, "🏷️ 占位符")

        # Tab 4: 协议编号
        self.tab_agreements = self._create_agreements_tab()
        right_panel.addTab(self.tab_agreements, "📝 协议编号")

        # Tab 5: PDF分组规则
        self.tab_pdf_groups = self._create_pdf_groups_tab()
        right_panel.addTab(self.tab_pdf_groups, "📑 PDF分组")

        # Tab 6: 排版列宽定义
        self.tab_processed_widths = self._create_processed_widths_tab()
        right_panel.addTab(self.tab_processed_widths, "📐 排版列宽")

        # Tab 7: 对账单分组规则
        self.tab_statement_groups = self._create_statement_groups_tab()
        right_panel.addTab(self.tab_statement_groups, "📄 对账单分组")

        # Tab 8: 表头识别规则
        self.tab_header_recognition = self._create_header_recognition_tab()
        right_panel.addTab(self.tab_header_recognition, "🔍 表头识别")

        # Tab 9: 数据提取规则
        self.tab_data_extraction = self._create_data_extraction_tab()
        right_panel.addTab(self.tab_data_extraction, "📥 数据提取")

        # Tab 10: 费率计算配置
        self.tab_rate_calculation = self._create_rate_calculation_tab()
        right_panel.addTab(self.tab_rate_calculation, "💰 费率计算")

        # Tab 11: 特殊列配置
        self.tab_special_columns = self._create_special_columns_tab()
        right_panel.addTab(self.tab_special_columns, "📏 特殊列")

        # Tab 12: 行高配置
        self.tab_row_heights = self._create_row_heights_tab()
        right_panel.addTab(self.tab_row_heights, "📐 行高配置")

        splitter.addWidget(right_panel)
        splitter.setSizes([180, 600])

        layout.addWidget(splitter)

        # 底部保存按钮
        btn_save_row = QHBoxLayout()
        btn_save_row.addStretch()
        self.btn_save = QPushButton("💾 保存配置")
        self.btn_save.setObjectName("runBtn")
        self.btn_save.clicked.connect(self._save_config)
        btn_save_row.addWidget(self.btn_save)
        layout.addLayout(btn_save_row)

    def _create_basic_tab(self) -> QWidget:
        """创建基本信息标签页"""
        widget = QWidget()
        layout = QFormLayout(widget)
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        self.edit_id = QLineEdit()
        self.edit_id.setPlaceholderText("例: multimodal, hengli_pta")
        self.edit_display_name = QLineEdit()
        self.edit_display_name.setPlaceholderText("显示在界面上的名称")
        self.edit_description = QLineEdit()
        self.edit_description.setPlaceholderText("可选的描述信息")
        self.check_enabled = QCheckBox("启用此客户类型")
        self.check_enabled.setChecked(True)

        self.combo_extract_func = QComboBox()
        self.combo_extract_func.addItems([
            "extract_sheet_data",
            "extract_hengli_data",
            "extract_hengli_energy_data"
        ])
        self.combo_extract_func.setEditable(True)

        self.combo_process_func = QComboBox()
        self.combo_process_func.addItems([
            "process_multimodal_sheet",
            "process_hengli_sheet",
            "process_hengli_energy_sheet"
        ])
        self.combo_process_func.setEditable(True)

        self.check_use_formatter = QCheckBox("使用Excel格式化器 (ExcelFormatter)")
        self.check_use_formatter.setChecked(True)

        self.check_is_hengli = QCheckBox("恒力模式 (特殊标题格式)")

        self.combo_page_orientation = QComboBox()
        self.combo_page_orientation.addItems(["landscape", "portrait"])
        self.combo_page_orientation.setToolTip("landscape=横向打印, portrait=纵向打印")

        self.edit_log_format = QLineEdit()
        self.edit_log_format.setPlaceholderText("例: {sheet_name}: 费率={rate:.3f}%, 保费={new_premium:,.2f}")

        layout.addRow("标识符 (ID):", self.edit_id)
        layout.addRow("显示名称:", self.edit_display_name)
        layout.addRow("描述:", self.edit_description)
        layout.addRow("", self.check_enabled)
        layout.addRow("数据提取函数:", self.combo_extract_func)
        layout.addRow("数据处理函数:", self.combo_process_func)
        layout.addRow("", self.check_use_formatter)
        layout.addRow("", self.check_is_hengli)
        layout.addRow("打印方向:", self.combo_page_orientation)
        layout.addRow("日志格式:", self.edit_log_format)

        return widget

    def _create_headers_tab(self) -> QWidget:
        """创建表头列定义标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义汇总表的列，每列包含字段Key（对应数据字典的键）、列标题、宽度等属性")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        # 表格
        self.table_headers = QTableWidget()
        self.table_headers.setColumnCount(7)
        self.table_headers.setHorizontalHeaderLabels([
            "字段Key", "列标题", "宽度", "格式", "求和", "换行", "固定值"
        ])
        self.table_headers.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_headers.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table_headers.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table_headers)

        # 按钮行
        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加列")
        btn_del = QPushButton("🗑️ 删除列")
        btn_up = QPushButton("⬆️ 上移")
        btn_down = QPushButton("⬇️ 下移")
        btn_add.clicked.connect(self._add_header_column)
        btn_del.clicked.connect(self._del_header_column)
        btn_up.clicked.connect(self._move_header_up)
        btn_down.clicked.connect(self._move_header_down)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addWidget(btn_up)
        btn_row.addWidget(btn_down)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # 其他设置
        other_layout = QHBoxLayout()
        other_layout.addWidget(QLabel("合计行标签列:"))
        self.spin_sum_label_col = QSpinBox()
        self.spin_sum_label_col.setRange(1, 20)
        self.spin_sum_label_col.setValue(1)
        other_layout.addWidget(self.spin_sum_label_col)
        other_layout.addSpacing(20)
        other_layout.addWidget(QLabel("人民币大写列:"))
        self.spin_rmb_col = QSpinBox()
        self.spin_rmb_col.setRange(0, 20)
        self.spin_rmb_col.setSpecialValueText("无")
        other_layout.addWidget(self.spin_rmb_col)
        other_layout.addStretch()
        layout.addLayout(other_layout)

        return widget

    def _create_placeholders_tab(self) -> QWidget:
        """创建占位符标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义投保单模板中可用的占位符。占位符格式为 {名称}，会被替换为实际数据。")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self.table_placeholders = QTableWidget()
        self.table_placeholders.setColumnCount(4)
        self.table_placeholders.setHorizontalHeaderLabels([
            "占位符", "数据字段", "格式化", "转换函数"
        ])
        self.table_placeholders.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_placeholders.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table_placeholders)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加")
        btn_del = QPushButton("🗑️ 删除")
        btn_add.clicked.connect(self._add_placeholder)
        btn_del.clicked.connect(self._del_placeholder)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # 帮助信息
        help_text = QLabel(
            "转换函数说明:\n"
            "  • cn_currency: 转换为人民币大写\n"
            "  • format_date_slashes: 格式化日期为 YYYY/MM/DD"
        )
        help_text.setStyleSheet("color: #666; font-size: 10px; margin-top: 10px;")
        layout.addWidget(help_text)

        return widget

    def _create_agreements_tab(self) -> QWidget:
        """创建协议编号标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义公司名称到协议编号的映射，用于自动填充投保单中的协议编号。")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self.table_agreements = QTableWidget()
        self.table_agreements.setColumnCount(2)
        self.table_agreements.setHorizontalHeaderLabels(["公司名称", "协议编号"])
        self.table_agreements.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_agreements.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table_agreements)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加")
        btn_del = QPushButton("🗑️ 删除")
        btn_add.clicked.connect(self._add_agreement)
        btn_del.clicked.connect(self._del_agreement)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        return widget

    def _create_pdf_groups_tab(self) -> QWidget:
        """创建PDF分组规则标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义PDF导出时按Sheet名称分组的规则，每个分组会生成一个独立的PDF文件。")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self.table_pdf_groups = QTableWidget()
        self.table_pdf_groups.setColumnCount(3)
        self.table_pdf_groups.setHorizontalHeaderLabels([
            "分组名称", "匹配类型", "匹配模式"
        ])
        self.table_pdf_groups.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_pdf_groups.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table_pdf_groups)

        help_text = QLabel(
            "匹配类型说明:\n"
            "  • contains: 名称包含指定字符串\n"
            "  • contains_any: 名称包含任意一个指定字符串（用英文逗号分隔）\n"
            "  • startswith: 名称以指定字符串开头\n"
            "  • endswith: 名称以指定字符串结尾"
        )
        help_text.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(help_text)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加")
        btn_del = QPushButton("🗑️ 删除")
        btn_add.clicked.connect(self._add_pdf_group)
        btn_del.clicked.connect(self._del_pdf_group)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        return widget

    def _create_processed_widths_tab(self) -> QWidget:
        """创建排版列宽定义标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义 _processed.xlsx 排版文件的列宽，按 A、B、C、D 等列区分。")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self.table_processed_widths = QTableWidget()
        self.table_processed_widths.setColumnCount(2)
        self.table_processed_widths.setHorizontalHeaderLabels(["列名", "宽度"])
        self.table_processed_widths.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_processed_widths.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table_processed_widths)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加")
        btn_del = QPushButton("🗑️ 删除")
        btn_add.clicked.connect(self._add_processed_width)
        btn_del.clicked.connect(self._del_processed_width)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        help_text = QLabel(
            "说明:\n"
            "  • 列名使用 Excel 格式，如 A、B、C、D、E 等\n"
            "  • 宽度为数字，如 13.5、15.8 等\n"
            "  • 未定义的列将使用默认宽度"
        )
        help_text.setStyleSheet("color: #666; font-size: 10px; margin-top: 10px;")
        layout.addWidget(help_text)

        return widget

    def _create_statement_groups_tab(self) -> QWidget:
        """创建对账单分组规则标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义生成对账单时按Sheet名称分组的规则，每个分组会生成对应的付款通知书。")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self.table_statement_groups = QTableWidget()
        self.table_statement_groups.setColumnCount(3)
        self.table_statement_groups.setHorizontalHeaderLabels([
            "分组名称", "匹配类型", "匹配模式"
        ])
        self.table_statement_groups.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_statement_groups.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table_statement_groups)

        help_text = QLabel(
            "匹配类型说明:\n"
            "  • contains: 名称包含指定字符串\n"
            "  • contains_any: 名称包含任意一个指定字符串（用英文逗号分隔）\n"
            "  • startswith: 名称以指定字符串开头\n"
            "  • endswith: 名称以指定字符串结尾\n"
            "注意：若只有一个分组且匹配模式为空，则所有Sheet数据合并为一份对账单"
        )
        help_text.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(help_text)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加")
        btn_del = QPushButton("🗑️ 删除")
        btn_add.clicked.connect(self._add_statement_group)
        btn_del.clicked.connect(self._del_statement_group)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        return widget

    def _create_header_recognition_tab(self) -> QWidget:
        """创建表头识别规则标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义如何识别Excel源文件中的表头行和合计行")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        # 表头关键词
        group1 = QGroupBox("表头行识别")
        group1_layout = QFormLayout(group1)

        self.edit_header_keywords = QLineEdit()
        self.edit_header_keywords.setPlaceholderText("用逗号分隔，如: 序号,编号")
        group1_layout.addRow("关键词:", self.edit_header_keywords)

        self.spin_max_search_rows = QSpinBox()
        self.spin_max_search_rows.setRange(1, 50)
        self.spin_max_search_rows.setValue(10)
        group1_layout.addRow("最大搜索行数:", self.spin_max_search_rows)

        layout.addWidget(group1)

        # 合计行关键词
        group2 = QGroupBox("合计行识别")
        group2_layout = QFormLayout(group2)

        self.edit_total_keywords = QLineEdit()
        self.edit_total_keywords.setPlaceholderText("用逗号分隔，如: 合计,总计")
        group2_layout.addRow("关键词:", self.edit_total_keywords)

        layout.addWidget(group2)

        # 列名映射
        group3 = QGroupBox("列名映射 (用于识别不同名称的相同数据列)")
        group3_layout = QVBoxLayout(group3)

        self.table_column_mappings = QTableWidget()
        self.table_column_mappings.setColumnCount(2)
        self.table_column_mappings.setHorizontalHeaderLabels(["标准字段名", "可能的列名(逗号分隔)"])
        self.table_column_mappings.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_column_mappings.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        group3_layout.addWidget(self.table_column_mappings)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加")
        btn_del = QPushButton("🗑️ 删除")
        btn_add.clicked.connect(self._add_column_mapping)
        btn_del.clicked.connect(self._del_column_mapping)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        group3_layout.addLayout(btn_row)

        layout.addWidget(group3)
        layout.addStretch()

        return widget

    def _create_data_extraction_tab(self) -> QWidget:
        """创建数据提取规则标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义如何从Excel源文件中提取和处理数据")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        # 业务笔数计算
        group1 = QGroupBox("业务笔数计算")
        group1_layout = QFormLayout(group1)

        self.combo_business_count_method = QComboBox()
        self.combo_business_count_method.addItems(["sequence", "count"])
        self.combo_business_count_method.setToolTip("sequence: 取序号列最大值\ncount: 统计有效数据行数")
        group1_layout.addRow("计算方式:", self.combo_business_count_method)

        self.spin_sequence_column = QSpinBox()
        self.spin_sequence_column.setRange(1, 20)
        self.spin_sequence_column.setValue(1)
        group1_layout.addRow("序号列位置:", self.spin_sequence_column)

        layout.addWidget(group1)

        # 日期处理
        group2 = QGroupBox("日期处理")
        group2_layout = QFormLayout(group2)

        self.edit_date_columns = QLineEdit()
        self.edit_date_columns.setPlaceholderText("用逗号分隔，如: 起运日期,发货日期")
        group2_layout.addRow("日期列名:", self.edit_date_columns)

        self.edit_date_format = QLineEdit()
        self.edit_date_format.setPlaceholderText("如: %Y/%m/%d")
        group2_layout.addRow("日期格式:", self.edit_date_format)

        layout.addWidget(group2)

        # 数据行处理
        group3 = QGroupBox("数据行处理")
        group3_layout = QFormLayout(group3)

        self.check_skip_empty_rows = QCheckBox("跳过空行")
        self.check_skip_empty_rows.setChecked(True)
        group3_layout.addRow("", self.check_skip_empty_rows)

        self.spin_data_start_offset = QSpinBox()
        self.spin_data_start_offset.setRange(0, 10)
        self.spin_data_start_offset.setValue(1)
        self.spin_data_start_offset.setToolTip("数据起始行相对于表头行的偏移量")
        group3_layout.addRow("数据起始偏移:", self.spin_data_start_offset)

        layout.addWidget(group3)

        # 数值列格式
        group4 = QGroupBox("数值列格式")
        group4_layout = QVBoxLayout(group4)

        self.table_numeric_columns = QTableWidget()
        self.table_numeric_columns.setColumnCount(2)
        self.table_numeric_columns.setHorizontalHeaderLabels(["列名", "格式"])
        self.table_numeric_columns.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_numeric_columns.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        group4_layout.addWidget(self.table_numeric_columns)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加")
        btn_del = QPushButton("🗑️ 删除")
        btn_add.clicked.connect(self._add_numeric_column)
        btn_del.clicked.connect(self._del_numeric_column)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        group4_layout.addLayout(btn_row)

        layout.addWidget(group4)
        layout.addStretch()

        return widget

    def _create_rate_calculation_tab(self) -> QWidget:
        """创建费率计算配置标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("定义费率的计算方式和显示格式")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        # 计算模式
        group1 = QGroupBox("计算模式")
        group1_layout = QFormLayout(group1)

        self.combo_calculation_mode = QComboBox()
        self.combo_calculation_mode.addItems(["auto", "fixed", "manual"])
        self.combo_calculation_mode.setToolTip("auto: 自动计算费率\nfixed: 使用固定费率\nmanual: 手动输入")
        self.combo_calculation_mode.currentTextChanged.connect(self._on_calculation_mode_changed)
        group1_layout.addRow("计算模式:", self.combo_calculation_mode)

        self.spin_fixed_rate = QDoubleSpinBox()
        self.spin_fixed_rate.setRange(0, 1)
        self.spin_fixed_rate.setDecimals(8)
        self.spin_fixed_rate.setSingleStep(0.0001)
        self.spin_fixed_rate.setToolTip("当模式为fixed时使用的固定费率(小数形式)")
        group1_layout.addRow("固定费率:", self.spin_fixed_rate)

        layout.addWidget(group1)

        # 精度设置
        group2 = QGroupBox("精度设置")
        group2_layout = QFormLayout(group2)

        self.spin_rate_precision = QSpinBox()
        self.spin_rate_precision.setRange(2, 12)
        self.spin_rate_precision.setValue(8)
        group2_layout.addRow("费率精度(小数位):", self.spin_rate_precision)

        self.spin_premium_precision = QSpinBox()
        self.spin_premium_precision.setRange(0, 6)
        self.spin_premium_precision.setValue(2)
        group2_layout.addRow("保费精度(小数位):", self.spin_premium_precision)

        layout.addWidget(group2)

        # 列名配置
        group3 = QGroupBox("相关列名")
        group3_layout = QFormLayout(group3)

        self.edit_rate_column = QLineEdit()
        self.edit_rate_column.setPlaceholderText("如: 费率,综合费率")
        group3_layout.addRow("费率列:", self.edit_rate_column)

        self.edit_premium_column = QLineEdit()
        self.edit_premium_column.setPlaceholderText("如: 保费,保险费")
        group3_layout.addRow("保费列:", self.edit_premium_column)

        self.edit_insurance_amount_column = QLineEdit()
        self.edit_insurance_amount_column.setPlaceholderText("如: 保险金额,保额")
        group3_layout.addRow("保险金额列:", self.edit_insurance_amount_column)

        layout.addWidget(group3)

        # 显示设置
        group4 = QGroupBox("显示设置")
        group4_layout = QFormLayout(group4)

        self.combo_display_unit = QComboBox()
        self.combo_display_unit.addItems(["%", "‰"])
        group4_layout.addRow("显示单位:", self.combo_display_unit)

        self.spin_display_multiplier = QDoubleSpinBox()
        self.spin_display_multiplier.setRange(1, 10000)
        self.spin_display_multiplier.setValue(100)
        self.spin_display_multiplier.setToolTip("用于将小数费率转换为显示值\n例如: 100表示乘以100显示为百分比")
        group4_layout.addRow("显示乘数:", self.spin_display_multiplier)

        self.edit_formula = QLineEdit()
        self.edit_formula.setPlaceholderText("如: premium = insurance_amount * rate")
        group4_layout.addRow("计算公式:", self.edit_formula)

        layout.addWidget(group4)
        layout.addStretch()

        return widget

    def _create_special_columns_tab(self) -> QWidget:
        """创建特殊列配置标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("为特定列设置特殊的宽度、对齐方式和格式化规则")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self.table_special_columns = QTableWidget()
        self.table_special_columns.setColumnCount(7)
        self.table_special_columns.setHorizontalHeaderLabels([
            "列名", "宽度", "换行", "对齐", "字号", "首Sheet宽度", "其他Sheet宽度"
        ])
        self.table_special_columns.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_special_columns.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self.table_special_columns)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("➕ 添加")
        btn_del = QPushButton("🗑️ 删除")
        btn_add.clicked.connect(self._add_special_column)
        btn_del.clicked.connect(self._del_special_column)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        help_text = QLabel(
            "说明:\n"
            "  • 列名: 对应表头中的列标题\n"
            "  • 对齐: left/center/right\n"
            "  • 首Sheet宽度: 仅对第一个Sheet生效的列宽\n"
            "  • 其他Sheet宽度: 对其他Sheet生效的列宽"
        )
        help_text.setStyleSheet("color: #666; font-size: 10px; margin-top: 10px;")
        layout.addWidget(help_text)

        return widget

    def _create_row_heights_tab(self) -> QWidget:
        """创建行高配置标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(10, 10, 10, 10)

        help_label = QLabel("配置生成Excel文件时各类型行的高度")
        help_label.setStyleSheet("color: #888; font-size: 11px;")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        form = QFormLayout()
        form.setSpacing(12)

        self.spin_title_row_height = QSpinBox()
        self.spin_title_row_height.setRange(10, 100)
        self.spin_title_row_height.setValue(39)
        form.addRow("标题行高:", self.spin_title_row_height)

        self.spin_second_row_height = QSpinBox()
        self.spin_second_row_height.setRange(10, 100)
        self.spin_second_row_height.setValue(33)
        form.addRow("第二行高:", self.spin_second_row_height)

        self.spin_header_row_height = QSpinBox()
        self.spin_header_row_height.setRange(10, 100)
        self.spin_header_row_height.setValue(32)
        form.addRow("表头行高:", self.spin_header_row_height)

        self.spin_data_row_height = QSpinBox()
        self.spin_data_row_height.setRange(10, 100)
        self.spin_data_row_height.setValue(15)
        form.addRow("数据行高:", self.spin_data_row_height)

        self.spin_total_row_height = QSpinBox()
        self.spin_total_row_height.setRange(10, 100)
        self.spin_total_row_height.setValue(18)
        form.addRow("合计行高:", self.spin_total_row_height)

        self.spin_footer_row_height = QSpinBox()
        self.spin_footer_row_height.setRange(10, 100)
        self.spin_footer_row_height.setValue(18)
        form.addRow("页脚行高:", self.spin_footer_row_height)

        self.check_auto_fit_height = QCheckBox("自动调整行高 (根据内容)")
        self.check_auto_fit_height.setChecked(True)
        form.addRow("", self.check_auto_fit_height)

        layout.addLayout(form)
        layout.addStretch()

        return widget

    def _load_data(self):
        """加载配置数据到UI"""
        self.list_customer_types.clear()
        for name in self.config_manager.get_all_customer_types():
            config = self.config_manager.get_customer_config(name)
            item = QListWidgetItem(name)
            if config and not config.enabled:
                item.setForeground(Qt.GlobalColor.gray)
            self.list_customer_types.addItem(item)

        if self.list_customer_types.count() > 0:
            self.list_customer_types.setCurrentRow(0)

    def _on_customer_type_selected(self, row):
        """客户类型选中事件"""
        if row < 0:
            return

        name = self.list_customer_types.item(row).text()
        self.current_customer_type = name
        config = self.config_manager.get_customer_config(name)

        if config:
            self._load_customer_config(config)

    def _load_customer_config(self, config):
        """加载客户类型配置到编辑区"""
        # 基本信息
        self.edit_id.setText(config.id)
        self.edit_display_name.setText(config.display_name)
        self.edit_description.setText(config.description)
        self.check_enabled.setChecked(config.enabled)
        self.combo_extract_func.setCurrentText(config.extract_function)
        self.combo_process_func.setCurrentText(config.process_function)
        self.check_use_formatter.setChecked(config.use_excel_formatter)
        self.check_is_hengli.setChecked(config.is_hengli)
        self.combo_page_orientation.setCurrentText(config.page_orientation or "landscape")
        self.edit_log_format.setText(config.log_format)

        # 表头列定义
        self._load_headers(config.summary_headers)

        # 占位符
        self._load_placeholders(config.placeholders)

        # 协议编号
        self._load_agreements(config.agreement_codes)

        # PDF分组规则
        self._load_pdf_groups(config.pdf_export_groups)

        # 排版列宽
        self._load_processed_widths(config.processed_column_widths)

        # 对账单分组规则
        self._load_statement_groups(config.statement_groups)

        # 表头识别规则
        self._load_header_recognition(config.header_recognition)

        # 数据提取规则
        self._load_data_extraction(config.data_extraction)

        # 费率计算配置
        self._load_rate_calculation(config.rate_calculation)

        # 特殊列配置
        self._load_special_columns(config.special_columns)

        # 行高配置
        self._load_row_heights(config.row_heights)

    def _load_headers(self, summary_headers):
        """加载表头配置"""
        self.table_headers.setRowCount(0)
        if not summary_headers:
            return

        for col in summary_headers.columns:
            row = self.table_headers.rowCount()
            self.table_headers.insertRow(row)
            self.table_headers.setItem(row, 0, QTableWidgetItem(col.key))
            self.table_headers.setItem(row, 1, QTableWidgetItem(col.label))
            self.table_headers.setItem(row, 2, QTableWidgetItem(str(col.width)))
            self.table_headers.setItem(row, 3, QTableWidgetItem(col.format or ""))

            check_sum = QCheckBox()
            check_sum.setChecked(col.sum)
            self.table_headers.setCellWidget(row, 4, check_sum)

            check_wrap = QCheckBox()
            check_wrap.setChecked(col.wrap_text)
            self.table_headers.setCellWidget(row, 5, check_wrap)

            self.table_headers.setItem(row, 6, QTableWidgetItem(col.fixed_value or ""))

        self.spin_sum_label_col.setValue(summary_headers.sum_row_label_column)
        self.spin_rmb_col.setValue(summary_headers.rmb_total_column or 0)

    def _load_placeholders(self, placeholders):
        """加载占位符配置"""
        self.table_placeholders.setRowCount(0)
        mappings = placeholders.get("mappings", {}) if placeholders else {}

        for placeholder, mapping in mappings.items():
            row = self.table_placeholders.rowCount()
            self.table_placeholders.insertRow(row)
            self.table_placeholders.setItem(row, 0, QTableWidgetItem(placeholder))

            if isinstance(mapping, str):
                self.table_placeholders.setItem(row, 1, QTableWidgetItem(mapping))
                self.table_placeholders.setItem(row, 2, QTableWidgetItem(""))
                self.table_placeholders.setItem(row, 3, QTableWidgetItem(""))
            elif isinstance(mapping, dict):
                self.table_placeholders.setItem(row, 1, QTableWidgetItem(mapping.get("field", "")))
                self.table_placeholders.setItem(row, 2, QTableWidgetItem(mapping.get("format", "")))
                self.table_placeholders.setItem(row, 3, QTableWidgetItem(mapping.get("transform", "")))

    def _load_agreements(self, agreements):
        """加载协议编号配置"""
        self.table_agreements.setRowCount(0)
        for company, code in (agreements or {}).items():
            row = self.table_agreements.rowCount()
            self.table_agreements.insertRow(row)
            self.table_agreements.setItem(row, 0, QTableWidgetItem(company))
            self.table_agreements.setItem(row, 1, QTableWidgetItem(code))

    def _load_pdf_groups(self, pdf_groups):
        """加载PDF分组规则"""
        self.table_pdf_groups.setRowCount(0)
        for group_name, rule in (pdf_groups or {}).items():
            row = self.table_pdf_groups.rowCount()
            self.table_pdf_groups.insertRow(row)
            self.table_pdf_groups.setItem(row, 0, QTableWidgetItem(group_name))
            self.table_pdf_groups.setItem(row, 1, QTableWidgetItem(rule.get("match_type", "")))
            patterns = rule.get("patterns", [])
            self.table_pdf_groups.setItem(row, 2, QTableWidgetItem(",".join(patterns)))

    def _load_processed_widths(self, processed_widths):
        """加载排版列宽配置"""
        self.table_processed_widths.setRowCount(0)
        for cw in (processed_widths or []):
            row = self.table_processed_widths.rowCount()
            self.table_processed_widths.insertRow(row)
            self.table_processed_widths.setItem(row, 0, QTableWidgetItem(cw.column))
            self.table_processed_widths.setItem(row, 1, QTableWidgetItem(str(cw.width)))

    def _load_statement_groups(self, statement_groups):
        """加载对账单分组规则"""
        self.table_statement_groups.setRowCount(0)
        for group_name, rule in (statement_groups or {}).items():
            row = self.table_statement_groups.rowCount()
            self.table_statement_groups.insertRow(row)
            self.table_statement_groups.setItem(row, 0, QTableWidgetItem(group_name))
            self.table_statement_groups.setItem(row, 1, QTableWidgetItem(rule.get("match_type", "")))
            patterns = rule.get("patterns", [])
            self.table_statement_groups.setItem(row, 2, QTableWidgetItem(",".join(patterns)))

    def _load_header_recognition(self, header_recognition):
        """加载表头识别规则配置"""
        if header_recognition:
            keywords = header_recognition.header_keywords or ["序号"]
            self.edit_header_keywords.setText(",".join(keywords))
            self.spin_max_search_rows.setValue(header_recognition.max_search_rows or 10)
            total_kw = header_recognition.total_keywords or ["合计"]
            self.edit_total_keywords.setText(",".join(total_kw))

            # 列名映射
            self.table_column_mappings.setRowCount(0)
            for field, aliases in (header_recognition.column_mappings or {}).items():
                row = self.table_column_mappings.rowCount()
                self.table_column_mappings.insertRow(row)
                self.table_column_mappings.setItem(row, 0, QTableWidgetItem(field))
                self.table_column_mappings.setItem(row, 1, QTableWidgetItem(",".join(aliases) if isinstance(aliases, list) else aliases))
        else:
            self.edit_header_keywords.setText("序号")
            self.spin_max_search_rows.setValue(10)
            self.edit_total_keywords.setText("合计")
            self.table_column_mappings.setRowCount(0)

    def _load_data_extraction(self, data_extraction):
        """加载数据提取规则配置"""
        if data_extraction:
            self.combo_business_count_method.setCurrentText(data_extraction.business_count_method or "sequence")
            self.spin_sequence_column.setValue(data_extraction.sequence_column or 1)
            date_cols = data_extraction.date_columns or ["起运日期", "发货日期"]
            self.edit_date_columns.setText(",".join(date_cols))
            self.edit_date_format.setText(data_extraction.date_format or "%Y/%m/%d")
            self.check_skip_empty_rows.setChecked(data_extraction.skip_empty_rows)
            self.spin_data_start_offset.setValue(data_extraction.data_start_offset or 1)

            # 数值列格式
            self.table_numeric_columns.setRowCount(0)
            for col_name, fmt in (data_extraction.numeric_columns or {}).items():
                row = self.table_numeric_columns.rowCount()
                self.table_numeric_columns.insertRow(row)
                self.table_numeric_columns.setItem(row, 0, QTableWidgetItem(col_name))
                self.table_numeric_columns.setItem(row, 1, QTableWidgetItem(fmt))
        else:
            self.combo_business_count_method.setCurrentText("sequence")
            self.spin_sequence_column.setValue(1)
            self.edit_date_columns.setText("起运日期,发货日期")
            self.edit_date_format.setText("%Y/%m/%d")
            self.check_skip_empty_rows.setChecked(True)
            self.spin_data_start_offset.setValue(1)
            self.table_numeric_columns.setRowCount(0)

    def _load_rate_calculation(self, rate_calculation):
        """加载费率计算配置"""
        if rate_calculation:
            self.combo_calculation_mode.setCurrentText(rate_calculation.calculation_mode or "auto")
            self.spin_fixed_rate.setValue(rate_calculation.fixed_rate or 0)
            self.spin_rate_precision.setValue(rate_calculation.rate_precision or 8)
            self.spin_premium_precision.setValue(rate_calculation.premium_precision or 2)
            self.edit_rate_column.setText(rate_calculation.rate_column or "")
            self.edit_premium_column.setText(rate_calculation.premium_column or "")
            self.edit_insurance_amount_column.setText(rate_calculation.insurance_amount_column or "")
            self.combo_display_unit.setCurrentText(rate_calculation.display_unit or "%")
            self.spin_display_multiplier.setValue(rate_calculation.display_multiplier or 100)
            self.edit_formula.setText(rate_calculation.formula or "premium = insurance_amount * rate")
        else:
            self.combo_calculation_mode.setCurrentText("auto")
            self.spin_fixed_rate.setValue(0)
            self.spin_rate_precision.setValue(8)
            self.spin_premium_precision.setValue(2)
            self.edit_rate_column.clear()
            self.edit_premium_column.clear()
            self.edit_insurance_amount_column.clear()
            self.combo_display_unit.setCurrentText("%")
            self.spin_display_multiplier.setValue(100)
            self.edit_formula.setText("premium = insurance_amount * rate")

    def _load_special_columns(self, special_columns):
        """加载特殊列配置"""
        self.table_special_columns.setRowCount(0)
        for sc in (special_columns or []):
            row = self.table_special_columns.rowCount()
            self.table_special_columns.insertRow(row)
            self.table_special_columns.setItem(row, 0, QTableWidgetItem(sc.column_name))
            self.table_special_columns.setItem(row, 1, QTableWidgetItem(str(sc.width)))

            check_wrap = QCheckBox()
            check_wrap.setChecked(sc.wrap_text)
            self.table_special_columns.setCellWidget(row, 2, check_wrap)

            self.table_special_columns.setItem(row, 3, QTableWidgetItem(sc.alignment or "center"))
            self.table_special_columns.setItem(row, 4, QTableWidgetItem(str(sc.font_size) if sc.font_size else ""))
            self.table_special_columns.setItem(row, 5, QTableWidgetItem(str(sc.first_sheet_width) if sc.first_sheet_width else ""))
            self.table_special_columns.setItem(row, 6, QTableWidgetItem(str(sc.other_sheet_width) if sc.other_sheet_width else ""))

    def _load_row_heights(self, row_heights):
        """加载行高配置"""
        if row_heights:
            self.spin_title_row_height.setValue(row_heights.title_row or 39)
            self.spin_second_row_height.setValue(row_heights.second_row or 33)
            self.spin_header_row_height.setValue(row_heights.header_row or 32)
            self.spin_data_row_height.setValue(row_heights.data_row or 15)
            self.spin_total_row_height.setValue(row_heights.total_row or 18)
            self.spin_footer_row_height.setValue(row_heights.footer_row or 18)
            self.check_auto_fit_height.setChecked(row_heights.auto_fit)
        else:
            self.spin_title_row_height.setValue(39)
            self.spin_second_row_height.setValue(33)
            self.spin_header_row_height.setValue(32)
            self.spin_data_row_height.setValue(15)
            self.spin_total_row_height.setValue(18)
            self.spin_footer_row_height.setValue(18)
            self.check_auto_fit_height.setChecked(True)

    # 表头列操作
    def _add_header_column(self):
        row = self.table_headers.rowCount()
        self.table_headers.insertRow(row)
        self.table_headers.setItem(row, 0, QTableWidgetItem("new_field"))
        self.table_headers.setItem(row, 1, QTableWidgetItem("新列"))
        self.table_headers.setItem(row, 2, QTableWidgetItem("15"))
        self.table_headers.setItem(row, 3, QTableWidgetItem(""))
        check_sum = QCheckBox()
        self.table_headers.setCellWidget(row, 4, check_sum)
        check_wrap = QCheckBox()
        self.table_headers.setCellWidget(row, 5, check_wrap)
        self.table_headers.setItem(row, 6, QTableWidgetItem(""))

    def _del_header_column(self):
        row = self.table_headers.currentRow()
        if row >= 0:
            self.table_headers.removeRow(row)

    def _move_header_up(self):
        row = self.table_headers.currentRow()
        if row > 0:
            self._swap_table_rows(self.table_headers, row, row - 1)
            self.table_headers.setCurrentCell(row - 1, 0)

    def _move_header_down(self):
        row = self.table_headers.currentRow()
        if row < self.table_headers.rowCount() - 1:
            self._swap_table_rows(self.table_headers, row, row + 1)
            self.table_headers.setCurrentCell(row + 1, 0)

    def _swap_table_rows(self, table, row1, row2):
        """交换表格两行"""
        for col in range(table.columnCount()):
            widget1 = table.cellWidget(row1, col)
            widget2 = table.cellWidget(row2, col)
            if widget1 and widget2 and isinstance(widget1, QCheckBox) and isinstance(widget2, QCheckBox):
                checked1 = widget1.isChecked()
                checked2 = widget2.isChecked()
                widget1.setChecked(checked2)
                widget2.setChecked(checked1)
            else:
                item1 = table.takeItem(row1, col)
                item2 = table.takeItem(row2, col)
                if item1:
                    table.setItem(row2, col, item1)
                if item2:
                    table.setItem(row1, col, item2)

    # 占位符操作
    def _add_placeholder(self):
        row = self.table_placeholders.rowCount()
        self.table_placeholders.insertRow(row)
        self.table_placeholders.setItem(row, 0, QTableWidgetItem("{新占位符}"))
        self.table_placeholders.setItem(row, 1, QTableWidgetItem("field_name"))
        self.table_placeholders.setItem(row, 2, QTableWidgetItem(""))
        self.table_placeholders.setItem(row, 3, QTableWidgetItem(""))

    def _del_placeholder(self):
        row = self.table_placeholders.currentRow()
        if row >= 0:
            self.table_placeholders.removeRow(row)

    # 协议编号操作
    def _add_agreement(self):
        row = self.table_agreements.rowCount()
        self.table_agreements.insertRow(row)
        self.table_agreements.setItem(row, 0, QTableWidgetItem("公司名称"))
        self.table_agreements.setItem(row, 1, QTableWidgetItem("协议编号"))

    def _del_agreement(self):
        row = self.table_agreements.currentRow()
        if row >= 0:
            self.table_agreements.removeRow(row)

    # PDF分组操作
    def _add_pdf_group(self):
        row = self.table_pdf_groups.rowCount()
        self.table_pdf_groups.insertRow(row)
        self.table_pdf_groups.setItem(row, 0, QTableWidgetItem("分组名称"))
        self.table_pdf_groups.setItem(row, 1, QTableWidgetItem("contains"))
        self.table_pdf_groups.setItem(row, 2, QTableWidgetItem("匹配文本"))

    def _del_pdf_group(self):
        row = self.table_pdf_groups.currentRow()
        if row >= 0:
            self.table_pdf_groups.removeRow(row)

    # 排版列宽操作
    def _add_processed_width(self):
        row = self.table_processed_widths.rowCount()
        self.table_processed_widths.insertRow(row)
        # 自动计算下一个列名
        next_col = chr(ord('A') + row) if row < 26 else f"A{row - 25}"
        self.table_processed_widths.setItem(row, 0, QTableWidgetItem(next_col))
        self.table_processed_widths.setItem(row, 1, QTableWidgetItem("10.0"))

    def _del_processed_width(self):
        row = self.table_processed_widths.currentRow()
        if row >= 0:
            self.table_processed_widths.removeRow(row)

    # 对账单分组操作
    def _add_statement_group(self):
        row = self.table_statement_groups.rowCount()
        self.table_statement_groups.insertRow(row)
        self.table_statement_groups.setItem(row, 0, QTableWidgetItem("分组名称"))
        self.table_statement_groups.setItem(row, 1, QTableWidgetItem("contains"))
        self.table_statement_groups.setItem(row, 2, QTableWidgetItem("匹配文本"))

    def _del_statement_group(self):
        row = self.table_statement_groups.currentRow()
        if row >= 0:
            self.table_statement_groups.removeRow(row)

    # 列名映射操作
    def _add_column_mapping(self):
        row = self.table_column_mappings.rowCount()
        self.table_column_mappings.insertRow(row)
        self.table_column_mappings.setItem(row, 0, QTableWidgetItem("field_name"))
        self.table_column_mappings.setItem(row, 1, QTableWidgetItem("列名1,列名2"))

    def _del_column_mapping(self):
        row = self.table_column_mappings.currentRow()
        if row >= 0:
            self.table_column_mappings.removeRow(row)

    # 数值列格式操作
    def _add_numeric_column(self):
        row = self.table_numeric_columns.rowCount()
        self.table_numeric_columns.insertRow(row)
        self.table_numeric_columns.setItem(row, 0, QTableWidgetItem("列名"))
        self.table_numeric_columns.setItem(row, 1, QTableWidgetItem("#,##0.00"))

    def _del_numeric_column(self):
        row = self.table_numeric_columns.currentRow()
        if row >= 0:
            self.table_numeric_columns.removeRow(row)

    # 费率计算模式变更
    def _on_calculation_mode_changed(self, mode):
        self.spin_fixed_rate.setEnabled(mode == "fixed")

    # 特殊列操作
    def _add_special_column(self):
        row = self.table_special_columns.rowCount()
        self.table_special_columns.insertRow(row)
        self.table_special_columns.setItem(row, 0, QTableWidgetItem("列名"))
        self.table_special_columns.setItem(row, 1, QTableWidgetItem("15"))

        check_wrap = QCheckBox()
        self.table_special_columns.setCellWidget(row, 2, check_wrap)

        self.table_special_columns.setItem(row, 3, QTableWidgetItem("center"))
        self.table_special_columns.setItem(row, 4, QTableWidgetItem(""))
        self.table_special_columns.setItem(row, 5, QTableWidgetItem(""))
        self.table_special_columns.setItem(row, 6, QTableWidgetItem(""))

    def _del_special_column(self):
        row = self.table_special_columns.currentRow()
        if row >= 0:
            self.table_special_columns.removeRow(row)

    def _save_config(self):
        """保存配置"""
        if not self.current_customer_type:
            return

        # 从UI收集配置
        config = self._collect_config_from_ui()

        # 更新配置
        if self.config_manager.update_customer_type(self.current_customer_type, config):
            QMessageBox.information(self, "成功", "配置已保存")
            self.config_changed.emit()
            # 刷新列表显示
            self._load_data()
            # 重新选中当前项
            for i in range(self.list_customer_types.count()):
                if self.list_customer_types.item(i).text() == self.current_customer_type:
                    self.list_customer_types.setCurrentRow(i)
                    break
        else:
            QMessageBox.critical(self, "错误", "配置保存失败")

    def _collect_config_from_ui(self) -> dict:
        """从UI收集配置数据"""
        # 收集表头配置
        columns = []
        for row in range(self.table_headers.rowCount()):
            key_item = self.table_headers.item(row, 0)
            label_item = self.table_headers.item(row, 1)
            width_item = self.table_headers.item(row, 2)

            if not key_item or not label_item:
                continue

            col = {
                "key": key_item.text(),
                "label": label_item.text(),
                "width": int(width_item.text() or 15) if width_item else 15,
            }

            fmt_item = self.table_headers.item(row, 3)
            if fmt_item and fmt_item.text():
                col["format"] = fmt_item.text()

            check_sum = self.table_headers.cellWidget(row, 4)
            if check_sum and isinstance(check_sum, QCheckBox) and check_sum.isChecked():
                col["sum"] = True

            check_wrap = self.table_headers.cellWidget(row, 5)
            if check_wrap and isinstance(check_wrap, QCheckBox) and check_wrap.isChecked():
                col["wrap_text"] = True

            fixed_item = self.table_headers.item(row, 6)
            if fixed_item and fixed_item.text():
                col["fixed_value"] = fixed_item.text()

            columns.append(col)

        # 收集占位符配置
        mappings = {}
        common_placeholders = []
        for row in range(self.table_placeholders.rowCount()):
            ph_item = self.table_placeholders.item(row, 0)
            field_item = self.table_placeholders.item(row, 1)
            if not ph_item:
                continue

            placeholder = ph_item.text()
            field = field_item.text() if field_item else ""
            fmt_item = self.table_placeholders.item(row, 2)
            transform_item = self.table_placeholders.item(row, 3)
            fmt = fmt_item.text() if fmt_item else ""
            transform = transform_item.text() if transform_item else ""

            common_placeholders.append(placeholder)

            if fmt or transform:
                mapping = {"field": field}
                if fmt:
                    mapping["format"] = fmt
                if transform:
                    mapping["transform"] = transform
                mappings[placeholder] = mapping
            else:
                mappings[placeholder] = field

        # 收集协议编号
        agreements = {}
        for row in range(self.table_agreements.rowCount()):
            company_item = self.table_agreements.item(row, 0)
            code_item = self.table_agreements.item(row, 1)
            if company_item and code_item:
                company = company_item.text()
                code = code_item.text()
                if company and code:
                    agreements[company] = code

        # 收集PDF分组规则
        pdf_groups = {}
        for row in range(self.table_pdf_groups.rowCount()):
            name_item = self.table_pdf_groups.item(row, 0)
            type_item = self.table_pdf_groups.item(row, 1)
            pattern_item = self.table_pdf_groups.item(row, 2)
            if name_item:
                name = name_item.text()
                match_type = type_item.text() if type_item else "contains"
                patterns_str = pattern_item.text() if pattern_item else ""
                patterns = [p.strip() for p in patterns_str.split(",") if p.strip()]
                if name and patterns:
                    pdf_groups[name] = {
                        "match_type": match_type,
                        "patterns": patterns
                    }

        # 收集排版列宽
        processed_widths = []
        for row in range(self.table_processed_widths.rowCount()):
            col_item = self.table_processed_widths.item(row, 0)
            width_item = self.table_processed_widths.item(row, 1)
            if col_item and width_item:
                col = col_item.text().strip().upper()
                try:
                    width = float(width_item.text())
                except ValueError:
                    width = 10.0
                if col:
                    processed_widths.append({"column": col, "width": width})

        # 收集对账单分组规则
        statement_groups = {}
        for row in range(self.table_statement_groups.rowCount()):
            name_item = self.table_statement_groups.item(row, 0)
            type_item = self.table_statement_groups.item(row, 1)
            pattern_item = self.table_statement_groups.item(row, 2)
            if name_item:
                name = name_item.text()
                match_type = type_item.text() if type_item else "contains"
                patterns_str = pattern_item.text() if pattern_item else ""
                patterns = [p.strip() for p in patterns_str.split(",")]
                if name:
                    statement_groups[name] = {
                        "match_type": match_type,
                        "patterns": patterns
                    }

        # 收集表头识别规则
        header_keywords_str = self.edit_header_keywords.text().strip()
        header_keywords = [k.strip() for k in header_keywords_str.split(",") if k.strip()]
        total_keywords_str = self.edit_total_keywords.text().strip()
        total_keywords = [k.strip() for k in total_keywords_str.split(",") if k.strip()]

        column_mappings = {}
        for row in range(self.table_column_mappings.rowCount()):
            field_item = self.table_column_mappings.item(row, 0)
            aliases_item = self.table_column_mappings.item(row, 1)
            if field_item and aliases_item:
                field = field_item.text().strip()
                aliases_str = aliases_item.text()
                aliases = [a.strip() for a in aliases_str.split(",") if a.strip()]
                if field and aliases:
                    column_mappings[field] = aliases

        header_recognition = {
            "header_keywords": header_keywords or ["序号"],
            "total_keywords": total_keywords or ["合计"],
            "max_search_rows": self.spin_max_search_rows.value(),
            "column_mappings": column_mappings
        }

        # 收集数据提取规则
        date_columns_str = self.edit_date_columns.text().strip()
        date_columns = [c.strip() for c in date_columns_str.split(",") if c.strip()]

        numeric_columns = {}
        for row in range(self.table_numeric_columns.rowCount()):
            col_item = self.table_numeric_columns.item(row, 0)
            fmt_item = self.table_numeric_columns.item(row, 1)
            if col_item and fmt_item:
                col_name = col_item.text().strip()
                fmt = fmt_item.text()
                if col_name:
                    numeric_columns[col_name] = fmt

        data_extraction = {
            "business_count_method": self.combo_business_count_method.currentText(),
            "sequence_column": self.spin_sequence_column.value(),
            "date_columns": date_columns or ["起运日期", "发货日期"],
            "date_format": self.edit_date_format.text() or "%Y/%m/%d",
            "numeric_columns": numeric_columns,
            "skip_empty_rows": self.check_skip_empty_rows.isChecked(),
            "data_start_offset": self.spin_data_start_offset.value()
        }

        # 收集费率计算配置
        rate_calculation = {
            "calculation_mode": self.combo_calculation_mode.currentText(),
            "fixed_rate": self.spin_fixed_rate.value() if self.spin_fixed_rate.value() > 0 else None,
            "rate_precision": self.spin_rate_precision.value(),
            "premium_precision": self.spin_premium_precision.value(),
            "rate_column": self.edit_rate_column.text() or None,
            "premium_column": self.edit_premium_column.text() or None,
            "insurance_amount_column": self.edit_insurance_amount_column.text() or None,
            "formula": self.edit_formula.text() or "premium = insurance_amount * rate",
            "display_unit": self.combo_display_unit.currentText(),
            "display_multiplier": self.spin_display_multiplier.value()
        }

        # 收集特殊列配置
        special_columns = []
        for row in range(self.table_special_columns.rowCount()):
            name_item = self.table_special_columns.item(row, 0)
            if not name_item:
                continue
            width_item = self.table_special_columns.item(row, 1)
            wrap_widget = self.table_special_columns.cellWidget(row, 2)
            align_item = self.table_special_columns.item(row, 3)
            font_size_item = self.table_special_columns.item(row, 4)
            first_width_item = self.table_special_columns.item(row, 5)
            other_width_item = self.table_special_columns.item(row, 6)

            sc = {
                "column_name": name_item.text(),
                "width": float(width_item.text()) if width_item and width_item.text() else 15,
                "wrap_text": wrap_widget.isChecked() if wrap_widget and isinstance(wrap_widget, QCheckBox) else False,
                "alignment": align_item.text() if align_item else "center"
            }

            if font_size_item and font_size_item.text():
                try:
                    sc["font_size"] = int(font_size_item.text())
                except ValueError:
                    pass

            if first_width_item and first_width_item.text():
                try:
                    sc["first_sheet_width"] = float(first_width_item.text())
                except ValueError:
                    pass

            if other_width_item and other_width_item.text():
                try:
                    sc["other_sheet_width"] = float(other_width_item.text())
                except ValueError:
                    pass

            special_columns.append(sc)

        # 收集行高配置
        row_heights = {
            "title_row": self.spin_title_row_height.value(),
            "second_row": self.spin_second_row_height.value(),
            "header_row": self.spin_header_row_height.value(),
            "data_row": self.spin_data_row_height.value(),
            "total_row": self.spin_total_row_height.value(),
            "footer_row": self.spin_footer_row_height.value(),
            "auto_fit": self.check_auto_fit_height.isChecked()
        }

        return {
            "id": self.edit_id.text(),
            "display_name": self.edit_display_name.text(),
            "description": self.edit_description.text(),
            "enabled": self.check_enabled.isChecked(),
            "extract_function": self.combo_extract_func.currentText(),
            "process_function": self.combo_process_func.currentText(),
            "use_excel_formatter": self.check_use_formatter.isChecked(),
            "is_hengli": self.check_is_hengli.isChecked(),
            "page_orientation": self.combo_page_orientation.currentText(),
            "log_format": self.edit_log_format.text(),
            "summary_headers": {
                "columns": columns,
                "sum_row_label_column": self.spin_sum_label_col.value(),
                "rmb_total_column": self.spin_rmb_col.value() if self.spin_rmb_col.value() > 0 else None
            },
            "placeholders": {
                "common": common_placeholders,
                "mappings": mappings
            },
            "agreement_codes": agreements,
            "pdf_export_groups": pdf_groups,
            "processed_column_widths": processed_widths,
            "statement_groups": statement_groups,
            "header_recognition": header_recognition,
            "data_extraction": data_extraction,
            "rate_calculation": rate_calculation,
            "special_columns": special_columns,
            "row_heights": row_heights
        }

    def _add_customer_type(self):
        """添加新客户类型"""
        dialog = AddCustomerTypeDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name = dialog.get_name()
            if not name:
                QMessageBox.warning(self, "警告", "客户类型名称不能为空")
                return

            if name in self.config_manager.get_all_customer_types():
                QMessageBox.warning(self, "警告", f"客户类型 '{name}' 已存在")
                return

            default_config = {
                "id": name.lower().replace(" ", "_").replace("-", "_"),
                "display_name": name,
                "description": "",
                "enabled": True,
                "extract_function": "extract_sheet_data",
                "process_function": "process_multimodal_sheet",
                "use_excel_formatter": True,
                "is_hengli": False,
                "summary_headers": {
                    "columns": [
                        {"key": "file_name", "label": "文件名", "width": 25},
                        {"key": "sheet_name", "label": "Sheet名", "width": 30}
                    ],
                    "sum_row_label_column": 1
                },
                "placeholders": {"common": [], "mappings": {}},
                "agreement_codes": {},
                "pdf_export_groups": {}
            }

            if self.config_manager.add_customer_type(name, default_config):
                self._load_data()
                # 选中新添加的项
                for i in range(self.list_customer_types.count()):
                    if self.list_customer_types.item(i).text() == name:
                        self.list_customer_types.setCurrentRow(i)
                        break
                self.config_changed.emit()
            else:
                QMessageBox.critical(self, "错误", "添加客户类型失败")

    def _delete_customer_type(self):
        """删除客户类型"""
        if not self.current_customer_type:
            return

        # 不允许删除最后一个
        if self.list_customer_types.count() <= 1:
            QMessageBox.warning(self, "警告", "至少保留一个客户类型")
            return

        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除客户类型 '{self.current_customer_type}' 吗？\n此操作不可撤销。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.config_manager.delete_customer_type(self.current_customer_type):
                self._load_data()
                self.config_changed.emit()
            else:
                QMessageBox.critical(self, "错误", "删除客户类型失败")

    def _import_config(self):
        """导入配置"""
        path, _ = QFileDialog.getOpenFileName(
            self, "导入配置",
            "", "JSON文件 (*.json)"
        )
        if path:
            success, errors = self.config_manager.import_config(path)
            if success:
                self._load_data()
                QMessageBox.information(self, "成功", "配置已导入")
                self.config_changed.emit()
            else:
                QMessageBox.critical(self, "导入失败", "\n".join(errors))

    def _export_config(self):
        """导出配置"""
        path, _ = QFileDialog.getSaveFileName(
            self, "导出配置",
            "customer_config.json", "JSON文件 (*.json)"
        )
        if path:
            if self.config_manager.export_config(path):
                QMessageBox.information(self, "成功", f"配置已导出到:\n{path}")
            else:
                QMessageBox.critical(self, "错误", "导出配置失败")

    def _reset_config(self):
        """重置为默认配置"""
        reply = QMessageBox.question(
            self, "确认重置",
            "确定要重置为默认配置吗？\n所有自定义配置将丢失，此操作不可撤销。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # 删除配置文件，重新加载会生成默认配置
            import os
            try:
                if os.path.exists(self.config_manager.config_path):
                    os.remove(self.config_manager.config_path)
                self.config_manager.load()
                self._load_data()
                QMessageBox.information(self, "成功", "已重置为默认配置")
                self.config_changed.emit()
            except Exception as e:
                QMessageBox.critical(self, "错误", f"重置失败: {e}")


class AddCustomerTypeDialog(QDialog):
    """添加客户类型对话框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("添加客户类型")
        self.setMinimumWidth(300)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        form = QFormLayout()
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("输入客户类型名称")
        form.addRow("客户类型名称:", self.edit_name)
        layout.addLayout(form)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_name(self) -> str:
        return self.edit_name.text().strip()
