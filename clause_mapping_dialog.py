# -*- coding: utf-8 -*-
"""
用户自定义条款映射管理对话框

功能：
- 查看/添加/编辑/删除用户自定义映射
- 从条款库选择条款名称（带搜索）
- 批量导入/导出

Author: Claude
Date: 2025-01-04
"""

import logging
from typing import List, Optional
from pathlib import Path

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QFileDialog, QCompleter, QAbstractItemView,
    QGroupBox, QFormLayout, QWidget
)
from PyQt5.QtCore import Qt, pyqtSignal, QStringListModel
from PyQt5.QtGui import QFont, QColor

from clause_mapping_manager import get_mapping_manager, ClauseMapping

logger = logging.getLogger(__name__)


class ClauseMappingDialog(QDialog):
    """条款映射管理对话框"""

    mappings_changed = pyqtSignal()  # 映射变更信号

    def __init__(self, parent=None, library_clauses: List[str] = None):
        super().__init__(parent)
        self.library_clauses = library_clauses or []
        self.manager = get_mapping_manager()
        self.manager.load()

        self._setup_ui()
        self._load_mappings()

    def _setup_ui(self):
        self.setWindowTitle("条款映射管理 - 人工匹配设置")
        self.setMinimumSize(900, 600)
        # Anthropic UI 色彩系统
        BG_PRIMARY = "#faf9f5"      # 主背景/奶油白
        BG_CARD = "#f0eee6"         # 卡片背景/浅米色
        ACCENT = "#d97757"          # 主强调色/陶土色
        TEXT_PRIMARY = "#141413"    # 主要文字
        TEXT_SECONDARY = "#b0aea5"  # 次要文字
        TEXT_LIGHT = "#faf9f5"      # 深色背景上的文字
        BORDER = "#e5e3db"          # 浅边框
        BG_DARK = "#141413"         # 深色区域

        self.setStyleSheet(f"""
            QDialog {{
                background: {BG_PRIMARY};
            }}
            QLabel {{
                color: {TEXT_PRIMARY};
                font-size: 13px;
            }}
            QLineEdit {{
                background: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 8px 12px;
                color: {TEXT_PRIMARY};
                font-size: 13px;
            }}
            QLineEdit:focus {{
                border-color: {ACCENT};
            }}
            QTableWidget {{
                background: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 8px;
                color: {TEXT_PRIMARY};
                gridline-color: {BORDER};
            }}
            QTableWidget::item {{
                padding: 8px;
            }}
            QTableWidget::item:selected {{
                background: rgba(217, 119, 87, 0.2);
            }}
            QHeaderView::section {{
                background: {BG_DARK};
                color: {TEXT_LIGHT};
                padding: 10px;
                border: none;
                font-weight: bold;
            }}
            QPushButton {{
                background: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 8px 16px;
                color: {TEXT_PRIMARY};
                font-weight: 500;
            }}
            QPushButton:hover {{
                background: {BG_DARK};
                color: {TEXT_LIGHT};
                border-color: {BG_DARK};
            }}
            QPushButton:disabled {{
                color: {TEXT_SECONDARY};
            }}
            QGroupBox {{
                color: {TEXT_PRIMARY};
                border: 1px solid {BORDER};
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 8px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 标题
        title = QLabel("人工匹配设置")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #d97757;")
        layout.addWidget(title)

        desc = QLabel("当自动匹配不准确时，您可以在此设置正确的映射关系。设置后，下次比对将优先使用您的映射。")
        desc.setStyleSheet("color: #b0aea5; font-size: 12px;")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # 添加映射区域
        add_group = QGroupBox("添加新映射")
        add_layout = QVBoxLayout(add_group)

        form_layout = QHBoxLayout()

        # 客户条款输入
        client_layout = QVBoxLayout()
        client_layout.addWidget(QLabel("客户条款名称:"))
        self.client_input = QLineEdit()
        self.client_input.setPlaceholderText("输入客户文档中的条款名称，如：72小时条款")
        client_layout.addWidget(self.client_input)
        form_layout.addLayout(client_layout, 1)

        # 箭头
        arrow_label = QLabel("  →  ")
        arrow_label.setStyleSheet("font-size: 20px; color: #d97757;")
        arrow_label.setAlignment(Qt.AlignCenter)
        form_layout.addWidget(arrow_label)

        # 条款库名称输入（带自动完成）
        library_layout = QVBoxLayout()
        library_layout.addWidget(QLabel("条款库名称:"))
        self.library_input = QLineEdit()
        self.library_input.setPlaceholderText("输入或选择条款库中的标准名称")

        # 设置自动完成
        if self.library_clauses:
            completer = QCompleter(self.library_clauses)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            completer.setFilterMode(Qt.MatchContains)
            completer.setMaxVisibleItems(10)
            self.library_input.setCompleter(completer)

        library_layout.addWidget(self.library_input)
        form_layout.addLayout(library_layout, 2)

        add_layout.addLayout(form_layout)

        # 添加按钮
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.add_btn = QPushButton("添加映射")
        self.add_btn.setStyleSheet("""
            QPushButton {
                background: #141413;
                color: #faf9f5;
                border: none;
                padding: 10px 25px;
            }
            QPushButton:hover {
                background: #d97757;
            }
        """)
        self.add_btn.clicked.connect(self._add_mapping)
        btn_row.addWidget(self.add_btn)

        add_layout.addLayout(btn_row)
        layout.addWidget(add_group)

        # 现有映射列表
        list_group = QGroupBox(f"已保存的映射 ({self.manager.get_mapping_count()} 条)")
        self.list_group = list_group
        list_layout = QVBoxLayout(list_group)

        # 搜索框
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("搜索:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入关键词筛选...")
        self.search_input.textChanged.connect(self._filter_mappings)
        search_layout.addWidget(self.search_input)
        list_layout.addLayout(search_layout)

        # 表格 (v18.2: 删除创建时间列，简化界面)
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["客户条款", "条款库名称", "操作"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.table.setColumnWidth(2, 180)  # 操作列固定宽度
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(55)  # v18.2: 增加行高以完整显示按钮
        list_layout.addWidget(self.table)

        # 批量操作按钮
        batch_layout = QHBoxLayout()

        self.import_btn = QPushButton("📥 从报告导入")
        self.import_btn.setToolTip("从已修正的比对报告中导入映射（自动去除编号、限额等信息）")
        self.import_btn.setStyleSheet("""
            QPushButton {
                background: rgba(90, 154, 122, 0.15);
                border: 1px solid #5a9a7a;
                color: #5a9a7a;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background: #5a9a7a;
                color: #faf9f5;
            }
        """)
        self.import_btn.clicked.connect(self._import_from_corrected_report)
        batch_layout.addWidget(self.import_btn)

        self.export_btn = QPushButton("导出映射")
        self.export_btn.clicked.connect(self._export_mappings)
        batch_layout.addWidget(self.export_btn)

        batch_layout.addStretch()

        self.delete_all_btn = QPushButton("清空所有")
        self.delete_all_btn.setStyleSheet("color: #c75050; border-color: #c75050;")
        self.delete_all_btn.clicked.connect(self._delete_all)
        batch_layout.addWidget(self.delete_all_btn)

        list_layout.addLayout(batch_layout)
        layout.addWidget(list_group, 1)

        # 底部按钮
        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        bottom_layout.addWidget(close_btn)

        layout.addLayout(bottom_layout)

    def _load_mappings(self):
        """加载映射到表格"""
        mappings = self.manager.get_all_mappings()
        self.table.setRowCount(len(mappings))

        for i, mapping in enumerate(mappings):
            self.table.setItem(i, 0, QTableWidgetItem(mapping.client_name))
            self.table.setItem(i, 1, QTableWidgetItem(mapping.library_name))

            # 操作按钮容器 (v18.2: 扩大按钮尺寸)
            btn_container = QWidget()
            btn_layout = QHBoxLayout(btn_container)
            btn_layout.setContentsMargins(4, 4, 4, 4)
            btn_layout.setSpacing(8)

            # 修改按钮 (v18.2: 扩大按钮)
            edit_btn = QPushButton("修改")
            edit_btn.setMinimumWidth(70)
            edit_btn.setMinimumHeight(28)
            edit_btn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    color: #5a9a7a;
                    border: 1px solid #5a9a7a;
                    border-radius: 4px;
                    padding: 6px 16px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background: #5a9a7a;
                    color: #faf9f5;
                }
            """)
            edit_btn.clicked.connect(lambda checked, name=mapping.client_name, lib=mapping.library_name: self._edit_mapping(name, lib))
            btn_layout.addWidget(edit_btn)

            # 删除按钮 (v18.2: 扩大按钮)
            delete_btn = QPushButton("删除")
            delete_btn.setMinimumWidth(70)
            delete_btn.setMinimumHeight(28)
            delete_btn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    color: #c75050;
                    border: 1px solid #c75050;
                    border-radius: 4px;
                    padding: 6px 16px;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background: #c75050;
                    color: #faf9f5;
                }
            """)
            delete_btn.clicked.connect(lambda checked, name=mapping.client_name: self._delete_mapping(name))
            btn_layout.addWidget(delete_btn)

            self.table.setCellWidget(i, 2, btn_container)

        self.list_group.setTitle(f"已保存的映射 ({len(mappings)} 条)")

    def _filter_mappings(self, text: str):
        """筛选映射"""
        text = text.lower()
        for i in range(self.table.rowCount()):
            client = self.table.item(i, 0).text().lower()
            library = self.table.item(i, 1).text().lower()
            match = text in client or text in library
            self.table.setRowHidden(i, not match)

    def _add_mapping(self):
        """添加映射"""
        client = self.client_input.text().strip()
        library = self.library_input.text().strip()

        if not client:
            QMessageBox.warning(self, "提示", "请输入客户条款名称")
            return
        if not library:
            QMessageBox.warning(self, "提示", "请输入条款库名称")
            return

        self.manager.add_mapping(client, library)
        self.manager.save()

        self.client_input.clear()
        self.library_input.clear()
        self._load_mappings()
        self.mappings_changed.emit()

        QMessageBox.information(self, "成功", f"已添加映射:\n{client}\n→\n{library}")

    def _edit_mapping(self, client_name: str, current_library: str):
        """修改映射"""
        # 弹出输入对话框
        from PyQt5.QtWidgets import QInputDialog
        new_library, ok = QInputDialog.getText(
            self, "修改映射",
            f"客户条款: {client_name}\n\n请输入新的条款库名称:",
            text=current_library
        )
        if ok and new_library.strip():
            new_library = new_library.strip()
            if new_library != current_library:
                # 先删除旧映射，再添加新映射
                self.manager.remove_mapping(client_name)
                self.manager.add_mapping(client_name, new_library)
                self.manager.save()
                self._load_mappings()
                self.mappings_changed.emit()
                QMessageBox.information(self, "成功", f"已修改映射:\n{client_name}\n→\n{new_library}")

    def _delete_mapping(self, client_name: str):
        """删除映射"""
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除映射 '{client_name}' 吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.manager.remove_mapping(client_name)
            self.manager.save()
            self._load_mappings()
            self.mappings_changed.emit()

    def _delete_all(self):
        """清空所有映射"""
        if self.manager.get_mapping_count() == 0:
            QMessageBox.information(self, "提示", "没有映射可以删除")
            return

        reply = QMessageBox.question(
            self, "确认清空",
            f"确定要删除所有 {self.manager.get_mapping_count()} 条映射吗？\n此操作不可恢复！",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            for mapping in self.manager.get_all_mappings():
                self.manager.remove_mapping(mapping.client_name)
            self.manager.save()
            self._load_mappings()
            self.mappings_changed.emit()

    def _import_from_corrected_report(self):
        """从已修正的比对报告导入映射"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择已修正的比对报告", "",
            "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        # 确认导入
        reply = QMessageBox.question(
            self, "确认导入",
            "将从报告中导入所有条款映射：\n\n"
            "• 读取「客户条款(原)」和「匹配条款库名称」列\n"
            "• 自动去除编号（如1.、2.）和限额信息\n"
            "• 建立客户条款与库内条款的映射关系\n\n"
            "是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        if reply != QMessageBox.Yes:
            return

        try:
            imported, skipped = self.manager.import_from_corrected_report(
                file_path,
                clean_names=True  # 自动清理条款名称
            )

            if imported > 0:
                self.manager.save()
                self._load_mappings()
                self.mappings_changed.emit()
                QMessageBox.information(
                    self, "导入成功",
                    f"成功导入 {imported} 条映射\n跳过 {skipped} 条（空值或无效）"
                )
            else:
                QMessageBox.information(
                    self, "提示",
                    f"没有可以导入的映射\n（检查了 {skipped} 条记录）"
                )

        except Exception as e:
            logger.exception("导入报告失败")
            QMessageBox.critical(self, "错误", f"导入失败: {e}")

    def _export_mappings(self):
        """导出映射"""
        if self.manager.get_mapping_count() == 0:
            QMessageBox.information(self, "提示", "没有映射可以导出")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出映射", "clause_mappings.json",
            "JSON Files (*.json)"
        )
        if file_path:
            import json
            data = self.manager.export_to_dict()
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            QMessageBox.information(self, "成功", f"已导出 {len(data)} 条映射")


class ImportPreviewDialog(QDialog):
    """导入预览对话框 - 允许用户修正映射"""

    def __init__(self, parent, items: List[tuple], library_clauses: List[str]):
        super().__init__(parent)
        self.items = items
        self.library_clauses = library_clauses
        self.selected_mappings = {}

        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("导入预览 - 修正低分匹配")
        self.setMinimumSize(1000, 500)
        # Anthropic UI 色彩系统
        BG_PRIMARY = "#faf9f5"
        BG_CARD = "#f0eee6"
        ACCENT = "#d97757"
        TEXT_PRIMARY = "#141413"
        TEXT_LIGHT = "#faf9f5"
        BORDER = "#e5e3db"
        BG_DARK = "#141413"

        self.setStyleSheet(f"""
            QDialog {{ background: {BG_PRIMARY}; }}
            QLabel {{ color: {TEXT_PRIMARY}; }}
            QLineEdit {{
                background: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 4px;
                padding: 6px;
                color: {TEXT_PRIMARY};
            }}
            QLineEdit:focus {{
                border-color: {ACCENT};
            }}
            QTableWidget {{
                background: {BG_CARD};
                border: 1px solid {BORDER};
                color: {TEXT_PRIMARY};
            }}
            QHeaderView::section {{
                background: {BG_DARK};
                color: {TEXT_LIGHT};
                padding: 8px;
                border: none;
            }}
            QPushButton {{
                background: {BG_CARD};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 8px 16px;
                color: {TEXT_PRIMARY};
            }}
            QPushButton:hover {{
                background: {BG_DARK};
                color: {TEXT_LIGHT};
            }}
            QCheckBox {{
                color: {TEXT_PRIMARY};
            }}
            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 说明
        desc = QLabel(f"发现 {len(self.items)} 个低分匹配项，请修正后导入：")
        desc.setStyleSheet("font-size: 14px; font-weight: bold; color: #d97757;")
        layout.addWidget(desc)

        hint = QLabel("提示：在「正确的条款库名称」列中输入正确的名称，支持自动补全。勾选要导入的项目。")
        hint.setStyleSheet("color: #b0aea5; font-size: 12px;")
        layout.addWidget(hint)

        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["导入", "客户条款", "当前匹配（错误）", "匹配度", "正确的条款库名称"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setRowCount(len(self.items))

        self.checkboxes = []
        self.inputs = []

        for i, (client, matched, score) in enumerate(self.items):
            # 复选框
            from PyQt5.QtWidgets import QCheckBox
            cb = QCheckBox()
            cb.setChecked(True)
            self.checkboxes.append(cb)
            cb_widget = QWidget()
            cb_layout = QHBoxLayout(cb_widget)
            cb_layout.addWidget(cb)
            cb_layout.setAlignment(Qt.AlignCenter)
            cb_layout.setContentsMargins(0, 0, 0, 0)
            self.table.setCellWidget(i, 0, cb_widget)

            # 客户条款
            self.table.setItem(i, 1, QTableWidgetItem(client))

            # 当前匹配
            item = QTableWidgetItem(matched)
            item.setForeground(QColor("#c75050"))  # Anthropic ERROR color
            self.table.setItem(i, 2, item)

            # 匹配度
            score_item = QTableWidgetItem(f"{score:.2f}")
            score_item.setForeground(QColor("#d9a557"))  # Anthropic WARNING color
            self.table.setItem(i, 3, score_item)

            # 正确名称输入（带自动完成）
            input_widget = QLineEdit()
            input_widget.setPlaceholderText("输入正确的条款库名称...")
            if self.library_clauses:
                completer = QCompleter(self.library_clauses)
                completer.setCaseSensitivity(Qt.CaseInsensitive)
                completer.setFilterMode(Qt.MatchContains)
                input_widget.setCompleter(completer)
            self.inputs.append(input_widget)
            self.table.setCellWidget(i, 4, input_widget)

        layout.addWidget(self.table, 1)

        # 按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        import_btn = QPushButton("导入选中项")
        import_btn.setStyleSheet("""
            QPushButton {
                background: #141413;
                color: #faf9f5;
                border: none;
            }
            QPushButton:hover {
                background: #d97757;
            }
        """)
        import_btn.clicked.connect(self._do_import)
        btn_layout.addWidget(import_btn)

        layout.addLayout(btn_layout)

    def _do_import(self):
        """执行导入"""
        self.selected_mappings = {}

        for i, (client, _, _) in enumerate(self.items):
            if self.checkboxes[i].isChecked():
                correct_name = self.inputs[i].text().strip()
                if correct_name:
                    self.selected_mappings[client] = correct_name

        if not self.selected_mappings:
            QMessageBox.warning(self, "提示", "请至少填写一个正确的条款库名称")
            return

        self.accept()

    def get_selected_mappings(self) -> dict:
        return self.selected_mappings
