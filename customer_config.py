# -*- coding: utf-8 -*-
"""
客户类型配置管理模块
提供客户类型的加载、保存、验证和管理功能
"""

import json
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Callable
from pathlib import Path
from datetime import datetime


@dataclass
class ColumnConfig:
    """汇总表列配置"""
    key: str
    label: str
    width: int = 15
    format: Optional[str] = None
    sum: bool = False
    wrap_text: bool = False
    fixed_value: Optional[str] = None
    transform: Optional[str] = None


@dataclass
class SummaryHeaderConfig:
    """汇总表表头配置"""
    columns: List[ColumnConfig] = field(default_factory=list)
    sum_row_label_column: int = 1
    rmb_total_column: Optional[int] = None


@dataclass
class ProcessedColumnWidth:
    """排版文件列宽配置"""
    column: str  # A, B, C, D, ...
    width: float = 10.0


@dataclass
class VisualSettings:
    """可视化设置"""
    font: Dict[str, Any] = field(default_factory=dict)
    row_height: Dict[str, int] = field(default_factory=dict)
    header_style: Dict[str, Any] = field(default_factory=dict)
    data_style: Dict[str, Any] = field(default_factory=dict)
    border: Dict[str, Any] = field(default_factory=dict)
    print_settings: Dict[str, Any] = field(default_factory=dict)
    number_formats: Dict[str, str] = field(default_factory=dict)


@dataclass
class SourceFieldConfig:
    """来源字段配置"""
    enabled: bool = False
    column_keywords: List[str] = field(default_factory=list)
    row: str = "first_data"
    label: str = "来源"


@dataclass
class HeaderRecognitionConfig:
    """表头识别规则配置"""
    header_keywords: List[str] = field(default_factory=lambda: ["序号"])  # 表头行识别关键词
    total_keywords: List[str] = field(default_factory=lambda: ["合计"])  # 合计行识别关键词
    max_search_rows: int = 10  # 最大搜索行数
    column_mappings: Dict[str, List[str]] = field(default_factory=dict)  # 列名映射，如 {"cargo_type": ["货种", "保险货物"]}


@dataclass
class DataExtractionConfig:
    """数据提取规则配置"""
    business_count_method: str = "sequence"  # 业务笔数计算方式: sequence(序号), count(行数)
    sequence_column: int = 1  # 序号所在列
    date_columns: List[str] = field(default_factory=lambda: ["起运日期", "发货日期"])  # 日期列关键词
    date_format: str = "%Y/%m/%d"  # 日期格式
    numeric_columns: Dict[str, str] = field(default_factory=dict)  # 数值列格式，如 {"tonnage": "#,##0.000"}
    skip_empty_rows: bool = True  # 是否跳过空行
    data_start_offset: int = 1  # 数据起始行相对于表头的偏移


@dataclass
class RateCalculationConfig:
    """费率计算配置"""
    calculation_mode: str = "auto"  # 计算模式: auto(自动计算), fixed(固定费率), manual(手动)
    fixed_rate: Optional[float] = None  # 固定费率值
    rate_precision: int = 8  # 费率精度（小数位数）
    premium_precision: int = 2  # 保费精度（小数位数）
    rate_column: Optional[str] = None  # 费率列名
    premium_column: Optional[str] = None  # 保费列名
    insurance_amount_column: Optional[str] = None  # 保险金额列名
    formula: str = "premium = insurance_amount * rate"  # 计算公式
    display_unit: str = "%"  # 显示单位: %, ‰
    display_multiplier: float = 100  # 显示乘数


@dataclass
class SpecialColumnConfig:
    """特殊列配置"""
    column_name: str  # 列名
    width: float = 15  # 列宽
    wrap_text: bool = False  # 是否自动换行
    alignment: str = "center"  # 对齐方式: left, center, right
    font_size: Optional[int] = None  # 字号（None表示使用默认）
    first_sheet_width: Optional[float] = None  # 第一个sheet的特殊宽度
    other_sheet_width: Optional[float] = None  # 其他sheet的宽度


@dataclass
class RowHeightConfig:
    """行高配置"""
    title_row: int = 39  # 标题行高
    second_row: int = 33  # 第二行高
    header_row: int = 32  # 表头行高
    data_row: int = 15  # 数据行高
    total_row: int = 18  # 合计行高
    footer_row: int = 18  # 页脚行高
    auto_fit: bool = True  # 是否自动调整行高


@dataclass
class CustomerTypeConfig:
    """客户类型配置"""
    id: str
    display_name: str
    description: str = ""
    enabled: bool = True
    extract_function: str = "extract_sheet_data"
    process_function: str = "process_multimodal_sheet"
    use_excel_formatter: bool = True
    is_hengli: bool = False
    page_orientation: str = "landscape"  # 打印方向: landscape(横向) 或 portrait(纵向)
    summary_headers: Optional[SummaryHeaderConfig] = None
    placeholders: Dict[str, Any] = field(default_factory=dict)
    rate_config: Dict[str, Any] = field(default_factory=dict)
    log_format: str = ""
    agreement_codes: Dict[str, str] = field(default_factory=dict)
    pdf_export_groups: Dict[str, Any] = field(default_factory=dict)
    processed_column_widths: List[ProcessedColumnWidth] = field(default_factory=list)  # 排版文件列宽
    statement_groups: Dict[str, Any] = field(default_factory=dict)  # 对账单分组规则
    source_field: Optional[SourceFieldConfig] = None  # 来源字段配置
    visual_settings: Optional[VisualSettings] = None  # 可视化设置
    # 新增配置项
    header_recognition: Optional[HeaderRecognitionConfig] = None  # 表头识别规则
    data_extraction: Optional[DataExtractionConfig] = None  # 数据提取规则
    rate_calculation: Optional[RateCalculationConfig] = None  # 费率计算配置
    special_columns: List[SpecialColumnConfig] = field(default_factory=list)  # 特殊列配置
    row_heights: Optional[RowHeightConfig] = None  # 行高配置


class CustomerConfigManager:
    """客户类型配置管理器"""

    CONFIG_FILE = "customer_config.json"
    DEFAULT_CONFIG_VERSION = "1.0"

    def __init__(self, config_path: Optional[str] = None):
        self.config_path = config_path or self._get_default_config_path()
        self._config: Dict[str, Any] = {}
        self._customer_types: Dict[str, CustomerTypeConfig] = {}
        self._callbacks: List[Callable] = []
        self.load()

    def _get_default_config_path(self) -> str:
        """获取默认配置文件路径（与脚本同目录）"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(script_dir, self.CONFIG_FILE)

    def load(self) -> bool:
        """加载配置文件"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self._config = json.load(f)
                self._parse_customer_types()
                return True
            else:
                # 使用默认配置
                self._config = self._get_default_config()
                self.save()
                self._parse_customer_types()
                return True
        except Exception as e:
            print(f"加载配置失败: {e}")
            self._config = self._get_default_config()
            self._parse_customer_types()
            return False

    def save(self) -> bool:
        """保存配置文件"""
        try:
            # 更新元数据
            self._config.setdefault("_meta", {})
            self._config["_meta"]["last_modified"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)

            # 通知观察者
            for callback in self._callbacks:
                try:
                    callback()
                except Exception:
                    pass
            return True
        except Exception as e:
            print(f"保存配置失败: {e}")
            return False

    def validate(self) -> tuple:
        """验证配置有效性，返回 (是否有效, 错误列表)"""
        errors = []

        if "customer_types" not in self._config:
            errors.append("缺少 customer_types 配置节")
            return False, errors

        for name, config in self._config.get("customer_types", {}).items():
            # 验证必需字段
            required_fields = ["id", "display_name"]
            for fld in required_fields:
                if fld not in config:
                    errors.append(f"客户类型 '{name}' 缺少必需字段: {fld}")

            # 验证汇总表表头
            if "summary_headers" in config:
                headers = config["summary_headers"]
                if "columns" not in headers or not headers["columns"]:
                    errors.append(f"客户类型 '{name}' 的汇总表表头未定义列")

        return len(errors) == 0, errors

    def _parse_customer_types(self):
        """解析客户类型配置为数据类"""
        self._customer_types.clear()
        for name, config in self._config.get("customer_types", {}).items():
            # 解析汇总表表头
            summary_headers = None
            if "summary_headers" in config:
                sh = config["summary_headers"]
                columns = []
                for col in sh.get("columns", []):
                    columns.append(ColumnConfig(
                        key=col.get("key", ""),
                        label=col.get("label", ""),
                        width=col.get("width", 15),
                        format=col.get("format"),
                        sum=col.get("sum", False),
                        wrap_text=col.get("wrap_text", False),
                        fixed_value=col.get("fixed_value"),
                        transform=col.get("transform")
                    ))
                summary_headers = SummaryHeaderConfig(
                    columns=columns,
                    sum_row_label_column=sh.get("sum_row_label_column", 1),
                    rmb_total_column=sh.get("rmb_total_column")
                )

            # 解析排版文件列宽配置
            processed_widths = []
            for cw in config.get("processed_column_widths", []):
                processed_widths.append(ProcessedColumnWidth(
                    column=cw.get("column", "A"),
                    width=cw.get("width", 10.0)
                ))

            # 解析来源字段配置
            source_field = None
            if "source_field" in config:
                sf = config["source_field"]
                source_field = SourceFieldConfig(
                    enabled=sf.get("enabled", False),
                    column_keywords=sf.get("column_keywords", []),
                    row=sf.get("row", "first_data"),
                    label=sf.get("label", "来源")
                )

            # 解析可视化设置
            visual_settings = None
            if "visual_settings" in config:
                vs = config["visual_settings"]
                visual_settings = VisualSettings(
                    font=vs.get("font", {}),
                    row_height=vs.get("row_height", {}),
                    header_style=vs.get("header_style", {}),
                    data_style=vs.get("data_style", {}),
                    border=vs.get("border", {}),
                    print_settings=vs.get("print_settings", {}),
                    number_formats=vs.get("number_formats", {})
                )

            # 解析表头识别规则
            header_recognition = None
            if "header_recognition" in config:
                hr = config["header_recognition"]
                header_recognition = HeaderRecognitionConfig(
                    header_keywords=hr.get("header_keywords", ["序号"]),
                    total_keywords=hr.get("total_keywords", ["合计"]),
                    max_search_rows=hr.get("max_search_rows", 10),
                    column_mappings=hr.get("column_mappings", {})
                )

            # 解析数据提取规则
            data_extraction = None
            if "data_extraction" in config:
                de = config["data_extraction"]
                data_extraction = DataExtractionConfig(
                    business_count_method=de.get("business_count_method", "sequence"),
                    sequence_column=de.get("sequence_column", 1),
                    date_columns=de.get("date_columns", ["起运日期", "发货日期"]),
                    date_format=de.get("date_format", "%Y/%m/%d"),
                    numeric_columns=de.get("numeric_columns", {}),
                    skip_empty_rows=de.get("skip_empty_rows", True),
                    data_start_offset=de.get("data_start_offset", 1)
                )

            # 解析费率计算配置
            rate_calculation = None
            if "rate_calculation" in config:
                rc = config["rate_calculation"]
                rate_calculation = RateCalculationConfig(
                    calculation_mode=rc.get("calculation_mode", "auto"),
                    fixed_rate=rc.get("fixed_rate"),
                    rate_precision=rc.get("rate_precision", 8),
                    premium_precision=rc.get("premium_precision", 2),
                    rate_column=rc.get("rate_column"),
                    premium_column=rc.get("premium_column"),
                    insurance_amount_column=rc.get("insurance_amount_column"),
                    formula=rc.get("formula", "premium = insurance_amount * rate"),
                    display_unit=rc.get("display_unit", "%"),
                    display_multiplier=rc.get("display_multiplier", 100)
                )

            # 解析特殊列配置
            special_columns = []
            for sc in config.get("special_columns", []):
                special_columns.append(SpecialColumnConfig(
                    column_name=sc.get("column_name", ""),
                    width=sc.get("width", 15),
                    wrap_text=sc.get("wrap_text", False),
                    alignment=sc.get("alignment", "center"),
                    font_size=sc.get("font_size"),
                    first_sheet_width=sc.get("first_sheet_width"),
                    other_sheet_width=sc.get("other_sheet_width")
                ))

            # 解析行高配置
            row_heights = None
            if "row_heights" in config:
                rh = config["row_heights"]
                row_heights = RowHeightConfig(
                    title_row=rh.get("title_row", 39),
                    second_row=rh.get("second_row", 33),
                    header_row=rh.get("header_row", 32),
                    data_row=rh.get("data_row", 15),
                    total_row=rh.get("total_row", 18),
                    footer_row=rh.get("footer_row", 18),
                    auto_fit=rh.get("auto_fit", True)
                )

            self._customer_types[name] = CustomerTypeConfig(
                id=config.get("id", ""),
                display_name=config.get("display_name", name),
                description=config.get("description", ""),
                enabled=config.get("enabled", True),
                extract_function=config.get("extract_function", "extract_sheet_data"),
                process_function=config.get("process_function", "process_multimodal_sheet"),
                use_excel_formatter=config.get("use_excel_formatter", True),
                is_hengli=config.get("is_hengli", False),
                page_orientation=config.get("page_orientation", "landscape"),
                summary_headers=summary_headers,
                placeholders=config.get("placeholders", {}),
                rate_config=config.get("rate_config", {}),
                log_format=config.get("log_format", ""),
                agreement_codes=config.get("agreement_codes", {}),
                pdf_export_groups=config.get("pdf_export_groups", {}),
                processed_column_widths=processed_widths,
                statement_groups=config.get("statement_groups", {}),
                source_field=source_field,
                visual_settings=visual_settings,
                header_recognition=header_recognition,
                data_extraction=data_extraction,
                rate_calculation=rate_calculation,
                special_columns=special_columns,
                row_heights=row_heights
            )

    def get_customer_types(self) -> List[str]:
        """获取所有启用的客户类型名称列表"""
        return [
            name for name, config in self._customer_types.items()
            if config.enabled
        ]

    def get_all_customer_types(self) -> List[str]:
        """获取所有客户类型名称（包括禁用的）"""
        return list(self._customer_types.keys())

    def get_customer_config(self, name: str) -> Optional[CustomerTypeConfig]:
        """获取指定客户类型的配置"""
        return self._customer_types.get(name)

    def get_raw_config(self, name: str) -> Optional[Dict]:
        """获取指定客户类型的原始配置字典"""
        return self._config.get("customer_types", {}).get(name)

    def add_customer_type(self, name: str, config: Dict[str, Any]) -> bool:
        """添加新客户类型"""
        if name in self._config.get("customer_types", {}):
            return False

        self._config.setdefault("customer_types", {})[name] = config
        self._parse_customer_types()
        return self.save()

    def update_customer_type(self, name: str, config: Dict[str, Any]) -> bool:
        """更新客户类型配置"""
        if name not in self._config.get("customer_types", {}):
            return False

        self._config["customer_types"][name].update(config)
        self._parse_customer_types()
        return self.save()

    def delete_customer_type(self, name: str) -> bool:
        """删除客户类型"""
        if name not in self._config.get("customer_types", {}):
            return False

        del self._config["customer_types"][name]
        self._parse_customer_types()
        return self.save()

    def rename_customer_type(self, old_name: str, new_name: str) -> bool:
        """重命名客户类型"""
        if old_name not in self._config.get("customer_types", {}):
            return False
        if new_name in self._config.get("customer_types", {}):
            return False

        config = self._config["customer_types"].pop(old_name)
        config["display_name"] = new_name
        self._config["customer_types"][new_name] = config
        self._parse_customer_types()
        return self.save()

    def get_headers(self, customer_type: str) -> List[str]:
        """获取指定客户类型的汇总表表头标签列表"""
        config = self.get_customer_config(customer_type)
        if config and config.summary_headers:
            return [col.label for col in config.summary_headers.columns]
        return []

    def get_header_columns(self, customer_type: str) -> List[ColumnConfig]:
        """获取指定客户类型的汇总表列配置"""
        config = self.get_customer_config(customer_type)
        if config and config.summary_headers:
            return config.summary_headers.columns
        return []

    def get_placeholders(self, customer_type: str) -> List[str]:
        """获取指定客户类型的占位符列表"""
        config = self.get_customer_config(customer_type)
        if config and config.placeholders:
            return config.placeholders.get("common", [])
        return []

    def get_placeholder_mappings(self, customer_type: str) -> Dict[str, Any]:
        """获取指定客户类型的占位符映射"""
        config = self.get_customer_config(customer_type)
        if config and config.placeholders:
            return config.placeholders.get("mappings", {})
        return {}

    def get_agreement_codes(self, customer_type: str = None) -> Dict[str, str]:
        """获取协议编号映射"""
        if customer_type:
            config = self.get_customer_config(customer_type)
            if config:
                return config.agreement_codes
            return {}
        # 合并所有客户类型的协议编号
        all_codes = {}
        for config in self._customer_types.values():
            all_codes.update(config.agreement_codes)
        return all_codes

    def get_pdf_export_groups(self, customer_type: str) -> Dict[str, Callable]:
        """获取PDF导出分组规则（转换为可调用函数）"""
        config = self.get_customer_config(customer_type)
        if not config or not config.pdf_export_groups:
            return {}

        groups = {}
        for group_name, rule in config.pdf_export_groups.items():
            match_type = rule.get("match_type", "contains")
            patterns = rule.get("patterns", [])

            if match_type == "contains" and patterns:
                pattern = patterns[0]
                groups[group_name] = lambda name, p=pattern: p in name
            elif match_type == "contains_any" and patterns:
                groups[group_name] = lambda name, ps=patterns: any(p in name for p in ps)
            elif match_type == "startswith" and patterns:
                pattern = patterns[0]
                groups[group_name] = lambda name, p=pattern: name.startswith(p)
            elif match_type == "endswith" and patterns:
                pattern = patterns[0]
                groups[group_name] = lambda name, p=pattern: name.endswith(p)

        return groups

    def get_pdf_export_groups_raw(self, customer_type: str) -> Dict[str, Any]:
        """获取PDF导出分组规则的原始配置"""
        config = self.get_customer_config(customer_type)
        if config:
            return config.pdf_export_groups
        return {}

    def register_change_callback(self, callback: Callable):
        """注册配置变更回调函数"""
        if callback not in self._callbacks:
            self._callbacks.append(callback)

    def unregister_change_callback(self, callback: Callable):
        """注销配置变更回调函数"""
        if callback in self._callbacks:
            self._callbacks.remove(callback)

    def export_config(self, path: str) -> bool:
        """导出配置到指定路径"""
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"导出配置失败: {e}")
            return False

    def import_config(self, path: str) -> tuple:
        """从指定路径导入配置，返回 (是否成功, 错误信息)"""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                new_config = json.load(f)

            # 验证导入的配置
            temp_config = self._config
            self._config = new_config
            valid, errors = self.validate()

            if not valid:
                self._config = temp_config
                return False, errors

            self.save()
            self._parse_customer_types()
            return True, []
        except Exception as e:
            return False, [str(e)]

    def get_global_settings(self) -> Dict[str, Any]:
        """获取全局设置"""
        return self._config.get("global_settings", {})

    def update_global_settings(self, settings: Dict[str, Any]) -> bool:
        """更新全局设置"""
        self._config.setdefault("global_settings", {}).update(settings)
        return self.save()

    def get_visual_settings(self, customer_type: str = None) -> Dict[str, Any]:
        """获取可视化设置（合并全局设置和客户类型特定设置）"""
        # 获取全局可视化设置
        global_vs = self._config.get("global_settings", {}).get("visual_settings", {})

        if not customer_type:
            return global_vs

        # 获取客户类型特定的可视化设置
        config = self.get_customer_config(customer_type)
        if not config or not config.visual_settings:
            return global_vs

        # 合并设置（客户类型设置优先）
        merged = {}
        for key in ['font', 'row_height', 'header_style', 'data_style', 'border', 'print_settings', 'number_formats']:
            global_val = global_vs.get(key, {})
            customer_val = getattr(config.visual_settings, key, {}) or {}
            if isinstance(global_val, dict) and isinstance(customer_val, dict):
                merged[key] = {**global_val, **customer_val}
            else:
                merged[key] = customer_val if customer_val else global_val

        return merged

    def get_source_field_config(self, customer_type: str) -> Optional[SourceFieldConfig]:
        """获取来源字段配置"""
        config = self.get_customer_config(customer_type)
        if config:
            return config.source_field
        return None

    def _get_default_config(self) -> Dict[str, Any]:
        """获取默认配置（将现有硬编码转换为配置）"""
        return {
            "_meta": {
                "version": self.DEFAULT_CONFIG_VERSION,
                "description": "货运险投保工具 - 客户类型配置文件",
                "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            },

            "customer_types": {
                "多式联运": {
                    "id": "multimodal",
                    "display_name": "多式联运",
                    "description": "标准多式联运保险业务",
                    "enabled": True,
                    "extract_function": "extract_sheet_data",
                    "process_function": "process_multimodal_sheet",
                    "use_excel_formatter": True,
                    "is_hengli": False,
                    "summary_headers": {
                        "columns": [
                            {"key": "file_name", "label": "文件名", "width": 25},
                            {"key": "sheet_name", "label": "Sheet名", "width": 30},
                            {"key": "ship_voyage", "label": "船名/航次", "width": 15},
                            {"key": "business_count", "label": "业务笔数", "width": 10, "sum": True},
                            {"key": "departure_date", "label": "起运日期", "width": 12},
                            {"key": "cargo_type", "label": "货种", "width": 50},
                            {"key": "tonnage", "label": "实载吨位", "width": 15, "format": "#,##0.000", "sum": True},
                            {"key": "insurance_amount", "label": "保险金额", "width": 18, "format": "#,##0.00", "sum": True},
                            {"key": "rate_percent", "label": "综合费率", "width": 12},
                            {"key": "rate_permille", "label": "千分费率", "width": 10},
                            {"key": "new_premium", "label": "保费", "width": 15, "format": "#,##0.00", "sum": True},
                            {"key": "rmb_premium", "label": "人民币保费", "width": 20},
                            {"key": "special_terms", "label": "非标准化特约", "width": 60, "wrap_text": True}
                        ],
                        "sum_row_label_column": 1,
                        "rmb_total_column": 12
                    },
                    "placeholders": {
                        "common": [
                            "{Sheet名}", "{船名/航次}", "{业务笔数}", "{起运日期}",
                            "{货种}", "{实载吨位}", "{保险金额}", "{总保额}",
                            "{综合费率}", "{千分费率}", "{保费}", "{总保费}",
                            "{保费大写}", "{非标准化特约}", "{特约}"
                        ],
                        "mappings": {
                            "{Sheet名}": "sheet_name",
                            "{船名/航次}": "ship_voyage",
                            "{业务笔数}": "business_count",
                            "{起运日期}": "departure_date",
                            "{货种}": "cargo_type",
                            "{实载吨位}": {"field": "tonnage", "format": ",.3f"},
                            "{保险金额}": {"field": "insurance_amount", "format": ",.2f"},
                            "{总保额}": {"field": "insurance_amount", "format": ",.2f"},
                            "{综合费率}": {"field": "rate", "format": ".6%"},
                            "{千分费率}": {"field": "rate", "multiplier": 1000, "format": ".6f"},
                            "{保费}": {"field": "new_premium", "format": ",.2f"},
                            "{总保费}": {"field": "new_premium", "format": ",.2f"},
                            "{保费大写}": {"field": "new_premium", "transform": "cn_currency"},
                            "{非标准化特约}": "special_terms",
                            "{特约}": "special_terms"
                        }
                    },
                    "rate_config": {
                        "display_unit": "%",
                        "multiplier": 100,
                        "precision": 6
                    },
                    "log_format": "{sheet_name}: 费率={rate:.6f}, 保费={new_premium:,.2f}"
                },

                "恒力PTA": {
                    "id": "hengli_pta",
                    "display_name": "恒力PTA",
                    "description": "恒力PTA运输保险业务",
                    "enabled": True,
                    "extract_function": "extract_hengli_data",
                    "process_function": "process_hengli_sheet",
                    "use_excel_formatter": True,
                    "is_hengli": True,
                    "page_orientation": "landscape",
                    "summary_headers": {
                        "columns": [
                            {"key": "file_name", "label": "文件名", "width": 25},
                            {"key": "sheet_name", "label": "Sheet名", "width": 30},
                            {"key": "transport_tool", "label": "运输工具", "width": 20},
                            {"key": "business_count", "label": "业务笔数", "width": 10, "sum": True},
                            {"key": "departure_date", "label": "起运日期", "width": 12},
                            {"key": "tonnage", "label": "装货数量（吨）", "width": 15, "format": "#,##0.00", "sum": True},
                            {"key": "insurance_amount", "label": "保险金额", "width": 18, "format": "#,##0.00", "sum": True},
                            {"key": "rate_permille", "label": "千分费率", "width": 10, "fixed_value": "0.083"},
                            {"key": "new_premium", "label": "保费", "width": 15, "format": "#,##0.00", "sum": True},
                            {"key": "rmb_premium", "label": "人民币保费", "width": 20}
                        ],
                        "sum_row_label_column": 1,
                        "rmb_total_column": 10
                    },
                    "placeholders": {
                        "common": [
                            "{Sheet名}", "{运输工具}", "{业务笔数}", "{起运日期}",
                            "{装货数量（吨）}", "{装货数量}", "{保险金额}", "{千分费率}",
                            "{保费}", "{保费大写}"
                        ],
                        "mappings": {
                            "{Sheet名}": "sheet_name",
                            "{运输工具}": "transport_tool",
                            "{业务笔数}": "business_count",
                            "{起运日期}": "departure_date",
                            "{装货数量（吨）}": {"field": "tonnage", "format": ",.2f"},
                            "{装货数量}": {"field": "tonnage", "format": ",.2f"},
                            "{保险金额}": {"field": "insurance_amount", "format": ",.2f"},
                            "{千分费率}": {"fixed": "0.083"},
                            "{保费}": {"field": "new_premium", "format": ",.2f"},
                            "{保费大写}": {"field": "new_premium", "transform": "cn_currency"}
                        }
                    },
                    "rate_config": {
                        "fixed_rate": 0.000083,
                        "display_unit": "‰",
                        "display_value": "0.083"
                    },
                    "log_format": "{sheet_name}: 费率=0.083‰, 保费={new_premium:,.2f}",
                    "statement_groups": {
                        "PTA船运": {"match_type": "contains", "patterns": ["PTA船运"]},
                        "BA": {"match_type": "contains", "patterns": ["BA"]},
                        "PTA车运": {"match_type": "contains", "patterns": ["PTA车运"]}
                    }
                },

                "惠州PTA": {
                    "id": "huizhou_pta",
                    "display_name": "惠州PTA",
                    "description": "惠州PTA运输保险业务（纵向打印）",
                    "enabled": True,
                    "extract_function": "extract_hengli_data",
                    "process_function": "process_hengli_sheet",
                    "use_excel_formatter": True,
                    "is_hengli": True,
                    "page_orientation": "portrait",
                    "summary_headers": {
                        "columns": [
                            {"key": "file_name", "label": "文件名", "width": 25},
                            {"key": "sheet_name", "label": "Sheet名", "width": 30},
                            {"key": "transport_tool", "label": "运输工具", "width": 20},
                            {"key": "business_count", "label": "业务笔数", "width": 10, "sum": True},
                            {"key": "departure_date", "label": "起运日期", "width": 12},
                            {"key": "tonnage", "label": "装货数量（吨）", "width": 15, "format": "#,##0.00", "sum": True},
                            {"key": "insurance_amount", "label": "保险金额", "width": 18, "format": "#,##0.00", "sum": True},
                            {"key": "rate_permille", "label": "千分费率", "width": 10, "fixed_value": "0.083"},
                            {"key": "new_premium", "label": "保费", "width": 15, "format": "#,##0.00", "sum": True},
                            {"key": "rmb_premium", "label": "人民币保费", "width": 20}
                        ],
                        "sum_row_label_column": 1,
                        "rmb_total_column": 10
                    },
                    "placeholders": {
                        "common": [
                            "{Sheet名}", "{运输工具}", "{业务笔数}", "{起运日期}",
                            "{装货数量（吨）}", "{装货数量}", "{保险金额}", "{千分费率}",
                            "{保费}", "{保费大写}"
                        ],
                        "mappings": {
                            "{Sheet名}": "sheet_name",
                            "{运输工具}": "transport_tool",
                            "{业务笔数}": "business_count",
                            "{起运日期}": "departure_date",
                            "{装货数量（吨）}": {"field": "tonnage", "format": ",.2f"},
                            "{装货数量}": {"field": "tonnage", "format": ",.2f"},
                            "{保险金额}": {"field": "insurance_amount", "format": ",.2f"},
                            "{千分费率}": {"fixed": "0.083"},
                            "{保费}": {"field": "new_premium", "format": ",.2f"},
                            "{保费大写}": {"field": "new_premium", "transform": "cn_currency"}
                        }
                    },
                    "rate_config": {
                        "fixed_rate": 0.000083,
                        "display_unit": "‰",
                        "display_value": "0.083"
                    },
                    "log_format": "{sheet_name}: 费率=0.083‰, 保费={new_premium:,.2f}",
                    "processed_column_widths": [
                        {"column": "D", "width": 13.5},
                        {"column": "E", "width": 15.8},
                        {"column": "F", "width": 24.7},
                        {"column": "G", "width": 15.8}
                    ],
                    "statement_groups": {
                        "全部数据": {"match_type": "contains_any", "patterns": [""]}
                    }
                },

                "恒力能源销售": {
                    "id": "hengli_energy",
                    "display_name": "恒力能源销售",
                    "description": "恒力能源销售保险业务",
                    "enabled": True,
                    "extract_function": "extract_hengli_energy_data",
                    "process_function": "process_hengli_energy_sheet",
                    "use_excel_formatter": False,
                    "is_hengli": False,
                    "summary_headers": {
                        "columns": [
                            {"key": "comp", "label": "申报公司名称", "width": 38},
                            {"key": "no", "label": "车船号", "width": 14},
                            {"key": "date", "label": "发货日期", "width": 18, "transform": "format_date_slashes"},
                            {"key": "latest_date", "label": "申报止期", "width": 14},
                            {"key": "mat", "label": "物料名称", "width": 38},
                            {"key": "business_count", "label": "业务笔数", "width": 10, "sum": True},
                            {"key": "amt", "label": "开单量", "width": 14, "format": "0.000", "sum": True},
                            {"key": "money", "label": "金额（元）", "width": 18, "format": "#,##0.00", "sum": True},
                            {"key": "prem", "label": "保费（元）", "width": 16, "format": "#,##0.00", "sum": True}
                        ],
                        "sum_row_label_column": 1,
                        "rmb_total_column": None
                    },
                    "placeholders": {
                        "common": [
                            "{申报公司名称}", "{申报公司}", "{车船号}", "{发货日期}",
                            "{物料名称}", "{业务笔数}", "{开单量}", "{金额（元）}",
                            "{金额(元)}", "{金额}", "{保费（元）}", "{保费(元)}",
                            "{保费}", "{保费大写}", "{协议编号}", "{起始日期}",
                            "{截止日期}", "{申报周期}"
                        ],
                        "mappings": {
                            "{申报公司名称}": "comp",
                            "{申报公司}": "comp",
                            "{车船号}": "no",
                            "{发货日期}": {"field": "date", "transform": "format_date_slashes"},
                            "{物料名称}": "mat",
                            "{业务笔数}": "business_count",
                            "{开单量}": {"field": "amt", "format": ",.3f"},
                            "{金额（元）}": {"field": "money", "format": ",.2f"},
                            "{金额(元)}": {"field": "money", "format": ",.2f"},
                            "{金额}": {"field": "money", "format": ",.2f"},
                            "{保费（元）}": {"field": "prem", "format": ",.2f"},
                            "{保费(元)}": {"field": "prem", "format": ",.2f"},
                            "{保费}": {"field": "prem", "format": ",.2f"},
                            "{保费大写}": {"field": "prem", "transform": "cn_currency"},
                            "{协议编号}": {"lookup": "agreement_codes", "key_field": "comp"},
                            "{起始日期}": {"source": "period", "field": "start"},
                            "{截止日期}": {"source": "period", "field": "end"},
                            "{申报周期}": {"source": "period", "field": "str"}
                        }
                    },
                    "rate_config": {},
                    "log_format": "{sheet_name}: 金额={money:,.2f}, 保费={prem:,.2f}",
                    "agreement_codes": {
                        "恒力能源（苏州）有限公司": "CSHHHYX2025Q000337",
                        "苏州恒力精细化工销售有限公司": "CSHHHYX2025Q000360",
                        "恒力石化销售有限公司": "CSHHHYX2025Q000356",
                        "恒力油品销售（苏州）有限公司": "CSHHHYX2025Q000361",
                        "恒力华南石化销售有限公司": "CSHHHYX2025Q000358"
                    },
                    "pdf_export_groups": {
                        "能源苏州": {"match_type": "contains", "patterns": ["能源苏州"]},
                        "华南石化": {"match_type": "contains", "patterns": ["华南石化"]},
                        "其他业务": {"match_type": "contains_any", "patterns": ["精细化工", "恒力石化", "油品销售"]}
                    }
                }
            },

            "global_settings": {
                "default_customer_type": "多式联运",
                "date_format": "yyyy/MM/dd",
                "currency_locale": "zh_CN"
            }
        }


# 全局配置管理器实例（延迟初始化）
_config_manager_instance = None


def get_config_manager() -> CustomerConfigManager:
    """获取全局配置管理器实例"""
    global _config_manager_instance
    if _config_manager_instance is None:
        _config_manager_instance = CustomerConfigManager()
    return _config_manager_instance
