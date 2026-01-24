# -*- coding: utf-8 -*-
"""
用户自定义条款映射管理器

功能：
- 存储和管理用户自定义的条款映射（客户条款 -> 条款库名称）
- 持久化到JSON文件
- 支持模糊查找

Author: Claude
Date: 2025-01-04
"""

import json
import logging
from pathlib import Path
from typing import Dict, Optional, List, Tuple
from dataclasses import dataclass, asdict
from datetime import datetime

logger = logging.getLogger(__name__)


def _get_mapping_file_path() -> Path:
    """获取映射文件路径（支持PyInstaller打包）"""
    import sys
    import os

    # 映射文件名
    MAPPING_FILENAME = "clause_mappings.json"

    # 1. 优先使用脚本同目录的映射文件（开发模式）
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包模式
        bundle_dir = getattr(sys, '_MEIPASS', Path(sys.executable).parent)
        bundle_file = Path(bundle_dir) / MAPPING_FILENAME
    else:
        # 开发模式：使用脚本同目录
        bundle_file = Path(__file__).parent / MAPPING_FILENAME

    if bundle_file.exists():
        logger.info(f"使用映射文件: {bundle_file}")
        return bundle_file

    # 2. 回退到用户数据目录
    if sys.platform == 'darwin':
        user_data_dir = Path.home() / "Library" / "Application Support" / "ClauseNexus"
    elif sys.platform == 'win32':
        user_data_dir = Path(os.environ.get('APPDATA', Path.home())) / "ClauseNexus"
    else:
        user_data_dir = Path.home() / ".config" / "ClauseNexus"

    user_data_dir.mkdir(parents=True, exist_ok=True)
    user_mapping_file = user_data_dir / MAPPING_FILENAME

    # 如果用户目录没有文件，创建空文件
    if not user_mapping_file.exists():
        import json
        with open(user_mapping_file, 'w', encoding='utf-8') as f:
            json.dump({'version': '1.0', 'updated_at': '', 'mappings': []}, f, ensure_ascii=False)
        logger.info(f"已创建空映射文件: {user_mapping_file}")

    return user_mapping_file


# 映射文件路径
MAPPING_FILE = _get_mapping_file_path()


@dataclass
class ClauseMapping:
    """单条映射记录"""
    client_name: str          # 客户条款名称
    library_name: str         # 条款库名称
    created_at: str = ""      # 创建时间
    notes: str = ""           # 备注

    def __post_init__(self):
        if not self.created_at:
            self.created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class ClauseMappingManager:
    """用户自定义映射管理器"""

    def __init__(self, mapping_file: Path = MAPPING_FILE):
        self.mapping_file = mapping_file
        self._mappings: Dict[str, ClauseMapping] = {}  # key: normalized client_name
        self._loaded = False

    def _normalize(self, text: str) -> str:
        """标准化文本用于匹配"""
        if not text:
            return ""
        import re
        text = text.lower().strip()

        # 全角转半角（数字、字母、符号）
        text = self._fullwidth_to_halfwidth(text)

        # 移除空白和括号等
        text = re.sub(r'[\s\(\)（）\[\]【】\'\"\'\'\"\"]', '', text)
        return text

    @staticmethod
    def _fullwidth_to_halfwidth(text: str) -> str:
        """全角字符转半角"""
        result = []
        for char in text:
            code = ord(char)
            # 全角空格
            if code == 0x3000:
                result.append(' ')
            # 全角字符范围 (！到～)
            elif 0xFF01 <= code <= 0xFF5E:
                result.append(chr(code - 0xFEE0))
            else:
                result.append(char)
        return ''.join(result)

    def load(self) -> int:
        """从文件加载映射"""
        if not self.mapping_file.exists():
            logger.info(f"映射文件不存在，将创建新文件: {self.mapping_file}")
            self._mappings = {}
            self._loaded = True
            return 0

        try:
            with open(self.mapping_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            self._mappings = {}
            for item in data.get('mappings', []):
                mapping = ClauseMapping(**item)
                key = self._normalize(mapping.client_name)
                self._mappings[key] = mapping

            self._loaded = True
            logger.info(f"加载了 {len(self._mappings)} 条用户映射")
            return len(self._mappings)

        except Exception as e:
            logger.error(f"加载映射文件失败: {e}")
            self._mappings = {}
            self._loaded = True
            return 0

    def save(self) -> bool:
        """保存映射到文件"""
        try:
            data = {
                'version': '1.0',
                'updated_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'mappings': [asdict(m) for m in self._mappings.values()]
            }

            with open(self.mapping_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            logger.info(f"保存了 {len(self._mappings)} 条映射到 {self.mapping_file}")
            return True

        except Exception as e:
            logger.error(f"保存映射文件失败: {e}")
            return False

    def add_mapping(self, client_name: str, library_name: str, notes: str = "") -> bool:
        """添加或更新映射"""
        if not client_name or not library_name:
            return False

        key = self._normalize(client_name)
        self._mappings[key] = ClauseMapping(
            client_name=client_name.strip(),
            library_name=library_name.strip(),
            notes=notes
        )
        logger.info(f"添加映射: '{client_name}' -> '{library_name}'")
        return True

    def remove_mapping(self, client_name: str) -> bool:
        """删除映射"""
        key = self._normalize(client_name)
        if key in self._mappings:
            del self._mappings[key]
            logger.info(f"删除映射: '{client_name}'")
            return True
        return False

    def get_library_name(self, client_name: str) -> Optional[str]:
        """根据客户条款名称获取对应的条款库名称"""
        if not self._loaded:
            self.load()

        key = self._normalize(client_name)

        # 精确匹配
        if key in self._mappings:
            return self._mappings[key].library_name

        # 包含匹配（客户名称包含在映射的key中，或反过来）
        for stored_key, mapping in self._mappings.items():
            if key in stored_key or stored_key in key:
                return mapping.library_name

        return None

    def get_all_mappings(self) -> List[ClauseMapping]:
        """获取所有映射"""
        if not self._loaded:
            self.load()
        return list(self._mappings.values())

    def get_mapping_count(self) -> int:
        """获取映射数量"""
        if not self._loaded:
            self.load()
        return len(self._mappings)

    def apply_to_config(self, config) -> int:
        """将用户映射应用到配置对象"""
        if not self._loaded:
            self.load()

        count = 0
        for mapping in self._mappings.values():
            # 添加到语义别名映射
            if hasattr(config, 'add_semantic_alias'):
                config.add_semantic_alias(mapping.client_name, mapping.library_name)
                count += 1
        return count

    def import_from_dict(self, mappings: Dict[str, str]) -> int:
        """从字典批量导入映射"""
        count = 0
        for client, library in mappings.items():
            if self.add_mapping(client, library):
                count += 1
        return count

    def export_to_dict(self) -> Dict[str, str]:
        """导出为字典"""
        if not self._loaded:
            self.load()
        return {m.client_name: m.library_name for m in self._mappings.values()}

    def import_from_report_excel(self, excel_path: str, min_score: float = 0.8,
                                  match_levels: List[str] = None) -> Tuple[int, int]:
        """
        从条款比对报告Excel导入映射

        Args:
            excel_path: 报告Excel文件路径
            min_score: 最低匹配分数阈值（默认0.8）
            match_levels: 要导入的匹配级别列表（默认只导入精确匹配和语义匹配）

        Returns:
            (导入数量, 跳过数量)
        """
        import pandas as pd

        if match_levels is None:
            match_levels = ["精确匹配", "语义匹配"]

        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            logger.error(f"读取报告Excel失败: {e}")
            return 0, 0

        # 查找列名（支持不同格式）
        client_col = None
        library_col = None
        score_col = None
        level_col = None

        for col in df.columns:
            col_str = str(col).lower()
            if '客户条款' in col_str and '原' in col_str:
                client_col = col
            elif '匹配条款库名称' in col_str or '匹配' in col_str and '名称' in col_str:
                library_col = col
            elif '综合匹配度' in col_str or '匹配度' in col_str or 'score' in col_str:
                score_col = col
            elif '匹配级别' in col_str or '级别' in col_str:
                level_col = col

        if not client_col or not library_col:
            logger.error(f"无法识别报告列名: {list(df.columns)}")
            return 0, 0

        imported = 0
        skipped = 0

        for _, row in df.iterrows():
            client_name = str(row.get(client_col, '')).strip()
            library_name = str(row.get(library_col, '')).strip()

            if not client_name or not library_name or library_name in ['无匹配', '', 'nan']:
                skipped += 1
                continue

            # 检查分数
            if score_col:
                try:
                    score = float(row.get(score_col, 0))
                    if score < min_score:
                        skipped += 1
                        continue
                except (ValueError, TypeError):
                    pass

            # 检查匹配级别
            if level_col and match_levels:
                level = str(row.get(level_col, ''))
                if level not in match_levels:
                    skipped += 1
                    continue

            # 添加映射
            if self.add_mapping(client_name, library_name, notes=f"从报告导入"):
                imported += 1
            else:
                skipped += 1

        logger.info(f"从报告导入映射: 成功 {imported}, 跳过 {skipped}")
        return imported, skipped

    @staticmethod
    def _clean_clause_name(name: str) -> str:
        """
        清理条款名称，去除编号、限额等信息

        清理规则：
        1. 去除开头的序号（如 "1.", "2.", "8."）
        2. 去除限额信息（如 "（分项限额：总申报金额的10%）"）
        3. 去除金额信息（如 "（每次事故分项限额：人民币20,000,000元）"）
        4. 保留核心条款名称
        """
        import re
        import pandas as pd

        if not name or (isinstance(name, float) and pd.isna(name)):
            return ""

        name = str(name).strip()

        # 1. 去除开头的序号 (如 "1.", "8.", "10.")
        name = re.sub(r'^\d+\.\s*', '', name)

        # 2. 去除各种限额描述的括号内容
        # 匹配如：（分项限额：xxx）、（每次事故分项限额：xxx）等
        patterns = [
            r'[（(][^）)]*限额[^）)]*[）)]',  # 包含"限额"的括号
            r'[（(][^）)]*人民币[\d,，]+[万元]+[^）)]*[）)]',  # 包含金额的括号
            r'[（(]分项限额[：:][^）)]+[）)]',  # 分项限额
            r'[（(]每次事故[^）)]*[）)]',  # 每次事故限额
            r'[（(]总[申报]*金额的\d+%[）)]',  # 总申报金额百分比
        ]

        for pattern in patterns:
            name = re.sub(pattern, '', name)

        # 3. 清理多余空格
        name = re.sub(r'\s+', ' ', name).strip()

        return name

    def import_from_corrected_report(self, excel_path: str,
                                      clean_names: bool = True) -> Tuple[int, int]:
        """
        从已修正的条款比对报告导入映射

        直接读取报告中的客户条款名称和匹配条款库名称，建立映射关系。
        不过滤分数或匹配级别，适用于用户已手动修正过的报告。

        Args:
            excel_path: 报告Excel文件路径
            clean_names: 是否清理条款名称（去除编号、限额等）

        Returns:
            (导入数量, 跳过数量)
        """
        import pandas as pd

        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            logger.error(f"读取报告Excel失败: {e}")
            return 0, 0

        # 查找列名 (v18.3: 支持多种格式)
        client_col = None
        client_col_backup = None  # 备选：客户条款(译)
        library_col = None

        for col in df.columns:
            col_str = str(col)
            # 客户条款列
            if '客户条款' in col_str and '原' in col_str:
                client_col = col
            elif '客户条款' in col_str and '译' in col_str:
                client_col_backup = col
            # 匹配条款库列 - 支持多种格式
            elif '匹配条款库名称' in col_str:
                library_col = col
            elif col_str == '匹配1_条款名称':  # v17.1 多结果格式
                library_col = col

        # v18.3: 如果客户条款(原)为空，使用客户条款(译)作为备选
        if client_col and client_col_backup:
            # 检查客户条款(原)是否有数据
            if df[client_col].isna().all():
                logger.info("客户条款(原)列为空，使用客户条款(译)作为备选")
                client_col = client_col_backup
        elif not client_col and client_col_backup:
            client_col = client_col_backup

        if not client_col or not library_col:
            logger.error(f"无法识别报告列名，需要'客户条款(原/译)'和'匹配条款库名称/匹配1_条款名称'列")
            return 0, 0

        logger.info(f"导入列映射: 客户条款='{client_col}', 条款库='{library_col}'")

        imported = 0
        skipped = 0

        for _, row in df.iterrows():
            client_name = str(row.get(client_col, '')).strip()
            library_name = str(row.get(library_col, '')).strip()

            # 跳过空值
            if not client_name or not library_name:
                skipped += 1
                continue

            if library_name.lower() in ['无匹配', 'nan', '']:
                skipped += 1
                continue

            # 清理条款名称
            if clean_names:
                client_name_cleaned = self._clean_clause_name(client_name)
                library_name_cleaned = self._clean_clause_name(library_name)
            else:
                client_name_cleaned = client_name
                library_name_cleaned = library_name

            # 跳过清理后为空的
            if not client_name_cleaned or not library_name_cleaned:
                skipped += 1
                continue

            # 添加映射（使用清理后的名称）
            if self.add_mapping(client_name_cleaned, library_name_cleaned,
                               notes=f"从修正报告导入"):
                imported += 1
                logger.debug(f"导入映射: '{client_name_cleaned}' -> '{library_name_cleaned}'")
            else:
                skipped += 1

        logger.info(f"从修正报告导入映射: 成功 {imported}, 跳过 {skipped}")
        return imported, skipped


# 单例模式
_manager_instance: Optional[ClauseMappingManager] = None


def get_mapping_manager() -> ClauseMappingManager:
    """获取全局映射管理器实例"""
    global _manager_instance
    if _manager_instance is None:
        _manager_instance = ClauseMappingManager()
    return _manager_instance
