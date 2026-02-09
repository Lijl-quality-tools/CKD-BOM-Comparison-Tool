# -*- coding: utf-8 -*-
"""
Excel 文件读取模块 v5.0 — 归一化版

核心函数：
    load_excel_secure   — 上传→落盘→xlwings 解密→返回二维列表→清理
    parse_bom           — 根据 MappingConfig 从 raw_data 解析 BOM
    parse_generic_list  — 根据 MappingConfig 从 raw_data 解析任意清单
"""

import os
import tempfile
import threading
import logging
from io import BytesIO
from typing import List, Dict, Any, Optional, Tuple, Union
from dataclasses import dataclass, field

try:
    import xlwings as xw
except ImportError:
    xw = None

from .config import MappingConfig, ParseMode, BOX_MARKER_PATTERNS
from .utils import (
    clean_part_number,
    extract_substitute_ids,
    safe_eval_expression,
    normalize_box_number,
    extract_box_number_from_text,
    is_empty_row,
)

logger = logging.getLogger(__name__)

# 全局锁——同一时间只允许一个 Excel COM 进程
_excel_lock = threading.Lock()


# ============================================================
# 数据结构
# ============================================================
@dataclass
class BOMItem:
    """BOM 单项"""
    main_part_id: str
    quantity: float
    substitute_ids: List[str]
    name: str = ''
    row_index: int = 0


@dataclass
class ListItem:
    """清单单项（品质/生产通用）"""
    part_id: str
    quantity: float
    box_number: str = ''
    row_index: int = 0


@dataclass
class ParseDiagnostics:
    """解析诊断"""
    file_type: str
    header_row: int
    header_score: int
    parse_mode: str
    total_rows: int
    parsed_items: int
    skipped_rows: int
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            '文件类型': self.file_type,
            '表头行号': self.header_row,
            '表头得分': self.header_score,
            '解析模式': self.parse_mode,
            '总行数': self.total_rows,
            '解析项数': self.parsed_items,
            '跳过行数': self.skipped_rows,
            '警告': self.warnings,
        }


# ============================================================
# 核心：安全读取 Excel（解密服务器）
# ============================================================
def load_excel_secure(
    file_data: Union[BytesIO, bytes],
    original_filename: str = 'upload.xlsx',
    sheet_index: int = 0,
    sheet_name: Optional[str] = None,
) -> List[List[Any]]:
    """
    上传→物理落盘→xlwings 解密读取→销毁临时文件。

    * numbers=str  防止科学计数法丢精度
    * threading.Lock  防止并发 COM 冲突
    * try…finally  无论成败都清理
    """
    if xw is None:
        raise ImportError('xlwings 未安装，请运行: pip install xlwings')

    _, ext = os.path.splitext(original_filename)
    if not ext:
        ext = '.xlsx'

    tmp_path: Optional[str] = None
    app = None
    workbook = None

    try:
        # 1. 物理落盘
        with tempfile.NamedTemporaryFile(suffix=ext, prefix='ckd_', delete=False) as tmp:
            if isinstance(file_data, BytesIO):
                file_data.seek(0)
                tmp.write(file_data.read())
            else:
                tmp.write(file_data)
            tmp_path = tmp.name
        logger.info('[解密] 落盘: %s', tmp_path)

        # 2. 加锁 + xlwings 读取
        with _excel_lock:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False

            workbook = app.books.open(tmp_path)
            sheet = workbook.sheets[sheet_name] if sheet_name else workbook.sheets[sheet_index]

            used = sheet.used_range
            if used is None:
                return []

            data = used.options(numbers=str).value

            if data is None:
                return []
            if not isinstance(data, list):
                return [[data]]
            if data and not isinstance(data[0], list):
                return [data]
            return data

    finally:
        if workbook:
            try:
                workbook.close()
            except Exception:
                pass
        if app:
            try:
                app.quit()
            except Exception:
                pass
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
                logger.info('[解密] 已删除: %s', tmp_path)
            except Exception as e:
                logger.warning('[解密] 删除失败: %s', e)


# ============================================================
# BOM 解析
# ============================================================
def parse_bom(
    raw_data: List[List[Any]],
    config: MappingConfig,
) -> Tuple[List[BOMItem], 'pd.DataFrame', ParseDiagnostics]:
    """根据用户确认的列映射，从缓存的 raw_data 中解析 BOM。"""
    import pandas as pd

    header_row_idx = config.header_row
    data_rows = raw_data[header_row_idx + 1:]

    bom_items: List[BOMItem] = []
    skipped = 0

    for row_idx, row in enumerate(data_rows):
        if is_empty_row(row):
            skipped += 1
            continue
        if len(row) <= max(config.part_col, config.qty_col):
            skipped += 1
            continue

        part_id = clean_part_number(row[config.part_col] if config.part_col < len(row) else None)
        if not part_id:
            skipped += 1
            continue

        qty = safe_eval_expression(row[config.qty_col] if config.qty_col < len(row) else None)
        if qty is None or qty <= 0:
            skipped += 1
            continue

        substitutes: List[str] = []
        if config.substitute_col is not None and config.substitute_col < len(row):
            substitutes = extract_substitute_ids(row[config.substitute_col])

        name = ''
        if config.name_col is not None and config.name_col < len(row):
            name = str(row[config.name_col]).strip() if row[config.name_col] else ''

        bom_items.append(BOMItem(
            main_part_id=part_id,
            quantity=qty,
            substitute_ids=substitutes,
            name=name,
            row_index=header_row_idx + 1 + row_idx + 1,
        ))

    diag = ParseDiagnostics(
        file_type='BOM',
        header_row=header_row_idx + 1,
        header_score=0,
        parse_mode=ParseMode.STANDARD,
        total_rows=len(data_rows),
        parsed_items=len(bom_items),
        skipped_rows=skipped,
    )

    rows = [
        {'料号': i.main_part_id, '名称': i.name, '需求数量': i.quantity,
         '替代料': '; '.join(i.substitute_ids) if i.substitute_ids else '', '行号': i.row_index}
        for i in bom_items
    ]
    df = pd.DataFrame(rows) if rows else pd.DataFrame()
    return bom_items, df, diag


# ============================================================
# 通用清单解析（品质 / 生产统一入口）
# ============================================================
def parse_generic_list(
    raw_data: List[List[Any]],
    config: MappingConfig,
    file_label: str = '清单',
) -> Tuple[List[ListItem], 'pd.DataFrame', ParseDiagnostics]:
    """
    根据 MappingConfig 解析任意清单。

    * stream_parse=False → 标准列模式（直接按 box_col 取箱号）
    * stream_parse=True  → 流式分箱模式（状态机扫描"第X箱"标记行）
    """
    if config.stream_parse:
        return _parse_stream(raw_data, config, file_label)
    return _parse_standard(raw_data, config, file_label)


# ---------- 标准列模式 ----------
def _parse_standard(
    raw_data: List[List[Any]],
    config: MappingConfig,
    file_label: str,
) -> Tuple[List['ListItem'], 'pd.DataFrame', 'ParseDiagnostics']:
    import pandas as pd

    header_row_idx = config.header_row
    data_rows = raw_data[header_row_idx + 1:]

    items: List[ListItem] = []
    skipped = 0

    for row_idx, row in enumerate(data_rows):
        if is_empty_row(row):
            skipped += 1
            continue
        if len(row) <= max(config.part_col, config.qty_col):
            skipped += 1
            continue

        part_id = clean_part_number(row[config.part_col] if config.part_col < len(row) else None)
        if not part_id:
            skipped += 1
            continue

        qty = safe_eval_expression(row[config.qty_col] if config.qty_col < len(row) else None)
        if qty is None:
            skipped += 1
            continue

        box = ''
        if config.box_col is not None and config.box_col < len(row):
            box = normalize_box_number(row[config.box_col])

        items.append(ListItem(
            part_id=part_id, quantity=qty,
            box_number=box, row_index=header_row_idx + 1 + row_idx + 1,
        ))

    diag = ParseDiagnostics(
        file_type=file_label, header_row=header_row_idx + 1,
        header_score=0, parse_mode=ParseMode.STANDARD,
        total_rows=len(data_rows), parsed_items=len(items), skipped_rows=skipped,
    )

    rows = [{'料号': i.part_id, '数量': i.quantity, '箱号': i.box_number, '行号': i.row_index} for i in items]
    df = pd.DataFrame(rows) if rows else pd.DataFrame()
    return items, df, diag


# ---------- 流式分箱模式 ----------
def _parse_stream(
    raw_data: List[List[Any]],
    config: MappingConfig,
    file_label: str,
) -> Tuple[List['ListItem'], 'pd.DataFrame', 'ParseDiagnostics']:
    import pandas as pd

    header_row_idx = config.header_row
    box_patterns = BOX_MARKER_PATTERNS

    items: List[ListItem] = []
    skipped = 0
    current_box = ''

    for row_idx in range(header_row_idx + 1, len(raw_data)):
        row = raw_data[row_idx]

        if not row:
            skipped += 1
            continue

        # 检查箱号标记行
        row_text = ' '.join(str(c) if c else '' for c in row)
        box_match = extract_box_number_from_text(row_text, box_patterns)
        if box_match:
            current_box = box_match
            skipped += 1
            continue

        if is_empty_row(row):
            skipped += 1
            continue

        if len(row) <= max(config.part_col, config.qty_col):
            skipped += 1
            continue

        part_id = clean_part_number(row[config.part_col] if config.part_col < len(row) else None)
        if not part_id:
            skipped += 1
            continue

        qty = safe_eval_expression(row[config.qty_col] if config.qty_col < len(row) else None)
        if qty is None:
            skipped += 1
            continue

        items.append(ListItem(
            part_id=part_id, quantity=qty,
            box_number=current_box, row_index=row_idx + 1,
        ))

    diag = ParseDiagnostics(
        file_type=file_label, header_row=header_row_idx + 1,
        header_score=0, parse_mode=ParseMode.STREAM,
        total_rows=len(raw_data), parsed_items=len(items), skipped_rows=skipped,
    )

    rows = [{'料号': i.part_id, '数量': i.quantity, '箱号': i.box_number, '行号': i.row_index} for i in items]
    df = pd.DataFrame(rows) if rows else pd.DataFrame()
    return items, df, diag
