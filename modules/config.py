# -*- coding: utf-8 -*-
"""
全局配置模块 v5.0 — 归一化极简版

只保留：判定状态、解析模式、列映射结构、关键词池、箱号正则、中文数字表。
所有"可配置输入框"已移除，关键词仅用于自动预判。
"""

from dataclasses import dataclass
from typing import List, Optional


# ============================================================
# 判定状态常量
# ============================================================
class JudgmentStatus:
    """五种全场景状态"""
    OK = "OK"
    OK_WITH_SUB = "OK (含替料)"
    NG_QTY_DIFF = "NG (数量差异)"
    NG_MISSING = "NG (缺料)"
    NG_NOT_IN_BOM = "NG (BOM无)"


# ============================================================
# 解析模式
# ============================================================
class ParseMode:
    STANDARD = "标准列模式"
    STREAM = "流式分箱模式"


# ============================================================
# 列映射配置（由 UI 层 selectbox 生成，传入读取器）
# ============================================================
@dataclass
class MappingConfig:
    header_row: int                          # 表头行索引（0-based）
    part_col: int                            # 料号列索引
    qty_col: int                             # 数量列索引
    box_col: Optional[int] = None            # 箱号列索引（None = 不使用）
    name_col: Optional[int] = None           # 名称列（BOM 专用）
    substitute_col: Optional[int] = None     # 替代料列（BOM 专用）
    stream_parse: bool = False               # 是否启用流式分箱


# ============================================================
# 关键词池 —— 仅用于 Smart Anchor 自动预判，不暴露到侧边栏
# ============================================================
PART_KEYWORDS: List[str] = [
    '编号', '料号', '物料编号', '物料号', '零件号', '材料编号',
    'Part No', 'PartNo', 'P/N', 'PN', 'Part Number', 'Material No',
]

QTY_KEYWORDS: List[str] = [
    '数量', '需求数', '用量', '需求量', '单机用量',
    '实收数', '实收', '收货数', '来货数', '发货数',
    'Qty', 'QTY', 'Quantity', 'Required Qty', 'Usage',
]

SUBSTITUTE_KEYWORDS: List[str] = [
    '替代状况', '替代料', '替代', '可替代', '代用料', '替换料',
    'Substitute', 'Alt', 'Alternative', 'Replacement',
]

NAME_KEYWORDS: List[str] = [
    '名称', '品名', '物料名称', '零件名称', '材料名称', '品名规格',
    'Description', 'Name', 'Part Name', 'Material Name',
]

BOX_KEYWORDS: List[str] = [
    '箱号', '箱别', '箱序号', '箱编号',
    'Box', 'Carton', 'Box No', 'Carton No', 'Package',
]

# 合并池（表头行打分用）
ALL_BOM_KEYWORDS = PART_KEYWORDS + QTY_KEYWORDS + SUBSTITUTE_KEYWORDS + NAME_KEYWORDS
ALL_LIST_KEYWORDS = PART_KEYWORDS + QTY_KEYWORDS + BOX_KEYWORDS


# ============================================================
# 箱号识别正则（流式分箱用）
# ============================================================
BOX_MARKER_PATTERNS: List[str] = [
    r'第?\s*(\d+)\s*号?\s*箱',
    r'第?\s*([一二三四五六七八九十百]+)\s*号?\s*箱',
    r'[Bb]ox\s*[#№]?\s*(\d+)',
    r'[Cc]arton\s*[#№]?\s*(\d+)',
]


# ============================================================
# 中文数字映射
# ============================================================
CHINESE_NUM_MAP = {
    '零': 0, '一': 1, '二': 2, '三': 3, '四': 4,
    '五': 5, '六': 6, '七': 7, '八': 8, '九': 9,
    '十': 10, '百': 100, '千': 1000,
    '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5,
    '陆': 6, '柒': 7, '捌': 8, '玖': 9, '拾': 10,
}


# ============================================================
# 扫描常量
# ============================================================
HEADER_SCAN_ROWS = 20
MIN_HEADER_SCORE = 2


# ============================================================
# UI 哨兵值
# ============================================================
STREAM_PARSE_LABEL = "🔄 【流式解析】(从行间标题抓取)"
NO_COLUMN_LABEL = "➖ （不使用）"
