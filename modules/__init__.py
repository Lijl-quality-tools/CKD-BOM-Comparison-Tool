# -*- coding: utf-8 -*-
"""
CKD清单核对工具 — 模块包 v5.0（归一化）
"""

from .config import (
    JudgmentStatus,
    ParseMode,
    MappingConfig,
    PART_KEYWORDS, QTY_KEYWORDS, SUBSTITUTE_KEYWORDS, NAME_KEYWORDS, BOX_KEYWORDS,
    ALL_BOM_KEYWORDS, ALL_LIST_KEYWORDS,
    BOX_MARKER_PATTERNS,
    CHINESE_NUM_MAP,
    HEADER_SCAN_ROWS, MIN_HEADER_SCORE,
    STREAM_PARSE_LABEL, NO_COLUMN_LABEL,
)

from .file_reader import (
    BOMItem,
    ListItem,
    load_excel_secure,
    parse_bom,
    parse_generic_list,
)

from .ui_helper import (
    ensure_file_loaded,
    render_bom_mapping,
    render_list_mapping,
    auto_predict_column,
)

from .data_processor import (
    compare_bom_and_list,
    get_abnormal_results,
    get_ok_results,
    generate_summary,
    export_results_to_excel,
    validate_data,
)

from .utils import (
    clean_part_number,
    clean_cell_value,
    extract_substitute_ids,
    safe_eval_expression,
    calculate_header_score,
    smart_find_header_row,
    find_column_by_keywords,
    chinese_to_arabic,
    extract_box_number_from_text,
    is_empty_row,
    filter_empty_rows,
    normalize_box_number,
    format_number,
    merge_box_numbers,
    is_valid_part_number,
)

__all__ = [
    # config
    'JudgmentStatus', 'ParseMode', 'MappingConfig',
    'PART_KEYWORDS', 'QTY_KEYWORDS', 'SUBSTITUTE_KEYWORDS',
    'NAME_KEYWORDS', 'BOX_KEYWORDS',
    'ALL_BOM_KEYWORDS', 'ALL_LIST_KEYWORDS',
    'BOX_MARKER_PATTERNS', 'CHINESE_NUM_MAP',
    'HEADER_SCAN_ROWS', 'MIN_HEADER_SCORE',
    'STREAM_PARSE_LABEL', 'NO_COLUMN_LABEL',
    # file_reader
    'BOMItem', 'ListItem',
    'load_excel_secure', 'parse_bom', 'parse_generic_list',
    # ui_helper
    'ensure_file_loaded', 'render_bom_mapping', 'render_list_mapping', 'auto_predict_column',
    # data_processor
    'compare_bom_and_list', 'get_abnormal_results', 'get_ok_results',
    'generate_summary', 'export_results_to_excel', 'validate_data',
    # utils
    'clean_part_number', 'clean_cell_value', 'extract_substitute_ids',
    'safe_eval_expression', 'calculate_header_score', 'smart_find_header_row',
    'find_column_by_keywords', 'chinese_to_arabic', 'extract_box_number_from_text',
    'is_empty_row', 'filter_empty_rows', 'normalize_box_number',
    'format_number', 'merge_box_numbers', 'is_valid_part_number',
]
