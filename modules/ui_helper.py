# -*- coding: utf-8 -*-
"""
UI 辅助模块 v5.0 — 智能列映射确认区（极简版）

职责：
    1. ensure_file_loaded  — 上传文件 → xlwings 解密 → 缓存 raw_data
    2. render_bom_mapping   — BOM 列映射 UI（极简：仅下拉框）
    3. render_list_mapping  — 通用清单列映射 UI（极简：仅下拉框）
    4. auto_predict_column  — 基于关键词打分自动预判最佳列
"""

import streamlit as st
from io import BytesIO
from typing import List, Any, Optional

from .config import (
    MappingConfig,
    PART_KEYWORDS, QTY_KEYWORDS, SUBSTITUTE_KEYWORDS,
    NAME_KEYWORDS, BOX_KEYWORDS,
    ALL_BOM_KEYWORDS, ALL_LIST_KEYWORDS,
    STREAM_PARSE_LABEL, NO_COLUMN_LABEL,
    HEADER_SCAN_ROWS, MIN_HEADER_SCORE,
    BOX_MARKER_PATTERNS,
)
from .utils import smart_find_header_row, extract_box_number_from_text
from .file_reader import load_excel_secure


# ============================================================
# 工具函数
# ============================================================
def _col_letter(idx: int) -> str:
    """0-based → Excel 列字母 (A, B, ..., Z, AA, ...)"""
    result = ''
    i = idx
    while True:
        result = chr(65 + i % 26) + result
        i = i // 26 - 1
        if i < 0:
            break
    return result


def _build_options(headers: List[str]) -> List[str]:
    """构建带列字母的显示选项"""
    return [f"{_col_letter(i)}列: {h}" for i, h in enumerate(headers)]


# ============================================================
# 自动预判
# ============================================================
def auto_predict_column(headers: List[str], keywords: List[str]) -> Optional[int]:
    """关键词匹配打分，返回得分最高的列索引"""
    best_idx: Optional[int] = None
    best_score = 0
    for col_idx, cell in enumerate(headers):
        text = str(cell).lower() if cell else ''
        score = sum(1 for kw in keywords if kw.lower() in text)
        if score > best_score:
            best_score = score
            best_idx = col_idx
    return best_idx if best_score > 0 else None


def _has_stream_markers(raw_data: List[List[Any]], header_row: int) -> bool:
    """快速扫描表头后数据区，判断是否存在流式分箱标记"""
    end = min(len(raw_data), header_row + 60)
    for idx in range(header_row + 1, end):
        row = raw_data[idx]
        if not row:
            continue
        text = ' '.join(str(c) if c else '' for c in row)
        if extract_box_number_from_text(text, BOX_MARKER_PATTERNS):
            return True
    return False


# ============================================================
# 文件缓存（避免重复 xlwings 解密）
# ============================================================
def ensure_file_loaded(uploaded_file, cache_key: str) -> Optional[List[List[Any]]]:
    """若文件为新上传，则解密读取并缓存"""
    raw_key = f'{cache_key}_raw'
    fp_key = f'{cache_key}_fp'

    if uploaded_file is None:
        st.session_state.pop(raw_key, None)
        st.session_state.pop(fp_key, None)
        return None

    fp = f'{uploaded_file.name}_{uploaded_file.size}'

    if st.session_state.get(fp_key) != fp:
        with st.spinner(f'🔐 正在解密读取 **{uploaded_file.name}** …'):
            try:
                uploaded_file.seek(0)
                raw = load_excel_secure(BytesIO(uploaded_file.read()), uploaded_file.name)
                if not raw:
                    st.error(f'❌ 文件为空: {uploaded_file.name}')
                    return None
                st.session_state[raw_key] = raw
                st.session_state[fp_key] = fp
                st.session_state['processed'] = False
            except Exception as e:
                st.error(f'❌ 读取失败: {e}')
                return None

    return st.session_state.get(raw_key)


# ============================================================
# BOM 映射 UI（极简版）
# ============================================================
def render_bom_mapping(
    raw_data: List[List[Any]],
    key_prefix: str = 'bom',
    show_title: bool = True,
) -> Optional[MappingConfig]:
    """渲染 BOM 列映射（极简：仅下拉框，完全依赖 Smart Anchor）"""
    if not raw_data:
        return None

    # Smart Anchor 自动检测表头行
    auto_idx, _ = smart_find_header_row(
        raw_data, ALL_BOM_KEYWORDS,
        max_rows=HEADER_SCAN_ROWS, min_score=MIN_HEADER_SCORE,
    )
    if auto_idx is None:
        auto_idx = 0

    if show_title:
        st.markdown('##### 📑 BOM 清单')

    header_row = auto_idx
    raw_headers = raw_data[header_row] if header_row < len(raw_data) else []
    headers = [str(c) if c else f'列{i+1}' for i, c in enumerate(raw_headers)]
    n = len(headers)
    if n == 0:
        st.warning('表头行无有效列')
        return None

    opts = _build_options(headers)

    # 料号列
    pred_part = auto_predict_column(headers, PART_KEYWORDS) or 0
    part_col = st.selectbox(
        '料号列', options=range(n), format_func=lambda i: opts[i],
        index=min(pred_part, n - 1), key=f'{key_prefix}_part',
    )

    # 数量列
    pred_qty = auto_predict_column(headers, QTY_KEYWORDS)
    qty_default = pred_qty if pred_qty is not None else min(1, n - 1)
    qty_col = st.selectbox(
        '数量列', options=range(n), format_func=lambda i: opts[i],
        index=min(qty_default, n - 1), key=f'{key_prefix}_qty',
    )

    # 可选列
    none_opts = [-1] + list(range(n))

    def _opt_fmt(i):
        return NO_COLUMN_LABEL if i == -1 else opts[i]

    pred_sub = auto_predict_column(headers, SUBSTITUTE_KEYWORDS)
    sub_default = (pred_sub + 1) if pred_sub is not None else 0
    sub_val = st.selectbox(
        '替代料列', options=none_opts, format_func=_opt_fmt,
        index=min(sub_default, len(none_opts) - 1), key=f'{key_prefix}_sub',
    )

    pred_name = auto_predict_column(headers, NAME_KEYWORDS)
    name_default = (pred_name + 1) if pred_name is not None else 0
    name_val = st.selectbox(
        '名称列', options=none_opts, format_func=_opt_fmt,
        index=min(name_default, len(none_opts) - 1), key=f'{key_prefix}_name',
    )

    return MappingConfig(
        header_row=header_row,
        part_col=part_col,
        qty_col=qty_col,
        substitute_col=sub_val if sub_val >= 0 else None,
        name_col=name_val if name_val >= 0 else None,
    )


# ============================================================
# 通用清单映射 UI（极简版）
# ============================================================
def render_list_mapping(
    raw_data: List[List[Any]],
    key_prefix: str,
    label: str = '清单',
    show_title: bool = False,
) -> Optional[MappingConfig]:
    """渲染清单列映射（极简：仅下拉框，完全依赖 Smart Anchor）"""
    if not raw_data:
        return None

    # Smart Anchor 自动检测表头行
    auto_idx, _ = smart_find_header_row(
        raw_data, ALL_LIST_KEYWORDS,
        max_rows=HEADER_SCAN_ROWS, min_score=MIN_HEADER_SCORE,
    )
    if auto_idx is None:
        auto_idx = 0

    if show_title:
        st.markdown(f'##### 📑 {label}')

    header_row = auto_idx
    raw_headers = raw_data[header_row] if header_row < len(raw_data) else []
    headers = [str(c) if c else f'列{i+1}' for i, c in enumerate(raw_headers)]
    n = len(headers)
    if n == 0:
        st.warning('表头行无有效列')
        return None

    opts = _build_options(headers)

    # 料号 & 数量
    pred_part = auto_predict_column(headers, PART_KEYWORDS) or 0
    part_col = st.selectbox(
        '料号列', options=range(n), format_func=lambda i: opts[i],
        index=min(pred_part, n - 1), key=f'{key_prefix}_part',
    )

    pred_qty = auto_predict_column(headers, QTY_KEYWORDS)
    qty_default = pred_qty if pred_qty is not None else min(1, n - 1)
    qty_col = st.selectbox(
        '数量列', options=range(n), format_func=lambda i: opts[i],
        index=min(qty_default, n - 1), key=f'{key_prefix}_qty',
    )

    # 箱号（含流式解析选项）
    STREAM_VAL = -2
    NONE_VAL = -1
    box_opts = [STREAM_VAL, NONE_VAL] + list(range(n))

    def _box_fmt(v):
        if v == STREAM_VAL:
            return STREAM_PARSE_LABEL
        if v == NONE_VAL:
            return NO_COLUMN_LABEL
        return opts[v]

    pred_box = auto_predict_column(headers, BOX_KEYWORDS)
    has_markers = _has_stream_markers(raw_data, header_row)

    if pred_box is not None:
        box_default_idx = pred_box + 2
    elif has_markers:
        box_default_idx = 0
    else:
        box_default_idx = 1

    box_val = st.selectbox(
        '箱号列', options=box_opts, format_func=_box_fmt,
        index=min(box_default_idx, len(box_opts) - 1),
        key=f'{key_prefix}_box',
    )

    stream = (box_val == STREAM_VAL)
    box_col = box_val if box_val >= 0 else None

    return MappingConfig(
        header_row=header_row,
        part_col=part_col,
        qty_col=qty_col,
        box_col=box_col,
        stream_parse=stream,
    )
