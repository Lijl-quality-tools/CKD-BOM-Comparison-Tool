# -*- coding: utf-8 -*-
"""
CKD清单核对工具 v5.0 — 极致精简版

交互流程：
    1. 上传 BOM + 待核对清单（支持多选）
    2. 智能列映射确认（自动预判）
    3. 点击"开始核对" → 结果展示
    4. 下载 Excel 报告
"""

import sys
from io import BytesIO
from datetime import datetime
from pathlib import Path

import streamlit as st
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

from modules.config import JudgmentStatus
from modules.file_reader import (
    parse_bom,
    parse_generic_list,
)
from modules.ui_helper import (
    ensure_file_loaded,
    render_bom_mapping,
    render_list_mapping,
)
from modules.data_processor import (
    compare_bom_and_list,
    get_abnormal_results,
    generate_summary,
    export_results_to_excel,
    validate_data,
)


# ============================================================
# 页面 & 样式
# ============================================================
st.set_page_config(page_title='CKD清单核对 v5.0', page_icon='📋', layout='wide',
                   initial_sidebar_state='expanded')

st.markdown("""
<style>
.main .block-container{padding-top:1.5rem;padding-bottom:1.5rem}
.main-title{font-size:2rem;font-weight:700;text-align:center;padding:0.8rem 0;
  background:linear-gradient(135deg,#667eea,#764ba2);-webkit-background-clip:text;
  -webkit-text-fill-color:transparent;background-clip:text}
.stat-card{background:linear-gradient(145deg,#fff,#f0f0f3);border-radius:10px;
  padding:1rem;text-align:center;box-shadow:0 4px 15px rgba(0,0,0,.08);border:1px solid #e8e8e8}
.stat-value{font-size:1.8rem;font-weight:700;color:#2d3436}
.stat-label{font-size:.85rem;color:#636e72;margin-top:.2rem}
.pass-indicator{color:#00b894;font-weight:600}
.fail-indicator{color:#d63031;font-weight:600}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#f8f9fc,#e8ecf3)}
.list-card{background:#f8f9fa;border:1px solid #dee2e6;border-radius:8px;padding:0.8rem;margin-bottom:0.8rem}
div[data-testid="stSelectbox"] label{font-size:0.85rem;margin-bottom:0.2rem}
div[data-testid="stSelectbox"]{margin-bottom:0.5rem}
</style>
""", unsafe_allow_html=True)


# ============================================================
# Session State 初始化
# ============================================================
_DEFAULTS = dict(
    processed=False, work_order='', batch='',
    bom_items=None, bom_df=None,
    fact_lists_data=[],
)
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ============================================================
# 侧边栏（极简：基本信息 + 使用说明）
# ============================================================
with st.sidebar:
    st.markdown('### 📋 基本信息')
    st.session_state.work_order = st.text_input('工单号', st.session_state.work_order,
                                                 placeholder='请输入工单号')
    st.session_state.batch = st.text_input('批量', st.session_state.batch,
                                            placeholder='请输入批量')
    st.markdown('---')
    with st.expander('📖 使用说明', expanded=False):
        st.markdown("""
**操作步骤：**
1. 上传 BOM + 待核对清单（可多选）
2. 确认列映射（系统自动预判）
3. 点击 **开始核对**
4. 查看结果 → 下载报告

**箱号列选项：**
- 选择某一列 → 标准列模式
- 选择"流式解析" → 扫描"第X箱"标记

**五种判定状态：**
OK · OK(含替料) · NG(差异) · NG(缺料) · NG(BOM无)
        """)


# ============================================================
# 标题 & 文件上传
# ============================================================
st.markdown('<h1 class="main-title">🔍 CKD清单核对工具 v5.0</h1>', unsafe_allow_html=True)

st.markdown('### 📁 文件上传')

col_up = st.columns([1, 1])
with col_up[0]:
    bom_up = st.file_uploader('📤 BOM 清单', type=['xlsx', 'xls'], key='up_bom')
with col_up[1]:
    fact_ups = st.file_uploader('📤 待核对清单 (支持多选)', type=['xlsx', 'xls'],
                                 accept_multiple_files=True, key='up_fact_lists')


# ============================================================
# 解密读取
# ============================================================
bom_raw = ensure_file_loaded(bom_up, 'bom')

fact_raws = []
if fact_ups:
    for idx, f_up in enumerate(fact_ups):
        raw = ensure_file_loaded(f_up, f'fact_{idx}')
        if raw:
            fact_raws.append((f_up.name, raw))


# ============================================================
# 列映射确认区（极简版）
# ============================================================
bom_cfg = None
fact_cfgs = []

if bom_raw or fact_raws:
    st.markdown('---')
    st.markdown('### 📋 列映射配置')

    map_cols = st.columns([1, 1])

    with map_cols[0]:
        if bom_raw:
            bom_cfg = render_bom_mapping(bom_raw, 'bom')

    with map_cols[1]:
        if fact_raws:
            for idx, (fname, raw) in enumerate(fact_raws):
                with st.container():
                    st.markdown(f'<div class="list-card"><b>📄 {fname}</b></div>',
                                unsafe_allow_html=True)
                    cfg = render_list_mapping(raw, f'fact_{idx}', fname)
                    if cfg:
                        fact_cfgs.append((fname, cfg))


# ============================================================
# 开始核对
# ============================================================
st.markdown('<br>', unsafe_allow_html=True)
_, btn_col, _ = st.columns([1, 2, 1])
with btn_col:
    do_compare = st.button('🚀 开始核对', use_container_width=True, type='primary')

if do_compare:
    if bom_raw is None or bom_cfg is None:
        st.error('❌ 请先上传 BOM 文件并确认列映射')
        st.stop()
    if not fact_raws:
        st.error('❌ 请至少上传一个待核对清单文件')
        st.stop()
    if len(fact_cfgs) != len(fact_raws):
        st.error('❌ 部分清单的列映射配置失败，请检查')
        st.stop()
    if bom_cfg.part_col == bom_cfg.qty_col:
        st.error('❌ BOM 的料号列和数量列不能相同')
        st.stop()

    progress = st.progress(0)
    status = st.empty()

    try:
        status.text('📖 解析 BOM…')
        progress.progress(5)
        bom_items, bom_df, _ = parse_bom(bom_raw, bom_cfg)
        st.session_state.update(bom_items=bom_items, bom_df=bom_df)
        progress.progress(20)

        fact_data = []
        total_lists = len(fact_cfgs)
        for i, ((fname, raw), (_, cfg)) in enumerate(zip(fact_raws, fact_cfgs)):
            status.text(f'📖 解析 {fname}…')
            items, df, _ = parse_generic_list(raw, cfg, fname)
            for w in validate_data(bom_items, items, fname):
                st.warning(w)
            status.text(f'🔍 比对 {fname}…')
            result_df, stats = compare_bom_and_list(bom_items, items, fname,
                                                     st.session_state.work_order)
            fact_data.append({
                'file_name': fname,
                'items': items,
                'df': df,
                'result_df': result_df,
                'stats': stats,
            })
            progress.progress(int(20 + 75 * (i + 1) / total_lists))

        st.session_state.fact_lists_data = fact_data
        progress.progress(100)
        status.text('✅ 核对完成！')
        st.session_state.processed = True

    except ImportError as e:
        st.error('❌ xlwings 未安装或 Excel/WPS 环境异常，请检查环境配置')
        st.code(str(e))
    except Exception as e:
        st.error(f'❌ 处理出错: {e}')
        import traceback
        st.code(traceback.format_exc())
    finally:
        progress.empty()


# ============================================================
# 结果展示
# ============================================================
def _stat_card(value, label, color=''):
    cls = f'class="{color}"' if color else ''
    st.markdown(f'<div class="stat-card"><div class="stat-value" {cls}>{value}</div>'
                f'<div class="stat-label">{label}</div></div>', unsafe_allow_html=True)


def _result_table(df, title, stats):
    if stats:
        ok_m = stats.get('ok_main_only', 0)
        ok_s = stats.get('ok_with_substitute', 0)
        ng_q = stats.get('ng_qty_difference', 0)
        ng_m = stats.get('ng_missing', 0)
        ng_b = stats.get('ng_not_in_bom', 0)
        st.markdown(
            f"**{title}** — 通过率: **{stats.get('pass_rate',0):.1f}%** "
            f"| OK: {stats.get('ok_count',0)} (主料:{ok_m}, 替料:{ok_s}) "
            f"| NG: {stats.get('ng_count',0)} (差异:{ng_q}, 缺料:{ng_m}, BOM无:{ng_b})"
        )
    show_all = st.checkbox('显示全部', key=f'all_{title}')
    view = df if show_all else get_abnormal_results(df)
    if view.empty:
        st.success('🎉 全部通过，无异常！')
    else:
        def _hl(val):
            v = str(val)
            if v == JudgmentStatus.OK:
                return 'background-color:#C6EFCE;color:#006100'
            if v == JudgmentStatus.OK_WITH_SUB:
                return 'background-color:#DDEBF7;color:#1F4E79'
            if v == JudgmentStatus.NG_NOT_IN_BOM:
                return 'background-color:#FCE4D6;color:#C65911'
            if v.startswith('NG'):
                return 'background-color:#FFC7CE;color:#9C0006'
            return ''
        rc = next((c for c in ('判定结果', '结果') if c in view.columns), None)
        styled = view.style.applymap(_hl, subset=[rc]) if rc else view.style
        st.dataframe(styled, use_container_width=True, height=400, key=f'df_{title}')


if st.session_state.processed:
    st.markdown('---')
    st.markdown('### 📊 核对结果')

    fact_data = st.session_state.fact_lists_data
    bom_n = len(st.session_state.bom_items) if st.session_state.bom_items else 0

    sc = st.columns(2 + len(fact_data))
    with sc[0]:
        _stat_card(bom_n, 'BOM物料总数')

    for i, fd in enumerate(fact_data):
        with sc[1 + i]:
            sts = fd['stats']
            clr = 'pass-indicator' if sts.get('ng_count', 0) == 0 else 'fail-indicator'
            short_name = fd['file_name'][:12] + '...' if len(fd['file_name']) > 15 else fd['file_name']
            _stat_card(f"{sts['ok_count']}/{sts['total_items']}", short_name, clr)

    with sc[-1]:
        sub_n = sum(fd['stats'].get('substitute_used_count', 0) for fd in fact_data)
        _stat_card(sub_n, '含替料OK数')

    st.markdown('<br>', unsafe_allow_html=True)

    if fact_data:
        if len(fact_data) == 1:
            fd = fact_data[0]
            _result_table(fd['result_df'], fd['file_name'], fd['stats'])
        else:
            tab_names = [f"📋 {fd['file_name']}" for fd in fact_data]
            tabs = st.tabs(tab_names)
            for t, fd in zip(tabs, fact_data):
                with t:
                    _result_table(fd['result_df'], fd['file_name'], fd['stats'])


# ============================================================
# 导出
# ============================================================
if st.session_state.processed:
    st.markdown('---')
    st.markdown('### 💾 导出报告')
    try:
        buf = BytesIO()
        fact_data = st.session_state.fact_lists_data

        ls = [(fd['file_name'], fd['stats']) for fd in fact_data]
        summary = generate_summary(st.session_state.bom_items or [], ls,
                                    st.session_state.work_order, st.session_state.batch)

        sheets = [(fd['file_name'], fd['result_df']) for fd in fact_data]

        export_results_to_excel(buf, summary, sheets, st.session_state.bom_df,
                                 st.session_state.work_order, st.session_state.batch)
        buf.seek(0)

        wo = st.session_state.work_order or '无工单'
        fname = f"CKD核对报告_{wo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        ec1, ec2 = st.columns([1, 3])
        with ec1:
            st.download_button('📥 下载 Excel 报告', buf, fname,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               use_container_width=True, type='primary')
        with ec2:
            sheet_list = ', '.join([fd['file_name'] for fd in fact_data])
            st.caption(f'包含汇总、{sheet_list}、BOM数据等多 Sheet')
    except Exception as e:
        st.error(f'❌ 生成报告失败: {e}')


# ============================================================
# 页脚
# ============================================================
st.markdown('---')
st.markdown('<p style="text-align:center;color:#888;font-size:.75rem">'
            'CKD清单核对工具 v5.0 | 极致精简 · 智能映射 · 五种状态判定</p>',
            unsafe_allow_html=True)
