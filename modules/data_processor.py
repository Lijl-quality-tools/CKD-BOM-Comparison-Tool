# -*- coding: utf-8 -*-
"""
核心业务逻辑模块 v5.0 — 归一化版

职责：
    compare_bom_and_list  — BOM ↔ 清单双向比对（五种状态判定）
    generate_summary      — 生成汇总 DataFrame
    export_results_to_excel — 多 Sheet Excel 导出
    validate_data / get_abnormal_results / get_ok_results — 辅助

接口变化（相对 v4）：
    - compare_bom_and_list  的 list_type 改为 list_label（任意字符串标签）
    - generate_summary      接收 list_stats: [(label, stats_dict), ...]
    - export_results_to_excel 接收 list_sheets: [(sheet_name, DataFrame), ...]
"""

from typing import List, Dict, Tuple, Optional, Set
from dataclasses import dataclass
import pandas as pd

from .file_reader import BOMItem, ListItem
from .utils import clean_part_number, merge_box_numbers, format_number
from .config import JudgmentStatus


# ============================================================
# 比对结果数据结构
# ============================================================
@dataclass
class CompareResult:
    work_order: str
    part_id: str
    part_name: str
    bom_quantity: float
    actual_quantity: float
    difference: float
    status: str
    box_sources: List[str]
    matched_substitutes: List[str]
    remark: str

    @property
    def is_pass(self) -> bool:
        return self.status.startswith('OK')

    @property
    def is_ng(self) -> bool:
        return self.status.startswith('NG')


@dataclass
class MatchResult:
    compare_result: CompareResult
    matched_part_ids: Set[str]


# ============================================================
# 索引构建
# ============================================================
def build_part_lookup(list_items: List[ListItem]) -> Dict[str, List[ListItem]]:
    lookup: Dict[str, List[ListItem]] = {}
    for item in list_items:
        pid = clean_part_number(item.part_id)
        if pid:
            lookup.setdefault(pid, []).append(item)
    return lookup


# ============================================================
# 单项匹配（替代料深度融合）
# ============================================================
def match_bom_item(
    bom_item: BOMItem,
    part_lookup: Dict[str, List[ListItem]],
    work_order: str = '',
) -> MatchResult:
    main_part = clean_part_number(bom_item.main_part_id)

    substitutes = [clean_part_number(s) for s in bom_item.substitute_ids if s]
    substitutes = [s for s in substitutes if s and s != main_part]

    all_parts = [main_part] + substitutes

    matched_items: List[ListItem] = []
    matched_sub_ids: Set[str] = set()
    matched_list_ids: Set[str] = set()

    for pid in all_parts:
        if pid in part_lookup:
            matched_items.extend(part_lookup[pid])
            matched_list_ids.add(pid)
            if pid != main_part:
                matched_sub_ids.add(pid)

    actual = sum(i.quantity for i in matched_items)
    boxes = [i.box_number for i in matched_items if i.box_number]
    diff = actual - bom_item.quantity

    remarks: List[str] = []

    if not matched_items:
        status = JudgmentStatus.NG_MISSING
        remarks.append('清单中未找到该料号及其替代料')
    elif abs(diff) < 0.001:
        if matched_sub_ids:
            status = JudgmentStatus.OK_WITH_SUB
            remarks.append(f"使用替代料: {', '.join(sorted(matched_sub_ids))}")
        else:
            status = JudgmentStatus.OK
    else:
        status = JudgmentStatus.NG_QTY_DIFF
        remarks.append(f"{'超量 +' if diff > 0 else '欠量 '}{format_number(diff)}")
        if matched_sub_ids:
            remarks.append(f"含替代料: {', '.join(sorted(matched_sub_ids))}")

    return MatchResult(
        compare_result=CompareResult(
            work_order=work_order, part_id=main_part, part_name=bom_item.name,
            bom_quantity=bom_item.quantity, actual_quantity=actual, difference=diff,
            status=status, box_sources=boxes, matched_substitutes=list(matched_sub_ids),
            remark='; '.join(remarks) if remarks else '',
        ),
        matched_part_ids=matched_list_ids,
    )


# ============================================================
# 反向补漏
# ============================================================
def find_unmatched_list_items(
    list_items: List[ListItem],
    bom_all_parts: Set[str],
    work_order: str = '',
) -> List[CompareResult]:
    unmatched: Dict[str, List[ListItem]] = {}
    for item in list_items:
        pid = clean_part_number(item.part_id)
        if pid and pid not in bom_all_parts:
            unmatched.setdefault(pid, []).append(item)

    results: List[CompareResult] = []
    for pid, items in unmatched.items():
        total = sum(i.quantity for i in items)
        boxes = [i.box_number for i in items if i.box_number]
        results.append(CompareResult(
            work_order=work_order, part_id=pid, part_name='',
            bom_quantity=0, actual_quantity=total, difference=total,
            status=JudgmentStatus.NG_NOT_IN_BOM, box_sources=boxes,
            matched_substitutes=[], remark='疑似技术变更或异常混料，请核实',
        ))
    return results


# ============================================================
# 核心比对（双向）
# ============================================================
def compare_bom_and_list(
    bom_items: List[BOMItem],
    list_items: List[ListItem],
    list_label: str = '清单',
    work_order: str = '',
) -> Tuple[pd.DataFrame, Dict]:
    """
    返回 (结果 DataFrame, 统计 dict)。
    list_label 用于在 stats['list_type'] 中标记。
    """
    lookup = build_part_lookup(list_items)

    bom_all: Set[str] = set()
    for b in bom_items:
        mp = clean_part_number(b.main_part_id)
        if mp:
            bom_all.add(mp)
        for s in b.substitute_ids:
            sc = clean_part_number(s)
            if sc:
                bom_all.add(sc)

    bom_results: List[CompareResult] = []
    for b in bom_items:
        mr = match_bom_item(b, lookup, work_order)
        bom_results.append(mr.compare_result)

    extra = find_unmatched_list_items(list_items, bom_all, work_order)
    all_res = bom_results + extra

    total = len(all_res)
    ok = sum(1 for r in all_res if r.is_pass)
    ng = sum(1 for r in all_res if r.is_ng)

    stats = {
        'total_items': total,
        'bom_items_count': len(bom_items),
        'ok_count': ok,
        'ng_count': ng,
        'pass_rate': (ok / total * 100) if total else 0,
        'ok_main_only': sum(1 for r in all_res if r.status == JudgmentStatus.OK),
        'ok_with_substitute': sum(1 for r in all_res if r.status == JudgmentStatus.OK_WITH_SUB),
        'substitute_used_count': sum(1 for r in all_res if r.status == JudgmentStatus.OK_WITH_SUB),
        'ng_qty_difference': sum(1 for r in all_res if r.status == JudgmentStatus.NG_QTY_DIFF),
        'ng_missing': sum(1 for r in all_res if r.status == JudgmentStatus.NG_MISSING),
        'ng_not_in_bom': sum(1 for r in all_res if r.status == JudgmentStatus.NG_NOT_IN_BOM),
        'list_type': list_label,
    }

    rows = [{
        '工单号': r.work_order or '-', '料号': r.part_id, '名称': r.part_name or '-',
        'BOM数量': format_number(r.bom_quantity), '清单实收': format_number(r.actual_quantity),
        '差异': format_number(r.difference), '判定结果': r.status,
        '箱号溯源': merge_box_numbers(r.box_sources), '备注': r.remark,
    } for r in all_res]

    df = pd.DataFrame(rows)
    col_order = ['工单号', '料号', '名称', 'BOM数量', '清单实收', '差异', '判定结果', '箱号溯源', '备注']
    if not df.empty:
        df = df[col_order]
    return df, stats


# ============================================================
# 筛选
# ============================================================
def get_abnormal_results(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    for c in ('判定结果', '结果'):
        if c in df.columns:
            return df[df[c].str.contains('NG', case=False, na=False)].reset_index(drop=True)
    return df


def get_ok_results(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    for c in ('判定结果', '结果'):
        if c in df.columns:
            return df[df[c].str.startswith('OK', na=False)].reset_index(drop=True)
    return df


# ============================================================
# 汇总（泛化）
# ============================================================
def generate_summary(
    bom_items: List[BOMItem],
    list_stats: Optional[List[Tuple[str, Dict]]] = None,
    work_order: str = '',
    batch: str = '',
) -> pd.DataFrame:
    """
    list_stats: [(清单文件名, stats_dict), ...]
    """
    data = [
        {'项目': '工单号', '值': work_order or '-'},
        {'项目': '批量', '值': batch or '-'},
        {'项目': 'BOM物料总数', '值': len(bom_items)},
    ]

    def _add(stats: Dict, prefix: str):
        data.extend([
            {'项目': f'{prefix}-核对总数', '值': stats.get('total_items', 0)},
            {'项目': f'{prefix}-OK数量', '值': stats.get('ok_count', 0)},
            {'项目': f'{prefix}-OK(仅主料)', '值': stats.get('ok_main_only', 0)},
            {'项目': f'{prefix}-OK(含替料)', '值': stats.get('ok_with_substitute', 0)},
            {'项目': f'{prefix}-NG数量', '值': stats.get('ng_count', 0)},
            {'项目': f'{prefix}-NG(数量差异)', '值': stats.get('ng_qty_difference', 0)},
            {'项目': f'{prefix}-NG(缺料)', '值': stats.get('ng_missing', 0)},
            {'项目': f'{prefix}-NG(BOM无)', '值': stats.get('ng_not_in_bom', 0)},
            {'项目': f'{prefix}-通过率', '值': f"{stats.get('pass_rate', 0):.1f}%"},
        ])

    if list_stats:
        for label, st_dict in list_stats:
            _add(st_dict, label)

    return pd.DataFrame(data)


# ============================================================
# Excel 导出（泛化）
# ============================================================
def export_results_to_excel(
    output_path,
    summary_df: pd.DataFrame,
    list_sheets: Optional[List[Tuple[str, pd.DataFrame]]] = None,
    bom_df: Optional[pd.DataFrame] = None,
    work_order: str = '',
    batch: str = '',
):
    """
    list_sheets: [(清单文件名, result_df), ...]
    """
    from datetime import datetime

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        wb = writer.book

        hdr_fmt = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
                                  'border': 1, 'align': 'center', 'valign': 'vcenter'})
        ok_fmt = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
        ok_sub_fmt = wb.add_format({'bg_color': '#DDEBF7', 'font_color': '#1F4E79', 'border': 1})
        ng_fmt = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})
        ng_bom_fmt = wb.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911', 'border': 1})
        cell_fmt = wb.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
        title_fmt = wb.add_format({'bold': True, 'font_size': 12, 'font_color': '#2F5496'})
        label_fmt = wb.add_format({'bold': True, 'font_color': '#404040'})
        val_fmt = wb.add_format({'font_color': '#1F4E79'})

        # ---- 汇总 ----
        summary_df.to_excel(writer, sheet_name='汇总', index=False, startrow=1)
        ws = writer.sheets['汇总']
        ws.write(0, 0, 'CKD清单核对报告 - 汇总', wb.add_format({'bold': True, 'font_size': 14}))
        ws.set_column('A:A', 25)
        ws.set_column('B:B', 20)

        # ---- 清单结果 Sheet（泛化循环） ----
        def _write_sheet(df: pd.DataFrame, sheet_name: str):
            if df is None or df.empty:
                return
            df_ex = df.drop(columns=['工单号'], errors='ignore')
            data_row = 4
            df_ex.to_excel(writer, sheet_name=sheet_name, index=False, startrow=data_row)
            ws2 = writer.sheets[sheet_name]
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ws2.write(0, 0, f'📋 {sheet_name}', title_fmt)
            ws2.write(1, 0, '工单号:', label_fmt)
            ws2.write(1, 1, work_order or '-', val_fmt)
            ws2.write(1, 3, '批量:', label_fmt)
            ws2.write(1, 4, batch or '-', val_fmt)
            ws2.write(2, 0, '导出时间:', label_fmt)
            ws2.write(2, 1, now, val_fmt)
            for ci, cn in enumerate(df_ex.columns):
                ws2.write(data_row, ci, cn, hdr_fmt)
            res_ci = -1
            for cn in ('判定结果', '结果'):
                if cn in df_ex.columns:
                    res_ci = df_ex.columns.get_loc(cn)
                    break
            for ri in range(len(df_ex)):
                for ci in range(len(df_ex.columns)):
                    v = df_ex.iloc[ri, ci]
                    if res_ci >= 0:
                        rv = str(df_ex.iloc[ri, res_ci])
                        if rv == JudgmentStatus.OK:
                            f = ok_fmt
                        elif rv == JudgmentStatus.OK_WITH_SUB:
                            f = ok_sub_fmt
                        elif rv == JudgmentStatus.NG_NOT_IN_BOM:
                            f = ng_bom_fmt
                        elif rv.startswith('NG'):
                            f = ng_fmt
                        else:
                            f = cell_fmt
                    else:
                        f = cell_fmt
                    ws2.write(data_row + 1 + ri, ci, v, f)
            for ci, cn in enumerate(df_ex.columns):
                ml = max(len(str(cn)),
                         df_ex.iloc[:, ci].astype(str).str.len().max() if len(df_ex) else 0)
                ws2.set_column(ci, ci, min(ml + 2, 50))

        if list_sheets:
            for sn, sdf in list_sheets:
                _write_sheet(sdf, sn)

        # ---- BOM 数据 ----
        if bom_df is not None and not bom_df.empty:
            bom_df.to_excel(writer, sheet_name='BOM数据', index=False)
            ws3 = writer.sheets['BOM数据']
            for ci, cn in enumerate(bom_df.columns):
                ws3.write(0, ci, cn, hdr_fmt)
                ml = max(len(str(cn)),
                         bom_df.iloc[:, ci].astype(str).str.len().max() if len(bom_df) else 0)
                ws3.set_column(ci, ci, min(ml + 2, 50))


# ============================================================
# 数据验证
# ============================================================
def validate_data(
    bom_items: List[BOMItem],
    list_items: List[ListItem],
    list_name: str = '清单',
) -> List[str]:
    warnings: List[str] = []
    if not bom_items:
        warnings.append('⚠️ BOM数据为空，请检查文件')
    if not list_items:
        warnings.append(f'⚠️ {list_name}数据为空，请检查文件')
    pids = [clean_part_number(i.main_part_id) for i in bom_items]
    dups = [p for p in set(pids) if pids.count(p) > 1]
    if dups:
        warnings.append(f"⚠️ BOM重复料号: {', '.join(dups[:5])}{'...' if len(dups) > 5 else ''}")
    z_bom = sum(1 for i in bom_items if i.quantity == 0)
    if z_bom:
        warnings.append(f'⚠️ BOM中有 {z_bom} 项数量为0')
    z_list = sum(1 for i in list_items if i.quantity == 0)
    if z_list:
        warnings.append(f'⚠️ {list_name}中有 {z_list} 项数量为0')
    return warnings
