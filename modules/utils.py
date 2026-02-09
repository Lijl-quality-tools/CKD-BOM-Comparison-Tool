# -*- coding: utf-8 -*-
"""
通用工具函数模块 v3.0
包含：智能表头打分定位、正则提取替代料、文本清洗、安全算式计算

功能升级：
    - 智能表头打分逻辑（Smart Anchor）
    - 中文数字转换
    - 脏数据深度清洗（前缀单引号、不可见字符）
    - 安全算式计算（支持 100*9 格式）
"""

import re
from typing import List, Optional, Union, Any, Tuple
from .config import CHINESE_NUM_MAP, ParseMode


# ============================================================
# 字符串清洗函数
# ============================================================
def clean_part_number(value: Any) -> str:
    """
    清洗料号，统一转换为干净的字符串格式
    
    处理规则：
    1. 转换为字符串
    2. 去除首尾空格
    3. 去除隐藏的单引号（Excel常见的前缀单引号）
    4. 去除不可见字符和控制字符
    5. 去除浮点数的 .0 后缀
    
    Args:
        value: 原始料号值（可能是int, float, str等类型）
        
    Returns:
        清洗后的字符串料号
    """
    if value is None:
        return ''
    
    # 转换为字符串
    str_val = str(value)
    
    # 处理浮点数的 .0 后缀（如 123.0 -> 123）
    if isinstance(value, float):
        if value != value:  # NaN 检查
            return ''
        if value == int(value):
            str_val = str(int(value))
    elif str_val.endswith('.0'):
        str_val = str_val[:-2]
    
    # 去除首尾空格
    str_val = str_val.strip()
    
    # 去除隐藏的单引号（Excel常见）- 包括各种引号变体
    str_val = str_val.lstrip("'").lstrip("'").lstrip("'").lstrip("`")
    str_val = str_val.replace("'", '').replace("'", '').replace("'", '')
    
    # 去除非打印字符（保留数字、字母、常见分隔符）
    # 移除控制字符、零宽字符、BOM等
    str_val = re.sub(r'[\x00-\x1f\x7f-\x9f\u200b-\u200d\ufeff\u00a0]', '', str_val)
    
    # 去除首尾空格（再次清洗，防止移除不可见字符后残留）
    str_val = str_val.strip()
    
    return str_val


def clean_cell_value(value: Any) -> str:
    """
    清洗单元格值为干净的字符串
    
    Args:
        value: 原始单元格值
        
    Returns:
        清洗后的字符串
    """
    if value is None:
        return ''
    
    str_val = str(value).strip()
    
    # 去除不可见字符
    str_val = re.sub(r'[\x00-\x1f\x7f-\x9f\u200b-\u200d\ufeff]', '', str_val)
    
    return str_val.strip()


# ============================================================
# 替代料提取函数
# ============================================================
def extract_substitute_ids(substitute_text: Any) -> List[str]:
    """
    从替代状况文本中提取所有替代料号
    
    使用正则表达式 \\d+ 提取所有连续数字序列
    
    Args:
        substitute_text: 替代状况列的值，可能是：
            - 单个料号: "456"
            - 多个分隔料号: "456;789" 或 "456/789" 或 "456,789"
            - 包含描述: "替代料456和789"
            - 完整料号: "ABC-456-XYZ"
            
    Returns:
        替代料号列表（字符串格式）
    """
    if substitute_text is None:
        return []
    
    text = str(substitute_text).strip()
    
    # 空值检查
    if not text or text.lower() in ('nan', 'none', '无', '-', '', 'null'):
        return []
    
    # 使用正则提取所有连续数字序列
    pattern = r'\d+'
    matches = re.findall(pattern, text)
    
    # 去重并保持顺序
    seen = set()
    result = []
    for m in matches:
        if m not in seen and len(m) >= 3:  # 至少3位数字才算料号
            seen.add(m)
            result.append(m)
    
    return result


# ============================================================
# 安全算式计算
# ============================================================
def safe_eval_expression(expression: Any) -> Optional[float]:
    """
    安全计算数学表达式（仅支持基本四则运算）
    
    用于处理清单中的数量算式，如 "100*9"
    
    Args:
        expression: 数学表达式字符串或数值
        
    Returns:
        计算结果（浮点数），计算失败返回None
    """
    if expression is None:
        return None
    
    # 如果已经是数值类型，直接返回
    if isinstance(expression, (int, float)):
        if isinstance(expression, float) and (expression != expression):  # NaN检查
            return None
        return float(expression)
    
    expr_str = str(expression).strip()
    
    # 去除前缀单引号
    expr_str = expr_str.lstrip("'").lstrip("'").lstrip("'")
    
    # 空值检查
    if not expr_str or expr_str.lower() in ('nan', 'none', '', '-'):
        return None
    
    # 尝试直接转换为数值
    try:
        return float(expr_str)
    except ValueError:
        pass
    
    # 安全计算表达式
    # 只允许数字和基本运算符
    allowed_pattern = r'^[\d\s\+\-\*\/\.\(\)]+$'
    if not re.match(allowed_pattern, expr_str):
        return None
    
    try:
        # 使用eval计算，但限制命名空间，防止代码注入
        result = eval(expr_str, {"__builtins__": {}}, {})
        return float(result)
    except (SyntaxError, NameError, TypeError, ZeroDivisionError, ValueError):
        return None


# ============================================================
# 智能表头打分定位（Smart Anchor）
# ============================================================
def calculate_header_score(row: List[Any], keywords: List[str]) -> int:
    """
    计算某一行作为表头的得分（关键词命中数）
    
    Args:
        row: 行数据
        keywords: 要匹配的关键词列表
        
    Returns:
        得分（命中的关键词数量）
    """
    if not row:
        return 0
    
    # 将行数据转换为文本
    row_text = ' '.join([str(cell) if cell else '' for cell in row])
    row_text_lower = row_text.lower()
    
    score = 0
    matched_keywords = set()
    
    for keyword in keywords:
        kw_lower = keyword.lower()
        if kw_lower in row_text_lower and kw_lower not in matched_keywords:
            score += 1
            matched_keywords.add(kw_lower)
    
    return score


def smart_find_header_row(data: List[List[Any]], 
                          keywords: List[str], 
                          max_rows: int = 20,
                          min_score: int = 2) -> Tuple[Optional[int], int]:
    """
    智能定位表头行 - 基于打分逻辑
    
    扫描前N行，计算每行的关键词命中得分，得分最高的行判定为表头行
    
    Args:
        data: 二维数据列表
        keywords: 要匹配的关键词列表
        max_rows: 最大搜索行数
        min_score: 最低识别得分阈值
        
    Returns:
        (表头行索引, 得分)，未找到返回 (None, 0)
    """
    if not data or not keywords:
        return None, 0
    
    search_rows = min(len(data), max_rows)
    
    best_row_idx = None
    best_score = 0
    
    for row_idx in range(search_rows):
        row = data[row_idx]
        if not row:
            continue
        
        # 跳过明显的空行
        non_empty_cells = sum(1 for cell in row if cell is not None and str(cell).strip())
        if non_empty_cells < 2:
            continue
        
        score = calculate_header_score(row, keywords)
        
        if score > best_score:
            best_score = score
            best_row_idx = row_idx
    
    # 检查是否达到最低得分阈值
    if best_score >= min_score:
        return best_row_idx, best_score
    
    return None, 0


def find_column_by_keywords(header_row: List[Any], 
                            keywords: List[str]) -> Optional[int]:
    """
    在表头行中查找包含指定关键词的列
    
    Args:
        header_row: 表头行数据
        keywords: 要匹配的关键词列表（按优先级排序）
        
    Returns:
        列索引（0-based），未找到返回None
    """
    if not header_row or not keywords:
        return None
    
    for keyword in keywords:
        kw_lower = keyword.lower()
        for col_idx, cell in enumerate(header_row):
            cell_str = str(cell).lower() if cell else ''
            if kw_lower in cell_str:
                return col_idx
    
    return None


# ============================================================
# 中文数字转换
# ============================================================
def chinese_to_arabic(chinese_str: str) -> Optional[int]:
    """
    将中文数字转换为阿拉伯数字
    
    支持: 一、二、三...十、百（简单形式）
    
    Args:
        chinese_str: 中文数字字符串
        
    Returns:
        阿拉伯数字，转换失败返回None
    """
    if not chinese_str:
        return None
    
    chinese_str = chinese_str.strip()
    
    # 简单的单字符转换
    if len(chinese_str) == 1 and chinese_str in CHINESE_NUM_MAP:
        return CHINESE_NUM_MAP[chinese_str]
    
    # 处理 "十一" 到 "十九"
    if chinese_str.startswith('十'):
        if len(chinese_str) == 1:
            return 10
        rest = chinese_str[1:]
        if rest in CHINESE_NUM_MAP:
            return 10 + CHINESE_NUM_MAP[rest]
    
    # 处理 "二十" 到 "九十九"
    result = 0
    temp = 0
    
    for char in chinese_str:
        if char in CHINESE_NUM_MAP:
            num = CHINESE_NUM_MAP[char]
            if num == 10:
                if temp == 0:
                    temp = 1
                result += temp * 10
                temp = 0
            elif num == 100:
                if temp == 0:
                    temp = 1
                result += temp * 100
                temp = 0
            else:
                temp = num
        else:
            return None
    
    result += temp
    return result if result > 0 else None


def extract_box_number_from_text(text: str, 
                                  patterns: List[str]) -> Optional[str]:
    """
    从文本中提取箱号
    
    支持阿拉伯数字和中文数字的箱号标识
    
    Args:
        text: 待匹配的文本
        patterns: 正则模式列表
        
    Returns:
        提取到的箱号字符串（如 "1号箱"），未匹配返回None
    """
    if not text:
        return None
    
    text = str(text).strip()
    
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            box_num = match.group(1)
            
            # 如果是中文数字，转换为阿拉伯数字
            if re.match(r'^[一二三四五六七八九十百]+$', box_num):
                arabic_num = chinese_to_arabic(box_num)
                if arabic_num:
                    return f"{arabic_num}号箱"
            else:
                return f"{box_num}号箱"
    
    return None


# ============================================================
# 空行检测
# ============================================================
def is_empty_row(row: List[Any]) -> bool:
    """
    判断是否为空行（全空或全为空白字符）
    
    Args:
        row: 行数据
        
    Returns:
        是否为空行
    """
    if not row:
        return True
    
    for cell in row:
        if cell is not None:
            cell_str = str(cell).strip()
            if cell_str and cell_str.lower() not in ('nan', 'none', ''):
                return False
    
    return True


def filter_empty_rows(data: List[List[Any]], 
                       start_row: int = 0) -> List[Tuple[int, List[Any]]]:
    """
    过滤空行，返回非空行及其原始索引
    
    Args:
        data: 二维数据列表
        start_row: 起始行索引
        
    Returns:
        [(原始行索引, 行数据), ...] 的列表
    """
    result = []
    
    for idx, row in enumerate(data[start_row:], start=start_row):
        if not is_empty_row(row):
            result.append((idx, row))
    
    return result


# ============================================================
# 箱号格式化
# ============================================================
def normalize_box_number(box_value: Any) -> str:
    """
    标准化箱号格式
    
    Args:
        box_value: 原始箱号值
        
    Returns:
        标准化的箱号字符串
    """
    if box_value is None:
        return ''
    
    box_str = str(box_value).strip()
    
    # 处理浮点数
    if isinstance(box_value, float):
        if box_value != box_value:  # NaN
            return ''
        if box_value == int(box_value):
            box_str = str(int(box_value))
    elif box_str.endswith('.0'):
        box_str = box_str[:-2]
    
    # 去除不可见字符
    box_str = re.sub(r'[\x00-\x1f\x7f-\x9f\u200b-\u200d\ufeff]', '', box_str)
    
    return box_str.strip()


def format_number(value: Optional[float], decimals: int = 2) -> str:
    """
    格式化数字显示
    
    Args:
        value: 数值
        decimals: 小数位数
        
    Returns:
        格式化的字符串
    """
    if value is None:
        return '-'
    
    # 如果是整数，不显示小数
    if value == int(value):
        return str(int(value))
    
    return f"{value:.{decimals}f}"


def merge_box_numbers(box_list: List[str]) -> str:
    """
    合并箱号列表为显示字符串
    
    Args:
        box_list: 箱号列表
        
    Returns:
        合并后的字符串，如 "1号箱, 2号箱"
    """
    if not box_list:
        return '-'
    
    # 去重并排序
    unique_boxes = sorted(set(b for b in box_list if b), 
                          key=lambda x: (len(x), x))
    
    if not unique_boxes:
        return '-'
    
    return ', '.join(unique_boxes)


# ============================================================
# 料号有效性检查
# ============================================================
def is_valid_part_number(part_no: str) -> bool:
    """
    检查料号是否有效（非空且包含有效字符）
    
    Args:
        part_no: 料号字符串
        
    Returns:
        是否有效
    """
    if not part_no:
        return False
    
    cleaned = clean_part_number(part_no)
    if not cleaned:
        return False
    
    # 至少包含一个数字或字母
    return bool(re.search(r'[a-zA-Z0-9]', cleaned))
