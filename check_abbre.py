#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
检查 Word 文档中的英文缩写规范：
  - 首次出现时应给出中英文全称并用括号括起缩写（如“卷积神经网络（CNN）”）
  - 之后可直接使用缩写

功能模式：
  1. 检查模式 (--check，默认)   ：只输出未给出全称的缩写
  2. 列表模式 (--list)          ：列出所有首次出现的缩写及其位置
  3. 导出CSV模式 (--export-csv) ：将缩略词表导出为 CSV 文件，包含可能全称

使用方法：
  python check_abbr.py <docx文件> [选项]

示例：
  python check_abbr.py thesis.docx                      # 默认检查
  python check_abbr.py thesis.docx --list               # 列出所有缩写
  python check_abbr.py thesis.docx --check --min-len 3  # 检查长度≥3的缩写
  python check_abbr.py thesis.docx --export-csv out.csv # 导出CSV
  python check_abbr.py thesis.docx --lines 5-10 --ignore AI CNN
  python check_abbr.py thesis.docx --ignore-file ignore.txt
"""

import sys
import re
import argparse
import csv
from docx import Document


def read_paragraphs(docx_path):
    """从 .docx 读取所有段落文本，返回列表（保留原始顺序）。"""
    doc = Document(docx_path)
    return [para.text for para in doc.paragraphs]


def parse_line_range(range_str, total_paragraphs):
    """
    解析行范围字符串，返回段落索引列表（0-based）。
    支持格式："5", "5-10", "5-", "-8", ""（全部）
    """
    if not range_str:
        return list(range(total_paragraphs))

    range_str = range_str.strip()
    if '-' in range_str:
        parts = range_str.split('-')
        if len(parts) == 2:
            start_str, end_str = parts
            start = int(start_str) - 1 if start_str.strip() else 0
            end = int(end_str) if end_str.strip() else total_paragraphs
            start = max(0, start)
            end = min(total_paragraphs, end)
            if start >= total_paragraphs or end <= 0:
                return []
            return list(range(start, end))
        else:
            return list(range(total_paragraphs))
    else:
        try:
            idx = int(range_str) - 1
            if 0 <= idx < total_paragraphs:
                return [idx]
            else:
                return []
        except ValueError:
            return list(range(total_paragraphs))


def build_selected_text_and_mapping(paragraphs, selected_indices):
    """
    根据选中的段落索引构建合并文本（段落间用空格连接），
    并返回字符位置到原始段落索引的映射（列表，长度等于合并文本长度）。
    映射中每个元素为对应的段落索引（原始索引），用于定位。
    """
    selected_texts = []
    mapping = []
    for idx in selected_indices:
        para_text = paragraphs[idx]
        selected_texts.append(para_text)
        mapping.extend([idx] * len(para_text))
        if idx != selected_indices[-1]:
            selected_texts.append(' ')
            mapping.append(idx)  # 空格归属于前一段
    full_text = ''.join(selected_texts)
    return full_text, mapping


def find_bracket_pairs(text):
    """找出所有成对括号的区间（左闭右开），包括英文和中文括号。"""
    pairs = []
    stack_cn = []  # 中文左括号位置
    stack_en = []  # 英文左括号位置
    for i, ch in enumerate(text):
        if ch == '（':
            stack_cn.append(i)
        elif ch == '）' and stack_cn:
            start = stack_cn.pop()
            pairs.append((start, i))
        elif ch == '(':
            stack_en.append(i)
        elif ch == ')' and stack_en:
            start = stack_en.pop()
            pairs.append((start, i))
    return pairs


def is_in_brackets(pos, bracket_pairs):
    """判断位置是否在某个括号内（不含括号本身）。"""
    for start, end in bracket_pairs:
        if start < pos < end:
            return True
    return False


def find_bracket_containing_pos(pos, bracket_pairs):
    """
    返回包含 pos 的括号对 (start, end)（左闭右开），如果 pos 在括号内。
    注意：如果 pos 正好位于括号字符上，则不视为在内部，返回 None。
    """
    for start, end in bracket_pairs:
        if start < pos < end:
            return (start, end)
    return None


def find_abbreviation_matches(text, min_len=2):
    """匹配大写字母组成的缩写（长度 >= min_len）。"""
    pattern = re.compile(r'\b[A-Z]{' + str(min_len) + r',}\b')
    return list(pattern.finditer(text))


def get_first_occurrences_in_text(text, min_len=2):
    """返回字典 {缩写: 首次出现位置}，位置为在 text 中的字符索引。"""
    matches = find_abbreviation_matches(text, min_len)
    first_occ = {}
    for m in matches:
        abbr = m.group()
        if abbr not in first_occ:
            first_occ[abbr] = m.start()
    return first_occ


def load_ignore_list_from_file(file_path):
    """从文本文件读取忽略列表，每行一个缩写，返回集合。"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        return {line.strip() for line in lines if line.strip()}
    except Exception as e:
        print(f"警告：读取忽略文件失败 {file_path}: {e}")
        return set()


def get_context(text, pos, window=30):
    """获取位置前后的上下文片段。"""
    start = max(0, pos - window)
    end = min(len(text), pos + window)
    ctx = text[start:end].replace('\n', ' ')
    return ctx


def find_paragraph_and_offset(mapping, pos):
    """
    根据字符位置在映射中查找所属的段落索引（原始）和在该段落内的偏移。
    返回 (para_idx, offset)，如果 pos 无效则返回 (-1, -1)。
    """
    if pos < len(mapping):
        para_idx = mapping[pos]
        # 计算在该段落内的偏移：需要重新计算该段落在合并文本中的起始位置
        # 由于映射是逐字符的，我们可以向前搜索直到段落索引变化
        # 简单方法：从 pos 开始向前找到第一个不同的段落索引，得到该段落起始位置
        start_pos = pos
        while start_pos > 0 and mapping[start_pos - 1] == para_idx:
            start_pos -= 1
        offset = pos - start_pos
        return para_idx, offset
    return -1, -1


def main():
    parser = argparse.ArgumentParser(
        description='检查 Word 文档中英文缩写首次出现是否给出了全称（括号括起）',
        epilog='''使用示例：
  python check_abbr.py thesis.docx                      # 默认检查模式
  python check_abbr.py thesis.docx --list               # 列出所有首次出现的缩写
  python check_abbr.py thesis.docx --check --min-len 3  # 检查长度≥3的缩写
  python check_abbr.py thesis.docx --export-csv out.csv # 导出缩略词表CSV
  python check_abbr.py thesis.docx --lines 5-10 --ignore AI CNN
  python check_abbr.py thesis.docx --ignore-file ignore.txt'''
    )
    parser.add_argument('docx', help='Word 文档路径（.docx）')
    parser.add_argument('--list', action='store_true',
                        help='列出所有首次出现的缩写（不检查括号）')
    parser.add_argument('--check', action='store_true',
                        help='检查并仅输出未带全称的缩写（默认行为）')
    parser.add_argument('--export-csv', type=str, metavar='FILE',
                        help='导出缩略词表到 CSV 文件（包含可能全称）')
    parser.add_argument('--min-len', type=int, default=2,
                        help='最小缩写长度（默认2）')
    parser.add_argument('--ignore', nargs='*', default=[],
                        help='忽略的缩写列表（空格分隔），如 --ignore AI CNN')
    parser.add_argument('--ignore-file', action='append', default=[],
                        help='指定忽略列表文件（可多次使用），每行一个缩写')
    parser.add_argument('--lines', type=str, default='',
                        help='指定检查的段落范围（1-based），如 5-10, 3, 5-, -8')
    args = parser.parse_args()

    # 模式判断：如果指定了 --export-csv，则进入导出模式，忽略 --list 和 --check
    mode = 'export' if args.export_csv else ('list' if args.list else 'check')

    # 读取所有段落
    try:
        all_paragraphs = read_paragraphs(args.docx)
    except Exception as e:
        print(f"读取文档失败: {e}")
        return 1

    total = len(all_paragraphs)
    if total == 0:
        print("文档无段落内容。")
        return 0

    # 解析行范围
    selected_indices = parse_line_range(args.lines, total)
    if not selected_indices:
        print(f"指定的段落范围无效或超出文档段落总数（共 {total} 段）。")
        return 1

    # 构建选定段落的合并文本和位置映射
    text, mapping = build_selected_text_and_mapping(all_paragraphs, selected_indices)
    if not text.strip():
        print("选定段落内容为空。")
        return 0

    # 加载忽略列表
    ignore_set = set(args.ignore)
    for file_path in args.ignore_file:
        ignore_set.update(load_ignore_list_from_file(file_path))

    # 获取首次出现（在选定文本中）
    first_occ = get_first_occurrences_in_text(text, args.min_len)
    # 过滤忽略项
    for ign in ignore_set:
        first_occ.pop(ign, None)

    if not first_occ:
        print(f"在选定段落中未发现任何大写缩写（长度≥{args.min_len}）。")
        return 0

    # 获取括号对
    bracket_pairs = find_bracket_pairs(text)

    # ---------- 导出 CSV 模式 ----------
    if mode == 'export':
        csv_file = args.export_csv
        try:
            with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['Abbreviation', 'Paragraph', 'Position', 'Context', 'Possible_Full_Name'])
                sorted_items = sorted(first_occ.items(), key=lambda x: x[1])
                for abbr, pos in sorted_items:
                    para_idx, offset = find_paragraph_and_offset(mapping, pos)
                    para_num = para_idx + 1 if para_idx >= 0 else '?'
                    ctx = get_context(text, pos, window=30)
                    full_name = ''
                    bracket = find_bracket_containing_pos(pos, bracket_pairs)
                    if bracket:
                        start, end = bracket
                        inside = text[start+1:end]
                        full_name = inside.strip()
                    writer.writerow([abbr, para_num, offset, ctx, full_name])
            print(f"缩略词表已导出至: {csv_file} (共 {len(first_occ)} 个缩写)")
        except Exception as e:
            print(f"导出 CSV 失败: {e}")
            return 1
        return 0

    # ---------- 列表模式 ----------
    if mode == 'list':
        print("在选定段落中首次出现的缩写（按出现位置排序）：")
        sorted_items = sorted(first_occ.items(), key=lambda x: x[1])
        for abbr, pos in sorted_items:
            para_idx, offset = find_paragraph_and_offset(mapping, pos)
            para_num = para_idx + 1 if para_idx >= 0 else '?'
            ctx = get_context(text, pos, window=20)
            print(f"  {abbr} (段落 {para_num}, 偏移 {offset}) 周围: ...{ctx}...")
        return 0

    # ---------- 检查模式 ----------
    issues = {}
    for abbr, pos in first_occ.items():
        if not is_in_brackets(pos, bracket_pairs):
            issues[abbr] = pos

    if issues:
        print("以下缩写首次出现时未用括号括起，请补充中英文全称（例如：全称（缩写））：")
        sorted_issues = sorted(issues.items(), key=lambda x: x[1])
        for abbr, pos in sorted_issues:
            para_idx, offset = find_paragraph_and_offset(mapping, pos)
            para_num = para_idx + 1 if para_idx >= 0 else '?'
            ctx = get_context(text, pos, window=30)
            print(f"  - {abbr} (段落 {para_num}, 偏移 {offset}) 周围: ...{ctx}...")
        return 1
    else:
        print("在选定段落中，所有缩写首次出现均位于括号内，检查通过。")
        return 0


if __name__ == '__main__':
    sys.exit(main())