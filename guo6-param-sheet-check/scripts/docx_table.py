"""读取与写入 .docx 表格的最小工具，仅依赖 Python 标准库。

设计约束：
- 只读取 word/document.xml 中的表格文本，不解析样式。
- 写入时只改动目标单元格的文字，其余内容与压缩包其余条目原样保留。
- 不执行任何网络、子进程或动态求值操作。
"""

import re
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
DASH_SET = {'--', '---', '—', '——', '-', '－', '', '无'}


def normalize_label(text):
    """统一标签写法，便于跨文档匹配。"""
    s = (text or '').strip()
    for ch in (' ', '\u3000', '\n', '\r', '\t'):
        s = s.replace(ch, '')
    s = s.replace('（', '(').replace('）', ')')
    s = s.replace('•', '·').replace('‧', '·')
    return s


def is_dash(value):
    return normalize_label(value) in DASH_SET


def dash_normalize(value):
    """报告里的未安装写法统一成参数表要求的破折号。"""
    return '——' if is_dash(value) else value.strip()


def _cell_text(tc):
    parts = []
    for p in tc.iter(W + 'p'):
        text = ''.join(t.text or '' for t in p.iter(W + 't'))
        if text.strip():
            parts.append(text.strip())
    return '\n'.join(parts).strip()


def _open_root(path):
    with zipfile.ZipFile(path) as zf:
        return ET.fromstring(zf.read('word/document.xml'))


def read_tables(path):
    """返回文档中所有表格：list[list[list[str]]]。"""
    root = _open_root(path)
    result = []
    for tbl in root.iter(W + 'tbl'):
        rows = []
        for tr in tbl.findall(W + 'tr'):
            rows.append([_cell_text(tc) for tc in tr.findall(W + 'tc')])
        if rows:
            result.append(rows)
    return result


def _pick_table(root, must_have):
    for tbl in root.iter(W + 'tbl'):
        rows = []
        has_ratio = False
        for tr in tbl.findall(W + 'tr'):
            cells = [_cell_text(tc) for tc in tr.findall(W + 'tc')]
            rows.append(cells)
            if any(normalize_label(c) == '速比' for c in cells):
                has_ratio = True
        joined = normalize_label(''.join(''.join(r) for r in rows))
        if all(normalize_label(k) in joined for k in must_have):
            yield rows, has_ratio


def parse_sectioned(path, must_have, sections):
    """解析「段名 | 标签 | 值 | 标签 | 值」型表格。

    返回 (pairs, positions)：
    - pairs: {(段名, 标签): 值}
    - positions: {(段名, 标签): (行号, 单元格序号)}
    """
    root = _open_root(path)
    want = {normalize_label(s) for s in sections}
    for rows, _ in _pick_table(root, must_have):
        pairs, positions = {}, {}
        section = ''
        for ri, cells in enumerate(rows):
            cells = list(cells)
            head = normalize_label(cells[0]) if cells else ''
            offset = 0
            if head in want:
                section = head
                offset = 1
            elif head == '' and len(cells) >= 3:
                offset = 1
            body = cells[offset:]
            if not body:
                continue
            labels = [normalize_label(c) for c in body]
            limit = labels.index('速比') if '速比' in labels else len(body) - 1
            for ci in range(0, min(limit, len(body) - 1), 2):
                label = labels[ci]
                if not label:
                    continue
                value = body[ci + 1].strip()
                key = (section, label)
                if key not in pairs:
                    pairs[key] = value
                    positions[key] = (ri, ci + 1 + offset)
            if '速比' in labels:
                idx = labels.index('速比')
                rest = [c.strip() for c in body[idx + 1:] if c.strip() and not is_dash(c)]
                pairs[(section, '速比')] = ','.join(rest)
                positions[(section, '速比')] = (ri, idx + 1 + offset)
        if pairs:
            return pairs, positions
    return {}, {}


def parse_contam(path, header_label='催化转化器(DOC)'):
    """解析「装置 | 型号 | 生产厂」三列表格。"""
    root = _open_root(path)
    for tbl in root.iter(W + 'tbl'):
        rows = [[_cell_text(tc) for tc in tr.findall(W + 'tc')] for tr in tbl.findall(W + 'tr')]
        joined = ''.join(''.join(r) for r in rows)
        if header_label not in normalize_label(joined):
            continue
        pairs = {}
        for cells in rows:
            if len(cells) < 2:
                continue
            label = normalize_label(cells[0])
            if not label:
                continue
            maker = cells[2].strip() if len(cells) > 2 else ''
            pairs[('污染控制装置', label)] = (cells[1].strip(), maker)
        if pairs:
            return pairs
    return {}


def parse_fuel(path):
    """解析「柴油 / 润滑油 / 反应剂」三行燃料表。"""
    root = _open_root(path)
    for tbl in root.iter(W + 'tbl'):
        rows = [[_cell_text(tc) for tc in tr.findall(W + 'tc')] for tr in tbl.findall(W + 'tr')]
        joined = ''.join(''.join(r) for r in rows)
        if '反应剂' not in joined or '柴油' not in joined:
            continue
        pairs = {}
        for cells in rows:
            if len(cells) < 3:
                continue
            name = normalize_label(cells[0])
            if name in ('柴油', '润滑油', '反应剂'):
                pairs[(name, '生产厂')] = cells[1].strip()
                pairs[(name, '规格')] = cells[2].strip() if len(cells) > 2 else ''
        if pairs:
            return pairs
    return {}


def _set_cell_text(tc, value):
    """把单元格文字替换为 value，保留第一个 run 的字体格式。"""
    paragraphs = tc.findall(W + 'p')
    if not paragraphs:
        p = ET.SubElement(tc, W + 'p')
        paragraphs = [p]
    for extra in paragraphs[1:]:
        tc.remove(extra)
    p = paragraphs[0]
    runs = p.findall(W + 'r')
    if runs:
        first = runs[0]
        for extra in runs[1:]:
            p.remove(extra)
    else:
        first = ET.SubElement(p, W + 'r')
    for t in list(first.findall(W + 't')):
        first.remove(t)
    for t in list(p.findall(W + 't')):
        p.remove(t)
    t = ET.SubElement(first, W + 't')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = value


def write_cells(src_path, dst_path, updates):
    """把 updates 写入 dst_path 副本。

    updates: list[(行号, 单元格序号, 值)]
    返回实际写入条数。
    """
    if str(src_path) == str(dst_path):
        raise ValueError('目标路径不能与源文件相同')
    with zipfile.ZipFile(src_path) as zf:
        names = zf.namelist()
        blobs = {n: zf.read(n) for n in names}
    root = ET.fromstring(blobs['word/document.xml'])
    tables = [t for t in root.iter(W + 'tbl')]
    written = 0
    for row_idx, cell_idx, value in updates:
        for tbl in tables:
            trs = tbl.findall(W + 'tr')
            if row_idx >= len(trs):
                continue
            tcs = trs[row_idx].findall(W + 'tc')
            if cell_idx >= len(tcs):
                continue
            _set_cell_text(tcs[cell_idx], value)
            written += 1
            break
    blobs['word/document.xml'] = ET.tostring(root, encoding='UTF-8', xml_declaration=True)
    with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as out:
        for name in names:
            out.writestr(name, blobs[name])
    return written
