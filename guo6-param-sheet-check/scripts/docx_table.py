"""读取与写入 .docx 表格的最小工具，仅依赖 Python 标准库。

设计约束：
- 只读取 word/document.xml 中的表格文本，不解析样式。
- 写入时只改动目标单元格的文字，其余内容与压缩包其余条目原样保留。
- 不执行任何网络、子进程或动态求值操作。
"""

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
DASH_SET = {'--', '---', '—', '——', '-', '－', '', '无'}


class DocxError(Exception):
    """文件不是可解析的 .docx：给人看得懂的提示，不抛栈。"""


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
    try:
        with zipfile.ZipFile(path) as zf:
            data = zf.read('word/document.xml')
    except zipfile.BadZipFile as exc:
        raise DocxError(f'{Path(path).name} 不是有效的 .docx，可能是改了扩展名的 .doc，请用 WPS／Word 另存为 .docx') from exc
    except KeyError as exc:
        raise DocxError(f'{Path(path).name} 里没有 word/document.xml，不是 Word 文档') from exc
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise DocxError(f'{Path(path).name} 的 document.xml 无法解析，文件可能已损坏') from exc


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
    """按文档中的出现顺序产出 (表序号, 行数据, 是否含速比)。"""
    for idx, tbl in enumerate(root.iter(W + 'tbl')):
        rows = []
        has_ratio = False
        for tr in tbl.findall(W + 'tr'):
            cells = [_cell_text(tc) for tc in tr.findall(W + 'tc')]
            rows.append(cells)
            if any(normalize_label(c) == '速比' for c in cells):
                has_ratio = True
        joined = normalize_label(''.join(''.join(r) for r in rows))
        if all(normalize_label(k) in joined for k in must_have):
            yield idx, rows, has_ratio


def parse_sectioned(path, must_have, sections):
    """解析「段名 | 标签 | 值 | 标签 | 值」型表格。

    返回 (pairs, positions)：
    - pairs: {(段名, 标签): 值}
    - positions: {(段名, 标签): (表序号, 行号, 单元格序号)}
    """
    root = _open_root(path)
    want = {normalize_label(s) for s in sections}
    for tbl_idx, rows, _ in _pick_table(root, must_have):
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
                    positions[key] = (tbl_idx, ri, ci + 1 + offset)
            if '速比' in labels:
                idx = labels.index('速比')
                rest = [c.strip() for c in body[idx + 1:] if c.strip() and not is_dash(c)]
                pairs[(section, '速比')] = ','.join(rest)
                positions[(section, '速比')] = (tbl_idx, ri, idx + 1 + offset)
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


def _locate(tables, tbl_idx, row_idx, cell_idx, label=''):
    """按表序号精确定位单元格，越界即报错，不做模糊匹配。"""
    where = f'第{tbl_idx + 1}张表 第{row_idx + 1}行 第{cell_idx + 1}列{("（" + label + "）") if label else ""}'
    if not 0 <= tbl_idx < len(tables):
        raise ValueError(f'{where}：文档只有 {len(tables)} 张表，位置越界')
    trs = tables[tbl_idx].findall(W + 'tr')
    if not 0 <= row_idx < len(trs):
        raise ValueError(f'{where}：该表只有 {len(trs)} 行，位置越界')
    tcs = trs[row_idx].findall(W + 'tc')
    if not 0 <= cell_idx < len(tcs):
        raise ValueError(f'{where}：该行只有 {len(tcs)} 个单元格，位置越界')
    return tcs[cell_idx], where


def write_cells(src_path, dst_path, updates):
    """把 updates 写入 dst_path 副本。

    updates: list[(表序号, 行号, 单元格序号, 值)]，表序号由 parse_sectioned 给出。
    先整体校验定位再写，越界直接报错；写完回读逐格核对，落错格子会抛错而不是静默成功。
    返回实际写入条数。
    """
    if str(src_path) == str(dst_path):
        raise ValueError('目标路径不能与源文件相同')
    with zipfile.ZipFile(src_path) as zf:
        names = zf.namelist()
        blobs = {n: zf.read(n) for n in names}
    root = ET.fromstring(blobs['word/document.xml'])
    tables = [t for t in root.iter(W + 'tbl')]

    targets = []
    for tbl_idx, row_idx, cell_idx, value in updates:
        tc, where = _locate(tables, tbl_idx, row_idx, cell_idx)
        targets.append((tc, value, where))

    for tc, value, _ in targets:
        _set_cell_text(tc, value)

    bad = [(where, _cell_text(tc), value) for tc, value, where in targets
           if normalize_label(_cell_text(tc)) != normalize_label(value)]
    if bad:
        detail = '；'.join(f'{w} 写入后读回「{got}」，应为「{want}」' for w, got, want in bad[:5])
        raise ValueError(f'回读校验失败，未生成副本：{detail}')

    blobs['word/document.xml'] = ET.tostring(root, encoding='UTF-8', xml_declaration=True)
    with zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED) as out:
        for name in names:
            out.writestr(name, blobs[name])
    return len(targets)
