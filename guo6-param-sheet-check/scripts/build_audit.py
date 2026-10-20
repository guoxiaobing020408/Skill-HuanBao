"""按规则表生成本车参数核对表。

用法：
    python build_audit.py --sheet <参数表.docx> --pm <排放视同报告.docx> --zs <噪声视同报告.docx>
                          --rules <field-rules.tsv> --out-dir <输出目录>

产物：
    audit.md    给人看的核对表
    audit.json  给 fill_sheet.py 用的结构化结果

只读源文件，不写入任何原始目录。
"""

import argparse
import csv
import json
import re
import sys
from pathlib import Path

import docx_table as dt

SHEET_SECTIONS = ['整车参数', '发动机参数', '变速箱参数', '底盘参数', '污染物控制装置参数']
REPORT_SECTIONS = ['整车参数', '发动机', '变速箱', '底盘']

STATUS_AUTO = '自动'
STATUS_CHECK = '自动·待核对'
STATUS_MANUAL = '人工'
STATUS_CONFLICT = '冲突待定'
STATUS_MISSING = '报告缺此项'
STATUS_SKIP = '不填'

ENGINE_SECTIONS = {'发动机', '变速箱'}


def load_rules(path):
    rows = []
    with open(path, encoding='utf-8', newline='') as fh:
        for row in csv.DictReader(fh, delimiter='\t'):
            rows.append(row)
    return rows


def load_report(path):
    pairs, _ = dt.parse_sectioned(path, ['车辆型号名称', '发动机型号'], REPORT_SECTIONS)
    data = dict(pairs)
    for key, val in dt.parse_contam(path).items():
        data.setdefault(key, val)
    for key, val in dt.parse_fuel(path).items():
        data[key] = val
    return data


def report_id(path):
    rows = dt.read_tables(path)
    text = ''.join(''.join(''.join(r) for r in t) for t in rows)
    m = re.search(r'报告编号[:：]\s*([A-Z0-9\-_]+)', text)
    return m.group(1) if m else Path(path).stem


def as_text(value):
    """统一取值为参数表要写的字符串。"""
    if isinstance(value, tuple):
        model, maker = value
        model = model or ''
        maker = maker or ''
        if dt.is_dash(model):
            return '——'
        if dt.is_dash(maker):
            return model.strip()
        return f"{model.strip()}/{maker.strip()}"
    if value is None:
        return ''
    return dt.dash_normalize(value)


def fetch(data, ref_section, ref_label, mode):
    sec = dt.normalize_label(ref_section)
    lab = dt.normalize_label(ref_label)
    if mode == 'fuel_maker':
        return data.get((sec, '生产厂'))
    if mode == 'fuel_spec':
        return data.get((sec, '规格'))
    return data.get((sec, lab))


def main(argv=None):
    ap = argparse.ArgumentParser(description='生成国六环保样品参数核对表')
    ap.add_argument('--sheet', required=True)
    ap.add_argument('--pm', required=True)
    ap.add_argument('--zs', required=True)
    ap.add_argument('--rules', required=True)
    ap.add_argument('--out-dir', required=True)
    args = ap.parse_args(argv)

    for p in (args.sheet, args.pm, args.zs, args.rules):
        if not Path(p).exists():
            print(f'找不到文件：{p}', file=sys.stderr)
            return 2

    rules = load_rules(args.rules)
    sheet, positions = dt.parse_sectioned(args.sheet, ['车辆型号名称', '发动机型号'], SHEET_SECTIONS)
    pm = load_report(args.pm)
    zs = load_report(args.zs)

    items = []
    for rule in rules:
        section = rule['参数表分段']
        field = rule['参数表字段']
        source = rule['取数方式']
        mode = rule['取值处理'] or 'value'
        note = rule['说明']
        key = (dt.normalize_label(section), dt.normalize_label(field))
        row, col = positions.get(key, ('', ''))
        current = sheet.get(key, '')
        value, used, status = '', '', STATUS_MANUAL
        cands = []
        texts = []

        if source == '不填':
            status = STATUS_SKIP
            used = '本表不填'
        elif source != '人工':
            if source in ('排放', '两款', '自动'):
                v = fetch(pm, rule['报告分段'], rule['报告字段'], mode)
                if v not in (None, '', []):
                    cands.append(('pm', '排放报告', v))
            if source in ('噪声', '两款', '自动'):
                v = fetch(zs, rule['报告分段'], rule['报告字段'], mode)
                if v not in (None, '', []):
                    cands.append(('zs', '噪声报告', v))
            texts = [(k, n, as_text(v)) for k, n, v in cands]
            texts = [t for t in texts if t[2]]
            if not texts:
                status = STATUS_MISSING
            elif len({t[2] for t in texts}) == 1:
                value = texts[0][2]
                used = texts[0][1]
                if len(texts) == 2:
                    used = '两款报告一致'
                status = STATUS_CHECK if dt.normalize_label(rule['报告分段']) in ENGINE_SECTIONS else STATUS_AUTO
            else:
                status = STATUS_CONFLICT
                used = ' / '.join(f'{n}:{t}' for _, n, t in texts)
                value = ''
        else:
            ref = fetch(pm, rule['报告分段'], rule['报告字段'], mode) if rule['报告分段'] else None
            if ref not in (None, '', []):
                used = f"参考：{as_text(ref)}"

        cmp_note = ''
        if value:
            same = dt.normalize_label(value) == dt.normalize_label(current)
            cmp_note = '与表内一致' if same else f'与表内不同（表内 {current or "空"}）'

        items.append({
            'section': section,
            'field': field,
            'value': value,
            'source': used,
            'status': status,
            'note': note,
            'current': current,
            'compare': cmp_note,
            'row': row,
            'col': col,
            # 溯源信息：给「数据来源核对表」用
            'src_kind': ''.join(t[0] for t in texts),
            'src_section': rule['报告分段'],
            'src_field': rule['报告字段'],
            'src_cands': [{'kind': k, 'file': n, 'value': as_text(v)} for k, n, v in texts],
        })

    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    meta = {
        'sheet': Path(args.sheet).name,
        'pm_report': report_id(args.pm),
        'zs_report': report_id(args.zs),
        'sheet_file': Path(args.sheet).name,
        'pm_file': Path(args.pm).name,
        'zs_file': Path(args.zs).name,
        'total': len(items),
    }
    (out / 'audit.json').write_text(
        json.dumps({'meta': meta, 'items': items}, ensure_ascii=False, indent=2), encoding='utf-8')
    (out / 'audit.md').write_text(render_md(meta, items), encoding='utf-8')
    print(f'核对表已生成：{out / "audit.md"}')
    print(f'共 {len(items)} 项，自动 {count(items, STATUS_AUTO)} 项，待核对 {count(items, STATUS_CHECK)} 项，'
          f'人工 {count(items, STATUS_MANUAL)} 项，不填 {count(items, STATUS_SKIP)} 项，'
          f'冲突 {count(items, STATUS_CONFLICT)} 项，缺项 {count(items, STATUS_MISSING)} 项')
    return 0


def count(items, status):
    return sum(1 for i in items if i['status'] == status)


def render_md(meta, items):
    lines = [
        '# 样品参数表核对表',
        '',
        f"- 参数表：{meta['sheet']}",
        f"- 排放视同报告（PEMS）：{meta['pm_report']}",
        f"- 噪声视同报告：{meta['zs_report']}",
        f"- 字段总数：{meta['total']}",
        '',
        '## 一、逐字段核对',
        '',
        '| 分段 | 字段 | 建议取值 | 来源 | 判定 | 表内现值比对 | 说明 |',
        '|---|---|---|---|---|---|---|',
    ]
    for it in items:
        value = it['value'] or '（待填）'
        lines.append('| {section} | {field} | {value} | {source} | {status} | {compare} | {note} |'.format(
            section=it['section'], field=it['field'], value=value,
            source=it['source'] or '—', status=it['status'],
            compare=it['compare'] or '—', note=it['note'] or ''))

    lines += ['', '## 二、待你填写的字段', '']
    manual = [i for i in items if i['status'] == STATUS_MANUAL]
    if manual:
        lines.append('| 分段 | 字段 | 表内现值 | 怎么写 |')
        lines.append('|---|---|---|---|')
        for it in manual:
            lines.append(f"| {it['section']} | {it['field']} | {it['current'] or '空'} | {it['note'] or '—'} |")
    else:
        lines.append('无')
    skipped = [i for i in items if i['status'] == STATUS_SKIP]
    if skipped:
        lines.append('')
        lines.append('本表不填、留空即可：' + '、'.join(i['field'] for i in skipped) + '。')

    lines += ['', '## 三、两份报告不一致，等你定', '']
    conflicts = [i for i in items if i['status'] == STATUS_CONFLICT]
    if conflicts:
        lines.append('| 分段 | 字段 | 两份报告各自的值 |')
        lines.append('|---|---|---|')
        for it in conflicts:
            lines.append(f"| {it['section']} | {it['field']} | {it['source']} |")
    else:
        lines.append('无')

    lines += ['', '## 四、报告里找不到的字段', '']
    missing = [i for i in items if i['status'] == STATUS_MISSING]
    if missing:
        for it in missing:
            lines.append(f"- {it['section']} / {it['field']}：{it['note'] or '请人工确认'}")
    else:
        lines.append('无')

    lines += ['', '## 五、与表内现值的差异', '']
    diff = [i for i in items if it_has_diff(i)]
    if diff:
        lines.append('| 分段 | 字段 | 建议取值 | 表内现值 |')
        lines.append('|---|---|---|---|')
        for it in diff:
            lines.append(f"| {it['section']} | {it['field']} | {it['value']} | {it['current'] or '空'} |")
    else:
        lines.append('无')
    lines.append('')
    return '\n'.join(lines)


def it_has_diff(item):
    return bool(item['value']) and dt.normalize_label(item['value']) != dt.normalize_label(item['current'])


if __name__ == '__main__':
    raise SystemExit(main())
