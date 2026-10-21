"""核对正式检验报告里的参数，与自己提供的样品参数表是否一致。

用法：
    python compare_report.py --sheet <参数表.docx> --report <报告.docx> [--report <另一份> ...]
                             [--rules <field-rules.tsv>] --out-dir <目录> [--docx <路径>]

适用场景：参数表交给检测机构后，对方据此出了 PEMS 排放报告、噪声报告，
需要确认报告里抄录的车辆参数与自己提供的参数表有没有出入。
与填表方向相反，本脚本只读，不改任何单据。

判定分五类：一致 / 仅写法不同 / 实质不同 / 报告留空 / 报告未列，
另有「本报告不涉及」——噪声专属字段出现在排放报告里、或反之，不算缺漏。
"""

import argparse
import re
import sys
from pathlib import Path

import docx_table as dt
import build_audit as ba
from docx_write import Doc

# 参数表分段 → 报告里的段落名，用于人工字段按同名兜底
SEG_MAP = {'整车参数': '整车参数', '发动机参数': '发动机', '变速箱参数': '变速箱',
           '底盘参数': '底盘', '污染物控制装置参数': '污染控制装置'}

# 核对某类报告时，取数方式为该值的字段属于「另一份报告才有」，不算缺漏
EXCLUSIVE = {'pm': '噪声', 'zs': '排放'}
KIND_NAME = {'pm': '排放报告（PEMS）', 'zs': '噪声报告', 'unknown': '检验报告'}

SAME, FORMAT, DIFF, BLANK, NOTFOUND, NA = '一致', '仅写法不同', '实质不同', '报告留空', '报告未列', '本报告不涉及'
ORDER = [DIFF, BLANK, NOTFOUND, FORMAT]


def detect_kind(path, tables):
    """从文件名与正文判断这是排放报告还是噪声报告。"""
    name = Path(path).name.upper()
    text = ''.join(''.join(''.join(r) for r in t) for t in tables)
    if 'PM6' in name or '实际道路车载法排放' in text:
        return 'pm'
    if 'ZS6' in name or '加速行驶车外噪声' in text:
        return 'zs'
    return 'unknown'


def loose(text):
    """宽松归一化：乘号统一、去空格、数字写法差异不敏感。"""
    s = dt.normalize_label(text).lower()
    for ch in ('×', 'x', '*', '＊'):
        s = s.replace(ch, '×')
    return s


def numbers(text):
    return re.findall(r'\d+(?:\.\d+)?', text or '')


def find_by_label(tables, label):
    """在报告所有表格里按标签找右侧的值，兼容「样品信息」这类两列信息页。"""
    target = dt.normalize_label(label)
    for rows in tables:
        for row in rows:
            for i, cell in enumerate(row):
                if dt.normalize_label(cell) != target:
                    continue
                for nxt in row[i + 1:]:
                    if nxt.strip():
                        return nxt.strip()
    return ''


def classify(s_val, r_val):
    s_blank = not s_val or dt.is_dash(s_val)
    r_blank = not r_val or dt.is_dash(r_val)
    if s_blank and r_blank:
        return SAME          # 两边都写「——」，未安装项一致
    if r_blank:
        return BLANK
    if s_blank:
        return DIFF          # 表内空、报告有值
    if dt.normalize_label(s_val) == dt.normalize_label(r_val) or loose(s_val) == loose(r_val):
        return SAME
    if numbers(s_val) and numbers(s_val) == numbers(r_val):
        return FORMAT
    a, b = loose(s_val), loose(r_val)
    if a and b and (a in b or b in a):
        return FORMAT
    return DIFF


def pick_value(report, tables, rule):
    """按 规则表映射 → 分段同名 → 全文同名 三级取报告值。"""
    sec, fld = rule['参数表分段'], rule['参数表字段']
    mode = rule['取值处理'] or 'value'
    if rule['报告分段']:
        raw = ba.fetch(report, rule['报告分段'], rule['报告字段'], mode)
        if raw not in (None, '', []):
            text = ba.as_text(raw).strip()
            if text:
                return text, f"规则表 → {rule['报告分段']}/{rule['报告字段']}"
    seg = SEG_MAP.get(sec, '')
    raw = report.get((seg, dt.normalize_label(fld)))
    if raw not in (None, '', []):
        text = ba.as_text(raw).strip()
        if text:
            return text, f'同名栏目 → {seg}/{fld}'
    hit = find_by_label(tables, fld)
    if hit:
        return hit, '报告其他页同名栏目'
    return '', ''


def report_scope(report, sheet):
    """判断报告是按本车出的，还是写基准车数据的视同报告。

    依据很简单：报告里的车辆型号名称与本车一致，才是按本车出具，
    差一个字的写法差异也算一致；不一致说明报告写的是基准车。
    """
    s_name = dt.normalize_label(sheet.get(('整车参数', '车辆型号名称')) or '')
    r_name = dt.normalize_label(report.get(('整车参数', '车辆型号名称')) or '')
    if not s_name or not r_name:
        return 'unknown'
    return 'self' if (s_name in r_name or r_name in s_name) else 'baseline'


def compare(sheet_path, report_path, rules):
    tables = dt.read_tables(report_path)
    kind = detect_kind(report_path, tables)
    report = ba.load_report(report_path)
    if report is None:
        return None
    sheet, _ = dt.parse_sectioned(sheet_path, ['车辆型号名称', '发动机型号'], ba.SHEET_SECTIONS)
    items = []
    for rule in rules:
        if rule['取数方式'] == '不填':
            continue
        sec, fld = rule['参数表分段'], rule['参数表字段']
        s_val = (sheet.get((dt.normalize_label(sec), dt.normalize_label(fld))) or '').strip()
        r_val, where = pick_value(report, tables, rule)
        if not r_val and rule['取数方式'] == EXCLUSIVE.get(kind):
            verdict = NA
        elif not r_val:
            verdict = NOTFOUND
        else:
            verdict = classify(s_val, r_val)
        items.append({'section': sec, 'field': fld, 'sheet': s_val,
                      'report': r_val, 'where': where, 'verdict': verdict})
    return {'path': Path(report_path), 'kind': kind, 'items': items,
            'scope': report_scope(report, sheet)}


def render(results, sheet_path, day):
    lines = ['# 报告与参数表对照', '',
             f'- 参数表：{Path(sheet_path).name}',
             f'- 对照日期：{day}', '']
    for res in results:
        name = res['path'].name
        items = res['items']
        scope = res['scope']
        bucket = {k: [i for i in items if i['verdict'] == k] for k in ORDER + [NA]}
        title = KIND_NAME[res['kind']] + ('，按本车出具' if scope == 'self' else
                                          '，写的是基准车数据' if scope == 'baseline' else '')
        lines += [f'## {name}（{title}）', '',
                  f'比对 {len(items)} 项：一致 {len([i for i in items if i["verdict"] == SAME])}'
                  f'、仅写法不同 {len(bucket[FORMAT])}、实质不同 {len(bucket[DIFF])}'
                  f'、报告留空 {len(bucket[BLANK])}、报告未列 {len(bucket[NOTFOUND])}'
                  f'、本报告不涉及 {len(bucket[NA])}', '']
        if scope == 'baseline':
            lines += ['这份报告里的车辆参数写的是基准车，与本车不同属正常，只作参考。', '']
        head = bucket[DIFF] or bucket[BLANK] or bucket[NOTFOUND]
        if bucket[DIFF]:
            title = ('### 与基准车的差异（正常）' if scope == 'baseline'
                     else '### 实质不同，需要处理')
            lines += [title, '', '| 分段 | 字段 | 参数表 | 报告 |', '|---|---|---|---|']
            for i in bucket[DIFF]:
                lines.append(f'| {i["section"]} | {i["field"]} | {i["sheet"] or "（空）"} | {i["report"]} |')
            lines.append('')
        if scope == 'self' and bucket[BLANK]:
            lines += ['### 报告留空', '', '| 分段 | 字段 | 参数表 | 报告 |', '|---|---|---|---|']
            for i in bucket[BLANK]:
                lines.append(f'| {i["section"]} | {i["field"]} | {i["sheet"] or "（空）"} | {i["report"]} |')
            lines.append('')
        if scope == 'self' and bucket[NOTFOUND]:
            lines += ['### 报告里没有这一项（不算错，但没核到）', '',
                      '| 分段 | 字段 | 参数表 |', '|---|---|---|']
            for i in bucket[NOTFOUND]:
                lines.append(f'| {i["section"]} | {i["field"]} | {i["sheet"] or "（空）"} |')
            lines += ['', '这类字段的依据通常是视同报告、合格证或实车，检验报告本身不列，'
                          '不能据此认为报告漏写。', '']
        if bucket[FORMAT] and scope != 'baseline':
            lines += ['### 内容相同，只是写法不同', '', '| 字段 | 参数表 | 报告 |', '|---|---|---|']
            for i in bucket[FORMAT]:
                lines.append(f'| {i["field"]} | {i["sheet"]} | {i["report"]} |')
            lines.append('')
        if not head and scope != 'baseline':
            lines += ['实质没有出入。', '']
        if bucket[NA]:
            names = '、'.join(i['field'] for i in bucket[NA])
            lines += [f'本报告不涉及（另一份报告专属）：{names}', '']
    return '\n'.join(lines)


def verdict_line(results):
    def cnt(rs, verdict):
        return sum(1 for r in rs for i in r['items'] if i['verdict'] == verdict)

    selfs = [r for r in results if r['scope'] == 'self']
    bases = [r for r in results if r['scope'] == 'baseline']
    diff = cnt(selfs, DIFF)
    if diff:
        return f'按本车出具的检验报告有 {diff} 处实质不同，必须先解决'
    parts = []
    if selfs:
        parts.append(f'检验报告实质一致（报告留空 {cnt(selfs, BLANK)} 项、'
                     f'未列 {cnt(selfs, NOTFOUND)} 项、写法不同 {cnt(selfs, FORMAT)} 项）')
    if bases:
        parts.append(f'{len(bases)} 份为基准车报告，差异属正常')
    return '；'.join(parts) or '完全一致'


def main(argv=None):
    ap = argparse.ArgumentParser(description='核对正式检验报告与参数表是否一致')
    ap.add_argument('--sheet', required=True)
    ap.add_argument('--report', action='append', required=True, help='可重复，给几份核几份')
    ap.add_argument('--rules', default=str(Path(__file__).resolve().parent.parent / 'references' / 'field-rules.tsv'))
    ap.add_argument('--out-dir', required=True)
    ap.add_argument('--docx', default='', help='可选，同时输出 Word 版')
    args = ap.parse_args(argv)

    for p in [args.sheet, args.rules] + args.report:
        if not Path(p).exists():
            print(f'找不到文件：{p}', file=sys.stderr)
            return 2
    rules = ba.load_rules(args.rules)
    try:
        results = []
        for path in args.report:
            res = compare(args.sheet, path, rules)
            if res is None:
                print(f'报告解析失败：{Path(path).name} 读不出表格结构，未生成对照。', file=sys.stderr)
                return 3
            results.append(res)
    except dt.DocxError as exc:
        print(f'读取失败：{exc}', file=sys.stderr)
        return 3

    from datetime import date
    day = date.today().strftime('%Y年%m月%d日')
    out = Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    md = render(results, args.sheet, day)
    md_path = out / '报告对照.md'
    md_path.write_text(md, encoding='utf-8')

    if args.docx:
        doc = Doc(landscape=True)
        doc.title('报告与参数表对照')
        doc.p(verdict_line(results), size=24, bold=True, after=200)
        for res in results:
            label = KIND_NAME[res['kind']] + (
                '·按本车出具' if res['scope'] == 'self' else
                '·写的是基准车数据' if res['scope'] == 'baseline' else '')
            doc.h1(f"{res['path'].name}（{label}）")
            wanted = [DIFF] if res['scope'] == 'baseline' else ORDER
            rows = [[i['section'], i['field'], i['sheet'] or '（空）', i['report'] or '—', i['verdict']]
                    for i in res['items'] if i['verdict'] in wanted]
            if rows:
                doc.table(['分段', '字段', '参数表', '报告', '判定'], rows,
                          widths=[1500, 2600, 2600, 2600, 1400], size=16)
            else:
                doc.p('实质没有出入。', size=18)
        doc.save(str(args.docx))

    print(f'对照表已生成：{md_path}')
    print(verdict_line(results))
    for res in results:
        cnt = {}
        for i in res['items']:
            cnt[i['verdict']] = cnt.get(i['verdict'], 0) + 1
        scope = {'self': '按本车出具', 'baseline': '基准车数据'}.get(res['scope'], '未知')
        print(f"  {res['path'].name}：{KIND_NAME[res['kind']]}，{scope}，"
              + '，'.join(f'{k} {v}' for k, v in cnt.items()))
    return 1 if any(i['verdict'] == DIFF for r in results if r['scope'] == 'self'
                    for i in r['items']) else 0


if __name__ == '__main__':
    raise SystemExit(main())
