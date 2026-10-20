"""生成《数据来源核对表》Word，逐字段标明数据取自哪个文件的哪一栏。

用法：
    python make_source_doc.py --audit <audit.json> [--sheet <已填好的参数表.docx>]
                              [--out <输出.docx>] [--vehicle <车型>] [--date <日期>]

--sheet 可选：给了就读出表里实际写入的值，逐条标注「一致 / 不一致」；
不给就用核对表的建议取值，只显示来源。
"""

import argparse
import json
import sys
from datetime import date
from pathlib import Path

import docx_table as dt
from docx_write import Doc

SHEET_SECTIONS = ['整车参数', '发动机参数', '变速箱参数', '底盘参数', '污染物控制装置参数']

CODE = {'pm': 'A', 'zs': 'B'}
ROLE = {'pm': '排放视同报告（PEMS）', 'zs': '噪声视同报告'}

STATUS_ORDER = ['自动', '自动·待核对', '人工', '不填', '冲突待定', '报告缺此项']


def load_sheet(path):
    """读参数表实际值；读不出来就报错，不拿核对表里的旧值冒充表内值。"""
    if not path:
        return None
    pairs, positions = dt.parse_sectioned(path, ['车辆型号名称', '发动机型号'], SHEET_SECTIONS)
    if not positions:
        raise dt.DocxError(f'{Path(path).name} 里找不到「段名/标签/值」结构的表格')
    return dict(pairs)


def actual(sheet, section, field):
    if sheet is None:
        return None
    return sheet.get((dt.normalize_label(section), dt.normalize_label(field)), '')


def src_label(item):
    kind = (item.get('src_kind') or '').strip()
    if not kind:
        return '—'
    has_pm, has_zs = 'pm' in kind, 'zs' in kind
    if has_pm and has_zs:
        return 'A+B 两份一致'
    if has_pm:
        return 'A 排放报告'
    if has_zs:
        return 'B 噪声报告'
    return '—'


def src_place(item):
    sec = (item.get('src_section') or '').strip()
    lab = (item.get('src_field') or '').strip()
    if not sec and not lab:
        return '—'
    return f'{sec} › {lab}' if sec else lab


def build_out(data, sheet, vehicle, day):
    meta = data['meta']
    items = data['items']
    d = Doc(landscape=True)

    d.title('重型柴油车样品参数表（国六环保）')
    d.p('数 据 来 源 核 对 表', size=28, bold=True, align='center', after=200)

    info = []
    if vehicle:
        info.append(('车 型', vehicle))
    info.append(('参数表', meta.get('sheet', '')))
    info.append(('核对日期', day))
    for k, v in info:
        d.p(f'{k}：{v}', size=21, after=80)

    # 一、文件清单
    d.h1('一、本次用到哪几个文件')
    rows = [
        ['A', meta.get('pm_file', ''), meta.get('pm_report', ''), ROLE['pm'], '自动取数的主来源'],
        ['B', meta.get('zs_file', ''), meta.get('zs_report', ''), ROLE['zs'],
         '速比/档位数/主减速器速比、进排气消声器等只认这份'],
    ]
    d.table(['代号', '文件名', '报告编号', '是什么', '作用'], rows,
            widths=[500, 3000, 1600, 2200, 3200], size=18,
            aligns=['center', 'left', 'left', 'left', 'left'])
    d.p('提示：两台文件名很像，认准代号 A / B 就不会拿错。报告中同名字段的值可能不同，'
        '下表「来源」一栏写明每个值到底抄的是哪一份。', size=18, color='7F7F7F')

    # 二、读法
    d.h1('二、怎么核对这张表')
    for line in [
        '1. 先看第三部分：这是系统自动写进参数表的每一个值，逐条对着原件看「来源」和「报告位置」两栏。',
        '2. 「来源」里的 A 是排放报告，B 是噪声报告；写着「A+B 两份一致」的，说明两份报告都有且相同。',
        '3. 「报告位置」写的是该值在报告里的栏目路径，例如「发动机 › 最大净功率/转速」= 报告「检验车辆基本参数」表中发动机段的那一格。',
        '4. 「与表内」一栏是回读比对结果：一致 = 参数表里确实写着这个值；不一致 = 参数表里是别的值，需要立刻查。',
        '5. 第四部分是系统没取数、必须你手填的字段，列在这里是为了让你知道哪些格子还是空的。',
    ]:
        d.p(line, size=19, after=60)

    # 三、自动取数
    auto = [i for i in items if i['status'] in ('自动', '自动·待核对')]
    d.h1(f'三、自动写进参数表的字段（共 {len(auto)} 项，逐条溯源）')
    rows = []
    bad = 0
    for n, it in enumerate(auto, 1):
        got = actual(sheet, it['section'], it['field'])
        if got is None:
            chk = '—'
        elif dt.normalize_label(got) == dt.normalize_label(it['value']):
            chk = '一致'
        elif not got.strip():
            chk = '表内为空'
            bad += 1
        else:
            chk = '不一致'
            bad += 1
        rows.append([str(n), it['section'], it['field'], it['value'],
                     src_label(it), src_place(it), it['status'], chk])
    d.table(['#', '参数表分段', '参数表字段', '写入的值', '来源', '报告位置', '判定', '与表内'],
            rows, widths=[430, 1150, 2500, 2700, 1050, 3000, 1000, 1000], size=17,
            aligns=['center', 'left', 'left', 'left', 'center', 'left', 'center', 'center'])
    if sheet is not None:
        msg = '回读结果：%d 项一致' % (len(auto) - bad) if not bad else \
              '回读结果：%d 项一致，%d 项需要复查（看「与表内」列为「不一致」或「表内为空」的行）' % (len(auto) - bad, bad)
        d.p(msg, size=19, bold=bool(bad), color='C00000' if bad else '375623')

    # 四、人工
    manual = [i for i in items if i['status'] == '人工']
    d.h1(f'四、没有自动取数、需要你填的字段（共 {len(manual)} 项）')
    rows = []
    for n, it in enumerate(manual, 1):
        got = actual(sheet, it['section'], it['field'])
        shown = got if got is not None else it['current']
        rows.append([str(n), it['section'], it['field'], (shown or '').strip() or '空',
                     it.get('basis') or '—', it['note'] or '—'])
    d.table(['#', '参数表分段', '参数表字段', '表内当前值', '依据文件', '怎么写'],
            rows, widths=[380, 1400, 2900, 2100, 1900, 3900], size=17,
            aligns=['center', 'left', 'left', 'left', 'left', 'left'])
    d.p('「依据文件」写明这一格该去哪份材料上抄：底盘合格证、整车合格证/铭牌、公告或实车。'
        '报告里同名字段是基准车数据，只能当参考，不能照抄。',
        size=18, color='7F7F7F')

    # 五、不一致
    conflict = [i for i in items if i['status'] == '冲突待定']
    d.h1(f'五、两份报告不一致、等你定的字段（共 {len(conflict)} 项）')
    if conflict:
        rows = []
        for it in conflict:
            cands = {c['kind']: c['value'] for c in it.get('src_cands', [])}
            rows.append([it['section'], it['field'],
                         cands.get('pm', '—'), cands.get('zs', '—')])
        d.table(['参数表分段', '参数表字段', 'A 排放报告', 'B 噪声报告'],
                rows, widths=[1600, 3200, 3600, 3600], size=17,
                aligns=['left', 'left', 'left', 'left'])
        d.p('系统不猜：没有规则依据的冲突一律留空，由你按公告或实车定，定了再写进参数表。',
            size=18, color='C00000')
    else:
        d.p('无。两份报告没有互斥取值。', size=19)

    # 六、缺项
    missing = [i for i in items if i['status'] == '报告缺此项']
    d.h1(f'六、两份报告里都查不到的字段（共 {len(missing)} 项）')
    if missing:
        rows = [[it['section'], it['field'], src_place(it), it['note'] or '需人工确认']
                for it in missing]
        d.table(['参数表分段', '参数表字段', '找过的位置', '说明'], rows,
                widths=[1600, 3000, 3400, 3400], size=17,
                aligns=['left', 'left', 'left', 'left'])
    else:
        d.p('无。每个需要取数的字段都在报告里找到了。', size=19)

    # 七、不填
    skip = [i for i in items if i['status'] == '不填']
    if skip:
        d.h1('七、本表不填、保持留空的字段')
        d.p('、'.join(i['field'] for i in skip) + '。', size=19)

    # 汇总
    d.h1('八、字段统计')
    counts = {s: sum(1 for i in items if i['status'] == s) for s in STATUS_ORDER}
    total = len(items)
    rows = [['自动取数（含待核对）', str(counts['自动'] + counts['自动·待核对']),
             '已写入，来源见第三部分'],
            ['其中待核对', str(counts['自动·待核对']), '发动机/变速箱段，值取自报告，建议对铭牌复核'],
            ['人工填写', str(counts['人工']), '见第四部分'],
            ['不填', str(counts['不填']), '保持留空'],
            ['冲突待定', str(counts['冲突待定']), '见第五部分'],
            ['报告缺此项', str(counts['报告缺此项']), '见第六部分'],
            ['合计', str(total), '规则表字段总数']]
    d.table(['类别', '数量', '说明'], rows, widths=[2600, 900, 6000], size=18,
            aligns=['left', 'center', 'left'])

    d.p('本表由 guo6-param-sheet-check 依据 references/field-rules.tsv 自动生成，'
        '不含任何人工推测值。', size=17, color='7F7F7F', after=0)
    return d


def main(argv=None):
    ap = argparse.ArgumentParser(description='生成数据来源核对表 Word')
    ap.add_argument('--audit', required=True)
    ap.add_argument('--sheet', default='')
    ap.add_argument('--out', default='')
    ap.add_argument('--vehicle', default='')
    ap.add_argument('--date', default='')
    args = ap.parse_args(argv)

    audit_path = Path(args.audit)
    if not audit_path.exists():
        print(f'找不到核对表：{audit_path}', file=sys.stderr)
        return 2
    data = json.loads(audit_path.read_text(encoding='utf-8'))

    try:
        sheet = load_sheet(args.sheet) if args.sheet else None
    except dt.DocxError as exc:
        print(f'读取参数表失败：{exc}', file=sys.stderr)
        return 3
    day = args.date or date.today().strftime('%Y年%m月%d日')
    out = Path(args.out) if args.out else audit_path.parent / '数据来源核对表.docx'

    doc = build_out(data, sheet, args.vehicle, day)
    doc.save(str(out))
    print(f'数据来源核对表已生成：{out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
