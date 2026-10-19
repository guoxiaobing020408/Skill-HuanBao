"""核查已填好的参数表：回读比对 + 一致性校验。

用法：
    python verify_sheet.py --sheet <成品参数表> --audit <audit.json> [--out <核查报告路径>]

退出码：0 全部通过（可能有警告）；1 存在错误；2 输入文件缺失。
只读，不修改任何文件。
"""

import argparse
import json
import re
import sys
from pathlib import Path

import docx_table as dt

SHEET_SECTIONS = ['整车参数', '发动机参数', '变速箱参数', '底盘参数', '污染物控制装置参数']
ERROR = '错误'
WARN = '警告'
PASS = '通过'

AXLE_TIRES = {'4×2': 6, '4*2': 6, '6×4': 10, '6*4': 10, '4×4': 4, '4*4': 4, '8×4': 12, '8*4': 12}
POST_KEYS = {'DOC': '催化转化器(DOC)型号及厂家', 'SCR': '催化转化器(SCR)型号及厂家',
             'ASC': '催化转化器(ASC)型号及厂家', 'DPF': '颗粒捕集器(DPF)型号及厂家',
             'LNT': '催化转化器(LNT)型号及厂家'}


def tag(item):
    return f"{item['section']}/{item['field']}"


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))


def get(sheet, section, field):
    return sheet.get((dt.normalize_label(section), dt.normalize_label(field)), '')


def num(text):
    m = re.search(r'-?\d+(?:\.\d+)?', text or '')
    return float(m.group()) if m else None


def power(text):
    """从 114/3000 这类写法里取第一个数。"""
    return num((text or '').split('/')[0])


def check_readback(sheet, items, out):
    for it in items:
        if not it['value']:
            continue
        got = get(sheet, it['section'], it['field'])
        if dt.normalize_label(got) == dt.normalize_label(it['value']):
            out.append((PASS, tag(it), f"与报告取值一致：{it['value']}"))
        else:
            out.append((ERROR, tag(it), f"回读不一致：表内「{got or '空'}」，应为「{it['value']}」"))


def check_empty(sheet, items, out):
    for it in items:
        got = get(sheet, it['section'], it['field'])
        if got.strip():
            continue
        if it['status'] == '不填':
            continue
        if it['status'] == '人工':
            out.append((WARN, tag(it), '人工字段仍为空，等填写'))
        else:
            out.append((ERROR, tag(it), '应为自动取值却为空'))


def check_skip_fields(sheet, items, out):
    """标了不填的字段必须留空。"""
    hits = []
    for it in items:
        if it['status'] != '不填':
            continue
        got = get(sheet, it['section'], it['field'])
        if got.strip():
            hits.append(f"{it['field']}（表内「{got}」）")
    if hits:
        out.append((WARN, '不填字段', '本表不填却填了内容，建议清空：' + '、'.join(hits)))
    else:
        out.append((PASS, '不填字段', 'VIN号、发动机编号保持留空'))


def check_manual_vs_report(sheet, items, out):
    for it in items:
        if it['status'] != '人工' or not it['source'].startswith('参考：'):
            continue
        ref = it['source'][3:]
        got = get(sheet, it['section'], it['field'])
        if got.strip() and dt.normalize_label(got) != dt.normalize_label(ref):
            out.append((WARN, tag(it), f"表内「{got}」与报告参考「{ref}」不同，属人工判断项，请确认"))


def check_vehicle_copy(sheet, items, out):
    """参数表里的车辆型号不能直接等于报告的基准车型。"""
    name = get(sheet, '整车参数', '车辆型号名称')
    if not name:
        out.append((ERROR, '车辆型号名称', '为空'))
        return
    for it in items:
        if it['section'] == '整车参数' and it['field'] == '车辆型号名称' and it['source'].startswith('参考：'):
            if dt.normalize_label(name) == dt.normalize_label(it['source'][3:]):
                out.append((ERROR, '车辆型号名称', f'「{name}」是视同报告的基准车型，不是本车车型'))
                return
    out.append((PASS, '车辆型号名称', name))


def check_mass(sheet, out):
    base = num(get(sheet, '整车参数', '基准质量(kg)'))
    total = num(get(sheet, '整车参数', '最大总质量(kg)'))
    if base is None or total is None:
        out.append((WARN, '基准质量/最大总质量', '有值缺失或不是数字，无法比对'))
        return
    if base >= total:
        out.append((ERROR, '基准质量/最大总质量', f'基准质量 {base:g} 不小于最大总质量 {total:g}，不可能'))
    else:
        out.append((PASS, '基准质量/最大总质量', f'{base:g} < {total:g}'))
    if total > 12000 and 'N2' in get(sheet, '整车参数', '汽车分类'):
        out.append((ERROR, '汽车分类', f'最大总质量 {total:g} kg 超过 N2 上限 12000 kg，分类应改 N3'))
    elif total <= 12000 and 'N3' in get(sheet, '整车参数', '汽车分类'):
        out.append((ERROR, '汽车分类', f'最大总质量 {total:g} kg 未超 12000 kg，分类应改 N2'))
    else:
        out.append((PASS, '汽车分类', get(sheet, '整车参数', '汽车分类')))


def check_tires(sheet, out):
    drive = dt.normalize_label(get(sheet, '整车参数', '驱动方式'))
    count = num(get(sheet, '整车参数', '轮胎数量(个)'))
    expect = AXLE_TIRES.get(drive)
    if expect is None or count is None:
        out.append((WARN, '轮胎数量', f'驱动方式「{drive}」不在已知对照表内，无法校验'))
    elif int(count) != expect:
        out.append((ERROR, '轮胎数量', f'驱动方式 {drive} 应为 {expect} 个，表内 {int(count)} 个'))
    else:
        out.append((PASS, '轮胎数量', f'{drive} 对应 {expect} 个'))


def check_gears(sheet, out):
    gears = num(get(sheet, '变速箱参数', '档位数'))
    ratios = get(sheet, '变速箱参数', '速比')
    items = [x for x in re.split(r'[,，;；]', ratios) if x.strip()]
    if gears is None:
        out.append((WARN, '档位数', '不是数字，无法与速比个数比对'))
        return
    if len(items) != int(gears):
        out.append((ERROR, '速比', f'速比 {len(items)} 个，档位数 {int(gears)} 个，对不上'))
    else:
        out.append((PASS, '速比', f'{len(items)} 个速比对应 {int(gears)} 个档位'))


def check_power(sheet, out):
    rated = power(get(sheet, '发动机参数', '额定功率/转速(kW/r/min)'))
    net = power(get(sheet, '发动机参数', '最大净功率/转速(kW/r/min)'))
    if rated is None or net is None:
        out.append((WARN, '功率', '功率值格式无法解析'))
        return
    if net > rated:
        out.append((ERROR, '功率', f'最大净功率 {net:g} 大于额定功率 {rated:g}，不合常理'))
    else:
        out.append((PASS, '功率', f'净功率 {net:g} ≤ 额定功率 {rated:g}'))


def check_aftertreatment(sheet, out):
    types = get(sheet, '发动机参数', '后处理系统型式')
    mentioned = [k for k in POST_KEYS if k in types.upper()]
    if not mentioned:
        out.append((WARN, '后处理系统型式', f'「{types}」里认不出装置名称'))
        return
    for key in mentioned:
        value = get(sheet, '污染物控制装置参数', POST_KEYS[key])
        if not value.strip() or dt.is_dash(value):
            out.append((ERROR, POST_KEYS[key], f'后处理型式写了 {key}，该栏却为「{value or "空"}」'))
        else:
            out.append((PASS, POST_KEYS[key], value))
    for key, field in POST_KEYS.items():
        if key in mentioned:
            continue
        value = get(sheet, '污染物控制装置参数', field)
        if value.strip() and not dt.is_dash(value):
            out.append((WARN, field, f'该栏有型号，但后处理型式里没写 {key}'))


def check_pair_format(sheet, out):
    fields = [f for f in POST_KEYS.values() if f.startswith(('催化', '颗粒'))]
    fields += ['喷油泵型号及厂家', '共轨管型号及厂家', '喷油器型号及厂家', '增压器型号及厂家',
               'EGR型号及厂家', '曲轴箱气体再循环装置型号及厂家', 'EECU型号及厂家', 'OBD型号及厂家',
               '排气消声器型号及厂家', '排气管型号及厂家', '膨胀室型号及厂家', '空气滤清器型号及厂家',
               '进气消声器型号及厂家', '反应剂控制单元(DCU)型号及厂家', 'DPF控制单元型号及厂家']
    bad = []
    for field in fields:
        value = get(sheet, '污染物控制装置参数', field)
        if not value.strip() or dt.is_dash(value):
            continue
        if '/' not in value:
            bad.append(field)
    if bad:
        out.append((WARN, '型号/厂家格式', '这些字段缺厂家或没按「型号/厂家」写：' + '、'.join(bad)))
    else:
        out.append((PASS, '型号/厂家格式', '需要厂家的栏目都写了'))


def check_dash(sheet, out):
    bad = []
    for (section, field), value in sheet.items():
        if dt.is_dash(value) and dt.normalize_label(value) not in ('——', ''):
            bad.append(field)
    if bad:
        out.append((WARN, '未安装写法', '这些字段未统一成「——」：' + '、'.join(bad)))
    else:
        out.append((PASS, '未安装写法', '未安装项统一为「——」'))


def render(meta, results):
    errors = [r for r in results if r[0] == ERROR]
    warns = [r for r in results if r[0] == WARN]
    lines = [
        '# 参数表核查报告',
        '',
        f"- 核查对象：{meta['sheet']}",
        f"- 对照依据：{meta['audit']}",
        f"- 结论：{'不通过，有 ' + str(len(errors)) + ' 处错误' if errors else '通过'}",
        f"- 统计：错误 {len(errors)}、警告 {len(warns)}、通过 {len(results) - len(errors) - len(warns)}",
        '',
    ]
    for title, level in [('一、必须改（错误）', ERROR), ('二、请确认（警告）', WARN)]:
        lines += [f'## {title}', '']
        rows = [r for r in results if r[0] == level]
        if not rows:
            lines.append('无')
        else:
            lines.append('| 字段 | 说明 |')
            lines.append('|---|---|')
            for _, field, msg in rows:
                lines.append(f'| {field} | {msg} |')
        lines.append('')
    lines += ['## 三、通过的检查', '']
    passed = [r for r in results if r[0] == PASS]
    if not passed:
        lines.append('无')
    else:
        for _, field, msg in passed[:40]:
            lines.append(f'- {field}：{msg}')
        if len(passed) > 40:
            lines.append(f'- 另有 {len(passed) - 40} 项同类通过')
    lines.append('')
    return '\n'.join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description='核查已填好的样品参数表')
    ap.add_argument('--sheet', required=True)
    ap.add_argument('--audit', required=True)
    ap.add_argument('--out', help='核查报告输出路径，默认与 audit 同目录 verify.md')
    args = ap.parse_args(argv)

    if not Path(args.sheet).exists() or not Path(args.audit).exists():
        print('输入文件不存在', file=sys.stderr)
        return 2

    sheet, _ = dt.parse_sectioned(args.sheet, ['车辆型号名称', '发动机型号'], SHEET_SECTIONS)
    data = load_json(args.audit)
    items = data.get('items', [])

    results = []
    check_readback(sheet, items, results)
    check_empty(sheet, items, results)
    check_vehicle_copy(sheet, items, results)
    check_mass(sheet, results)
    check_tires(sheet, results)
    check_gears(sheet, results)
    check_power(sheet, results)
    check_aftertreatment(sheet, results)
    check_pair_format(sheet, results)
    check_dash(sheet, results)
    check_skip_fields(sheet, items, results)
    check_manual_vs_report(sheet, items, results)

    out = Path(args.out) if args.out else Path(args.audit).parent / 'verify.md'
    meta = {'sheet': Path(args.sheet).name, 'audit': Path(args.audit).name}
    out.write_text(render(meta, results), encoding='utf-8')

    errors = sum(1 for r in results if r[0] == ERROR)
    warns = sum(1 for r in results if r[0] == WARN)
    print(f'核查报告：{out}')
    print(f'错误 {errors}、警告 {warns}、通过 {len(results) - errors - warns}')
    return 1 if errors else 0


if __name__ == '__main__':
    raise SystemExit(main())
