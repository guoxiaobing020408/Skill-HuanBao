"""把核对结果写进参数表副本。

用法：
    python fill_sheet.py --sheet <参数表原件.docx> --audit <audit.json> --out <副本.docx> [--apply]

默认只打印将要写入的内容，不产生任何文件；加 --apply 才生成副本。
只写「自动」「自动·待核对」两类字段，人工字段一律不动；原件始终只读。
"""

import argparse
import json
import sys
from pathlib import Path

import docx_table as dt

WRITABLE = {'自动', '自动·待核对'}
STATUS_MANUAL = '人工'


def load_items(path):
    data = json.loads(Path(path).read_text(encoding='utf-8'))
    return data.get('items', [])


def main(argv=None):
    ap = argparse.ArgumentParser(description='把核对结果写入参数表副本')
    ap.add_argument('--sheet', required=True, help='参数表原件，只读')
    ap.add_argument('--audit', required=True, help='build_audit.py 生成的 audit.json')
    ap.add_argument('--out', required=True, help='副本路径，不能与原件相同')
    ap.add_argument('--apply', action='store_true', help='真正写盘；不加时只预览')
    ap.add_argument('--force', action='store_true', help='副本已存在时覆盖')
    ap.add_argument('--set', action='append', default=[], metavar='字段=值',
                    help='补充人工字段，可重复；支持 字段=值 或 分段/字段=值')
    ap.add_argument('--clear-manual', action='store_true',
                    help='把人工字段清空，避免旧车数据残留；默认不动')
    args = ap.parse_args(argv)

    if Path(args.out).resolve() == Path(args.sheet).resolve():
        print('拒绝执行：输出路径与原件相同，原件不允许被覆盖。', file=sys.stderr)
        return 2

    items = load_items(args.audit)
    index = {}
    for it in items:
        index[(it['section'], it['field'])] = it
        index.setdefault(it['field'], it)

    overrides = {}
    for raw in args.set:
        if '=' not in raw:
            print(f'--set 参数格式不对：{raw}', file=sys.stderr)
            return 4
        key, value = raw.split('=', 1)
        key = key.strip().replace('/', '\u0000')
        target = None
        if '\u0000' in key:
            section, field = key.split('\u0000', 1)
            target = index.get((section.strip(), field.strip()))
        else:
            target = index.get(key)
        if target is None:
            print(f'--set 找不到对应字段：{raw}', file=sys.stderr)
            return 4
        overrides[(target['section'], target['field'])] = value.strip()

    updates, skipped = [], []
    for it in items:
        key = (it['section'], it['field'])
        if key in overrides:
            updates.append((int(it['row']), int(it['col']), overrides[key], it))
            continue
        if it['status'] in WRITABLE and it['value'] and it['row'] != '' and it['col'] != '':
            updates.append((int(it['row']), int(it['col']), it['value'], it))
        elif it['status'] in WRITABLE and not it['value']:
            skipped.append(it)
        elif args.clear_manual and it['status'] == STATUS_MANUAL and it['row'] != '':
            updates.append((int(it['row']), int(it['col']), '', it))

    print(f'可写入 {len(updates)} 项，跳过 {len(skipped)} 项（判定为自动但没有取值）')
    for row, col, value, it in updates:
        print(f'  第{row + 1}行 第{col + 1}列  {it["section"]}/{it["field"]} = {value}')

    if not args.apply:
        print('\n预览模式，未写入任何文件。确认无误后加 --apply 生成副本：')
        print(f'  --out "{args.out}"')
        return 0

    if Path(args.out).exists() and not args.force:
        print(f'副本已存在：{args.out}\n如确需覆盖请加 --force。', file=sys.stderr)
        return 3

    payload = [(r, c, v) for r, c, v, _ in updates]
    written = dt.write_cells(args.sheet, args.out, payload)
    print(f'已生成副本：{args.out}（写入 {written} 个单元格）')
    print('原件未改动。人工填写项请按核对表第二节补齐。')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
