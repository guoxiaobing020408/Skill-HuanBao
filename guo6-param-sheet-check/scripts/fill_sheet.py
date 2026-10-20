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
    by_key, by_field = {}, {}
    for it in items:
        by_key[(dt.normalize_label(it['section']), dt.normalize_label(it['field']))] = it
        by_field.setdefault(dt.normalize_label(it['field']), []).append(it)

    overrides = {}
    for raw in args.set:
        if '=' not in raw:
            print(f'--set 参数格式不对：{raw}', file=sys.stderr)
            return 4
        key, value = raw.split('=', 1)
        key = key.strip()
        if '/' in key:
            section, field = key.split('/', 1)
            target = by_key.get((dt.normalize_label(section), dt.normalize_label(field)))
            if target is None:
                print(f'--set 找不到字段：{raw}（写成 参数表分段/参数表字段）', file=sys.stderr)
                return 4
        else:
            hits = by_field.get(dt.normalize_label(key), [])
            if not hits:
                print(f'--set 找不到字段：{raw}', file=sys.stderr)
                return 4
            if len(hits) > 1:
                names = '、'.join(f"{h['section']}/{h['field']}" for h in hits)
                print(f'--set 字段「{key}」在多个分段重名，请写成 参数表分段/参数表字段，候选：{names}',
                      file=sys.stderr)
                return 4
            target = hits[0]
        if target['status'] == '不填':
            print(f"--set 拒绝执行：「{target['section']}/{target['field']}」属本表不填项，不能赋值",
                  file=sys.stderr)
            return 4
        overrides[(target['section'], target['field'])] = value.strip()

    updates, skipped, unlocated = [], [], []
    for it in items:
        key = (it['section'], it['field'])
        has_pos = all(it.get(k) not in ('', None) for k in ('tbl', 'row', 'col'))
        if key in overrides:
            if not has_pos:
                unlocated.append(it)
                continue
            updates.append((int(it['tbl']), int(it['row']), int(it['col']), overrides[key], it))
            continue
        if not has_pos:
            unlocated.append(it)
            continue
        if it['status'] in WRITABLE and it['value']:
            updates.append((int(it['tbl']), int(it['row']), int(it['col']), it['value'], it))
        elif it['status'] in WRITABLE and not it['value']:
            skipped.append(it)
        elif args.clear_manual and it['status'] == STATUS_MANUAL:
            updates.append((int(it['tbl']), int(it['row']), int(it['col']), '', it))

    print(f'可写入 {len(updates)} 项，跳过 {len(skipped)} 项（判定为自动但没有取值）')
    for tbl, row, col, value, it in updates:
        print(f'  表{tbl + 1} 第{row + 1}行 第{col + 1}列  {it["section"]}/{it["field"]} = {value}')

    if unlocated:
        print(f'\n有 {len(unlocated)} 项在参数表里找不到位置，未写入：', file=sys.stderr)
        for it in unlocated[:10]:
            print(f"  {it['section']}/{it['field']}", file=sys.stderr)
        if 'tbl' not in unlocated[0]:
            print('核对结果缺少表序号，请用当前版本的 build_audit.py 重新生成 audit.json。', file=sys.stderr)
            return 5

    if not args.apply:
        print('\n预览模式，未写入任何文件。确认无误后加 --apply 生成副本：')
        print(f'  --out "{args.out}"')
        return 0

    if Path(args.out).exists() and not args.force:
        print(f'副本已存在：{args.out}\n如确需覆盖请加 --force。', file=sys.stderr)
        return 3

    payload = [(t, r, c, v) for t, r, c, v, _ in updates]
    try:
        written = dt.write_cells(args.sheet, args.out, payload)
    except dt.DocxError as exc:
        print(f'写表失败：{exc}', file=sys.stderr)
        return 3
    except ValueError as exc:
        print(f'写表失败：{exc}', file=sys.stderr)
        return 6
    print(f'已生成副本：{args.out}（写入并回读校验 {written} 个单元格）')
    print('原件未改动。人工填写项请按核对表第二节补齐。')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
