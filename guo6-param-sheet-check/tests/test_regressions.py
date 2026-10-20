"""回归用例：三个「错了还不报」的问题各一条。

    python tests/test_regressions.py

用例全部在临时目录里造最小 .docx，不读真实单据，不写任何用户目录。
"""

import contextlib
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
sys.path.insert(0, str(SKILL / 'scripts'))

import build_audit            # noqa: E402
import docx_table as dt       # noqa: E402
import fill_sheet             # noqa: E402
import verify_sheet           # noqa: E402

RULES_HEADER = ['参数表分段', '参数表字段', '取数方式', '报告分段', '报告字段', '取值处理', '说明', '依据文件']

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)
DOC_TMPL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>{body}</w:body></w:document>'
)


def _cell(text):
    esc = (text or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return f'<w:tc><w:p><w:r><w:t xml:space="preserve">{esc}</w:t></w:r></w:p></w:tc>'


def _table(rows):
    trs = ''.join('<w:tr>' + ''.join(_cell(c) for c in row) + '</w:tr>' for row in rows)
    return f'<w:tbl>{trs}</w:tbl>'


def make_docx(path, tables):
    """tables: list[list[list[str]]]，按顺序写成文档中的多张表。"""
    body = ''.join(_table(t) for t in tables)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('word/document.xml', DOC_TMPL.format(body=body))
    return Path(path)


def make_rules(path, rows):
    lines = ['\t'.join(RULES_HEADER)]
    for row in rows:
        lines.append('\t'.join(row))
    Path(path).write_text('\n'.join(lines) + '\n', encoding='utf-8')
    return Path(path)


SHEET_PARAM = [
    ['整车参数', '车辆型号名称', '占位原值', '发动机型号', 'YN25PLUS160B'],
]


class Base(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def run_main(self, func, argv):
        """跑一个脚本的 main，返回 (退出码, stdout, stderr)。"""
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = func(argv)
        return code, out.getvalue(), err.getvalue()


class MultiTableTest(Base):
    """bug 1：定位不带表序号，多表文档会写进第一张「行够列够」的表。"""

    def build(self):
        # 第 1 张：封面信息表，行列都比目标位置大，但整表不含识别关键词
        cover = [[f'封面{r}{c}' for c in range(6)] for r in range(3)]
        sheet = make_docx(self.tmp / 'sheet.docx', [cover, SHEET_PARAM])
        report = make_docx(self.tmp / 'pm.docx',
                           [[['整车参数', '车辆型号名称', 'HFC1042P23K1B4ZS', '发动机型号', 'YN25PLUS160B']]])
        rules = make_rules(self.tmp / 'rules.tsv',
                           [['整车参数', '车辆型号名称', '排放', '整车参数', '车辆型号名称', 'value', '取报告值', '']])
        return sheet, report, rules

    def test_position_carries_table_index(self):
        sheet, _, _ = self.build()
        _, positions = dt.parse_sectioned(sheet, ['车辆型号名称', '发动机型号'],
                                          ['整车参数', '发动机参数'])
        self.assertEqual(positions[('整车参数', '车辆型号名称')], (1, 0, 2))

    def test_writes_into_second_table_only(self):
        sheet, report, rules = self.build()
        out_dir = self.tmp / 'out'
        code, _, err = self.run_main(build_audit.main, [
            '--sheet', str(sheet), '--pm', str(report), '--zs', str(report),
            '--rules', str(rules), '--out-dir', str(out_dir)])
        self.assertEqual(code, 0, err)

        copy = self.tmp / 'copy.docx'
        code, _, err = self.run_main(fill_sheet.main, [
            '--sheet', str(sheet), '--audit', str(out_dir / 'audit.json'),
            '--out', str(copy), '--apply'])
        self.assertEqual(code, 0, err)

        tables = dt.read_tables(copy)
        self.assertEqual(tables[0][0][0], '封面00', '封面表被写坏了')
        self.assertTrue(all(c.startswith('封面') for row in tables[0] for c in row),
                        f'封面表出现非封面内容：{tables[0]}')
        self.assertEqual(tables[1][0][2], 'HFC1042P23K1B4ZS')
        self.assertEqual(tables[1][0][3], '发动机型号', '写到了错误的列')

    def test_out_of_range_position_raises(self):
        sheet, _, _ = self.build()
        with self.assertRaises(ValueError):
            dt.write_cells(sheet, self.tmp / 'x.docx', [(5, 0, 2, '值')])
        self.assertFalse((self.tmp / 'x.docx').exists(), '定位失败却生成了副本')


class AmbiguousSetTest(Base):
    """bug 2：--set 裸字段名命中第一个同名项，静默写错分段。"""

    def build(self):
        sheet = make_docx(self.tmp / 'sheet.docx', [[
            ['整车参数', '生产厂', '甲厂', '车辆型号名称', 'WRD5041THX'],
            ['底盘参数', '生产厂', '乙厂', '发动机型号', 'YN25'],
        ]])
        report = make_docx(self.tmp / 'pm.docx',
                           [[['整车参数', '车辆型号名称', 'HFC1042', '发动机型号', 'YN25']]])
        rules = make_rules(self.tmp / 'rules.tsv', [
            ['整车参数', '生产厂', '人工', '', '', 'value', '看铭牌', '整车合格证'],
            ['底盘参数', '生产厂', '人工', '', '', 'value', '看合格证', '底盘合格证'],
        ])
        out_dir = self.tmp / 'out'
        code, _, err = self.run_main(build_audit.main, [
            '--sheet', str(sheet), '--pm', str(report), '--zs', str(report),
            '--rules', str(rules), '--out-dir', str(out_dir)])
        self.assertEqual(code, 0, err)
        return sheet, out_dir / 'audit.json'

    def test_bare_duplicate_name_is_rejected(self):
        sheet, audit = self.build()
        code, _, err = self.run_main(fill_sheet.main, [
            '--sheet', str(sheet), '--audit', str(audit), '--out', str(self.tmp / 'c.docx'),
            '--set', '生产厂=丙厂'])
        self.assertEqual(code, 4, '重名字段被静默接受了')
        self.assertIn('重名', err)
        self.assertIn('整车参数/生产厂', err)
        self.assertIn('底盘参数/生产厂', err)

    def test_qualified_name_works(self):
        sheet, audit = self.build()
        code, out, err = self.run_main(fill_sheet.main, [
            '--sheet', str(sheet), '--audit', str(audit), '--out', str(self.tmp / 'c.docx'),
            '--set', '底盘参数/生产厂=丙厂'])
        self.assertEqual(code, 0, err)
        self.assertIn('丙厂', out)
        self.assertIn('底盘参数/生产厂', out)
        self.assertNotIn('整车参数/生产厂 = 丙厂', out)

    def test_skip_field_cannot_be_set(self):
        sheet, audit = self.build()
        data = json.loads(audit.read_text(encoding='utf-8'))
        for it in data['items']:
            if it['field'] == '生产厂':
                it['status'] = '不填'
                break
        audit.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
        code, _, err = self.run_main(fill_sheet.main, [
            '--sheet', str(sheet), '--audit', str(audit), '--out', str(self.tmp / 'c.docx'),
            '--set', '整车参数/生产厂=丙厂'])
        self.assertEqual(code, 4)
        self.assertIn('不填', err)


class ParseGuardTest(Base):
    """bug 3：解析出 0 项时照常往下跑，生成一张「报告全空」的核对表。"""

    def test_unrecognized_sheet_stops(self):
        sheet = make_docx(self.tmp / 'sheet.docx', [[['备注', '内容', '随便写']]])
        report = make_docx(self.tmp / 'pm.docx',
                           [[['整车参数', '车辆型号名称', 'HFC1042', '发动机型号', 'YN25']]])
        rules = make_rules(self.tmp / 'rules.tsv',
                           [['整车参数', '车辆型号名称', '排放', '整车参数', '车辆型号名称', 'value', '取报告值', '']])
        out_dir = self.tmp / 'out'
        code, _, err = self.run_main(build_audit.main, [
            '--sheet', str(sheet), '--pm', str(report), '--zs', str(report),
            '--rules', str(rules), '--out-dir', str(out_dir)])
        self.assertEqual(code, 3, '解析失败却继续生成了核对表')
        self.assertIn('参数表解析失败', err)
        self.assertFalse((out_dir / 'audit.json').exists())

    def test_binary_doc_renamed_gives_clean_error(self):
        fake = self.tmp / 'fake.docx'
        fake.write_bytes(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00' * 128)
        report = make_docx(self.tmp / 'pm.docx',
                           [[['整车参数', '车辆型号名称', 'HFC1042', '发动机型号', 'YN25']]])
        rules = make_rules(self.tmp / 'rules.tsv',
                           [['整车参数', '车辆型号名称', '排放', '整车参数', '车辆型号名称', 'value', '取报告值', '']])
        code, _, err = self.run_main(build_audit.main, [
            '--sheet', str(fake), '--pm', str(report), '--zs', str(report),
            '--rules', str(rules), '--out-dir', str(self.tmp / 'out')])
        self.assertEqual(code, 3)
        self.assertIn('不是有效的 .docx', err)
        self.assertNotIn('Traceback', err)

    def test_report_parse_failure_stops(self):
        sheet = make_docx(self.tmp / 'sheet.docx', [SHEET_PARAM])
        bad = make_docx(self.tmp / 'bad.docx', [[['甲', '乙', '丙']]])
        good = make_docx(self.tmp / 'pm.docx',
                         [[['整车参数', '车辆型号名称', 'HFC1042', '发动机型号', 'YN25']]])
        rules = make_rules(self.tmp / 'rules.tsv',
                           [['整车参数', '车辆型号名称', '排放', '整车参数', '车辆型号名称', 'value', '取报告值', '']])
        code, _, err = self.run_main(build_audit.main, [
            '--sheet', str(sheet), '--pm', str(bad), '--zs', str(good),
            '--rules', str(rules), '--out-dir', str(self.tmp / 'out')])
        self.assertEqual(code, 3)
        self.assertIn('报告解析失败', err)

    def test_verify_stops_on_unreadable_sheet(self):
        sheet = make_docx(self.tmp / 'sheet.docx', [[['备注', '内容', '随便写']]])
        audit = self.tmp / 'audit.json'
        audit.write_text(json.dumps({'meta': {}, 'items': []}, ensure_ascii=False), encoding='utf-8')
        code, _, err = self.run_main(verify_sheet.main, [
            '--sheet', str(sheet), '--audit', str(audit), '--out', str(self.tmp / 'v.md')])
        self.assertEqual(code, 3)
        self.assertIn('参数表解析失败', err)
        self.assertFalse((self.tmp / 'v.md').exists())


class SourceKindTest(Base):
    """取数方式只认五个值；写错的不能被当成人工悄悄放过。"""

    def test_unknown_source_rejected(self):
        sheet = make_docx(self.tmp / 'sheet.docx', [SHEET_PARAM])
        report = make_docx(self.tmp / 'pm.docx',
                           [[['整车参数', '车辆型号名称', 'HFC1042', '发动机型号', 'YN25']]])
        rules = make_rules(self.tmp / 'rules.tsv',
                           [['整车参数', '车辆型号名称', '自动', '整车参数', '车辆型号名称', 'value', '取报告值', '']])
        code, _, err = self.run_main(build_audit.main, [
            '--sheet', str(sheet), '--pm', str(report), '--zs', str(report),
            '--rules', str(rules), '--out-dir', str(self.tmp / 'out')])
        self.assertEqual(code, 4)
        self.assertIn('取数方式不认识', err)
        self.assertIn('两款', err)


class ManualRefTest(Base):
    """人工字段的参考值只给排放报告一边，会藏住噪声报告的口径差异。"""

    def build(self, pm_value='1-3', zs_value='2-3'):
        sheet = make_docx(self.tmp / 'sheet.docx', [[
            ['整车参数', '车辆型号名称', 'WRD5041THX', '设计乘员数(人)', ''],
            ['', '发动机型号', 'YN25', '', ''],
        ]])
        pm = make_docx(self.tmp / 'pm.docx', [[[
            '整车参数', '车辆型号名称', 'HFC1042', '设计乘员数(人)', pm_value,
            '发动机型号', 'YN25PLUS160B']]])
        zs = make_docx(self.tmp / 'zs.docx', [[[
            '整车参数', '车辆型号名称', 'HFC1042', '设计乘员数(人)', zs_value,
            '发动机型号', 'YN25PLUS160B']]])
        rules = make_rules(self.tmp / 'rules.tsv', [[
            '整车参数', '设计乘员数(人)', '人工', '整车参数', '设计乘员数(人)',
            'value', '看合格证', '整车合格证']])
        out_dir = self.tmp / 'out'
        code, _, err = self.run_main(build_audit.main, [
            '--sheet', str(sheet), '--pm', str(pm), '--zs', str(zs),
            '--rules', str(rules), '--out-dir', str(out_dir)])
        self.assertEqual(code, 0, err)
        data = json.loads((out_dir / 'audit.json').read_text(encoding='utf-8'))
        item = data['items'][0]
        md = (out_dir / 'audit.md').read_text(encoding='utf-8')
        return sheet, item, md

    def test_both_reports_shown_when_different(self):
        _, item, md = self.build('1-3', '2-3')
        self.assertEqual(item['ref'], '', '两份不一致时不应挑一个当参考值')
        self.assertEqual(item['ref_text'], 'A 排放:1-3 / B 噪声:2-3')
        self.assertEqual({c['value'] for c in item['ref_cands']}, {'1-3', '2-3'})
        self.assertIn('1-3', md)
        self.assertIn('2-3', md, 'audit.md 只显示了排放报告的值')

    def test_single_value_when_reports_agree(self):
        _, item, _ = self.build('2-3', '2-3')
        self.assertEqual(item['ref'], '2-3')
        self.assertEqual(item['ref_text'], '2-3')

    def test_verify_matches_either_candidate(self):
        sheet, item, _ = self.build('1-3', '2-3')
        del sheet
        res = []
        verify_sheet.check_manual_vs_report(
            {('整车参数', '设计乘员数(人)'): '2-3', '': ''}, [item], res)
        self.assertEqual(res, [], '表内值等于噪声报告的参考值却被报警')

    def test_verify_passes_certificate_basis(self):
        _, item, _ = self.build('1-3', '2-3')
        res = []
        verify_sheet.check_manual_vs_report({('整车参数', '设计乘员数(人)'): '3'}, [item], res)
        self.assertEqual(res[0][0], verify_sheet.PASS)
        self.assertIn('整车合格证', res[0][2])


if __name__ == '__main__':
    unittest.main(verbosity=2)
