"""从零生成 .docx 的最小工具，仅依赖 Python 标准库。

只支持本技能需要的排版：标题、一级小标题、正文段落、带表头底纹的网格表格。
不解析复杂样式，不读写外部文件，除 zipfile 外无其它依赖。

用法：
    from docx_write import Doc
    d = Doc(landscape=True)
    d.title('标题')
    d.h1('一、小节')
    d.p('正文')
    d.table(['列1', '列2'], [['a', 'b']], widths=[2000, 3000])
    d.save('out.docx')
"""

import re
import zipfile
from xml.sax.saxutils import escape

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

FONT_CN = '宋体'
FONT_CN_HEAD = '黑体'
FONT_EN = 'Times New Roman'

# A4 两种方向下的可用宽度（twips）
PAGE_W = {True: 16838, False: 11906}
PAGE_H = {True: 11906, False: 16838}
MARGIN = 1134

_CTRL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _clean(text):
    """去掉 XML 不接受的控制字符，并转义。"""
    s = _CTRL.sub('', str(text if text is not None else ''))
    return escape(s)


def _run(text, size=21, bold=False, font_cn=FONT_CN, color=None):
    props = [f'<w:rFonts w:ascii="{FONT_EN}" w:hAnsi="{FONT_EN}" w:eastAsia="{font_cn}"/>']
    if bold:
        props.append('<w:b/><w:bCs/>')
    if color:
        props.append(f'<w:color w:val="{color}"/>')
    props.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    return (f'<w:r><w:rPr>{"".join(props)}</w:rPr>'
            f'<w:t xml:space="preserve">{_clean(text)}</w:t></w:r>')


def _para(text='', size=21, bold=False, align='left', before=0, after=60,
          font_cn=FONT_CN, color=None, line=288, keep=False):
    ppr = [f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>',
           f'<w:jc w:val="{align}"/>']
    if keep:
        ppr.append('<w:keepNext/>')
    return f'<w:p><w:pPr>{"".join(ppr)}</w:pPr>{_run(text, size, bold, font_cn, color)}</w:p>'


def _cell(text, width, size=18, bold=False, fill=None, align='left'):
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ''
    p = (f'<w:p><w:pPr><w:spacing w:before="20" w:after="20" w:line="240" w:lineRule="auto"/>'
         f'<w:jc w:val="{align}"/></w:pPr>{_run(text, size, bold)}</w:p>')
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shd}'
            f'<w:vAlign w:val="center"/></w:tcPr>{p}</w:tc>')


class Doc:
    def __init__(self, landscape=True):
        self.landscape = landscape
        self.body = []
        self.content_w = PAGE_W[landscape] - 2 * MARGIN

    # ---------- 内容 ----------
    def title(self, text):
        self.body.append(_para(text, size=32, bold=True, align='center',
                              before=0, after=180, font_cn=FONT_CN_HEAD))

    def h1(self, text):
        self.body.append(_para(text, size=24, bold=True, align='left',
                              before=240, after=120, font_cn=FONT_CN_HEAD, keep=True))

    def p(self, text, size=21, bold=False, align='left', after=60, color=None):
        self.body.append(_para(text, size=size, bold=bold, align=align,
                              after=after, color=color))

    def gap(self):
        self.body.append('<w:p><w:pPr><w:spacing w:after="0" w:line="200" '
                         'w:lineRule="auto"/></w:pPr></w:p>')

    def table(self, headers, rows, widths, size=18, aligns=None, header_fill='D9E2F3',
              header_size=None):
        """widths 为相对权重，内部按可用宽度等比换算成 twips。"""
        total = sum(widths)
        cols = [max(400, int(self.content_w * w / total)) for w in widths]
        diff = self.content_w - sum(cols)
        cols[-1] += diff
        aligns = aligns or ['left'] * len(headers)
        hsize = header_size or size

        grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in cols)
        border = ''.join(
            f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="808080"/>'
            for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
        parts = [
            '<w:tbl><w:tblPr>',
            f'<w:tblW w:w="{self.content_w}" w:type="dxa"/>',
            '<w:tblLayout w:type="fixed"/>',
            f'<w:tblBorders>{border}</w:tblBorders>',
            '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="70" w:type="dxa"/>'
            '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar>',
            '</w:tblPr>',
            f'<w:tblGrid>{grid}</w:tblGrid>',
        ]
        # 表头：跨页重复
        head_cells = ''.join(_cell(h, cols[i], hsize, True, header_fill,
                                   'center' if aligns[i] == 'center' else 'left')
                             for i, h in enumerate(headers))
        parts.append(f'<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>{head_cells}</w:tr>')

        for row in rows:
            cells = ''.join(_cell(row[i] if i < len(row) else '', cols[i], size, False,
                                  None, aligns[i])
                            for i in range(len(headers)))
            parts.append(f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{cells}</w:tr>')
        parts.append('</w:tbl>')
        self.body.append(''.join(parts))
        self.body.append('<w:p><w:pPr><w:spacing w:after="0" w:line="120" '
                         'w:lineRule="auto"/></w:pPr></w:p>')

    # ---------- 打包 ----------
    def _document_xml(self):
        sect = (f'<w:sectPr><w:pgSz w:w="{PAGE_W[self.landscape]}" w:h="{PAGE_H[self.landscape]}"'
                + (' w:orient="landscape"' if self.landscape else '')
                + f'/><w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" '
                  f'w:left="{MARGIN}" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="{NS_W}" xmlns:r="{NS_R}"><w:body>'
                + ''.join(self.body) + sect + '</w:body></w:document>')

    def save(self, path):
        files = {
            '[Content_Types].xml': _CONTENT_TYPES,
            '_rels/.rels': _RELS,
            'word/_rels/document.xml.rels': _DOC_RELS,
            'word/styles.xml': _STYLES,
            'docProps/core.xml': _CORE,
            'docProps/app.xml': _APP,
            'word/document.xml': self._document_xml(),
        }
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, data in files.items():
                zf.writestr(name, data.encode('utf-8'))


_CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

_DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

_STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{NS_W}">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="{FONT_EN}" w:hAnsi="{FONT_EN}" w:eastAsia="{FONT_CN}" w:cs="{FONT_EN}"/>
<w:sz w:val="21"/><w:szCs w:val="21"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="60" w:line="288" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
<w:name w:val="Normal"/><w:qFormat/>
</w:style>
</w:styles>'''

_CORE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>国六环保样品参数表 数据来源核对表</dc:title>
<dc:creator>guo6-param-sheet-check</dc:creator>
</cp:coreProperties>'''

_APP = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>guo6-param-sheet-check</Application>
</Properties>'''
