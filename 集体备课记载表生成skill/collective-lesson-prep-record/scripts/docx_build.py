# -*- coding: utf-8 -*-
"""从零生成《教研活动记录》docx（纯标准库，不依赖 python-docx）。

版式：A4 纵向 / 标题黑体 16pt 居中 / 宋体 10.5pt 表格 / 教研内容 1.5 倍行距首行缩进。
只填事实性内容，不写反思、心得、收获。
"""

from __future__ import annotations

import os
import zipfile

A4_W, A4_H = 11906, 16838
MARGIN = 1440
GRID = [1700, 2700, 1300, 3326]  # 合计 9026 = A4 可用宽度

XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
NS_W = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

CONTENT_TYPES = (
    XML_HEAD
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)

RELS = (
    XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)

DOC_RELS = (
    XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    "</Relationships>"
)

STYLES = (
    XML_HEAD
    + f"<w:styles {NS_W}>"
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>'
    '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    "</w:styles>"
)

BORDERS = (
    "<w:tblBorders>"
    + "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    + "</w:tblBorders>"
)


def esc(s) -> str:
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def para(
    text: str = "",
    *,
    bold=False,
    size=21,
    font="宋体",
    align=None,
    first_line_indent=False,
    line=300,
    space_before=0,
    space_after=0,
) -> str:
    ppr = ['<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>' % (space_before, space_after, line)]
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    if first_line_indent:
        ppr.append('<w:ind w:firstLineChars="200" w:firstLine="420"/>')
    rpr = f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="{font}"/>'
    if bold:
        rpr += "<w:b/><w:bCs/>"
    rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    body = f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>' if text else ""
    return f"<w:p><w:pPr>{''.join(ppr)}</w:pPr>{body}</w:p>"


def cell(paragraphs_xml: str, width: int, span: int = 1, center=True, shade=None) -> str:
    tcpr = [f'<w:tcW w:w="{width}" w:type="dxa"/>']
    if span > 1:
        tcpr.append(f'<w:gridSpan w:val="{span}"/>')
    if shade:
        tcpr.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>')
    tcpr.append(f'<w:vAlign w:val="{"center" if center else "top"}"/>')
    return f"<w:tc><w:tcPr>{''.join(tcpr)}</w:tcPr>{paragraphs_xml}</w:tc>"


def label_cell(text: str, width: int = GRID[0]) -> str:
    return cell(para(text, align="center", bold=True), width)


def value_cell(text: str, width: int, span: int = 1) -> str:
    return cell(para(text, align="left"), width, span=span)


def row(cells_xml: str, height: int | None = None, rule="atLeast") -> str:
    trpr = f'<w:trPr><w:trHeight w:hRule="{rule}" w:val="{height}"/></w:trPr>' if height else ""
    return f"<w:tr>{trpr}{cells_xml}</w:tr>"


def build_document(rec: dict) -> str:
    title = rec.get("title") or "教研活动记录"
    rows = []

    rows.append(row(label_cell("活动主题") + value_cell(rec.get("活动主题", ""), 7326, span=3), 620))
    rows.append(
        row(
            label_cell("活动时间")
            + value_cell(rec.get("活动时间", ""), GRID[1])
            + label_cell("主持人", GRID[2])
            + value_cell(rec.get("主持人", ""), GRID[3]),
            620,
        )
    )
    rows.append(row(label_cell("教研主题") + value_cell(rec.get("教研主题", ""), 7326, span=3), 620))
    if rec.get("参加人员"):
        rows.append(row(label_cell("参加人员") + value_cell(rec["参加人员"], 7326, span=3), 620))

    content = str(rec.get("教研内容", "")).replace("\r\n", "\n").replace("\r", "\n")
    body_paras = "".join(
        para(line.strip(), first_line_indent=True, line=320) for line in content.split("\n")
    )
    content_cell = cell(body_paras, 7326, span=3, center=False)
    rows.append(row(label_cell("教研内容") + content_cell, 7800))

    tbl = (
        "<w:tbl><w:tblPr>"
        '<w:tblW w:w="9026" w:type="dxa"/><w:jc w:val="center"/>'
        + BORDERS
        + '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
        '<w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
        '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
        "</w:tblCellMar></w:tblPr>"
        "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in GRID) + "</w:tblGrid>"
        + "".join(rows)
        + "</w:tbl>"
    )

    sect = (
        "<w:sectPr>"
        f'<w:pgSz w:w="{A4_W}" w:h="{A4_H}"/>'
        f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:docGrid w:linePitch="312"/>'
        "</w:sectPr>"
    )

    body = (
        para(title, bold=True, size=32, font="黑体", align="center", space_after=240, line=240)
        + tbl
        + para("", line=240)  # 表格后空段，避免版面贴边
        + sect
    )
    return XML_HEAD + f"<w:document {NS_W}><w:body>{body}</w:body></w:document>"


def build_record_docx(rec: dict, out: str) -> str:
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    tmp = out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", build_document(rec))
    from docx_fill import _replace

    _replace(tmp, out)
    return out
