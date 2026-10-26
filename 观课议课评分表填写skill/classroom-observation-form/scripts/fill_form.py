# -*- coding: utf-8 -*-
"""
观课、议课量化表填充引擎
================================================================
作用：把一份「内容数据(JSON)」写进《观课、议课量化表.xlsx》模板的 6 个子表，
      保持模板结构完全不变（合并单元格、行高列宽、内嵌座位图、打印设置全部原样保留）。

      一节课 = 一份文件。模板里的 6 个子表由 6 位观课教师各填一张，
      因此 payload 里要为每个子表指定一位教师（teachers）。

为什么不用 openpyxl 保存：
    openpyxl 在 load→save 过程中会丢弃 xl/drawings 与 xl/media（本模板的座位图、
    小组观察图就嵌在那里），保存后座位图会消失。因此本脚本直接在 OOXML 包内部
    改单元格 XML，除被写入的 <c> 元素外，其余字节原样回写。

用法：
    python fill_form.py --payload data.json [--template 模板.xlsx] [--out 输出目录]
                        [--no-auto-height] [--verify]

payload 结构见 references/payload-schema.md
"""

import argparse
import json
import math
import os
import re
import sys
import zipfile
from datetime import date, datetime

# ----------------------------------------------------------------------------
# 1. 模板单元格映射（与 assets/template.xlsx 一一对应；模板若改版，只改这里）
# ----------------------------------------------------------------------------
# meta 条目字段说明：
#   cell    目标单元格（必须是合并区域左上角）
#   format  格式化模板，可用占位符 {date} {date_text} {content} {subject}
#           {class_name} {teacher}；有 format 时忽略 prefix/field
#   prefix  前缀 + field 取值，例如 "观课时间：" + date_text
#   kind    "text"（默认） | "number" | "date"（写 Excel 日期序列值）
SHEETS = {
    # ---------------- 子表1 师生问答（座位图 + 分析结论） ----------------
    "师生问答": {
        "meta": [
            {"cell": "B2", "format": "观课时间：{date_text}"},
            {"cell": "E2", "field": "content"},
            {"cell": "G2", "field": "subject"},
            {"cell": "H2", "format": "班级：{class_name}"},
            {"cell": "J2", "field": "teacher"},
        ],
        "fields": [("B24", "conclusion")],
        "table": None,
    },
    # ---------------- 子表2 教师提问有向、开放性 ----------------
    "教师提问": {
        "meta": [
            {"cell": "B2", "kind": "date", "field": "date"},
            {"cell": "C2", "format": "内容：{content}"},
            {"cell": "D2", "format": "       学科：{subject}"},
            {"cell": "E2", "format": "     班级：{class_name}"},
            {"cell": "F2", "format": "观课教师：{teacher}"},
        ],
        "fields": [],
        "table": {
            "key": "questions",
            "rows": [7, 8, 9, 10, 11],
            "cols": {"A": "no", "B": "q", "C": "target", "D": "answer",
                     "E": "openness", "F": "reconstruct"},
        },
    },
    # ---------------- 子表3 学生学习活动 ----------------
    "学生学习活动": {
        "meta": [
            {"cell": "B2", "kind": "date", "field": "date"},
            {"cell": "C2", "format": "内容：{content}"},
            {"cell": "D2", "format": "      学科：{subject}"},
            {"cell": "E2", "format": "   班级：{class_name}"},
            {"cell": "F2", "format": "观课教师：{teacher}"},
        ],
        "fields": [
            ("B5", "participation"),
            ("C5", "thinking"),
            ("D5", "cooperation"),
            ("E5", "generativity"),
            ("F5", "achievement"),
        ],
        "table": None,
    },
    # ---------------- 子表4 小组合作学习期间教师行为 ----------------
    "小组合作教师行为": {
        "meta": [
            {"cell": "B3", "kind": "date", "field": "date"},
            {"cell": "C3", "format": "内容：     {content}"},
            {"cell": "D3", "format": "  学科：{subject}"},
            {"cell": "E3", "format": "  班级：{class_name}"},
            {"cell": "F3", "format": "观课教师：{teacher}"},
        ],
        "fields": [("A18", "analysis")],
        "table": None,
    },
    # ---------------- 子表5 小组合作组员行为 ----------------
    "小组合作组员行为": {
        "meta": [
            {"cell": "A2", "format": "  观课时间：{date_text}"},
            {"cell": "B2", "format": "内容：{content}"},
            {"cell": "D2", "format": "       学科：{subject}"},
            {"cell": "E2", "format": "   班级：{class_name}"},
            {"cell": "F2", "format": "观课教师：{teacher}"},
        ],
        "fields": [
            ("C4", "task"),
            ("C5", "size", "number"),
            ("C6", "minutes", "number"),
        ],
        "members": {
            "duty_cells": ["A9", "B9", "D9", "E9", "F9"],
            "work_cells": ["A12", "B12", "D12", "E12", "F12"],
        },
        "table": None,
    },
    # ---------------- 子表6 课堂评价观察表 ----------------
    "课堂评价观察表": {
        "meta": [
            {"cell": "A2", "format": "  观课时间：{date_text}"},
            {"cell": "C2", "field": "content"},
            {"cell": "D2", "format": "学科：{subject}"},
            {"cell": "E2", "format": "班级：{class_name}   观课教师：{teacher}"},
        ],
        "fields": [],
        "table": {
            "key": "items",
            "rows": [7, 8, 9],
            "cols": {"A": "activity", "B": "statement", "C": "standard",
                     "D": "mode", "E": "effect"},
        },
    },
}

SHEET_ORDER = list(SHEETS.keys())

# ----------------------------------------------------------------------------
# 2. 基础工具
# ----------------------------------------------------------------------------
_ESCAPE = {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}


def esc(s):
    return "".join(_ESCAPE.get(ch, ch) for ch in str(s))


def col_index(ref):
    """'B24' -> 2"""
    n = 0
    for ch in re.match(r"([A-Z]+)", ref).group(1):
        n = n * 26 + (ord(ch) - 64)
    return n


def row_index(ref):
    return int(re.search(r"(\d+)$", ref).group(1))


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref.strip().upper())
    if not m:
        raise ValueError("非法单元格地址: %r" % ref)
    return m.group(1), int(m.group(2))


def excel_serial(d):
    """date -> Excel 1900 日期序列值"""
    return (d - date(1899, 12, 30)).days


def parse_date(value):
    """接受 2026-05-06 / 2026/5/6 / 5月6日 / 20260506，返回 date"""
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.match(r"^(\d{1,2})\s*月\s*(\d{1,2})\s*日?$", s)
    if m:
        return date(date.today().year, int(m.group(1)), int(m.group(2)))
    raise ValueError("无法识别的观课时间: %r（请用 2026-05-06 这样的格式）" % value)


def cn_date(d):
    return "%d月%d日" % (d.month, d.day)


# ----------------------------------------------------------------------------
# 3. OOXML 包读写
# ----------------------------------------------------------------------------
CELL_TMPL = r'<c r="%s"(?P<attrs>(?:\s[^>]*?)?)(?:/>|>(?P<inner>.*?)</c>)'
ROW_TMPL = r'<row\b[^>]*?\br="%d"(?=[\s/>])[^>]*?>'


class Book(object):
    def __init__(self, template_path):
        self.template = template_path
        with zipfile.ZipFile(template_path) as zf:
            self.order = [i.filename for i in zf.infolist()]
            self.parts = {i.filename: zf.read(i.filename) for i in zf.infolist()}
        self.sheet_paths = self._map_sheets()
        self.dirty = set()
        self.filled = []          # [(sheet, ref, text)] 供自动行高使用
        self._width_cache = {}
        self._merge_cache = {}
        self._h_cache = {}

    # -- 包结构解析 ----------------------------------------------------------
    def _map_sheets(self):
        wb = self.parts["xl/workbook.xml"].decode("utf-8")
        rels = self.parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rel = {}
        for tag in re.findall(r"<Relationship\b[^>]*/>", rels):
            rid = re.search(r'\bId="([^"]+)"', tag)
            tgt = re.search(r'\bTarget="([^"]+)"', tag)
            if not (rid and tgt):
                continue
            path = tgt.group(1).lstrip("/")
            path = path[3:] if path.startswith("../") else path
            rel[rid.group(1)] = path if path.startswith("xl/") else "xl/" + path
        out = {}
        for tag in re.findall(r"<sheet\b[^>]*/>", wb):
            name = re.search(r'\bname="([^"]*)"', tag)
            rid = re.search(r'\br:id="([^"]+)"', tag)
            if name and rid and rid.group(1) in rel:
                out[name.group(1)] = rel[rid.group(1)]
        return out

    def xml(self, sheet):
        return self.parts[self.sheet_paths[sheet]].decode("utf-8")

    def put(self, sheet, text):
        self.parts[self.sheet_paths[sheet]] = text.encode("utf-8")
        self.dirty.add(self.sheet_paths[sheet])

    # -- 写入单元格 ----------------------------------------------------------
    def _neighbour_style(self, sheet, xml, ref):
        col, row = split_ref(ref)
        target = col_index(ref)
        # 先在同一行找最近的单元格
        rm = re.search(ROW_TMPL % row, xml)
        best = None
        if rm:
            end = xml.find("</row>", rm.end())
            if end > 0:
                seg = xml[rm.end():end]
                for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)/?>', seg):
                    s = re.search(r'\bs="(\d+)"', cm.group(2))
                    if not s:
                        continue
                    d = abs(col_index(cm.group(1) + "1") - target)
                    if best is None or d < best[0]:
                        best = (d, s.group(1))
        if best:
            return best[1]
        # 退而在同一列找
        for m in re.finditer(r'<c r="%s(\d+)"([^>]*?)/?>' % col, xml):
            s = re.search(r'\bs="(\d+)"', m.group(2))
            if s:
                return s.group(1)
        return None

    def set_cell(self, sheet, ref, value, kind="text"):
        xml = self.xml(sheet)
        pat = re.compile(CELL_TMPL % re.escape(ref), re.S)
        m = pat.search(xml)
        style = None
        if m:
            s = re.search(r'\bs="(\d+)"', m.group("attrs") or "")
            style = s.group(1) if s else None
        else:
            style = self._neighbour_style(sheet, xml, ref)
        attr = (' s="%s"' % style) if style else ""

        empty = value is None or (isinstance(value, str) and value.strip() == "")
        if empty:
            new = '<c r="%s"%s/>' % (ref, attr)
        elif kind == "number":
            new = '<c r="%s"%s><v>%s</v></c>' % (ref, attr, value)
        else:
            new = ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s'
                   '</t></is></c>' % (ref, attr, esc(value)))

        if m:
            xml = xml[:m.start()] + new + xml[m.end():]
        else:
            xml = self._insert_cell(xml, ref, new)
        self.put(sheet, xml)

    def _insert_cell(self, xml, ref, cell_xml):
        col, row = split_ref(ref)
        target = col_index(ref)
        rm = re.search(ROW_TMPL % row, xml)
        if not rm:
            # 整行都不存在：按行号顺序插入新行
            new_row = '<row r="%d">%s</row>' % (row, cell_xml)
            pos = None
            for m in re.finditer(r'<row\b[^>]*?\br="(\d+)"', xml):
                if int(m.group(1)) > row:
                    pos = m.start()
                    break
            if pos is None:
                pos = xml.find("</sheetData>")
                if pos < 0:
                    raise RuntimeError("模板缺少 <sheetData>：%s" % self.sheet_paths)
            return xml[:pos] + new_row + xml[pos:]

        tag = rm.group(0)
        selfclose = tag.endswith("/>")
        if selfclose:
            head = tag[:-2].rstrip() + ">"
            return xml[:rm.start()] + head + cell_xml + "</row>" + xml[rm.end():]

        inner_start = rm.end()
        inner_end = xml.find("</row>", inner_start)
        inner = xml[inner_start:inner_end]
        cells = re.findall(r"<c\b.*?(?:/>|</c>)", inner, re.S)
        insert_at = len(inner)
        for cm in re.finditer(r"<c\b.*?(?:/>|</c>)", inner, re.S):
            cref = re.search(r'\br="([A-Z]+\d+)"', cm.group(0))
            if cref and col_index(cref.group(1)) > target:
                insert_at = cm.start()
                break
        new_inner = inner[:insert_at] + cell_xml + inner[insert_at:]
        return xml[:inner_start] + new_inner + xml[inner_end:]

    # -- 行高 ----------------------------------------------------------------
    def row_heights(self, xml):
        out = {}
        for m in re.finditer(r'<row\b[^>]*?\br="(\d+)"[^>]*?>', xml):
            h = re.search(r'\bht="([\d.]+)"', m.group(0))
            if h:
                out[int(m.group(1))] = float(h.group(1))
        return out

    def set_row_height(self, sheet, row, height):
        xml = self.xml(sheet)
        rm = re.search(ROW_TMPL % row, xml)
        if not rm:
            return
        tag = rm.group(0)
        selfclose = tag.endswith("/>")
        body = tag[:-2] if selfclose else tag[:-1]
        body = re.sub(r'\s*\bht="[^"]*"', "", body)
        body = re.sub(r'\s*\bcustomHeight="[^"]*"', "", body)
        body = body.rstrip() + ' ht="%.2f" customHeight="1"' % height
        newtag = body + ("/>" if selfclose else ">")
        self.put(sheet, xml[:rm.start()] + newtag + xml[rm.end():])
        self._h_cache[(sheet, row)] = height

    def col_widths(self, sheet):
        if sheet in self._width_cache:
            return self._width_cache[sheet]
        xml = self.xml(sheet)
        widths = {}
        for m in re.finditer(r"<col\b[^>]*?/>", xml):
            tag = m.group(0)
            try:
                lo = int(re.search(r'\bmin="(\d+)"', tag).group(1))
                hi = int(re.search(r'\bmax="(\d+)"', tag).group(1))
                w = float(re.search(r'\bwidth="([\d.]+)"', tag).group(1))
            except AttributeError:
                continue
            for c in range(lo, hi + 1):
                widths[c] = w
        self._width_cache[sheet] = widths
        return widths

    def merges(self, sheet):
        if sheet in self._merge_cache:
            return self._merge_cache[sheet]
        xml = self.xml(sheet)
        out = {}
        for m in re.finditer(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"/>', xml):
            a, b = m.group(1).split(":")
            c1, r1 = split_ref(a)
            c2, r2 = split_ref(b)
            out[a] = (col_index(c1 + "1"), r1, col_index(c2 + "1"), r2)
        self._merge_cache[sheet] = out
        return out

    # -- 自动行高（只在文字会溢出时把行加高，绝不缩小） ----------------------
    def default_row_height(self, sheet):
        xml = self.xml(sheet)
        m = re.search(r'<sheetFormatPr\b[^>]*\bdefaultRowHeight="([\d.]+)"', xml)
        return float(m.group(1)) if m else 14.25

    def row_h(self, sheet, row):
        key = (sheet, row)
        if key not in self._h_cache:
            heights = self.row_heights(self.xml(sheet))
            self._h_cache[key] = heights.get(row, self.default_row_height(sheet))
        return self._h_cache[key]

    def autofit(self):
        """按文字长度估算所需高度：
        单行单元格直接加高本行；纵向合并的单元格把差额加在合并区域最后一行，
        保证文字不会被裁掉，同时尽量少改动模板原有版面。"""
        deltas = {}
        for sheet, ref, text in self.filled:
            if not isinstance(text, str) or not text.strip():
                continue
            widths = self.col_widths(sheet)
            merges = self.merges(sheet)
            col1, row1 = col_index(ref), row_index(ref)
            col2, row2 = col1, row1
            if ref in merges:
                col1, row1, col2, row2 = merges[ref]
            width = sum(widths.get(c, 8.43) for c in range(col1, col2 + 1))
            cap = max(3.0, width)
            lines = 0
            for seg in str(text).split("\n"):
                units = sum(2.0 if ord(ch) > 127 else 1.05 for ch in seg)
                lines += max(1, int(math.ceil(units / cap)))
            need = lines * 13.8 + 4.5
            avail = sum(self.row_h(sheet, r) for r in range(row1, row2 + 1))
            if need > avail * 1.15:            # 留 15% 余量，避免轻微溢出就改版面
                key = (sheet, row2)
                deltas[key] = max(deltas.get(key, 0.0), need - avail)
        for (sheet, row), extra in deltas.items():
            self.set_row_height(sheet, row, min(self.row_h(sheet, row) + extra, 430.0))

    # -- 保存 ----------------------------------------------------------------
    def save(self, out_path):
        tmp = out_path + ".tmp"
        with zipfile.ZipFile(self.template) as src, \
                zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = self.parts.get(info.filename, b"")
                dst.writestr(info, data)
        os.replace(tmp, out_path)


# ----------------------------------------------------------------------------
# 4. 数据填充
# ----------------------------------------------------------------------------
def build_meta(payload):
    meta = {
        "date": None, "date_text": "", "content": "", "subject": "",
        "class_name": "", "teacher": "",
    }
    meta.update(payload.get("meta") or {})
    for k in ("content", "subject", "class_name", "date_text", "teacher"):
        if meta.get(k) is None:
            meta[k] = ""
    if meta.get("date"):
        d = parse_date(meta["date"])
        meta["date"] = d
        if not meta.get("date_text"):
            meta["date_text"] = cn_date(d)
    elif meta.get("date_text"):
        meta["date"] = parse_date(meta["date_text"])
    else:
        raise ValueError("缺少观课时间（meta.date 或 meta.date_text）")
    return meta


def resolve_teachers(payload):
    """把教师安排解析成 {子表名: 教师姓名}

    支持两种写法：
      "teachers": ["张三", "李四", ...]            按【子表1→6】的顺序依次分配
      "teachers": {"师生问答": "张三", ...}          显式指定每张表由谁填
    6 张表都要有人填；名单多于 6 位时多出来的不填（允许有多的老师不填）。
    """
    t = payload.get("teachers")
    if isinstance(t, dict):
        names = {k: str(v).strip() for k, v in t.items() if k in SHEETS}
        blank = [s for s in SHEET_ORDER if not names.get(s)]
        if blank:
            raise ValueError("这些子表还没有指定观课教师：%s" % "、".join(blank))
        return names
    if isinstance(t, list):
        names = [str(x).strip() for x in t if str(x).strip()]
        if len(names) < len(SHEET_ORDER):
            missing = SHEET_ORDER[len(names):]
            raise ValueError(
                "教师名单只有 %d 位，还差 %d 张表没人填（%s）。"
                "请补齐，或指定某位老师兼填多张表。"
                % (len(names), len(missing), "、".join(missing)))
        if len(names) > len(SHEET_ORDER):
            print("  ! 名单 %d 位，本次只用前 %d 位：%s"
                  % (len(names), len(SHEET_ORDER),
                     "、".join(names[len(SHEET_ORDER):])), file=sys.stderr)
        return dict(zip(SHEET_ORDER, names))
    raise ValueError("缺少 teachers（按子表顺序的教师名单，或 {子表: 姓名} 的映射）")


def fill(book, payload):
    meta = build_meta(payload)
    data = payload.get("sheets") or {}
    teachers = resolve_teachers(payload)

    for sheet in SHEET_ORDER:
        spec = SHEETS[sheet]
        content = data.get(sheet) or {}
        m = dict(meta)
        m["teacher"] = teachers.get(sheet, "")

        # 1) 表头信息（每张表的观课教师不同，用 m 而不是 meta）
        for item in spec["meta"]:
            cell = item["cell"]
            kind = item.get("kind", "text")
            if kind == "date":
                book.set_cell(sheet, cell, excel_serial(m["date"]), "number")
                continue
            if item.get("format"):
                value = item["format"].format(**m)
            else:
                value = str(item.get("prefix", "")) + str(m.get(item.get("field"), ""))
            book.set_cell(sheet, cell, value)
            book.filled.append((sheet, cell, value))

        # 2) 固定文本块
        for entry in spec["fields"]:
            cell, key = entry[0], entry[1]
            kind = entry[2] if len(entry) > 2 else "text"
            value = content.get(key)
            if kind == "number":
                if value in (None, ""):
                    continue
                book.set_cell(sheet, cell, value, "number")
                book.filled.append((sheet, cell, str(value)))
            else:
                book.set_cell(sheet, cell, value or "")
                if value:
                    book.filled.append((sheet, cell, value))

        # 3) 小组分工 / 工作情况（成员标签固定在模板里，只填内容）
        if "members" in spec:
            members = content.get("members") or []
            for i, cell in enumerate(spec["members"]["duty_cells"]):
                txt = members[i].get("duty") if i < len(members) else ""
                book.set_cell(sheet, cell, txt or "")
                if txt:
                    book.filled.append((sheet, cell, txt))
            for i, cell in enumerate(spec["members"]["work_cells"]):
                txt = members[i].get("work") if i < len(members) else ""
                book.set_cell(sheet, cell, txt or "")
                if txt:
                    book.filled.append((sheet, cell, txt))

        # 4) 列表型表格（教师提问 / 课堂评价）
        if spec.get("table"):
            tb = spec["table"]
            rows = tb["rows"]
            items = content.get(tb["key"]) or []
            if len(items) > len(rows):
                print("  ! %s 行数上限 %d，多余 %d 条已忽略"
                      % (sheet, len(rows), len(items) - len(rows)), file=sys.stderr)
                items = items[:len(rows)]
            for idx, row in enumerate(rows):
                item = items[idx] if idx < len(items) else {}
                for col, key in tb["cols"].items():
                    cell = "%s%d" % (col, row)
                    value = item.get(key)
                    if key == "no" and not value:
                        value = "" if not item else "问题%s" % "一二三四五六七八九十"[idx]
                    book.set_cell(sheet, cell, value if value is not None else "")
                    if value:
                        book.filled.append((sheet, cell, value))
    return meta, teachers


# ----------------------------------------------------------------------------
# 5. 校验：结构是否与模板完全一致
# ----------------------------------------------------------------------------
def verify(template, out_path):
    problems = []

    def x_count(text, tag):
        return len(re.findall(re.escape(tag), text))

    with zipfile.ZipFile(template) as t, zipfile.ZipFile(out_path) as o:
        tn, on = set(t.namelist()), set(o.namelist())
        if tn - on:
            problems.append("输出缺少部件: %s" % sorted(tn - on))
        if on - tn:
            problems.append("输出多出部件: %s" % sorted(on - tn))
        for name in ("xl/workbook.xml",):
            if name in tn and t.read(name) != o.read(name):
                problems.append("%s 被改动" % name)
        t_draw = {n: t.read(n) for n in tn if n.startswith("xl/drawings/") or n.startswith("xl/media/")}
        o_draw = {n: o.read(n) for n in on if n.startswith("xl/drawings/") or n.startswith("xl/media/")}
        if t_draw.keys() != o_draw.keys():
            problems.append("绘图/图片部件不一致")
        else:
            for k in t_draw:
                if t_draw[k] != o_draw[k]:
                    problems.append("%s 内容被改动（座位图可能受损）" % k)

        # 每个子表合并区域数量一致
        tb = Book(template)
        for sheet in SHEET_ORDER:
            path = tb.sheet_paths[sheet]
            txml = t.read(path).decode("utf-8")
            oxml = o.read(path).decode("utf-8")
            tm = sorted(re.findall(r'<mergeCell ref="([^"]+)"/>', txml))
            om = sorted(re.findall(r'<mergeCell ref="([^"]+)"/>', oxml))
            if tm != om:
                problems.append("%s 合并单元格被改动" % sheet)
            if "<pageSetup" in txml:
                tp = re.search(r"<pageSetup[^>]*/>", txml).group(0)
                op = re.search(r"<pageSetup[^>]*/>", oxml)
                if not op or op.group(0) != tp:
                    problems.append("%s 打印设置被改动" % sheet)

            def _cols(x):
                m = re.search(r"<cols>.*?</cols>", x, re.S)
                return m.group(0) if m else ""

            if _cols(txml) != _cols(oxml):
                problems.append("%s 列宽被改动" % sheet)
            for tag in ("<sheetFormatPr", "<dimension", "<drawing"):
                if x_count(txml, tag) != x_count(oxml, tag):
                    problems.append("%s 的 %s 节点数变化" % (sheet, tag))
    return problems


# ----------------------------------------------------------------------------
# 6. CLI
# ----------------------------------------------------------------------------
def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]', "_", str(s)).strip()


def main(argv=None):
    here = os.path.dirname(os.path.abspath(__file__))
    default_tpl = os.path.join(os.path.dirname(here), "assets", "template.xlsx")

    ap = argparse.ArgumentParser(description="填写观课、议课量化表")
    ap.add_argument("--payload", required=True, help="内容数据 JSON")
    ap.add_argument("--template", default=default_tpl, help="模板 xlsx 路径")
    ap.add_argument("--out", default=None, help="输出目录（默认 当前目录/观课议课输出）")
    ap.add_argument("--no-auto-height", action="store_true", help="不自动调整行高")
    ap.add_argument("--verify", action="store_true", help="生成后校验结构是否与模板一致")
    args = ap.parse_args(argv)

    if not os.path.exists(args.template):
        sys.exit("找不到模板：%s" % args.template)
    with open(args.payload, encoding="utf-8") as f:
        payload = json.load(f)

    out_dir = args.out or os.path.join(os.getcwd(), "观课议课输出")
    os.makedirs(out_dir, exist_ok=True)

    book = Book(args.template)
    try:
        meta, teachers = fill(book, payload)
    except ValueError as e:
        sys.exit("数据有问题：%s" % e)
    if not args.no_auto_height:
        book.autofit()

    fname = "%s_%s_观课议课量化表.xlsx" % (
        meta["date"].strftime("%Y%m%d"),
        safe_name(meta.get("content") or "观课"),
    )
    out_path = os.path.join(out_dir, fname)
    book.save(out_path)

    print("观课时间：%s　内容：%s　学科：%s　班级：%s"
          % (meta["date_text"], meta["content"], meta["subject"], meta["class_name"]))
    print("教师分工（每个子表一位）：")
    for i, sheet in enumerate(SHEET_ORDER, 1):
        print("  %d. %-10s ← %s" % (i, sheet, teachers.get(sheet, "")))

    status = ""
    if args.verify:
        problems = verify(args.template, out_path)
        status = "\n✓ 结构校验通过（合并单元格/列宽/打印设置/内嵌图片与模板一致）" \
            if not problems else "\n✗ " + "；".join(problems)
    print("\n已生成：%s%s" % (out_path, status))
    return 0


if __name__ == "__main__":
    sys.exit(main())
