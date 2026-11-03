# -*- coding: utf-8 -*-
"""在 OOXML 包内部替换 {{占位符}}，除被替换的文字外，模板逐字节保持原样。

只用标准库（zipfile + re），不依赖 python-docx。
适用于占位符被 Word 拆成多个 run 的情况（如 {{ / 年级 / }} 三段）。

用法（被 run.py 调用，也可单独用）：
    python docx_fill.py --template t.docx --out o.docx --data data.json
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import zipfile

# <w:p ...>  </w:p>   —— <w:pPr>/<w:pStyle> 之类不会被误匹配
PARA_RE = re.compile(r"<w:p(?:\s[^>]*)?>.*?</w:p>", re.S)
RUN_RE = re.compile(r"<w:r(?:\s[^>]*)?>.*?</w:r>", re.S)
T_RE = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)
BR_RE = re.compile(r"<w:(?:br|cr)\b[^>]*/>", re.S)
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}\n]{1,40}?)\s*\}\}")


class FileBusy(PermissionError):
    """目标 docx 正被 Word / 预览窗口占用，无法替换。"""

    def __init__(self, path: str):
        super().__init__(f"文件被占用：{path}")
        self.target = path


def _replace(tmp: str, out: str) -> None:
    try:
        os.replace(tmp, out)  # 原子替换：目标被占用时不会留下半截文件
    except PermissionError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise FileBusy(out) from None


def _unescape(s: str) -> str:
    return (
        s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def _escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _run_text(run_xml: str) -> str:
    """取一个 run 的可见文字；<w:br/> 视为换行。"""
    out = []
    for m in re.finditer(
        r"<w:t(?:\s[^>]*)?>.*?</w:t>|<w:(?:br|cr)\b[^>]*/>", run_xml, re.S
    ):
        seg = m.group(0)
        if seg.startswith("<w:t"):
            inner = T_RE.match(seg)
            out.append(_unescape(inner.group(1)) if inner else "")
        else:
            out.append("\n")
    return "".join(out)


def _run_props(run_xml: str) -> str:
    m = re.search(r"<w:rPr>.*?</w:rPr>", run_xml, re.S)
    return m.group(0) if m else ""


def _build_run(run_props: str, text: str) -> str:
    """把一段可能带换行的文字装进一个 run，换行转 <w:br/>。"""
    if not text:
        return f"<w:r>{run_props}</w:r>"
    parts = []
    for i, seg in enumerate(text.split("\n")):
        if i:
            parts.append("<w:br/>")
        if seg:
            parts.append(f'<w:t xml:space="preserve">{_escape(seg)}</w:t>')
    return f"<w:r>{run_props}{''.join(parts)}</w:r>"


def _fill_paragraph(para_xml: str, data: dict, used: set, missing: set) -> str:
    if "{{" not in para_xml:
        return para_xml

    runs = list(RUN_RE.finditer(para_xml))
    if not runs:
        return para_xml

    full = "".join(_run_text(r.group(0)) for r in runs)
    if "{{" not in full:
        return para_xml

    def repl(m):
        key = m.group(1).strip()
        if key in data:
            used.add(key)
            return str(data[key]).replace("\r\n", "\n").replace("\r", "\n")
        missing.add(key)
        return m.group(0)

    new_full = PLACEHOLDER_RE.sub(repl, full)
    if new_full == full:
        return para_xml

    # 保留首个 run 的字体格式，其余 run 一并并进去
    head = _build_run(_run_props(runs[0].group(0)), new_full)
    start, end = runs[0].start(), runs[-1].end()
    return para_xml[:start] + head + para_xml[end:]


def fill_document_xml(xml: str, data: dict) -> tuple[str, set, set]:
    used, missing = set(), set()

    def para_repl(m):
        return _fill_paragraph(m.group(0), data, used, missing)

    new_xml = PARA_RE.sub(para_repl, xml)
    return new_xml, used, missing


def fill_docx(template: str, out: str, data: dict) -> dict:
    """复制模板 → 只改 word/document.xml → 写回。返回替换统计。"""
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)

    used, missing = set(), set()
    tmp = out + ".tmp"
    with zipfile.ZipFile(template) as zin:
        names = zin.namelist()
        if "word/document.xml" not in names:
            raise SystemExit(f"[错误] 不是有效的 docx：{template}")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                raw = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    xml = raw.decode("utf-8")
                    xml, used, missing = fill_document_xml(xml, data)
                    raw = xml.encode("utf-8")
                zout.writestr(item, raw)
    _replace(tmp, out)

    left = placeholders_in_xml(_all_text(out))
    return {
        "out": out,
        "used": sorted(used),
        "missing": sorted(missing),
        "left": sorted(left),
    }


def _all_text(path: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def add_suffix_to_paragraph(path: str, anchor: str, suffix: str) -> bool:
    """在含 anchor 的段落末尾追加 suffix（用于把教师姓名标到附件标题上）。

    模板没有 {{教师姓名}} 占位符时使用；只在段落末尾加文字，不动任何格式。
    """
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    xml = dict((i.filename, raw) for i, raw in items)["word/document.xml"].decode("utf-8")

    for m in PARA_RE.finditer(xml):
        para = m.group(0)
        if anchor not in "".join(_run_text(r.group(0)) for r in RUN_RE.finditer(para)):
            continue
        runs = list(RUN_RE.finditer(para))
        last = runs[-1]
        run_xml = last.group(0)
        # 在最末一个 <w:t> 后追加文字
        t_matches = list(T_RE.finditer(run_xml))
        if not t_matches:
            continue
        t = t_matches[-1]
        new_t = (
            run_xml[: t.end() - len("</w:t>")]
            + _escape(suffix)
            + "</w:t>"
            + run_xml[t.end() :]
        )
        new_para = para[: last.start()] + new_t + para[last.end() :]
        new_xml = xml[: m.start()] + new_para + xml[m.end() :]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item, raw in items:
                zout.writestr(
                    item, new_xml.encode("utf-8") if item.filename == "word/document.xml" else raw
                )
        return True
    return False


def placeholders_in_xml(xml: str) -> list:
    """按「段落内 run 拼接后的文字」找占位符。

    必须这样做：Word 常把 {{ / 字段名 / }} 拆到多个 run 里，直接扫原始 XML 会漏。
    """
    keys, seen = [], set()
    for pm in PARA_RE.finditer(xml):
        para = pm.group(0)
        if "{{" not in para:
            continue
        text = "".join(_run_text(r.group(0)) for r in RUN_RE.finditer(para))
        for k in PLACEHOLDER_RE.findall(text):
            k = k.strip()
            if k not in seen:
                seen.add(k)
                keys.append(k)
    return keys


def extract_fields(template: str) -> list:
    """读出模板里的占位符（按出现顺序、去重）。"""
    return placeholders_in_xml(_all_text(template))


def _main():
    ap = argparse.ArgumentParser(description="按 {{占位符}} 填充 docx，保持模板格式不变")
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--data", required=True, help="JSON 文件，键名=占位符名")
    args = ap.parse_args()

    with open(args.data, encoding="utf-8") as f:
        data = json.load(f)
    res = fill_docx(args.template, args.out, data)
    print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    _main()
