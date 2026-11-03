# -*- coding: utf-8 -*-
"""集体备课记载表生成 · 一键出稿

输入  payload.json（结构见 references/payload-schema.md）
输出  1) 集体备课记载表-总表-XXX.docx
      2) 备课组成员教学重构-姓名.docx（每位参加教师一份）+ 可选汇编本
      3) 教研活动记录-XXX.docx

用法：
  python run.py --payload payload.json --out "D:/输出目录"
  python run.py --payload payload.json --out DIR --only main
  python run.py --payload payload.json --out DIR --merged
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import docx_build  # noqa: E402
import docx_fill  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
TPL_MAIN = os.path.join(SKILL, "assets", "模板-集体备课记载表-总表.docx")
TPL_ATT = os.path.join(SKILL, "assets", "模板-集体备课记载表-附件.docx")

MAIN_KEYS = ["教材解读", "教学目标", "教学重难点", "中心发言人发言内容", "备课组成员发言", "领导发言提要"]
ATT_KEYS = ["教学目标", "教学重难点", "教学设计", "学生活动"]
BAD_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')


def safe(name: str, fallback="未命名") -> str:
    name = BAD_CHARS.sub("", str(name or "")).strip(" .")
    return name or fallback


def pick(d: dict, *keys, default=""):
    for k in keys:
        if k in d and d[k] not in (None, ""):
            return d[k]
    return default


# ---------------------------------------------------------------- 校验


IDENTITY_KEYS = ("中心发言人", "蹲点领导", "参加人员")
REQUIRED_TOP = ["学科", "年级", "教材内容", "研究主题"]


def apply_defaults(p: dict) -> list[str]:
    """把「不填时的默认」补齐，返回做了什么默认的说明。"""
    notes = []
    if not str(pick(p, "时间", "date")).strip():
        d = __import__("datetime").date.today()
        p["时间"] = f"{d.year}年{d.month}月{d.day}日"
        notes.append(f"时间未给，用系统当天：{p['时间']}")

    grade, subject = pick(p, "年级"), pick(p, "学科")
    teachers = p.get("附件") or p.get("teachers") or []
    names = []
    for t in teachers:
        n = str(pick(t, "教师姓名", "姓名", "name")).strip()
        if n and n not in names:
            names.append(n)

    # 参加人员：一律写全名单上的每一位教师；用户自己写的只要已包含全部姓名就保留原写法
    cur = str(pick(p, "参加人员", "participants")).strip()
    if names:
        if not all(n in cur for n in names):
            p["参加人员"] = "、".join(names)
            notes.append(f"参加人员按名单写全 {len(names)} 人：{p['参加人员']}")
    elif not cur:
        p["参加人员"] = f"{grade}{subject}备课组教师"
        notes.append(f"未给教师名单，参加人员默认「{p['参加人员']}」")

    if not str(pick(p, "中心发言人", "speaker")).strip() and names:
        leader = str(pick(p, "蹲点领导", "supervisor")).strip()
        pool = [n for n in names if n != leader] or names
        p["中心发言人"] = __import__("random").choice(pool)
        notes.append(f"中心发言人未指定，随机取「{p['中心发言人']}」")
    return notes


def check_payload(p: dict, tpl_main: str, tpl_att: str) -> list[str]:
    errs, warns = [], []

    for k in REQUIRED_TOP:
        if not str(pick(p, k, k.lower())).strip():
            errs.append(f"缺少「{k}」——这一项必须给")

    main_data = build_main_data(p)
    main_keys = docx_fill.extract_fields(tpl_main)
    for k in main_keys:
        if not str(main_data.get(k, "")).strip():
            msg = f"总表字段为空：{k}"
            (warns if k in IDENTITY_KEYS else errs).append(msg)

    att_keys = docx_fill.extract_fields(tpl_att)
    skip = {"学科", "年级", "时间", "教材内容", "研究主题", "教师姓名"}
    teachers = p.get("附件") or p.get("teachers") or []
    if not teachers:
        warns.append("未提供参加教师名单，只生成总表与教研活动记录")
    else:
        for i, t in enumerate(teachers, 1):
            name = str(pick(t, "教师姓名", "姓名", "name")).strip()
            if not name:
                errs.append(f"第 {i} 位教师缺「教师姓名」")
                continue
            for k in att_keys:
                if k in skip:
                    continue
                if not str(pick(t, k, default="")).strip():
                    errs.append(f"{name} 缺字段：{k}")
    return errs + [f"[提示] {w}" for w in warns]


def build_main_data(p: dict) -> dict:
    src = p.get("总表") or p.get("main") or {}
    data = {
        "学科": pick(p, "学科", "subject"),
        "年级": pick(p, "年级", "grade"),
        "时间": pick(p, "时间", "date"),
        "教材内容": pick(p, "教材内容", "textbook", "content"),
        "研究主题": pick(p, "研究主题", "topic"),
        "中心发言人": pick(p, "中心发言人", "speaker"),
        "蹲点领导": pick(p, "蹲点领导", "supervisor"),
        "参加人员": pick(p, "参加人员", "participants"),
    }
    for k in MAIN_KEYS:
        data[k] = src.get(k, p.get(k, ""))
    return data


def build_att_data(p: dict, t: dict) -> dict:
    d = {
        "学科": pick(p, "学科", "subject"),
        "年级": pick(p, "年级", "grade"),
        "时间": pick(p, "时间", "date"),
        "教材内容": pick(p, "教材内容", "textbook", "content"),
        "研究主题": pick(p, "研究主题", "topic"),
        "教师姓名": pick(t, "教师姓名", "姓名", "name"),
    }
    for k in ATT_KEYS:
        d[k] = t.get(k, "")
    return d


# ---------------------------------------------------------------- 校验成品


def _parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def _xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def verify_same_layout(tpl: str, out: str) -> list[str]:
    """确认除 document.xml 外逐字节一致，且表格骨架未变。"""
    problems = []
    a, b = _parts(tpl), _parts(out)
    if set(a) != set(b):
        problems.append(f"包内文件不一致：{set(a) ^ set(b)}")
    for n in set(a) & set(b):
        if n != "word/document.xml" and a[n] != b[n]:
            problems.append(f"非正文部件被改动：{n}")

    xa, xb = _xml(tpl), _xml(out)
    for tag, pat in [
        ("表格", r"<w:tbl>"),
        ("行", r"<w:tr\b"),
        ("单元格", r"<w:tc\b"),
        ("列宽", r"<w:gridCol [^>]*/>"),
        ("页面设置", r"<w:sectPr\b.*?</w:sectPr>"),
    ]:
        ca = re.findall(pat, xa, re.S)
        cb = re.findall(pat, xb, re.S)
        if len(ca) != len(cb):
            problems.append(f"{tag}数量变化：模板 {len(ca)} → 成品 {len(cb)}")
        elif tag in ("列宽", "页面设置") and ca != cb:
            problems.append(f"{tag}被改动")

    left = docx_fill.placeholders_in_xml(xb)
    if left:
        problems.append("残留未填占位符：" + "、".join(left))
    return problems


def merge_attachments(docs: list[str], out: str) -> str:
    """把多份附件拼成一份，每人一页；沿用第一份的包结构。"""
    bodies, sect = [], ""
    for i, d in enumerate(docs):
        xml = _xml(d)
        m = re.search(r"<w:body>(.*)</w:body>", xml, re.S)
        inner = m.group(1)
        sm = re.search(r"<w:sectPr\b.*?</w:sectPr>\s*$", inner, re.S)
        if sm:
            if not sect:
                sect = sm.group(0)
            inner = inner[: sm.start()]
        if i:
            inner = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + inner
        bodies.append(inner)

    head = xml[: xml.index("<w:body>") + len("<w:body>")]
    merged = head + "".join(bodies) + sect + "</w:body></w:document>"
    with zipfile.ZipFile(docs[0]) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            raw = merged.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, raw)
    return out


# ---------------------------------------------------------------- 主流程


def main():
    ap = argparse.ArgumentParser(description="集体备课记载表 + 教研活动记录 一键生成")
    ap.add_argument("--payload", required=True)
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--only", choices=["all", "main", "attachments", "record"], default="all")
    ap.add_argument("--merged", action="store_true", help="附件另存一份汇编本（每人一页）")
    ap.add_argument("--tpl-main", default=TPL_MAIN)
    ap.add_argument("--tpl-att", default=TPL_ATT)
    ap.add_argument(
        "--no-name-in-title",
        action="store_true",
        help="附件标题不加教师姓名（默认会加成“……教学重构（张一）”）",
    )
    ap.add_argument("--check", action="store_true", help="只做校验，不写文件")
    args = ap.parse_args()

    with open(args.payload, encoding="utf-8") as f:
        p = json.load(f)

    for note in apply_defaults(p):
        print("·", note)

    issues = check_payload(p, args.tpl_main, args.tpl_att)
    hard = [i for i in issues if not i.startswith("[提示]")]
    if hard:
        print("✗ payload 不完整：")
        for i in hard:
            print("  -", i)
        raise SystemExit(2)
    for i in issues:
        print(" ", i)
    if args.check:
        print("✓ 校验通过")
        return

    os.makedirs(args.out, exist_ok=True)
    stem = safe(pick(p, "教材内容", default="集体备课"), "集体备课")
    made, reports = [], []

    if args.only in ("all", "main"):
        out = os.path.join(args.out, f"集体备课记载表-总表-{stem}.docx")
        r = docx_fill.fill_docx(args.tpl_main, out, build_main_data(p))
        made.append(out)
        reports.append(("总表", out, verify_same_layout(args.tpl_main, out), r))

    att_docs = []
    if args.only in ("all", "attachments"):
        teachers = p.get("附件") or p.get("teachers") or []
        if teachers:
            sub = os.path.join(args.out, "附件-教学重构")
            os.makedirs(sub, exist_ok=True)
            for t in teachers:
                name = safe(pick(t, "教师姓名", "姓名", "name"), "教师")
                out = os.path.join(sub, f"备课组成员教学重构-{name}.docx")
                r = docx_fill.fill_docx(args.tpl_att, out, build_att_data(p, t))
                # 附件模板没有 {{教师姓名}} 占位符，姓名标到标题上，保证一人一份可区分
                has_name_field = "教师姓名" in docx_fill.extract_fields(args.tpl_att)
                if not has_name_field and not args.no_name_in_title:
                    ok = docx_fill.add_suffix_to_paragraph(out, "备课组成员教学重构", f"（{name}）")
                    if not ok:
                        docx_fill.add_suffix_to_paragraph(out, "教学重构", f"（{name}）")
                att_docs.append(out)
                made.append(out)
                reports.append((f"附件·{name}", out, verify_same_layout(args.tpl_att, out), r))
            if args.merged:
                merged = os.path.join(args.out, "备课组成员教学重构-汇编.docx")
                merge_attachments(att_docs, merged)
                made.append(merged)

    if args.only in ("all", "record"):
        rec = dict(p.get("教研活动记录") or {})
        rec.setdefault("活动主题", f"{pick(p,'学科')}{pick(p,'年级')}备课组集体备课活动")
        rec.setdefault("活动时间", pick(p, "时间"))
        rec.setdefault("主持人", pick(p, "中心发言人"))
        rec.setdefault("教研主题", pick(p, "研究主题"))
        # 参加人员与总表保持一致：写全名单
        people = str(rec.get("参加人员", "")).strip()
        tnames = [
            str(pick(t, "教师姓名", "姓名", "name")).strip()
            for t in (p.get("附件") or p.get("teachers") or [])
        ]
        tnames = [n for n in tnames if n]
        if not people or (tnames and not all(n in people for n in tnames)):
            rec["参加人员"] = pick(p, "参加人员")
        if not str(rec.get("教研内容", "")).strip():
            print("✗ 缺“教研活动记录.教研内容”")
            raise SystemExit(2)
        # 教研组长补充只作写作用素材：给了要写进正文，没给就当没有这回事。脚本不补、不塞、不改写正文，只做提示。
        extra = str(pick(p, "教研组长补充", "组长补充", "leader_extra")).strip()
        if extra:
            probe = (extra.split("\n")[0].strip() or extra)[:10]
            if probe and probe not in str(rec.get("教研内容", "")):
                print("! 已给「教研组长补充」，但教研内容里没看到对应内容 —— 请回到 payload 把它写进记录（领导意见之前单立一节），脚本不会代写")
        out = os.path.join(args.out, f"教研活动记录-{stem}.docx")
        docx_build.build_record_docx(rec, out)
        made.append(out)

    print("\n================ 生成结果 ================")
    bad = 0
    for name, path, probs, r in reports:
        if probs:
            bad += 1
            print(f"✗ {name}：{os.path.basename(path)}")
            for x in probs:
                print("    -", x)
        else:
            print(f"✓ {name}：{os.path.basename(path)}  （填充 {len(r['used'])} 个字段）")
    for path in made:
        if not any(path == r[1] for r in reports):
            print(f"✓ {os.path.basename(path)}")
    print(f"\n共 {len(made)} 个文件 → {args.out}")
    if bad:
        raise SystemExit(3)


if __name__ == "__main__":
    try:
        main()
    except docx_fill.FileBusy as e:
        import os as _os

        print(f"✗ 文件被占用，写不进去：{_os.path.basename(e.target)}")
        print(f"  {_os.path.dirname(e.target)}")
        print("  这份 docx 正在 Word 或预览窗口里打开着，关掉它再跑一次。")
        raise SystemExit(4)
