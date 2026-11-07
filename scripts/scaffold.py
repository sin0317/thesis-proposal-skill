#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Build a blank (or partially filled) thesis-proposal skeleton with correct
Chinese typography: A4, 90/72 pt margins, SimSun 12pt body, 1.5 line spacing,
2-character first-line indent, HeiTi headings, three-line tables.

Usage:
    python scaffold.py --out 开题报告_骨架.docx --outline outline.json
    python scaffold.py --out draft.docx            # built-in seven-part skeleton
    python scaffold.py --out draft.docx --sections-only

outline.json schema:
{
  "meta": {"姓名": "", "学号": "", "专业": "", "指导教师": ""},
  "title": "本科毕业论文（设计）开题报告",
  "blocks": [
    {"type": "h1", "text": "一、本课题研究意义"},
    {"type": "h2", "text": "（一）理论意义"},
    {"type": "body", "text": "……"},
    {"type": "h3", "text": "1. 文献研究法"},
    {"type": "ref", "text": "[1] 作者. 题名[J]. 刊名, 年, 卷(期): 页."},
    {"type": "table", "caption": "表 1　关键问题与对应措施",
     "header": ["序号", "关键问题", "应对措施"], "rows": [["1", "…", "…"]]},
    {"type": "caption", "text": "图 1　技术路线图（作者自绘）"},
    {"type": "image", "path": "route.png", "width": 440}
  ]
}
"""
import argparse
import json
import os
import sys

try:
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
except ImportError:
    sys.exit("python-docx is required:  pip install python-docx")

DEFAULT = [
    ("h1", "一、本课题研究意义"),
    ("h2", "（一）理论意义"),
    ("body", "（此处写：现有证据为何不足、分歧在哪、为什么汇总的前提尚未建立。）"),
    ("h2", "（二）实践意义"),
    ("body", "（此处写：结果落到谁的动作上——监护重点、药事管理、制剂选择。）"),
    ("h1", "二、国内外有关本课题的研究动态"),
    ("h2", "（一）国外研究进展"),
    ("body", "（数据库信号检测 → 系统评价/定量合成 → 个案报道；以“综上，……但不充分之处在于……”收尾。）"),
    ("h2", "（二）国内研究进展"),
    ("body", "（总体特征 → 研究内容三方面 → 证据质量短板；以“综上，……尚不充分”收尾。）"),
    ("h2", "（三）研究空白"),
    ("body", "（“综上可见……因此，本研究拟……”——必须落到研究内容。）"),
    ("h1", "三、本课题研究的基本内容"),
    ("body", "1. （建立××数据集。）"),
    ("body", "2. （描述××特征。）"),
    ("body", "3. （识别××风险因素。）"),
    ("body", "4. （提出××建议。）"),
    ("h1", "四、本课题拟解决的主要问题"),
    ("body", "（只留 1 个核心问题；问题句 + 2—3 句“靠什么解决”，具体操作留到研究方法。）"),
    ("h1", "五、研究方法"),
    ("h3", "1. 文献研究法"),
    ("body", "（学术依据：方法体系中的位置、源流、公认规范。）"),
    ("body", "（定义 + 核心内涵 + 3 条基本特征，此处不引文献。）"),
    ("body", "（应用于研究内容第 1 项：具体做法与产出，引注可放这里。）"),
    ("h3", "2. 内容分析法"),
    ("body", "（同上三段式。）"),
    ("h3", "3. 统计分析法"),
    ("body", "（同上三段式。）"),
    ("body", "（文末一句：质量评分、抽查复核、一致性检验属质控措施，不单独构成研究方法。）"),
    ("h1", "六、主要创新点"),
    ("body", "（1）（与已有文献对比：别人怎么做 → 本研究怎么做。禁用“首次”“无人做过”。）"),
    ("h1", "七、主要参考文献"),
    ("ref", "[1] （按正文首次出现顺序编号，GB/T 7714—2015。）"),
    ("h1", "研究进度安排"),
    ("body", "2026年9月1日至11月15日　（工作内容 + 产出。）"),
]


def set_cn_font(run, cn="宋体", en="Times New Roman", size=12, bold=False):
    run.font.name = en
    run.font.size = Pt(size)
    run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), cn)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)


def para(doc, text="", size=12, bold=False, cn="宋体", align=None,
         indent=True, line=1.5):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if indent:
        pf.first_line_indent = Pt(size * 2)
    if align is not None:
        p.alignment = align
    if text:
        set_cn_font(p.add_run(text), cn=cn, size=size, bold=bold)
    return p


def _borders(tcPr, spec):
    tag = "w:tcBorders"
    el = tcPr.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        tcPr.append(el)
    for side, w in spec.items():
        e = el.find(qn("w:" + side))
        if e is None:
            e = OxmlElement("w:" + side)
            el.append(e)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(int(w * 8)))
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")


def three_line_table(doc, header, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + side)
        e.set(qn("w:val"), "nil")
        borders.append(e)
    t._tbl.tblPr.append(borders)

    for i, txt in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.0
        set_cn_font(p.add_run(str(txt)), size=10.5, bold=True)
        _borders(c._tc.get_or_add_tcPr(), {"top": 1.5, "bottom": 0.75})

    for r, row in enumerate(rows, start=1):
        for i, val in enumerate(row):
            if i >= len(row):
                break
            c = t.rows[r].cells[i]
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.line_spacing = 1.0
            set_cn_font(p.add_run(str(val)), size=10.5)

    for c in t.rows[-1].cells:
        _borders(c._tc.get_or_add_tcPr(), {"bottom": 1.5})
    return t


def render(doc, blocks):
    for b in blocks:
        typ = b.get("type", "body")
        if typ == "h1":
            para(doc, b["text"], size=14, bold=True, cn="黑体", indent=False)
        elif typ == "h2":
            para(doc, b["text"], size=12, bold=True, cn="黑体", indent=False)
        elif typ == "h3":
            para(doc, b["text"], size=12, bold=True, indent=False)
        elif typ == "caption":
            para(doc, b["text"], size=10.5, bold=True, indent=False,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
        elif typ == "image":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(b["path"], width=Pt(b.get("width", 440)))
        elif typ == "table":
            if b.get("caption"):
                para(doc, b["caption"], size=10.5, bold=True, indent=False,
                     align=WD_ALIGN_PARAGRAPH.CENTER)
            three_line_table(doc, b["header"], b["rows"])
        elif typ == "ref":
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.left_indent = Pt(24)
            pf.first_line_indent = Pt(-24)
            pf.line_spacing = 1.5
            set_cn_font(p.add_run(b["text"]), size=10.5)
        else:
            para(doc, b.get("text", ""))


def setup(doc):
    s = doc.sections[0]
    s.page_width = Pt(595)
    s.page_height = Pt(842)
    s.left_margin = Pt(90)
    s.right_margin = Pt(90)
    s.top_margin = Pt(72)
    s.bottom_margin = Pt(72)
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--outline")
    ap.add_argument("--sections-only", action="store_true",
                    help="only headings, no placeholder body text")
    args = ap.parse_args()

    doc = Document()
    setup(doc)

    if args.outline:
        cfg = json.load(open(args.outline, encoding="utf-8"))
        if cfg.get("title"):
            para(doc, cfg["title"], size=16, bold=True, cn="黑体",
                 indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        meta = cfg.get("meta") or {}
        if meta:
            line = "　　".join("%s：%s" % (k, v) for k, v in meta.items())
            para(doc, line, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        blocks = cfg.get("blocks") or [
            {"type": t, "text": x} for t, x in DEFAULT]
    else:
        blocks = [{"type": t, "text": x} for t, x in DEFAULT]

    if args.sections_only:
        blocks = [b for b in blocks
                  if b.get("type") in ("h1", "h2", "h3")]

    render(doc, blocks)
    if os.path.dirname(args.out):
        os.makedirs(os.path.dirname(args.out), exist_ok=True)
    doc.save(args.out)
    print("written: %s (%d blocks)" % (args.out, len(blocks)))


if __name__ == "__main__":
    main()
