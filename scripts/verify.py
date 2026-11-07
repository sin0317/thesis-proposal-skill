#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""pec-driven checker for a Chinese thesis proposal (.docx).

Usage:
    python verify.py --docx draft.docx --spec spec.json

Spec keys (all optional):
    sections         list of heading strings that must exist, e.g. ["一、","二、"]
    must_contain     list of substrings that must appear (paragraphs + table cells)
    must_not_contain list of substrings that must NOT appear in body scope
    body_tables      how many leading tables belong to the body (default: all)
    refs_start       heading that starts the reference list (default: "参考文献")
    refs             {"enabled": true, "min": 30}
    min_tables       minimum number of tables
    three_line       true -> check tables look like three-line tables
    superscript_refs true -> require [n] to be marked as superscript runs

Exit code 0 only when every check passes.
"""
import argparse
import json
import re
import sys

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("python-docx is required:  pip install python-docx")

WS = re.compile(r"\s+")  # includes U+3000 full-width space
REF = re.compile(r"\[(\d+)(?:[-–,，]\s*(\d+))?\]")


def norm(s):
    return WS.sub("", s)


def load(path):
    d = Document(path)
    return d


def full_text(d, body_tables=None):
    """Paragraphs + table cells. d.paragraphs alone misses every table cell."""
    parts = [p.text for p in d.paragraphs]
    tables = d.tables if body_tables is None else d.tables[:body_tables]
    for t in tables:
        for r in t.rows:
            for c in r.cells:
                parts.append(c.text)
    return "\n".join(parts)


def body_scope(d, body_tables=None, refs_start="参考文献"):
    """Everything before the reference list, used for must_not_contain checks."""
    paras = []
    for p in d.paragraphs:
        if refs_start and p.text.strip().startswith(refs_start):
            break
        paras.append(p.text)
    tables = d.tables if body_tables is None else d.tables[:body_tables]
    for t in tables:
        for r in t.rows:
            for c in r.cells:
                paras.append(c.text)
    return "\n".join(paras)


def ref_numbers(text):
    out = []
    for m in REF.finditer(text):
        a = int(m.group(1))
        out.append(a)
        if m.group(2):
            out.extend(range(a + 1, int(m.group(2)) + 1))
    return out


def first_appearance(d, refs_start="参考文献"):
    seq = []
    for p in d.paragraphs:
        if refs_start and p.text.strip().startswith(refs_start):
            break
        for n in ref_numbers(p.text):
            if n not in seq:
                seq.append(n)
    return seq


def is_three_line(table):
    tblPr = table._tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        return False
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = borders.find(qn("w:" + side))
        if e is None or e.get(qn("w:val")) not in ("nil", "none"):
            return False
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--docx", required=True)
    ap.add_argument("--spec", required=True)
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    spec = json.load(open(args.spec, encoding="utf-8"))
    d = load(args.docx)
    results = []

    def check(name, ok, detail=""):
        results.append((bool(ok), name, detail))

    all_text = full_text(d)
    scope = body_scope(d, spec.get("body_tables"))

    # 1. required sections
    for s in spec.get("sections", []):
        hit = any(p.text.strip().startswith(s) for p in d.paragraphs)
        check("章节 %s" % s, hit, "" if hit else "未找到该标题")

    # 2. must contain
    for s in spec.get("must_contain", []):
        ok = norm(s) in norm(all_text)
        check("含 %s" % s[:24], ok, "" if ok else "全文未命中")

    # 3. must NOT contain (body scope only, so appendix/reply tables do not count)
    for s in spec.get("must_not_contain", []):
        ok = norm(s) not in norm(scope)
        check("不含 %s" % s[:24], ok, "" if ok else "正文仍出现")

    # 4. tables / three-line
    n_tables = len(d.tables)
    if "min_tables" in spec:
        check("表数量 >= %d" % spec["min_tables"], n_tables >= spec["min_tables"],
              "实际 %d" % n_tables)
    if spec.get("three_line"):
        bad = [i for i, t in enumerate(d.tables) if not is_three_line(t)]
        check("三线表", not bad, "非三线表: %s" % bad if bad else "%d 张" % n_tables)

    # 5. superscript refs
    if spec.get("superscript_refs"):
        sup, plain = 0, 0
        for p in d.paragraphs:
            for r in p.runs:
                if REF.search(r.text):
                    if r.font.superscript:
                        sup += 1
                    else:
                        plain += 1
        check("上标引注", sup > 0 and plain == 0,
              "上标 %d / 普通 %d" % (sup, plain))

    # 6. references: order and coverage
    refs_cfg = spec.get("refs") or {}
    if refs_cfg.get("enabled"):
        start = spec.get("refs_start", "参考文献")
        seq = first_appearance(d, start)
        n_refs = len(re.findall(r"^\s*\[\d+\]", "\n".join(
            p.text for p in d.paragraphs), re.M))
        expected = list(range(1, len(seq) + 1))
        check("首次出现顺序 1..n", seq == expected,
              "实际 %s" % (seq[:12] + ["..."] if len(seq) > 12 else seq))
        if refs_cfg.get("min"):
            check("参考文献 >= %d 条" % refs_cfg["min"], n_refs >= refs_cfg["min"],
                  "实际 %d" % n_refs)
        if n_refs:
            missing = [i for i in range(1, n_refs + 1) if i not in seq]
            check("无孤条", not missing, "未引用: %s" % missing if missing else "")

    # report
    fails = [r for r in results if not r[0]]
    for ok, name, detail in results:
        if ok and not args.verbose:
            continue
        print("[%s] %-38s %s" % ("PASS" if ok else "FAIL", name, detail))
    print("\n%d 项检查，FAIL %d；段落 %d，表格 %d，图片 %d" % (
        len(results), len(fails), len(d.paragraphs), n_tables,
        len(d.inline_shapes)))
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
