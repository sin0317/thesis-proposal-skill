#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Renumber a reference list by the order in which each entry first appears in
the main text.

Usage
-----
    # order given on the command line (old numbers, comma separated)
    python refs_renumber.py --refs refs.txt --order 35,36,29,3,4 --out new_refs.txt

    # order read from a JSON file
    python refs_renumber.py --refs refs.json --out new_refs.txt

    # check only
    python refs_renumber.py --refs refs.json --check

Input
-----
refs.txt   one entry per line, optionally already numbered: "[12] 作者. 题名[J]..."
refs.json  {"refs": {"1": "作者. 题名[J]...", "2": "..."},
            "first_appearance": [35, 36, 29, 3, 4]}

Why
---
Ordering by idea ("database signals first, then meta-analysis, then case reports")
almost always breaks the numeric order. Renumbering by hand introduces duplicates
and orphan entries. This script keeps the mapping explicit and verifiable.
"""
import argparse
import json
import re
import sys

NUM = re.compile(r"^\s*\[(\d+)\]\s*(.*)$")


def load_refs(path):
    raw = open(path, encoding="utf-8").read()
    if path.lower().endswith(".json"):
        data = json.loads(raw)
        refs = {int(k): v.strip() for k, v in data["refs"].items()}
        order = data.get("first_appearance")
        return refs, order
    refs, order = {}, []
    for line in raw.splitlines():
        if not line.strip():
            continue
        m = NUM.match(line)
        if m:
            n = int(m.group(1))
            refs[n] = m.group(2).strip()
            order.append(n)
        else:
            refs[len(refs) + 1] = line.strip()
            order.append(len(refs))
    return refs, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--refs", required=True)
    ap.add_argument("--order", help="comma separated old numbers, in first-appearance order")
    ap.add_argument("--out", help="write renumbered list here")
    ap.add_argument("--map", help="write old->new mapping here (markdown)")
    ap.add_argument("--check", action="store_true")
    args = ap.parse_args()

    refs, file_order = load_refs(args.refs)
    if args.order:
        order = [int(x) for x in re.split(r"[,\s]+", args.order.strip()) if x]
    elif file_order:
        order = [int(x) for x in file_order]
    else:
        sys.exit("no order given: pass --order or use a JSON with first_appearance")

    errors = []
    dup = {n for n in order if order.count(n) > 1}
    if dup:
        errors.append("重复引用: %s" % sorted(dup))
    missing = sorted(set(refs) - set(order))
    extra = sorted(set(order) - set(refs))
    if missing:
        errors.append("未被引用的条目: %s" % missing)
    if extra:
        errors.append("引用了不存在的条目: %s" % extra)
    if len(order) != len(refs):
        errors.append("顺序表长度 %d != 文献条数 %d" % (len(order), len(refs)))

    print("文献 %d 条，首次出现顺序 %d 项" % (len(refs), len(order)))
    if errors:
        for e in errors:
            print("  [FAIL]", e)
    else:
        print("  [PASS] 顺序表与文献条目一一对应")

    lines, mapping = [], []
    for new, old in enumerate(order, start=1):
        body = refs.get(old, "（缺失）")
        lines.append("[%d] %s" % (new, body))
        mapping.append("| [%d] | [%d] | %s |" % (new, old, body[:40]))

    if not args.check:
        text = "\n".join(lines)
        if args.out:
            open(args.out, "w", encoding="utf-8").write(text + "\n")
            print("written: %s" % args.out)
        else:
            print()
            print(text)

    if args.map:
        md = "| 新编号 | 原编号 | 条目 |\n|---|---|---|\n" + "\n".join(mapping)
        open(args.map, "w", encoding="utf-8").write(md + "\n")
        print("mapping: %s" % args.map)

    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
