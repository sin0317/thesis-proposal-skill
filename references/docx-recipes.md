# python-docx 排版配方

本文件只放可直接复制的代码片段。全部在 python-docx 1.1+ 实测通过。

## 1. 中文字体（最容易错的一步）

`run.font.name` 只设西文字体，中文字符仍走默认字体。**必须同时设 `w:eastAsia`**：

```python
from docx.oxml.ns import qn

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
```

## 2. 正文段落（首行缩进 2 字符 + 1.5 倍行距）

```python
def body(doc, text, size=12):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.first_line_indent = Pt(size * 2)   # 小四 12pt → 24pt
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    set_cn_font(p.add_run(text), size=size)
    return p
```

## 3. 三线表

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt

def _borders(el_pr, tag, spec):
    """spec: {'top': 1.5, 'bottom': 0.75, ...} 单位 pt"""
    tag_el = el_pr.find(qn(tag))
    if tag_el is None:
        tag_el = OxmlElement(tag)
        el_pr.append(tag_el)
    for side, w in spec.items():
        e = tag_el.find(qn("w:" + side))
        if e is None:
            e = OxmlElement("w:" + side)
            tag_el.append(e)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(int(w * 8)))   # sz 单位是 1/8 pt
        e.set(qn("w:color"), "000000")

def three_line_table(doc, header, rows, widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    # 1) 整表先清空所有框线
    tblPr = t._tbl.tblPr
    none = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + side)
        e.set(qn("w:val"), "nil")
        none.append(e)
    tblPr.append(none)
    # 2) 表头
    for i, txt in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cn_font(p.add_run(txt), size=10.5, bold=True)
        _borders(c._tc.get_or_add_tcPr(), "w:tcBorders",
                 {"top": 1.5, "bottom": 0.75})
    # 3) 数据行
    for r, row in enumerate(rows, start=1):
        for i, val in enumerate(row):
            c = t.rows[r].cells[i]
            c.text = ""
            p = c.paragraphs[0]
            set_cn_font(p.add_run(str(val)), size=10.5)
    # 4) 底线
    last = t.rows[-1]
    for c in last.cells:
        _borders(c._tc.get_or_add_tcPr(), "w:tcBorders", {"bottom": 1.5})
    return t
```

## 4. 上标引注

不要写成普通文本 `[3]`，必须是上标 run：

```python
def add_ref(paragraph, num):
    r = paragraph.add_run(f"[{num}]")
    r.font.superscript = True
    return r
```

一段内混排：正文 run → `add_ref()` → 正文 run。**不要在长文本里用替换的方式补上标**，会破坏原有格式。

## 5. 参考文献悬挂缩进

```python
def reference(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Pt(24)
    pf.first_line_indent = Pt(-24)
    pf.line_spacing = 1.5
    set_cn_font(p.add_run(text), size=10.5)
```

## 6. 图片嵌入（技术路线图）

```python
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run().add_picture("route.png", width=Pt(440))   # 约 15.5 cm，A4 版心内
```

图题另起一段，居中、五号、加粗，位于图片**下方**。

## 7. 合并单元格的坑

`cell_a.merge(cell_b)` 后，被合并区域内的文字**会全部保留并串联**，导致标签被内容覆盖。

正确顺序：**先 merge，再写文字**。

```python
row = t.add_row()
c0 = row.cells[0].merge(row.cells[4])   # 先合并
c0.text = ""                            # 再清空
c0.paragraphs[0].add_run("合计").bold = True
```

反例（会出问题）：

```python
row.cells[0].text = "合计"
row.cells[0].merge(row.cells[4])   # 合并后"合计"可能被后面的空单元格吞掉或重复
```

## 8. 读取时不要把表格漏掉

```python
def full_text(path):
    d = Document(path)
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for r in t.rows:
            for c in r.cells:
                parts.append(c.text)
    return "\n".join(parts)
```

`d.paragraphs` 只含文档主体段落，**表格单元格里的段落不在其中**。用 `d.paragraphs` 做内容校验，会得到大量假阴性。

## 9. 检查图片是否真的嵌入

```python
import zipfile
names = zipfile.ZipFile(path).namelist()
print([n for n in names if n.startswith("word/media/")])
```

`len(doc.inline_shapes)` 统计的是 inline shape 数量，与图片数不一定相等（还有图表等类型），用 zip 里的 `word/media/` 更可靠。

## 10. 修改已有文档时的安全姿势

1. 先 `cp` 出副本，改副本
2. 修改前把目标段落的 `rPr`（字体属性）深拷贝保存，写新文字后套回去，避免字体丢失
3. 同一路径反复生成会命中编辑器缓存，改完用 `zipfile` 校验 `word/document.xml` 的修改时间
