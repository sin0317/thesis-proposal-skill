# thesis-proposal-cn

> 中文本科/硕士毕业论文**开题报告**的撰写与返修 skill：七部分写作规则、逐条返修流程、字段可得性核验、程序化验收脚本、python-docx 排版配方。

沉淀自药学专业文献分析类课题（19 篇中文病例系列 + 9 篇英文文献，四轮真实返修）。适用于医学、护理、教育、管理类的**文献研究型**课题；实验/问卷类课题可复用结构与排版部分。

## 这个 skill 解决什么

写开题报告难的不是"写"，是这四件事：

1. **把过程讲成计划**：研究方法要写成"学术依据 → 定义与特征 → 应用在哪一步"，而不是罗列七个方法名
2. **把结论讲出边界**：手里是分子没有分母，就只能说"在汇总的 X 例中占比"，不能说"发生率为"
3. **让引注站得住**：按首次出现顺序编号、上标标注、[1]—[n] 全覆盖、不留孤条
4. **让导师的每条意见都有落点**：A 类硬伤 / B 类应改 / C 类形式，逐条对应到"改哪一段、改成什么"

## 安装

**方式一：直接拷进用户技能目录**（推荐）

```bash
git clone https://github.com/sin0317/thesis-proposal-skill.git \
  ~/.workbuddy/skills/thesis-proposal-cn
```

Windows（Git Bash）：

```bash
git clone https://github.com/sin0317/thesis-proposal-skill.git \
  "C:/Users/<你的用户名>/.workbuddy/skills/thesis-proposal-cn"
```

**方式二：只装到当前项目**

```bash
git clone https://github.com/sin0317/thesis-proposal-skill.git \
  ./.workbuddy/skills/thesis-proposal-cn
```

**方式三：下载 ZIP**，解压后把整个目录改名为 `thesis-proposal-cn`，放进 `~/.workbuddy/skills/`。

装好后新建会话，输入"帮我写开题报告"或"按这份导师意见改开题报告"即可触发。

依赖：`python-docx`（脚本用）。缺了先装：

```bash
pip install python-docx
```

## 目录结构

```
thesis-proposal-cn/
├── SKILL.md                       # 技能主文件（frontmatter + 工作流）
├── README.md
├── LICENSE
├── references/
│   ├── structure.md               # 七部分骨架、字数配额、编号体系、表 A
│   ├── writing-rules.md           # 各部分写作细则与可套用句式
│   ├── style-and-citation.md      # 字体字号、三线表、GB/T 7714 核对表
│   ├── docx-recipes.md            # python-docx 配方（中文字体/三线表/上标/合并）
│   └── pitfalls.md                # 14 条真实踩坑与排查顺序
├── scripts/
│   ├── scaffold.py                # 生成骨架 docx（A4、宋体小四、1.5 倍行距）
│   ├── verify.py                  # 按 spec 逐项验收
│   └── refs_renumber.py           # 参考文献按首次出现顺序重排
└── examples/
    ├── 01-outline.json            # scaffold 输入示例
    ├── 02-verify-spec.json        # verify 输入示例
    ├── 03-teacher-opinion.md      # 导师意见原件（脱敏）
    ├── 04-revision-log.md         # 逐条回复表模板
    └── 05-real-case.md            # 四轮返修复盘
```

## 用法

### 1. 生成骨架

```bash
python scripts/scaffold.py --out 开题报告_骨架.docx
python scripts/scaffold.py --out 开题报告.docx --outline examples/01-outline.json
python scripts/scaffold.py --out 只留标题.docx --sections-only
```

输出已带 A4 页面、90/72 pt 页边距、宋体小四正文、1.5 倍行距、首行缩进 2 字符、黑体标题、三线表。

### 2. 参考文献重排

```bash
# 从 JSON 读顺序表
python scripts/refs_renumber.py --refs refs.json --out refs_new.txt --map mapping.md

# 或命令行直接给顺序
python scripts/refs_renumber.py --refs refs.txt --order 35,36,29,3,4

# 只检查，不改文件
python scripts/refs_renumber.py --refs refs.json --check
```

它会报"重复引用""未被引用的条目""引用了不存在的条目"三类错误，并在退出码上体现。

`refs.json` 格式：

```json
{
  "refs": {"1": "作者. 题名[J]. 刊名, 2020, 12(3):45-48.", "2": "..."},
  "first_appearance": [35, 36, 29, 3, 4]
}
```

### 3. 逐项验收

```bash
python scripts/verify.py --docx 修订稿.docx --spec examples/02-verify-spec.json
python scripts/verify.py --docx 修订稿.docx --spec spec.json --verbose
```

spec 各项含义：

| 键 | 作用 |
|---|---|
| `sections` | 必须存在的标题前缀，如 `["一、","二、"]` |
| `must_contain` | 全文（段落 + 表格）必须命中的字符串 |
| `must_not_contain` | 正文范围（不含参考文献之后、不含附录表）不得出现的字符串 |
| `body_tables` | 前多少张表属于正文，用于排除文末回复表 |
| `refs_start` | 参考文献起始标题，默认"参考文献" |
| `refs` | `{"enabled": true, "min": 30}`：检查顺序、条数、孤条 |
| `superscript_refs` | 要求 [n] 为**上标** run |
| `three_line` | 检查表格是否为三线表 |

退出码 0 表示全部通过，可直接进 CI。

## 三条最值钱的规则

1. **`d.paragraphs` 不含表格文本**。拿它做内容校验，表格条目全部假阴性。用 `full_text()` 把 `d.tables` 一并展开。
2. **例数口径必须逐篇核对**。题名标 126 例、正文可能只有 60 例；不核对就会出现"题目 1701、正文 1767"的打架。
3. **先怀疑校验串，再怀疑文档**。真实项目里 6 条 FAIL 全部是假阴性（全角空格、表格漏扫、措辞相近），没有一条是内容缺失。

## 已知限制

- 不能生成 Word 修订痕迹（`w:ins`/`w:del`）。要交带痕迹的稿子，用 Word「审阅 → 比较 → 比较文档」。
- 不替用户查重、不做文献检索代理，只处理"已有材料 → 合规文档"这一段。
- 学科差异较大时（如实验设计、量表编制），方法学部分需另配该领域专家，本 skill 只保证结构与表述合规。
- 案例来自医药类课题，非医药领域的示例表述需自行替换术语。

## 许可

MIT。见 `LICENSE`。

## 版本

1.0.0 — 首版。七部分写作规则、表 A 字段核验、三个脚本、五份示例。
