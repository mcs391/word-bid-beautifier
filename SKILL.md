---
name: ryusuke-bid-beautifier
description: "Word投标文档一键优化工具。当用户要求优化Word文档的标题样式、缩进排版、编号格式、字体间距等美化操作时触发，特别适用于ryusuke格式的投标技术文件。支持自动识别四级/五级标题、修复层级编号断层、清除重复编号、调整缩进值、根本性修复H2/H3编号分裂(三层面统一)、优化字体和段落间距。"
---

# Word 投标文档一键优化工具 (ryusuke-bid-beautifier) V3.1

## 概述

本 Skill 提供**一条龙**的 Word 投标文档全流程优化方案，整合**三大核心能力**：

### Phase 0 — 根本性编号修复（三层面统一）⭐ V2
> **核心价值：修复后手动套用 ryusuke标题2/3 样式时，Word 会自动继承正确的章节号计数**

解决「第14章下手动套用 H2 样式后编号显示 9.x 而非 14.x」的根本问题：

| 层面 | 文件 | 修复内容 | 解决的问题 |
|------|------|----------|-----------|
| **层面一** | `numbering.xml` | 补全主多级列表 `abstractNumId=0` 的 pStyle 链接 (ilvl=1→style13, ilvl=2→style14) | 多级列表定义缺少样式绑定 |
| **层面二** | `styles.xml` | 统一所有标题样式(12~16)的 numPr 指向同一个多级列表实例(numId=1→abs#0) | 不同样式指向不同 abstractNum 导致"编号分裂" |
| **层面三** | `document.xml` | 为缺失 numPr 的标题段落(12~16)插入正确编号引用 | 段落级缺少编号引用 |

利用 Word 多级列表定义（abstractNumId=1）：
```
ilvl=0: 第%1章 → style 12 (章节标题)
ilvl=1: %1.%2.   → style 13 (二级标题 H2)
ilvl=2: %1.%2.%3.→ style 14 (三级标题 H3)
ilvl=3: %1.%2.%3.%4. → style 15 (四级标题 H4)
ilvl=4: %1.%2.%3.%4.%5. → style 16 (五级标题 H5)
```

### Phase 1 — 编号层级修复
修复长篇投标文件中常见的**标题编号断层问题**：

| 层级 | 错误格式 | 修复后 |
|------|----------|--------|
| 四级 | ~~`1.文化传承与空间表达融合`~~ | **`9.1.1.1 文化传承...`** |
| 五级 | ~~`1.1传统意象转译`~~ | **`9.1.1.1.1 传统意象...`** |

同时清除正文中被误加的数字前缀。

### Phase 2 — 样式美化
基于中文技术文档排版最佳实践，全面优化视觉效果：

- **智能套用 ryusuke 标题样式** — 自动匹配 `X.X.X.X` / `X.X.X.X.X` 格式的四/五级标题
- **统一样式面板命名** — 将当前文档 1~5 级标题显示名统一为 `ryusuke标题1` 至 `ryusuke标题5`
- **清除重复编号** — 处理样式自动编号与文本硬编码冲突
- **正文首行缩进** — 为无样式的中文正文添加标准缩进
- **多级列表缩进优化** — 调整 numbering.xml 中各级 left/hanging 值
- **字体/间距精修** — 统一黑体、明确字号、优化段前/后距和行距
- **表格正文样式统一** — 自动识别并将历史表格正文样式重命名/固化为 `ryusuke-表格正文`，用于表格内容美化

**核心原则：零内容删除，仅修改格式属性。**

## 触发条件

当用户提出以下类型需求时使用此 Skill：

- "帮我优化这个投标文件" / "美化Word文档"
- "修一下编号" / "编号层级有问题" / "`1.` 应该是 `9.1.1.1`"
- "第10章的二级标题还是显示9.6" / "编号没有跟随章节更新"
- "套用ryusuke标题样式" / "应用四/五级标题格式"
- "标题缩进太多" / "标题太靠右了" / "调整左边距"
- "重复编号" / "编号出现两遍"
- "让整篇文档更美观" / "统一排版风格"

## 快速使用

### 方式一：脚本一键执行（推荐）

```bash
python scripts/bid_doc_optimizer.py <输入.docx> <输出.docx> [选项]
```

**参数说明：**

| 参数 | 必需 | 默认值 | 说明 |
|------|------|--------|------|
| `<输入.docx>` | ✅ | — | 待优化的 Word 文档 |
| `<输出.docx>` | ✅ | — | 输出路径 |
| `--phase {0\|1\|2\|all}` | ❌ | all | 执行阶段 (0=H2/H3编号, 1=层级编号, 2=样式美化) |
| `--indent-h4 N` | ❌ | 360 | H4 缩进 twips (~0.63cm) |
| `--indent-h5 N` | ❌ | 480 | H5 缩进 twips (~0.85cm) |
| `--ch-start N` | ❌ | auto | 编号修复起始章节号 |
| `--ch-end N` | ❌ | auto | 编号修复结束章节号 |
| `--dry-run` | ❌ | off | 仅分析不修改 |
| `--no-strip` | ❌ | off | 跳过清除多余编号 |

**常用命令示例：**

```bash
# 完整优化（推荐，默认执行全部流程 Phase 0+1+2）
python scripts/bid_doc_optimizer.py input.docx output.docx

# 仅修复H2/H3编号跟随问题
python scripts/bid_doc_optimizer.py input.docx output.docx --phase 0

# 只修H4/H5层级编号，不改样式
python scripts/bid_doc_optimizer.py input.docx output.docx --phase 1

# 只做样式美化
python scripts/bid_doc_optimizer.py input.docx output.docx --phase 2

# 更紧凑的缩进 + 只处理第9~12章
python scripts/bid_doc_optimizer.py input.docx output.docx \
  --indent-h4 240 --indent-h5 320 --ch-start 9 --ch-end 12
```

### 方式二：分步手动操作（需精细控制时）

#### Phase 0 分步详解：根本性编号修复（三层面统一）

**Step A — 层面一：修复 numbering.xml**
在 `abstractNumId=0`（主多级列表）中补全 pStyle 链接：
- 检查 ilvl=1 是否有 `<w:pStyle w:val="13"/>`，没有则插入
- 检查 ilvl=2 是否有 `<w:pStyle w:val="14"/>`，没有则插入

**Step B — 层面二：统一 styles.xml**
遍历所有标题样式(style12~16)：
- 检查每个样式的 numPr 是否指向 `numId=1`(→abstractNumId=0)
- 不正确或缺失的统一修正为正确的 numPr + ilvl
  - style12 → (numId=1, ilvl=0), style13 → (numId=1, ilvl=1)
  - style14 → (numId=1, ilvl=2), style15 → (numId=1, ilvl=3)
  - style16 → (numId=1, ilvl=4)

**Step C — 层面三：补全 document.xml**

遍历所有段落，对以下样式插入对应的 `<w:numPr>`：

| 样式 | styleId | 插入 ilvl | 效果 |
|------|---------|-----------|------|
| 章节标题 | 12 | ilvl=0 | 第N章 |
| 二级标题 (H2) | 13 | ilvl=1 | N.M. |
| 三级标题 (H3) | 14 | ilvl=2 | N.M.K. |
| 四级标题 (H4) | 15 | ilvl=3 | N.M.K.L. |
| 五级标题 (H5) | 16 | ilvl=4 | N.M.K.L.O. |

**Step D — 清除残留前缀**
删除文本中的旧编号前缀（如 `9.6.`），避免与自动编号重复显示。

#### Phase 1 分步详解

**Step D — 分析文档结构**
读取 document.xml，定位章节/H2/H3 标题，构建三级标题的完整前缀映射表。

**Step E — 查找并替换断裂编号**
四/五级标题区域内的编号补全（使用正则匹配）。

**Step F — 清除正文多余编号**
对非标题的长段落文本删除数字前缀。

#### Phase 2 分步详解

**Step G — 套用标题样式**
扫描识别 H4/H5 文本，优先使用当前文档真实的四/五级标题 styleId；若未找到再使用兼容默认值。

**Step G+ — 统一标题样式显示名**
将样式面板中的 1~5 级标题显示名统一为 `ryusuke标题1`、`ryusuke标题2`、`ryusuke标题3`、`ryusuke标题4`、`ryusuke标题5`。

**Step H — 清除重复编号前缀**
已标记 ryusuke标题4/5 的段落，去除文本中数字前缀。

**Step I — 正文首行缩进**
为无样式的中文正文插入缩进样式。

**Step J — 优化编号缩进（numbering.xml）**
修改 abstractNumId=1 中 ilvl=3(H4) 和 ilvl=4(H5)。

**Step K — 样式属性精修（styles.xml）**
字体/间距/行距/字号全面精修。

**Step L — 表格正文样式接管**
识别 Word 中已有的历史表格正文样式，并统一改为 `ryusuke-表格正文`；若文档缺失该样式，则自动补建一套表格正文样式，便于后续手动或批量套用。

## 关键经验教训

### ⚠️ 必须避开的坑

1. **`\w` 匹配中文问题**: Python 3 中 `\w` 包含 Unicode 中文。用 `[a-zA-Z0-9]` 判断纯字母数字字符串。
2. **禁止整体重构 XML**: 不要用 `re.sub(para_pattern, replacer, content)` 重构整个 document.xml！应使用**记录位置 + 切片替换**方式。
3. **numPr 自动编号冲突**: 如果样式定义了 `<w:numPr>`，文本不要再写数字前缀，否则显示两次。
4. **捕获组作用域问题**: 嵌套函数内修改外部变量时用字典 `counters={'key':0}`。
5. **XML 命名空间**: 正则匹配只需关心标签名如 `<w:p>`，无需完整命名空间 URI。
6. **段落匹配不可靠**: `<w:p>...</w:p>` 用正则 DOTALL 匹配时，内含子 `<w:p>` 会出错！改用**位置查找法** (`find_all_paras_positions`)。

### 📐 缩进速查表

| twips | cm | inch | 适用场景 |
|-------|-----|------|----------|
| 120 | ~0.21 | ~0.08 | 极小缩进 |
| 240 | ~0.42 | ~0.17 | 紧凑模式 |
| **360** | **~0.63** | **~0.25** | **H4 推荐** |
| **480** | **~0.85** | **~0.33** | **H5 推荐** |
| 720 | 1.27 | 0.50 | 标准 |
| 1440 | 2.54 | 1.00 | 宽松 |

> 1 cm ≈ 567 twips, 1 inch = 1440 twips

### 🎨 中文技术文档排版规范

- 标题优先**黑体**（醒目），正文**宋体**（易读）
- 多级标题段间距递减：一级→二级→三级 逐级缩小 1~2pt
- 行距建议 1.25~1.5 倍（lineRule=auto 时 line=280~360）
- 首行缩进 2 字符（约 480 twips 或 `w:firstLineChars="200"`）
- 表格正文建议使用居中对齐、小四字号（10.5pt）、宋体中文与 Times New Roman 西文组合，保证表内数据清晰统一

## 文件清单

```
ryusuke-bid-beautifier/
├── SKILL.md                          # 本文件（完整使用指南 V3）
├── scripts/
│   ├── bid_doc_optimizer.py          # ★ 主脚本（Phase 0+1+2 全流程）
│   └── word_bid_beautify.py          # Phase2 单独执行（仅样式美化）
└── references/
    └── ooxml-quickref.md             # Word OOXML 格式技术速查手册
```

## 输出报告模板

完成后向用户汇报：

```
✅ 优化完成!
📄 输出: <文件名>
📏 大小: X bytes (原 Y)
✅ 段落数: N → N 一致

═════════ 汇总报告 ══════════

📋 Phase 0 - 根本性编号修复（三层面统一）:
   • numbering.xml: 补全主列表pStyle链接
   • styles.xml: 统一 N 个样式的numPr
   • 章节标题(style12)插入numPr: N 处
   • 二级标题(style13)插入numPr: N 处
   • 三级标题(style14)插入numPr: N 处
   • 四级标题(style15)插入numPr: N 处
   • 五级标题(style16)插入numPr: N 处
   • 残留编号前缀清除:       N 处

📋 Phase 1 - 编号层级修复:
   • 层级编号修复:   N 处
   • 多余编号清除:   N 处
   
📋 Phase 2 - 样式美化:
   • 四级标题套用:   +N 处
   • 五级标题套用:   +N 处
   • 重复编号清除:   N 处
   • 正文首行缩进:   +N 处
   • 编号缩进优化:   N 项
   • 样式属性优化:   N 项
```
