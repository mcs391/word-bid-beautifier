# 📝 Word 投标文档一键美化工具

> **word-bid-beautifier** — 专为投标技术文件设计的自动化排版优化工具

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## ✨ 功能特性

本工具通过**三阶段流水线**，一条命令完成投标文档的全流程排版优化：

| 阶段 | 名称 | 功能说明 |
|:----:|------|---------|
| **Phase 0** | 根本性编号修复 ★V2★ | **三层面统一修复**：补全多级列表 pStyle 链接、统一样式 numPr、补全段落级编号引用。彻底解决「手动套用H2/H3后显示9.x而非14.x」的问题 |
| **Phase 1** | 编号层级修复 | 将断裂的四级/五级标题补全为完整多级编号（如 `1.` → `9.1.1.1`），并清除正文误加的数字前缀 |
| **Phase 2** | 样式美化 | 智能识别并套用 hik 标题样式、清除重复编号、添加首行缩进、优化缩进值和字体间距 |

### 详细功能清单

#### 📐 Phase 0 — 根本性编号修复（三层面统一）★V2★
> **核心价值**：修复后，**手动套用 hik标题2/3 样式时 Word 会自动继承正确的章节号计数**

- **层面一 — numbering.xml**：补全主多级列表 `abstractNumId=0` 的 pStyle 链接
  - 确保层级1绑定 style13(H2)、层级2绑定 style14(H3)
- **层面二 — styles.xml**：统一所有标题样式(12~16)的 numPr 指向同一个多级列表实例
  - 解决不同样式指向不同 abstractNum 导致的"编号分裂"
- **层面三 — document.xml**：为缺失 numPr 的标题段落(12~16)插入正确引用
- **清除残留前缀**：自动删除文本中旧的硬编码编号前缀
- 支持任意数量的章节，无需手动配置

#### 🔢 Phase 1 — 编号层级修复
- **智能补全断裂编号**：将 `1.文化传承...` → `9.1.1.1 文化传承...`
- **五级标题同理**：`1.1传统意象...` → `9.1.1.1.1 传统意象...`
- **清除多余前缀**：删除正文中被误加的数字前缀（如 `1.`、`2.`）
- **自动检测章节号范围**：默认从文档中自动推断起始/结束章节号

#### 🎨 Phase 2 — 样式美化
- **样式套用**：智能识别四级/五级标题，自动套用 `hik标题4` / `hik标题5` 样式
- **重复编号清理**：清除文本硬编码的编号前缀（避免与 Word 自动编号冲突显示两遍）
- **首行缩进**：为中文正文添加标准首行缩进（2字符 = 480 twips）
- **缩进优化**：H4 缩进降至 360 twips（~0.63cm），H5 降至 480 twips（~0.85cm）
- **字体精修**：统一中文字体为宋体，西文字体为 Times New Roman
- **间距优化**：调整标题段前段后间距和行距

---

## 🚀 快速开始

### 环境要求

- **Python 3.8+** （无需安装任何第三方依赖，仅使用标准库）

### 基本用法

```bash
# 一键全流程优化（推荐）
python bid_doc_optimizer.py 输入文件.docx 输出文件.docx

# 示例
python bid_doc_optimizer.py "投标方案.docx" "投标方案 - 已优化.docx"
```

### 高级选项

```bash
# 仅执行 Phase 0（H2/H3 编号跟随修复）
python bid_doc_optimizer.py input.docx output.docx --phase 0

# 仅执行 Phase 1（编号层级修复）
python bid_doc_optimizer.py input.docx output.docx --phase 1

# 仅执行 Phase 2（样式美化）
python bid_doc_optimizer.py input.docx output.docx --phase 2

# 自定义参数
python bid_doc_optimizer.py input.docx output.docx \
    --ch-start 5 \          # 从第5章开始编号修复
    --ch-end 12 \           # 到第12章结束
    --indent-h4 360 \       # H4 缩进 twips 值
    --indent-h5 480 \       # H5 缩进 twips 值
    --no-strip              # 跳过清除多余编号步骤

# 仅分析不修改（预览模式）
python bid_doc_optimizer.py input.docx output.docx --dry-run
```

### 参数一览

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--phase {0\|1\|2\|all}` | 执行阶段 | `all`（全部执行） |
| `--ch-start N` | 编号修复起始章节号 | 自动检测 |
| `--ch-end N` | 编号修复结束章节号 | 自动检测 |
| `--indent-h4 N` | H4 缩进值 (twips) | 360 (~0.63cm) |
| `--indent-h5 N` | H5 缩进值 (twips) | 480 (~0.85cm) |
| `--dry-run` | 仅分析不修改 | 关闭 |
| `--no-strip` | 跳过清除多余编号 | 关闭 |

---

## 📂 文件结构

```
word-bid-beautifier/
├── README.md                    # 本使用说明
├── SKILL.md                     # WorkBuddy Skill 定义文件
├── scripts/
│   ├── bid_doc_optimizer.py     # ★ 主脚本（三阶段一体化工具）V3
│   └── word_bid_beautify.py      # 辅助脚本（样式美化专用）
├── assets/                      # 资源目录
└── references/
    └── ooxml-quickref.md        # OOXML 参考手册
```

---

## 🔧 工作原理

### 技术架构

本工具直接操作 **OOXML**（Office Open XML）格式：

1. **解压** `.docx` → 提取 `word/document.xml` 和 `word/numbering.xml`
2. **正则解析**段落 XML → 识别样式 ID 和文本内容
3. **按规则修改**XML → 插入编号引用、修正文本、调整样式属性
4. **重新打包** → 生成新的 `.docx` 文件

### 样式 ID 对照表（海康/HIK 格式）

| 层级 | style ID | 样式名称 | 编号格式 |
|:----:|:--------:|----------|----------|
| 章标题 | 12 | hik标题1 | 第%1章 |
| 二级标题 (H2) | 13 | hik标题2 | %1.%2. |
| 三级标题 (H3) | 14 | hik标题3 | %1.%2.%3. |
| 四级标题 (H4) | 15 | hik标题4 | %1.%2.%3.%4. |
| 五级标题 (H5) | 16 | hik标题5 | %1.%2.%3.%4.%5. |
| 正文 | 17 | hik正文 | 无编号 |

### 编号系统

Word 的多级列表定义在 `numbering.xml` 中，核心结构：

```
abstractNumId=1 (主多级列表)
  ├── ilvl=0 → 第%1章        (style 12, 章标题)
  ├── ilvl=1 → %1.%2.        (style 13, 二级标题)
  ├── ilvl=2 → %1.%2.%3.     (style 14, 三级标题)
  ├── ilvl=3 → %1.%2.%3.%4.  (style 15, 四级标题)
  └── ilvl=4 → %1.%2.%3.%4.%5.(style 16, 五级标题)
```

**Phase 0 V2 的核心逻辑**：很多投标文档存在**编号分裂问题**——不同标题样式指向不同的 abstractNum（多级列表定义），导致计数器不同步。例如手动在第14章套用 hik标题2 样式时，编号显示为 9.x 而非 14.x。V2 通过三层面修复彻底解决：

1. **numbering.xml** — 在主列表的 ilvl=1/2 中补全 pStyle→style13/style14 链接
2. **styles.xml** — 将所有标题样式的 numPr 统一指向同一个 numId（→abstractNumId=0）
3. **document.xml** — 为缺失 numPr 的标题段落插入正确的编号引用

修复后，Word 能正确理解每个标题在多级列表中的位置，**手动添加的标题也能自动获得正确的章节号编号**。

---

## ⚠️ 使用注意事项

1. **备份原始文件**：运行前请务必备份原文件，虽然工具不会主动删除内容
2. **仅适用于 HIK/华数格式**：样式 ID (12–16) 是针对海康威视投标模板的，其他格式需先确认样式映射
3. **编号格式要求**：原文档需已包含 `abstractNumId=1` 的多级列表定义（大多数投标模板自带）
4. **编码设置**：Windows 环境下运行时建议设置 `PYTHONUTF8=1` 避免编码问题
5. **段落数验证**：每次运行结束后会验证段落数是否一致，确保无段落丢失

---

## 📋 典型应用场景

| 场景 | 推荐命令 |
|------|---------|
| 首次处理全新投标文件 | `python bid_doc_optimizer.py in.docx out.docx` |
| 只修复编号问题 | `python bid_doc_optimizer.py in.docx out.docx --phase 0` |
| 只做样式美化 | `python bid_doc_optimizer.py in.docx out.docx --phase 2` |
| 先预览再操作 | `python bid_doc_optimizer.py in.docx out.docx --dry-run` |
| 处理部分章节 | `python bid_doc_optimizer.py in.docx out.docx --ch-start 3 --ch-end 8` |

---

## 🔄 更新日志

### V3.1 — 2026-04-10
- 🔧 **Phase 0 升级为 V2（三层面根本性修复）**
  - 层面一：补全 `numbering.xml` 中主多级列表的 pStyle 链接（ilvl=1→style13, ilvl=2→style14）
  - 层面二：统一 `styles.xml` 中所有标题样式(12~16)的 numPr 指向同一个多级列表实例
  - 彻底解决「手动套用 H2/H3 样式后编号显示 9.x 而非当前章节号」的问题
- 扩展段落级 numPr 补全范围至所有标题级别(12~16)，不再仅限于 style12/13/14

### V3.0 — 2026-04-09
- ✨ 新增 **Phase 0: H2/H3 编号跟随修复**（V1）
- 解决第N章后的二级/三级标题仍显示旧章节号的经典问题
- 为章节(style12)/二级标题(style13)/三级标题(style14)自动插入 numPr
- 自动清除文本中的残留编号前缀

### V2.0
- 整合 Phase 1 + Phase 2 为一体化脚本 `bid_doc_optimizer.py`
- 新增四级/五级标题样式套用
- 新增缩进优化和字体精修

### V1.0
- 初始版本，基础编号修复和样式美化

---

## 📄 许可证

[MIT License](LICENSE)

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！如果你在使用中遇到以下情况，欢迎反馈：
- 其他格式的投标文档需要适配
- 发现编号或样式处理异常
- 有新的功能需求建议

---

<p align="center">
  Made with ❤️ for 投标人 | Powered by Python + OOXML
</p>
