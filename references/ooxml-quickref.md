# Word OOXML 格式技术速查

## 文档结构

```
document.docx (ZIP)
├── [Content_Types].xml     # 内容类型声明
├── _rels/.rels             # 根关系
├── word/
│   ├── document.xml        # ★ 主文档内容（段落/文本）
│   ├── styles.xml          # ★ 样式定义
│   ├── numbering.xml       # ★ 多级列表/编号定义
│   ├── fontTable.xml       # 字体表
│   └── ...
```

## 段落 XML 结构

```xml
<w:p>                         <!-- 段落 -->
  <w:pPr>                     <!-- 段落属性 -->
    <w:pStyle w:val="15"/>    <!-- 引用样式 ID=15 (hik标题4) -->
    <w:numPr>
      <w:ilvl w:val="3"/>     <!-- 编号级别 3 = 四级标题 -->
      <w:numId w:val="1"/>    <!-- 编号列表 ID=1 -->
    </w:numPr>
    <w:spacing w:before="60" w:after="40" w:line="280" w:lineRule="auto"/>
    <w:ind w:left="360" w:hanging="360"/>
  </w:pPr>
  <w:r>                       <!-- run (文本块) -->
    <w:rPr>                   <!-- run 属性 (字体等) -->
      <w:rFonts w:ascii="黑体" w:eastAsia="黑体"/>
      <w:b/>                   <!-- 加粗 -->
      <w:sz w:val="24"/>       <!-- 字号 12pt (half-points) -->
    </w:rPr>
    <w:t>文本内容</w:t>        <!-- 实际文本 -->
  </w:r>
</w:p>
```

## 常用单位换算

| 单位 | 说明 | 与其他单位关系 |
|------|------|---------------|
| twips | 1/20 pt | 1pt = 20 twips, 1inch = 1440 twips |
| half-points | 半磅 (字体大小) | sz="24" = 12pt |
| pt | 磅 | 1pt ≈ 0.353mm, 1cm ≈ 28.35pt |

### 缩进速查

| twips | cm | inch | EMU |
|-------|-----|------|-----|
| 120 | 0.21 | 0.08 | 152,400 |
| 240 | 0.42 | 0.17 | 304,800 |
| **360** | **0.63** | **0.25** | **457,200** |
| **480** | **0.85** | **0.33** | **609,600** |
| 720 | 1.27 | 0.50 | 914,400 |
| 864 | 1.52 | 0.34 | 1,097,280 |
| 1008 | 1.78 | 0.70 | 1,278,720 |
| 1440 | 2.54 | 1.00 | 1,828,800 |

## HIK 样式体系参考

本 Skill 针对华数/海康威视(HIK)投标文档的典型样式体系：

| styleId | 名称 | 类型 | 典型属性 |
|---------|------|------|----------|
| 12 | hik标题1 | 章标题 | 居中, 16pt加粗, chineseCounting(第N章) |
| 13 | hik标题2 | 二级标题 | 18pt, %1.%2. |
| 14 | hik标题3 | 三级标题 | 15pt加粗, %1.%2.%3. |
| **15** | **hik标题4** | **四级标题** | **12pt黑体加粗, %1.%2.%3.%4.** |
| **16** | **hik标题5** | **五级标题** | **加粗, %1.%2.%3.%4.%5.** |
| 17 | HIK-标题6 | 六级标题 | %1...%6. |
| 11 | hik正文 | 正文 | 宋体, 12pt |
| 25 | hik首行缩进两字符 | 缩进正文 | 首行缩进2字符 |

## 编号系统 (numbering.xml)

编号由 `abstractNum` 定义模板，`num` 引用实例：

```xml
<!-- 抽象编号模板 (abstractNumId=1 是 HIK 体系) -->
<w:abstractNum w:abstractNumId="1">
  <w:lvl w:ilvl="3">           <!-- 四级标题对应 ilvl=3 -->
    <w:start w:val="1"/>
    <w:numFmt w:val="decimal"/>
    <w:pStyle w:val="15"/>     <!-- 关联样式 hik标题4 -->
    <w:lvlText w:val="%1.%2.%3.%4."/>  <!-- 编号格式 -->
    <w:ind w:left="360" w:hanging="360"/> <!-- ★ 缩进控制 -->
  </w:lvl>
</w:abstractNum>

<!-- 编号实例引用 -->
<w:num w:numId="1">
  <w:abstractNumId w:val="1"/>
</w:num>
```

### 关键：left vs hanging

- **left**: 整个段落的左缩进位置
- **hanging**: "悬挂缩进"，即编号文字本身的宽度（从 left 向左延伸）

效果：文本实际起始位置 = **left - hanging + 编号宽度**

当 left == hanging 时，文本紧贴在编号后，视觉上最紧凑。

## 正则匹配模式

### 提取段落纯文本

```python
texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para_body)
full_text = ''.join(texts).strip()
```

### 匹配多级编号标题

```python
# 四级: X.X.X.X 标题文字
m4 = re.match(r'^(\d+\.\d+\.\d+\.\d+)[ \t]*(.+)$', text)

# 五级: X.X.X.X.X 标题文字  
m5 = re.match(r'^(\d+\.\d+\.\d+\.\d+\.\d+)[ \t]*(.+)$', text)
```

### 判断是否为标题（非型号/非句子）

```python
def is_likely_heading(text):
    if not text or len(text) > 60:
        return False
    if re.search(r'[。；！？]', text) and len(text) > 15:
        return False  # 长句含标点 → 正文
    if re.match(r'^[a-zA-Z0-9\-\+,./\\()[\]]{2,}$', text) and len(text) < 20:
        return False  # 纯字母数字 → 型号
    return True
```

### 精确定位替换（安全方式）

```python
# 1. 用带捕获组的正则定位所有段落
para_re = r'(<w:p[ >])(.*?)(</w:p>)'
paras = list(re.finditer(para_re, content, re.DOTALL))

# 2. 收集修改（不直接替换）
modifications = []
for m in paras:
    new_body = process(m.group(2))
    if new_body != m.group(2):
        modifications.append((m.start(2), m.end(2), new_body))

# 3. 从后往前替换（保持偏移正确）
content_list = list(content)
for start, end, repl in reversed(modifications):
    content_list[start:end] = list(repl)
result = ''.join(content_list)
```
