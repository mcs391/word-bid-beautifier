# -*- coding: utf-8 -*-
"""
Word 投标文档美化工具 — 一键优化脚本
========================================
功能：
  1. 识别并套用 ryusuke 四级/五级标题样式
  2. 清除重复编号前缀（样式自带自动编号时）
  3. 优化缩进值、间距、字体等排版参数
  4. 零内容删除，仅调整格式
  
用法:
  python word_bid_beautify.py <input.docx> <output.docx> [--indent-h4 N] [--indent-h5 N]
  
参数:
  input.docx   - 输入文件路径 (必须)
  output.docx - 输出文件路径 (必须)
  --indent-h4  - H4标题缩进 twips, 默认360
  --indent-h5  - H5标题缩进 twips, 默认480
"""
import os, re, sys, argparse, zipfile, shutil, tempfile, xml.etree.ElementTree as ET

# ============================================================
# 常量定义 — ryusuke 样式 ID
# ============================================================
STYLE_RYUSUKE_H4 = "15"      # ryusuke标题4
STYLE_RYUSUKE_H5 = "16"      # ryusuke标题5
STYLE_RYUSUKE_BODY = "11"    # ryusuke正文
STYLE_RYUSUKE_INDENT = "25"  # ryusuke首行缩进两字符

# 默认缩进值 (twips, 1cm ≈ 567 twips)
DEFAULT_INDENT_H4 = 360  # ~0.63cm
DEFAULT_INDENT_H5 = 480  # ~0.85cm


def unpack_doc(src_path, dest_dir):
    """解包 docx 到目录"""
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)
    with zipfile.ZipFile(src_path, 'r') as zf:
        zf.extractall(dest_dir)


def pack_doc(src_dir, out_path):
    """从目录打包为 docx"""
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(src_dir):
            for fname in files:
                fpath = os.path.join(root, fname)
                arcname = os.path.relpath(fpath, src_dir)
                zf.write(fpath, arcname)


def count_paragraphs(doc_xml_path):
    """统计段落数"""
    with open(doc_xml_path, 'r', encoding='utf-8') as f:
        content = f.read()
    return len(re.findall(r'<w:p[ >]', content))


def get_text_from_para(para_body):
    """提取段落纯文本"""
    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para_body)
    return ''.join(texts).strip()


def is_likely_heading(body_text):
    """判断是否像标题文本（非型号/非长句）"""
    t = body_text.strip()
    if not t or len(t) > 60:
        return False
    # 包含句号且较长 → 可能是正文句子
    if re.search(r'[。；！？]', t) and len(t) > 15:
        return False
    # 纯字母数字符号 → 型号列表，不是标题
    if re.match(r'^[a-zA-Z0-9\-\+,./\\()[\]]{2,}$', t) and len(t) < 20:
        return False
    return True


def insert_pStyle(p_body, style_id):
    """在 pPr 中插入 pStyle 标签"""
    if re.search(r'<w:pStyle\s', p_body):
        return p_body
    
    if '<w:pPr>' in p_body:
        return p_body.replace('<w:pPr>', f'<w:pPr><w:pStyle w:val="{style_id}"/>')
    
    return f'<w:pPr><w:pStyle w:val="{style_id}"/></w:pPr>{p_body}'


def step1_apply_heading_styles(content):
    """
    步骤1：识别并套用 H4/H5 标题样式
    匹配规则：
      - H4: X.X.X.X 标题文字 (4段数字 + 中文标题)
      - H5: X.X.X.X.X 标题文字 (5段数字 + 中文标题)
    返回: (修改后content, h4_count, h5_count)
    """
    para_pattern = r'(<w:p[ >])(.*?)(</w:p>)'
    paras = list(re.finditer(para_pattern, content, re.DOTALL))
    
    counters = {'h4': 0, 'h5': 0}
    modifications = []  # (start, end, replacement)
    
    for m in paras:
        p_open = m.group(1)
        p_body = m.group(2)
        p_close = m.group(3)
        
        # 已有样式的段落跳过
        if re.search(r'<w:pStyle', p_body):
            continue
        
        full_text = get_text_from_para(p_body)
        if not full_text:
            continue
        
        # === 检测五级编号: X.X.X.X.X 文字 (如 9.1.1.1.1 传统意象转译) ===
        m5 = re.match(
            r'^(\d+)\.(\d+)\.(\d+)\.(\d+)\.(\d+)([ \t]*)(.*)$',
            full_text
        )
        if m5 and is_likely_heading(m5.group(7)):
            new_pPr = insert_pStyle(p_body, STYLE_RYUSUKE_H5)
            if new_pPr != p_body:
                counters['h5'] += 1
                modifications.append((m.start(2), m.end(2), new_pPr))
            continue
        
        # === 检测四级编号: X.X.X.X 文字 (如 9.1.1.1 文化传承...) ===
        m4 = re.match(
            r'^(\d+)\.(\d+)\.(\d+)\.(\d+)([ \t]*)(.*)$',
            full_text
        )
        if m4 and is_likely_heading(m4.group(6)):
            new_pPr = insert_pStyle(p_body, STYLE_RYUSUKE_H4)
            if new_pPr != p_body:
                counters['h4'] += 1
                modifications.append((m.start(2), m.end(2), new_pPr))
    
    # 从后往前替换，保持偏移正确
    content_list = list(content)
    for start, end, repl in reversed(modifications):
        content_list[start:end] = list(repl)
    
    return ''.join(content_list), counters['h4'], counters['h5']


def step2_strip_number_prefixes(content):
    """
    步骤2：清除H4/H5标题中重复的数字前缀
    当样式带有 numPr 自动编号时，文本中的数字前缀会导致重复显示
    只处理第一个 <w:t> 节点
    返回: (修改后content, stripped_count)
    """
    para_pattern = r'(<w:p[ >])(.*?)(</w:p>)'
    paras = list(re.finditer(para_pattern, content, re.DOTALL))
    
    stripped = 0
    modifications = []
    
    for m in paras:
        p_body = m.group(2)
        
        # 只处理带 ryusuke标题4 或 ryusuke标题5 样式的段落
        if not re.search(rf'w:val="({STYLE_RYUSUKE_H4}|{STYLE_RYUSUKE_H5})"', p_body):
            continue
        
        # 找到第一个 <w:t> 节点
        t_match = re.search(r'(<w:t[^>]*>)([^<]*)(</w:t>)', p_body)
        if not t_match:
            continue
        
        text_val = t_match.group(2).strip()
        if not text_val:
            continue
        
        # 移除开头的数字编号前缀
        new_text = re.sub(
            r'^\d+(\.\d+){3,}[ \t]*(?=.)',
            '',
            text_val,
            count=1
        )
        
        if new_text != text_val and new_text.strip():
            # 保持原始空格/标签属性不变
            old_full = t_match.group(0)
            new_full = t_match.group(1) + new_text + t_match.group(3)
            
            abs_start = m.start(2) + t_match.start()
            abs_end = m.start(2) + t_match.end()
            modifications.append((abs_start, abs_end, new_full))
            stripped += 1
    
    # 从后往前替换
    content_list = list(content)
    for start, end, repl in reversed(modifications):
        content_list[start:end] = list(repl)
    
    return ''.join(content_list), stripped


def step3_apply_body_indent(content):
    """
    步骤3：为无样式的中文正文段落添加首行缩进样式
    跳过条件：已有样式 / 空段 / 纯数字符号 / 编号开头
    返回: (修改后content, indent_count)
    """
    para_pattern = r'(<w:p[ >])(.*?)(</w:p>)'
    paras = list(re.finditer(para_pattern, content, re.DOTALL))
    
    body_count = 0
    modifications = []
    
    for m in paras:
        p_body = m.group(2)
        
        # 已有样式则跳过
        if re.search(r'<w:pStyle', p_body):
            continue
        
        full_text = get_text_from_para(p_body)
        if not full_text:
            continue
        
        # 纯数字/符号跳过
        if re.match(r'^[\d\s.,;:\-+=\\/\\(){}\[\]\"\'<>]+$', full_text):
            continue
        
        # 编号开头的已处理
        if re.match(r'^\d+\.', full_text):
            continue
        
        # 为中文正文添加首行缩进样式
        new_pPr = insert_pStyle(p_body, STYLE_RYUSUKE_INDENT)
        if new_pPr != p_body:
            body_count += 1
            modifications.append((m.start(2), m.end(2), new_pPr))
    
    content_list = list(content)
    for start, end, repl in reversed(modifications):
        content_list[start:end] = list(repl)
    
    return ''.join(content_list), body_count


def step4_optimize_numbering(num_xml_content, indent_h4=DEFAULT_INDENT_H4, indent_h5=DEFAULT_INDENT_H5):
    """
    步骤4：优化 numbering.xml 中多级列表缩进值
    针对抽象编号定义 abstractNumId=1 (ryusuke标题体系, numId=1引用)
    调整 ilvl=3(H4) 和 ilvl=4(H5) 的 left/hanging 值
    """
    changes = []
    
    # 尝试精确匹配 abstractNumId=1 下的 ilvl=3 和 ilvl=4
    def replace_in_abstract_num(xml, abs_num_id, ilvl, new_left, label):
        pattern = rf'(<w:abstractNum w:abstractNumId="{abs_num_id}"[^>]*>.*?<w:lvl w:ilvl="{ilvl}"[^>]*>.*?<w:ind w:left=")(\d+)(" w:hanging=")(\d+)("/>)'
        match = re.search(pattern, xml, re.DOTALL)
        if match:
            old_left = int(match.group(2))
            old_hang = int(match.group(4))
            new_xml = re.sub(
                pattern,
                rf'\g<1>{new_left}\g<3>{new_left}\g<5>',
                xml, count=1, flags=re.DOTALL
            )
            changes.append(f'{label}: left={old_left}→{new_left}, hanging={old_hang}→{new_left}')
            return new_xml
        return xml
    
    result = num_xml_content
    result = replace_in_abstract_num(result, '1', '3', indent_h4, 'H4(ilvl=3)')
    result = replace_in_abstract_num(result, '1', '4', indent_h5, 'H5(ilvl=4)')
    
    return result, changes


def find_heading_style_id(style_xml, level):
    """按样式名称识别当前文档里的标题样式 ID，兼容 Word/WPS 内置标题。"""
    candidates = {
        f'heading {level}',
        f'title {level}',
        f'标题 {level}',
        f'标题{level}',
        f'ryusuke标题{level}',
    }
    style_pattern = r'<w:style\b(?=[^>]*w:type="paragraph")(?=[^>]*w:styleId="([^"]+)")[^>]*>.*?</w:style>'
    for m in re.finditer(style_pattern, style_xml, re.DOTALL):
        block = m.group(0)
        name = re.search(r'<w:name w:val="([^"]+)"', block)
        if name and name.group(1).strip().lower() in {c.lower() for c in candidates}:
            return m.group(1)

    builtin_style_ids = {1: "2", 2: "3", 3: "4", 4: "5", 5: "6"}
    fallback = builtin_style_ids.get(level)
    if fallback and re.search(
        rf'<w:style\b(?=[^>]*w:type="paragraph")(?=[^>]*w:styleId="{re.escape(fallback)}")',
        style_xml,
    ):
        return fallback
    return None


def resolve_ryusuke_heading_style_ids(style_xml):
    """把脚本内部 H4/H5 指向当前文档真实标题样式，避免固定 styleId 误伤。"""
    global STYLE_RYUSUKE_H4, STYLE_RYUSUKE_H5
    changes = []

    h4_id = find_heading_style_id(style_xml, 4)
    if h4_id and h4_id != STYLE_RYUSUKE_H4:
        STYLE_RYUSUKE_H4 = h4_id
        changes.append(f'四级标题样式绑定为 styleId={h4_id}')

    h5_id = find_heading_style_id(style_xml, 5)
    if h5_id and h5_id != STYLE_RYUSUKE_H5:
        STYLE_RYUSUKE_H5 = h5_id
        changes.append(f'五级标题样式绑定为 styleId={h5_id}')
    elif not h5_id and STYLE_RYUSUKE_H5 == "16" and 'w:styleId="16"' in style_xml:
        STYLE_RYUSUKE_H5 = "ryusukeHeading5"
        changes.append('styleId=16 已被非段落样式占用，改用 ryusukeHeading5 创建五级标题样式')

    return changes


def rename_ryusuke_heading_style_names(style_xml):
    """把 1-5 级标题样式在样式面板中的显示名统一为 ryusuke标题N。"""
    result = style_xml
    changes = []

    for level in range(1, 6):
        style_id = find_heading_style_id(result, level)
        if not style_id:
            continue

        style_re = rf'(<w:style\b(?=[^>]*w:type="paragraph")(?=[^>]*w:styleId="{re.escape(style_id)}")[^>]*>)(.*?)(</w:style>)'
        m = re.search(style_re, result, re.DOTALL)
        if not m:
            continue

        target_name = f'ryusuke标题{level}'
        block = m.group(2)
        name_re = r'<w:name w:val="([^"]+)"/>'
        name_match = re.search(name_re, block)
        if name_match and name_match.group(1) == target_name:
            continue

        if name_match:
            new_block = re.sub(name_re, f'<w:name w:val="{target_name}"/>', block, count=1)
        else:
            new_block = f'<w:name w:val="{target_name}"/>' + block

        result = result[:m.start(2)] + new_block + result[m.end(2):]
        changes.append(f'标题{level}: 显示名改为{target_name}(styleId={style_id})')

    return result, changes


def step5_optimize_styles(style_xml_content):
    """
    步骤5：优化 styles.xml 中的 ryusuke标题样式属性
    基于中文技术文档排版最佳实践
    """
    changes = []
    result = style_xml_content

    h5_style_re = rf'<w:style\b(?=[^>]*w:type="paragraph")(?=[^>]*w:styleId="{re.escape(STYLE_RYUSUKE_H5)}")'
    if not re.search(h5_style_re, result):
        based_on = '6' if 'w:styleId="6"' in result else '1'
        h5_style = (
            f'<w:style w:type="paragraph" w:customStyle="1" w:styleId="{STYLE_RYUSUKE_H5}">'
            '<w:name w:val="ryusuke标题5"/>'
            f'<w:basedOn w:val="{based_on}"/>'
            '<w:next w:val="1"/>'
            '<w:uiPriority w:val="9"/>'
            '<w:qFormat/>'
            '<w:pPr>'
            '<w:keepNext/>'
            '<w:keepLines/>'
            '<w:numPr><w:ilvl w:val="4"/><w:numId w:val="1"/></w:numPr>'
            '<w:spacing w:before="40" w:after="20" w:line="276" w:lineRule="auto"/>'
            '<w:ind w:left="480" w:hanging="480"/>'
            '<w:outlineLvl w:val="4"/>'
            '</w:pPr>'
            '<w:rPr>'
            '<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:cs="黑体"/>'
            '<w:b/>'
            '<w:sz w:val="21"/><w:szCs w:val="21"/>'
            '</w:rPr>'
            '</w:style>'
        )
        result = result.replace('</w:styles>', h5_style + '</w:styles>')
        changes.append(f'ryusuke标题5: 新增五级标题样式(styleId={STYLE_RYUSUKE_H5})')

    result, rename_changes = rename_ryusuke_heading_style_names(result)
    changes.extend(rename_changes)
    
    # --- ryusuke标题4 (styleId=15) 优化 ---
    # 段落间距：更紧凑
    old_h4_spacing = r'<w:spacing w:before="120" w:after="120" w:line="360" w:lineRule="auto"/>'
    new_h4_spacing = '<w:spacing w:before="60" w:after="40" w:line="280" w:lineRule="auto"/>'
    if re.search(old_h4_spacing, result):
        result = re.sub(old_h4_spacing, new_h4_spacing, result)
        changes.append('ryusuke标题4: 段前6pt→3pt, 段后6pt→2pt, 行距280')
    
    # 字体统一为黑体
    old_h4_font = r'<w:rFonts w:ascii="黑体" w:hAnsi="宋体" w:eastAsia="黑体" w:cs="Times New Roman"/>'
    new_h4_font = '<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:cs="黑体"/>'
    if re.search(old_h4_font, result):
        result = re.sub(old_h4_font, new_h4_font, result)
        changes.append('ryusuke标题4: 字体统一黑体')
    
    # --- ryusuke标题5 (styleId=16) 优化 ---
    # 添加/更新 spacing
    h5_ppr_match = re.search(
        rf'(<w:style\b(?=[^>]*w:type="paragraph")(?=[^>]*w:styleId="{re.escape(STYLE_RYUSUKE_H5)}")[^>]*>.*?<w:pPr>)(.*?)(</w:pPr>)',
        result, re.DOTALL
    )
    if h5_ppr_match:
        ppr_content = h5_ppr_match.group(2)
        if not re.search(r'<w:spacing', ppr_content):
            new_ppr = ppr_content + '<w:spacing w:before="40" w:after="20" w:line="276" w:lineRule="auto"/>'
            result = result[:h5_ppr_match.start(2)] + new_ppr + result[h5_ppr_match.end(2):]
            changes.append('ryusuke标题5: 新增段前2pt/段后1pt/行距')
    
    # 字体改为黑体
    old_h5_font = r'<w:rFonts w:ascii="宋体" w:hAnsi="黑体" w:eastAsia="宋体" w:cs="宋体"/>'
    new_h5_font = '<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:cs="黑体"/>'
    if re.search(old_h5_font, result):
        result = re.sub(old_h5_font, new_h5_font, result)
        changes.append('ryusuke标题5: 字体改为黑体')
    
    # 添加明确字号 10.5pt
    h5_rpr_match = re.search(
        rf'(<w:style\b(?=[^>]*w:type="paragraph")(?=[^>]*w:styleId="{re.escape(STYLE_RYUSUKE_H5)}")[^>]*>.*?<w:rPr>)(.*?)(</w:rPr>)',
        result, re.DOTALL
    )
    if h5_rpr_match:
        rpr_content = h5_rpr_match.group(2)
        if not re.search(r'<w:sz ', rpr_content):
            new_rpr = rpr_content + '<w:sz w:val="21"/><w:szCs w:val="21"/>'
            result = result[:h5_rpr_match.start(2)] + new_rpr + result[h5_rpr_match.end(2):]
            changes.append('ryusuke标题5: 新增字号10.5pt')
    
    return result, changes


def main():
    parser = argparse.ArgumentParser(description='Word 投标文档一键美化工具')
    parser.add_argument('input', help='输入 .docx 文件路径')
    parser.add_argument('output', help='输出 .docx 文件路径')
    parser.add_argument('--indent-h4', type=int, default=DEFAULT_INDENT_H4,
                        help=f'H4标题缩进 twips (默认 {DEFAULT_INDENT_H4}, ~0.63cm)')
    parser.add_argument('--indent-h5', type=int, default=DEFAULT_INDENT_H5,
                        help=f'H5标题缩进 twips (默认 {DEFAULT_INDENT_H5}, ~0.85cm)')
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f'❌ 输入文件不存在: {args.input}')
        sys.exit(1)
    
    tmp = tempfile.mkdtemp(prefix='wbb_')
    try:
        # 解包
        print(f'📦 解包: {os.path.basename(args.input)}')
        unpack_doc(args.input, tmp)
        
        orig_para_count = count_paragraphs(os.path.join(tmp, 'word', 'document.xml'))
        print(f'   原始段落数: {orig_para_count}')

        sty_path = os.path.join(tmp, 'word', 'styles.xml')
        if os.path.exists(sty_path):
            with open(sty_path, 'r', encoding='utf-8') as f:
                initial_sty_xml = f.read()
            for c in resolve_ryusuke_heading_style_ids(initial_sty_xml):
                print(f'   ⚠ {c}')
        
        # ---- Step 1: 套用标题样式 ----
        print('\n📌 Step 1: 识别并套用 H4/H5 标题样式...')
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            doc_xml = f.read()
        
        doc_xml, h4_cnt, h5_cnt = step1_apply_heading_styles(doc_xml)
        print(f'   ✓ 四级标题(ryusuke标题4): +{h4_cnt}')
        print(f'   ✓ 五级标题(ryusuke标题5): +{h5_cnt}')
        
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(doc_xml)
        
        # ---- Step 2: 清除重复编号 ----
        print('\n📌 Step 2: 清除标题中重复编号前缀...')
        with open(doc_path, 'r', encoding='utf-8') as f:
            doc_xml = f.read()
        
        doc_xml, stripped = step2_strip_number_prefixes(doc_xml)
        print(f'   ✓ 清除重复编号: {stripped} 处')
        
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(doc_xml)
        
        # ---- Step 3: 正文首行缩进 ----
        print('\n📌 Step 3: 应用正文首行缩进样式...')
        with open(doc_path, 'r', encoding='utf-8') as f:
            doc_xml = f.read()
        
        doc_xml, body_cnt = step3_apply_body_indent(doc_xml)
        print(f'   ✓ 正文首行缩进: +{body_cnt}')
        
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(doc_xml)
        
        # ---- Step 4: 优化编号缩进 ----
        print(f'\n📌 Step 4: 优化编号缩进 (H4={args.indent_h4}, H5={args.indent_h5} twips)...')
        num_path = os.path.join(tmp, 'word', 'numbering.xml')
        with open(num_path, 'r', encoding='utf-8') as f:
            num_xml = f.read()
        
        num_xml, num_changes = step4_optimize_numbering(num_xml, args.indent_h4, args.indent_h5)
        for c in num_changes:
            print(f'   ✓ {c}')
        if not num_changes:
            print('   ⚠ 未找到可优化的编号定义(可能文档不包含ryusuke编号体系)')
        
        with open(num_path, 'w', encoding='utf-8') as f:
            f.write(num_xml)
        
        # ---- Step 5: 优化样式 ----
        print('\n📌 Step 5: 优化标题样式属性...')
        with open(sty_path, 'r', encoding='utf-8') as f:
            sty_xml = f.read()
        
        sty_xml, sty_changes = step5_optimize_styles(sty_xml)
        for c in sty_changes:
            print(f'   ✓ {c}')
        
        with open(sty_path, 'w', encoding='utf-8') as f:
            f.write(sty_xml)
        
        # ---- 验证 & 打包 ----
        final_para_count = count_paragraphs(doc_path)
        print(f'\n📌 验证:')
        print(f'   段落数: {orig_para_count} → {final_para_count}', end='')
        if final_para_count == orig_para_count:
            print(' ✅ 一致')
        else:
            print(f' ⚠️ 差异={final_para_count - orig_para_count}')
        
        pack_doc(tmp, args.output)
        out_size = os.path.getsize(args.output)
        in_size = os.path.getsize(args.input)
        
        print(f'\n{"="*50}')
        print(f'✅ 美化完成!')
        print(f'   📄 输出: {args.output}')
        print(f'   📏 大小: {out_size:,} bytes (原 {in_size:,})')
        print(f'\n   汇总:')
        print(f'     • 四级标题套用:     +{h4_cnt}')
        print(f'     • 五级标题套用:     +{h5_cnt}')
        print(f'     • 重复编号清除:     {stripped}')
        print(f'     • 正文首行缩进:     +{body_cnt}')
        print(f'     • 编号缩进优化:     {len(num_changes)}项')
        print(f'     • 样式属性优化:     {len(sty_changes)}项')
    
    finally:
        shutil.rmtree(tmp)


if __name__ == '__main__':
    main()
