# -*- coding: utf-8 -*-
"""
Word 投标文档一键优化工具（完整版 V3）
==========================================
整合三大功能，一条命令搞定投标文件全流程优化。

功能：
  Phase 0 — H2/H3 编号跟随修复 ★NEW★
    - 为章节/H2/H3 段落插入 <w:numPr>，让 Word 自动按章节生成正确编号
      例: 第9章 → 第10章, 9.6 → 10.1, 9.6.1 → 10.1.1 ...
      解决"第10章后H2/H3仍显示旧章节号(如9.6)"的问题
    - 清除文本中残留的编号前缀避免重复显示

  Phase 1 — 编号层级修复
    - 将三级标题下断裂的四级/五级编号补全为完整层级
      例: "1.文化传承..." → "9.1.1.1 文化传承..."
          "1.1传统意象..." → "9.1.1.1.1 传统意象..."
    - 清除正文中被误加的数字前缀

  Phase 2 — 样式美化
    - 智能识别并套用 hik标题4 / hik标题5 样式
    - 清除重复编号（样式自动编号 vs 文本硬编码冲突）
    - 为中文正文添加首行缩进
    - 优化多级列表缩进值
    - 精修字体、间距、行距等排版参数

用法:
  python bid_doc_optimizer.py <input.docx> <output.docx> [选项]

选项:
  --phase {0|1|2|all} 执行阶段, 默认 all(全部执行)
  --ch-start N        编号修复起始章节号, 默认自动检测
  --ch-end N          编号修复结束章节号
  --indent-h4 N       H4缩进 twips, 默认360 (~0.63cm)
  --indent-h5 N       H5缩进 twips, 默认480 (~0.85cm)
  --dry-run           仅分析不修改
  --no-strip          跳过清除正文多余编号步骤
"""
import os, re, sys, argparse, zipfile, shutil, tempfile

# ============================================================
# 常量
# ============================================================
STYLE_HIK_H4 = "15"       # hik标题4
STYLE_HIK_H5 = "16"       # hik标题5
STYLE_HIK_BODY = "11"     # hik正文
STYLE_HIK_INDENT = "25"   # hik首行缩进两字符
DEFAULT_INDENT_H4 = 360
DEFAULT_INDENT_H5 = 480


# ============================================================
# 工具函数
# ============================================================

def unpack_docx(docx_path, output_dir):
    """解包 docx 到目录"""
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    with zipfile.ZipFile(docx_path, 'r') as zf:
        zf.extractall(output_dir)


def pack_docx(input_dir, output_path):
    """从目录打包为 docx"""
    if os.path.exists(output_path):
        os.remove(output_path)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(input_dir):
            for fname in files:
                fpath = os.path.join(root, fname)
                arcname = os.path.relpath(fpath, input_dir)
                zf.write(fpath, arcname)


def count_paragraphs(doc_xml_path):
    with open(doc_xml_path, 'r', encoding='utf-8') as f:
        content = f.read()
    return len(re.findall(r'<w:p[ >]', content))


def get_text_from_para(para_body):
    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para_body)
    return ''.join(texts).strip()


def is_likely_heading(body_text):
    t = body_text.strip()
    if not t or len(t) > 60:
        return False
    if re.search(r'[。；！？]', t) and len(t) > 15:
        return False
    # 关键: Python3中 \w 匹配中文! 用 [a-zA-Z0-9] 判断纯字母数字
    if re.match(r'^[a-zA-Z0-9\-\+,./\\()[\]]{2,}$', t) and len(t) < 20:
        return False
    return True


def insert_pStyle(p_body, style_id):
    if re.search(r'<w:pStyle\s', p_body):
        return p_body
    if '<w:pPr>' in p_body:
        return p_body.replace('<w:pPr>', f'<w:pPr><w:pStyle w:val="{style_id}"/>')
    return f'<w:pPr><w:pStyle w:val="{style_id}"/></w:pPr>{p_body}'


def safe_replace(content_list, modifications):
    """从后往前替换，返回新字符串和成功数"""
    success = 0
    for start, end, repl in reversed(modifications):
        actual = ''.join(content_list[start:end])
        expected_len = end - start
        if len(repl) <= expected_len * 3 + 100:  # 合理的替换长度
            content_list[start:end] = list(repl)
            success += 1
    return ''.join(content_list), success


# ================================================================
# PHASE 0: H2/H3 编号跟随修复
# 解决"第10章后H2/H3仍显示9.6等旧编号"的问题
# 根因: 章节(style12)/H2(style13)/H3(style14)段落缺少<w:numPr>，
#       Word无法自动生成多级列表编号
# ================================================================

def find_all_paras_positions(xml_content):
    """用标签级匹配找出所有 <w:p...>...</w:p> 段落的位置"""
    paras = []
    i = 0
    while i < len(xml_content):
        m = re.search(r'<w:p[ >]', xml_content[i:])
        if not m:
            break
        start = i + m.start()
        end_tag = xml_content.find('</w:p>', start)
        if end_tag == -1:
            break
        para_end = end_tag + len('</w:p>')
        paras.append((start, para_end, xml_content[start:para_end]))
        i = para_end
    return paras


def phase0_fix_heading_numbering(doc_xml, num_xml):
    """
    Phase 0: 为章节/H2/H3 段落插入 numPr，让 Word 自动生成正确编号。

    abstractNumId=1 的多级列表定义:
      ilvl=0: 第%1章 → style 12 (章标题)
      ilvl=1: %1.%2.   → style 13 (二级标题 H2)
      ilvl=2: %1.%2.%3.→ style 14 (三级标题 H3)

    返回: (modified_doc_xml, modified_num_xml, report_dict)
    """
    print('\n' + '='*55)
    print('  PHASE 0: H2/H3 编号跟随修复')
    print('='*55)

    # Step A: 找到引用 abstractNumId=1 的 numId
    print('\n  [A] 查找编号实例...')
    nums = re.findall(r'<w:num w:numId="(\d+)">(.*?)</w:num>', num_xml, re.DOTALL)
    target_num_id = None
    for nid, ncontent in nums:
        abs_m = re.search(r'<w:abstractNumId w:val="(\d+)"/>', ncontent)
        if abs_m and abs_m.group(1) == '1':
            target_num_id = nid
            break

    if target_num_id is None:
        max_nid = max(int(nid) for nid, _ in nums) if nums else 0
        target_num_id = str(max_nid + 1)
        new_num = f'<w:num w:numId="{target_num_id}"><w:abstractNumId w:val="1"/></w:num>'
        num_xml = num_xml.replace('</w:numbering>', new_num + '\n</w:numbering>')

    print(f'      ✓ numId={target_num_id} → abstractNumId=1 (主多级列表)')

    # Build numPr strings for each level
    numPr_ch = f'<w:numPr><w:numId w:val="{target_num_id}"/><w:ilvl w:val="0"/></w:numPr>'
    numPr_h2 = f'<w:numPr><w:numId w:val="{target_num_id}"/><w:ilvl w:val="1"/></w:numPr>'
    numPr_h3 = f'<w:numPr><w:numId w:val="{target_num_id}"/><w:ilvl w:val="2"/></w:numPr>'

    # Step B: 遍历所有段落，为 style 12/13/14 插入 numPr
    print('\n  [B] 为章节/H2/H3 插入编号引用...')
    all_paras = find_all_paras_positions(doc_xml)

    changes = {'ch': 0, 'h2': 0, 'h3': 0}
    strip_count = 0
    modifications = []

    for start, end, para in all_paras:
        style_m = re.search(r'<w:pStyle w:val="(12|13|14|15|16)"/>', para)
        if not style_m:
            continue

        style_val = style_m.group(1)
        has_numPr = bool(re.search(r'<w:numPr>', para))
        new_para = para
        modified = False

        # Insert numPr for styles 12/13/14 only (if missing)
        if not has_numPr and style_val in ('12', '13', '14'):
            if style_val == '12':
                numPr, key = numPr_ch, 'ch'
            elif style_val == '13':
                numPr, key = numPr_h2, 'h2'
            else:  # '14'
                numPr, key = numPr_h3, 'h3'

            if '<w:pPr>' in new_para:
                new_para = new_para.replace('<w:pPr>', f'<w:pPr>{numPr}', 1)
            elif '<w:pPr/>' in new_para:
                new_para = new_para.replace('<w:pPr/>', f'<w:pPr>{numPr}</w:pPr>', 1)
            else:
                new_para = re.sub(
                    r'(<w:p[ >][^>]*>)',
                    rf'\g<1><w:pPr>{numPr}</w:pPr>',
                    new_para, count=1)

            changes[key] += 1
            modified = True

        # Strip leading number prefix from text (all heading levels 12-16)
        t_match = re.search(r'(<w:t[^>]*>)(.*?)(</w:t>)', new_para)
        if t_match:
            old_text = t_match.group(2)
            # Remove leading digits+dots like "9.6." or "9.6.1."
            new_text = re.sub(r'^[\d\.]+[\s\u3000]*', '', old_text)
            if new_text != old_text:
                strip_count += 1
                new_para = (new_para[:t_match.start()] +
                           t_match.group(1) + new_text + t_match.group(3) +
                           new_para[t_match.end():])
                modified = True

        if modified:
            modifications.append((start, end, new_para))

    # Apply in reverse order to preserve positions
    for s, e, nc in reversed(modifications):
        doc_xml = doc_xml[:s] + nc + doc_xml[e:]

    print(f'      ✓ 章节标题(style12/ilvl0): +{changes["ch"]}')
    print(f'      ✓ 二级标题(style13/ilvl1): +{changes["h2"]}')
    print(f'      ✓ 三级标题(style14/ilvl2): +{changes["h3"]}')
    print(f'      ✓ 清除残留编号前缀: {strip_count} 处')

    report = {
        'num_chapter': changes['ch'],
        'num_h2': changes['h2'],
        'num_h3': changes['h3'],
        'strip_prefixes': strip_count,
    }
    total_fixed = changes['ch'] + changes['h2'] + changes['h3']
    print(f'\n  Phase 0 完成: 插入{total_fixed}个numPr + 去除{strip_count}个残留前缀')
    return doc_xml, num_xml, report


# ================================================================
# PHASE 1: 编号层级修复
# ================================================================

def analyze_hierarchy(xml_content):
    """
    分析文档结构，构建三级标题的层级映射表。
    返回: list of {'index','start_pos','end_pos','ch','h2','h3','full_prefix','text'}
    """
    ch_pattern = r'<w:pStyle w:val="12"/>'     # 章标题
    h2_pattern = r'<w:pStyle w:val="13"/>'     # 二级标题
    h3_pattern = r'<w:pStyle w:val="14"/>'     # 三级标题

    ch_matches = list(re.finditer(ch_pattern, xml_content))
    h2_matches = list(re.finditer(h2_pattern, xml_content))
    h3_matches = list(re.finditer(h3_pattern, xml_content))

    def get_heading_text(content, match_end):
        chunk = content[match_end:match_end + 800]
        texts = re.findall(r'<w:t>([^<]*)</w:t>', chunk)
        return ''.join(texts)

    # 章节边界检测
    chapter_keywords = [
        ('设计方案', 9), ('技术方案', 10), ('培训方案', 11),
        ('验收方案', 12), ('应急预案', 12), ('售后', 13),
    ]
    chapter_boundaries = []
    for m in ch_matches:
        text = get_heading_text(xml_content, m.end())
        ch_num = None
        for keyword, num in chapter_keywords:
            if keyword in text:
                ch_num = num
                break
        chapter_boundaries.append((m.start(), ch_num))

    # 构建层级映射
    hierarchy = []
    h2_idx = 0
    h3_idx_in_h2 = 0

    for i, h3_m in enumerate(h3_matches):
        pos = h3_m.start()

        while h2_idx < len(h2_matches) - 1 and h2_matches[h2_idx + 1].start() < pos:
            h2_idx += 1
            h3_idx_in_h2 = 0

        ch = 9
        for cb_pos, cb_num in chapter_boundaries:
            if pos >= cb_pos and cb_num is not None:
                ch = cb_num

        text = get_heading_text(xml_content, h3_m.end())
        hierarchy.append({
            'index': i,
            'start_pos': pos,
            'end_pos': h3_m.end(),
            'ch': ch,
            'h2': h2_idx + 1,
            'h3': h3_idx_in_h2 + 1,
            'full_prefix': f'{ch}.{h2_idx + 1}.{h3_idx_in_h2 + 1}',
            'text': text[:60],
        })
        h3_idx_in_h2 += 1

    return hierarchy, {
        'chapter_count': len(chapter_boundaries),
        'h2_count': len(h2_matches),
        'h3_count': len(h3_matches),
    }


def find_numbering_fixes(xml_content, hierarchy, ch_start=None, ch_end=None):
    """
    在每个三级标题下查找需要补全层级的四/五级标题。
    返回: [(abs_pos, old_text, new_text), ...]
    """
    replacements = []

    target = hierarchy
    if ch_start or ch_end:
        target = [item for item in hierarchy
                  if (ch_start is None or item['ch'] >= ch_start) and
                     (ch_end is None or item['ch'] <= ch_end)]

    for i, h3_info in enumerate(target):
        search_start = h3_info['end_pos']

        # 找下一个H3位置作为搜索边界
        next_h3_pos = None
        for j in range(hierarchy.index(h3_info) + 1, len(hierarchy)):
            if hierarchy[j]['start_pos'] > search_start:
                next_h3_pos = hierarchy[j]['start_pos']
                break
        search_end = next_h3_pos or (search_start + 50000)

        region = xml_content[search_start:search_end]
        t_matches = list(re.finditer(r'<w:t>([^<]+)</w:t>', region))

        for tm in t_matches:
            txt = tm.group(1)

            # 五级: N.N + 中文 → prefix.N.N 中文
            m5 = re.match(r'^(\d+)\.(\d+)([\u4e00-\u9fff].*)$', txt)
            if m5:
                new_txt = f"{h3_info['full_prefix']}.{int(m5.group(1))}.{int(m5.group(2))}{m5.group(3)}"
                replacements.append((search_start + tm.start(1), txt, new_txt))
                continue

            # 四级: N. + 中文 → prefix.N. 中文
            m4 = re.match(r'^(\d+)\.([\u4e00-\u9fff].*)$', txt)
            if m4:
                new_txt = f"{h3_info['full_prefix']}.{int(m4.group(1))}{m4.group(2)}"
                replacements.append((search_start + tm.start(1), txt, new_txt))

    return replacements


def strip_extra_numbers(xml_content):
    """
    清除正文中的多余数字前缀。
    规则: 保留 4/5级短标题文本; 删除其他带数字前缀的长段落文本中的编号。
    返回: (modified_content, stripped_count)
    """
    t_matches = list(re.finditer(r'<w:t>([^<]+)</w:t>', xml_content))
    strip_items = []

    for tm in t_matches:
        txt = tm.group(1)
        m = re.match(r'^(\d+(?:\.\d+)+)([\u4e00-\u9fff].*)$', txt)
        if not m:
            continue

        num_part = m.group(1)
        text_part = m.group(2)
        level = num_part.count('.') + 1

        # 只保留4/5级且短的标题文本
        is_valid = (level in (4, 5) and len(text_part) <= 28
                    and '。' not in text_part[:20] and '，' not in text_part[:15])

        if not is_valid:
            strip_items.append((tm.start(1), tm.end(1), text_part))

    content_list = list(xml_content)
    for start, end, repl in reversed(strip_items):
        content_list[start:end] = list(repl)

    return ''.join(content_list), len(strip_items)


def phase1_fix_numbering(xml_content, ch_start=None, ch_end=None, no_strip=False):
    """
    Phase 1 完整流程: 分析 → 查找修复项 → 执行替换 → 清除多余编号
    返回: (modified_xml, report_dict)
    """
    print('\n' + '='*55)
    print('  PHASE 1: 编号层级修复')
    print('='*55)

    # Step A: 分析结构
    print('\n  [A] 分析文档层级结构...')
    hierarchy, stats = analyze_hierarchy(xml_content)
    print(f'      章节标记: {stats["chapter_count"]} 个')
    print(f'      二级标题(pStyle=13): {stats["h2_count"]} 个')
    print(f'      三级标题(pStyle=14): {stats["h3_count"]} 个')

    if not hierarchy:
        print('      ⚠ 未找到三级标题，跳过Phase 1')
        return xml_content, {'fix_count': 0, 'strip_count': 0}

    chapters_found = sorted(set(item['ch'] for item in hierarchy))
    print(f'      涉及章节: {chapters_found}')

    # Step B: 查找需修复项
    print('\n  [B] 检查编号问题...')
    fixes = find_numbering_fixes(xml_content, hierarchy, ch_start, ch_end)
    print(f'      发现 {len(fixes)} 处需修复的编号')

    if not fixes:
        print('      ✅ 编号格式正确，无需修复')

    # 显示预览
    if fixes:
        print('      预览 (前8条):')
        for idx, (_, old, new) in enumerate(fixes[:8]):
            print(f'        {idx+1}. "{old[:35]}" → "{new[:45]}"')
        if len(fixes) > 8:
            print(f'        ... 还有 {len(fixes)-8} 条')

    # Step C: 执行替换
    fix_count = 0
    if fixes:
        print('\n  [C] 执行编号替换...')
        content_list = list(xml_content)
        sorted_fixes = sorted(fixes, key=lambda x: x[0], reverse=True)

        for abs_pos, old_txt, new_txt in sorted_fixes:
            actual = ''.join(content_list[abs_pos:abs_pos + len(old_txt)])
            if actual == old_txt:
                content_list[abs_pos:abs_pos + len(old_txt)] = list(new_txt)
                fix_count += 1

        xml_content = ''.join(content_list)
        print(f'      ✓ 编号修复: {fix_count}/{len(fixes)} 处成功')

    # Step D: 清除正文多余编号
    strip_count = 0
    if not no_strip:
        print('\n  [D] 清除正文多余编号...')
        xml_content, strip_count = strip_extra_numbers(xml_content)
        print(f'      ✓ 清除多余编号: {strip_count} 处')

    report = {'fix_count': fix_count, 'strip_count': strip_count}
    print(f'\n  Phase 1 完成: 修复{fix_count}处编号 + 清除{strip_count}处冗余')
    return xml_content, report


# ================================================================
# PHASE 2: 样式美化
# ================================================================

def step_apply_heading_styles(content):
    """识别并套用 H4/H5 标题样式"""
    para_pattern = r'(<w:p[ >])(.*?)(</w:p>)'
    paras = list(re.finditer(para_pattern, content, re.DOTALL))
    counters = {'h4': 0, 'h5': 0}
    modifications = []

    for m in paras:
        p_body = m.group(2)
        if re.search(r'<w:pStyle', p_body):
            continue

        full_text = get_text_from_para(p_body)
        if not full_text:
            continue

        # 五级: X.X.X.X.X 文字 (5个数字捕获组 + 1个文本 = group6)
        m5 = re.match(r'^(\d+)\.(\d+)\.(\d+)\.(\d+)\.(\d+)[ \t]*(.*)$', full_text)
        if m5 and is_likely_heading(m5.group(6)):
            new_pPr = insert_pStyle(p_body, STYLE_HIK_H5)
            if new_pPr != p_body:
                counters['h5'] += 1
                modifications.append((m.start(2), m.end(2), new_pPr))
            continue

        # 四级: X.X.X.X 文字
        m4 = re.match(r'^(\d+)\.(\d+)\.(\d+)\.(\d+)([ \t]*)(.*)$', full_text)
        if m4 and is_likely_heading(m4.group(6)):
            new_pPr = insert_pStyle(p_body, STYLE_HIK_H4)
            if new_pPr != p_body:
                counters['h4'] += 1
                modifications.append((m.start(2), m.end(2), new_pPr))

    content_list = list(content)
    result, _ = safe_replace(content_list, modifications)
    return result, counters['h4'], counters['h5']


def step_strip_duplicate_prefixes(content):
    """清除已套用样式的H4/H5段落的重复数字前缀"""
    para_pattern = r'(<w:p[ >])(.*?)(</w:p>)'
    paras = list(re.finditer(para_pattern, content, re.DOTALL))
    stripped = 0
    modifications = []

    for m in paras:
        p_body = m.group(2)
        if not re.search(rf'w:val="({STYLE_HIK_H4}|{STYLE_HIK_H5})"', p_body):
            continue

        t_match = re.search(r'(<w:t[^>]*>)([^<]*)(</w:t>)', p_body)
        if not t_match:
            continue

        text_val = t_match.group(2).strip()
        if not text_val:
            continue

        new_text = re.sub(r'^\d+(\.\d+){3,}[ \t]*(?=.)', '', text_val, count=1)
        if new_text != text_val and new_text.strip():
            old_full = t_match.group(0)
            new_full = t_match.group(1) + new_text + t_match.group(3)
            modifications.append((m.start(2) + t_match.start(),
                                 m.start(2) + t_match.end(), new_full))
            stripped += 1

    content_list = list(content)
    result, _ = safe_replace(content_list, modifications)
    return result, stripped


def step_apply_body_indent(content):
    """为无样式的中文正文添加首行缩进"""
    para_pattern = r'(<w:p[ >])(.*?)(</w:p>)'
    paras = list(re.finditer(para_pattern, content, re.DOTALL))
    body_count = 0
    modifications = []

    for m in paras:
        p_body = m.group(2)
        if re.search(r'<w:pStyle', p_body):
            continue

        full_text = get_text_from_para(p_body)
        if not full_text:
            continue
        if re.match(r'^[\d\s.,;:\-+=\\/\\(){}\[\]\"\'<>]+$', full_text):
            continue
        if re.match(r'^\d+\.', full_text):
            continue

        new_pPr = insert_pStyle(p_body, STYLE_HIK_INDENT)
        if new_pPr != p_body:
            body_count += 1
            modifications.append((m.start(2), m.end(2), new_pPr))

    content_list = list(content)
    result, _ = safe_replace(content_list, modifications)
    return result, body_count


def optimize_numbering(num_xml, indent_h4=DEFAULT_INDENT_H4, indent_h5=DEFAULT_INDENT_H5):
    """调整 numbering.xml 中 H4/H5 缩进值"""
    changes = []

    def replace_ind(xml, abs_id, ilvl, new_left, label):
        pat = (rf'(<w:abstractNum w:abstractNumId="{abs_id}"[^>]*>.*?'
               rf'<w:lvl w:ilvl="{ilvl}"[^>]*>.*?'
               r'<w:ind w:left=")(\d+)(" w:hanging=")(\d+)("/>)')
        m = re.search(pat, xml, re.DOTALL)
        if m:
            old_l, old_h = int(m.group(2)), int(m.group(4))
            xml = re.sub(pat,
                        rf'\g<1>{new_left}\g<3>{new_left}\g<5>',
                        xml, count=1, flags=re.DOTALL)
            changes.append(f'{label}: left={old_l}→{new_left}, hanging={old_h}→{new_left}')
        return xml

    result = num_xml
    result = replace_ind(result, '1', '3', indent_h4, 'H4缩进(ilvl=3)')
    result = replace_ind(result, '1', '4', indent_h5, 'H5缩进(ilvl=4)')
    return result, changes


def optimize_styles(style_xml):
    """精修 hik标题4/5 样式属性"""
    changes = []
    result = style_xml

    # --- hik标题4 ---
    old_s = r'<w:spacing w:before="120" w:after="120" w:line="360" w:lineRule="auto"/>'
    new_s = '<w:spacing w:before="60" w:after="40" w:line="280" w:lineRule="auto"/>'
    if re.search(old_s, result):
        result = re.sub(old_s, new_s, result)
        changes.append('hik标题4: 段前6pt→3pt, 段后6pt→2pt, 行距280')

    old_f = r'<w:rFonts w:ascii="黑体" w:hAnsi="宋体" w:eastAsia="黑体" w:cs="Times New Roman"/>'
    new_f = '<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:cs="黑体"/>'
    if re.search(old_f, result):
        result = re.sub(old_f, new_f, result)
        changes.append('hik标题4: 字体统一黑体')

    # --- hik标题5 ---
    m5ppr = re.search(
        r'(<w:style w:type="paragraph"[^>]*w:styleId="16"[^>]*>.*?<w:pPr>)(.*?)(</w:pPr>)',
        result, re.DOTALL)
    if m5ppr:
        pc = m5ppr.group(2)
        if not re.search('<w:spacing', pc):
            npc = pc + '<w:spacing w:before="40" w:after="20" w:line="276" w:lineRule="auto"/>'
            result = result[:m5ppr.start(2)] + npc + result[m5ppr.end(2):]
            changes.append('hik标题5: 新增段前2pt/段后1pt/行距')

    old_h5f = r'<w:rFonts w:ascii="宋体" w:hAnsi="黑体" w:eastAsia="宋体" w:cs="宋体"/>'
    new_h5f = '<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:cs="黑体"/>'
    if re.search(old_h5f, result):
        result = re.sub(old_h5f, new_h5f, result)
        changes.append('hik标题5: 字体改为黑体')

    m5rpr = re.search(
        r'(<w:style w:type="paragraph"[^>]*w:styleId="16"[^>]*>.*?<w:rPr>)(.*?)(</w:rPr>)',
        result, re.DOTALL)
    if m5rpr and not re.search('<w:sz ', m5rpr.group(2)):
        nr = m5rpr.group(2) + '<w:sz w:val="21"/><w:szCs w:val="21"/>'
        result = result[:m5rpr.start(2)] + nr + result[m5rpr.end(2):]
        changes.append('hik标题5: 新增字号10.5pt')

    return result, changes


def phase2_beautify(xml_content, num_xml, sty_xml,
                    indent_h4=DEFAULT_INDENT_H4, indent_h5=DEFAULT_INDENT_H5):
    """
    Phase 2 完整流程: 套用样式 → 清除重复编号 → 正文缩进 → 缩进优化 → 样式精修
    返回: (doc_xml, num_xml, sty_xml, report_dict)
    """
    print('\n' + '='*55)
    print('  PHASE 2: 样式美化')
    print('='*55)

    # Step A: 套用 H4/H5 标题样式
    print('\n  [A] 识别并套用 H4/H5 标题样式...')
    doc_xml, h4_cnt, h5_cnt = step_apply_heading_styles(xml_content)
    print(f'      ✓ 四级标题(hik标题4): +{h4_cnt}')
    print(f'      ✓ 五级标题(hik标题5): +{h5_cnt}')

    # Step B: 清除重复编号前缀
    print('\n  [B] 清除标题中重复编号前缀...')
    doc_xml, dup_stripped = step_strip_duplicate_prefixes(doc_xml)
    print(f'      ✓ 清除重复编号: {dup_stripped} 处')

    # Step C: 正文首行缩进
    print('\n  [C] 应用正文首行缩进样式...')
    doc_xml, body_cnt = step_apply_body_indent(doc_xml)
    print(f'      ✓ 正文首行缩进: +{body_cnt}')

    # Step D: 优化编号缩进
    print(f'\n  [D] 优化编号缩进 (H4={indent_h4}, H5={indent_h5} twips)...')
    num_xml, num_changes = optimize_numbering(num_xml, indent_h4, indent_h5)
    for c in num_changes:
        print(f'      ✓ {c}')
    if not num_changes:
        print(f'      ⚠ 未找到HIK编号体系定义')

    # Step E: 样式属性精修
    print('\n  [E] 优化标题样式属性...')
    sty_xml, sty_changes = optimize_styles(sty_xml)
    for c in sty_changes:
        print(f'      ✓ {c}')

    report = {
        'h4_count': h4_cnt,
        'h5_count': h5_cnt,
        'dup_stripped': dup_stripped,
        'body_indent': body_cnt,
        'num_optimizations': len(num_changes),
        'style_optimizations': len(sty_changes),
    }
    print(f'\n  Phase 2 完成: H4+{h4_cnt} / H5+{h5_cnt} / 去重{dup_stripped} / 缩进{body_cnt}')
    return doc_xml, num_xml, sty_xml, report


# ================================================================
# 主入口
# ================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Word 投标文档一键优化工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python bid_doc_optimizer.py input.docx output.docx
  python bid_doc_optimizer.py input.docx output.docx --phase 1
  python bid_doc_optimizer.py input.docx output.docx --indent-h4 240 --indent-h5 320
  python bid_doc_optimizer.py input.docx output.docx --dry-run
""")
    parser.add_argument('input', help='输入 .docx 文件路径')
    parser.add_argument('output', help='输出 .docx 文件路径')
    parser.add_argument('--phase', choices=['0', '1', '2', 'all'], default='all',
                        help='执行阶段: 0=仅H2/H3编号修复, 1=仅编号层级, 2=仅样式美化, all=全部 (默认)')
    parser.add_argument('--ch-start', type=int, default=None,
                        help='编号修复起始章节号')
    parser.add_argument('--ch-end', type=int, default=None,
                        help='编号修复结束章节号')
    parser.add_argument('--indent-h4', type=int, default=DEFAULT_INDENT_H4,
                        help=f'H4缩进 twips (默认{DEFAULT_INDENT_H4}, ~0.63cm)')
    parser.add_argument('--indent-h5', type=int, default=DEFAULT_INDENT_H5,
                        help=f'H5缩进 twips (默认{DEFAULT_INDENT_H5}, ~0.85cm)')
    parser.add_argument('--dry-run', action='store_true',
                        help='仅分析不修改')
    parser.add_argument('--no-strip', action='store_true',
                        help='跳过清除正文多余编号步骤')
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f'❌ 输入文件不存在: {args.input}')
        sys.exit(1)

    tmp = tempfile.mkdtemp(prefix='bidopt_')
    try:
        # 解包
        print(f'📦 解包: {os.path.basename(args.input)}')
        unpack_docx(args.input, tmp)

        orig_paras = count_paragraphs(os.path.join(tmp, 'word', 'document.xml'))
        print(f'   原始段落数: {orig_paras}')

        # 读取核心XML
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        num_path = os.path.join(tmp, 'word', 'numbering.xml')
        sty_path = os.path.join(tmp, 'word', 'styles.xml')

        with open(doc_path, 'r', encoding='utf-8') as f:
            doc_xml = f.read()
        with open(num_path, 'r', encoding='utf-8') as f:
            num_xml = f.read()
        with open(sty_path, 'r', encoding='utf-8') as f:
            sty_xml = f.read()

        # ---- Phase 0: H2/H3 编号跟随修复 ----
        rpt0 = {}
        if args.phase in ('0', 'all'):
            doc_xml, num_xml, rpt0 = phase0_fix_heading_numbering(doc_xml, num_xml)

            with open(doc_path, 'w', encoding='utf-8') as f:
                f.write(doc_xml)
            with open(num_path, 'w', encoding='utf-8') as f:
                f.write(num_xml)

        # ---- Phase 1: 编号修复 ----
        rpt1 = {}
        if args.phase in ('1', 'all'):
            doc_xml, rpt1 = phase1_fix_numbering(
                doc_xml, args.ch_start, args.ch_end, args.no_strip)

            with open(doc_path, 'w', encoding='utf-8') as f:
                f.write(doc_xml)

        # ---- Phase 2: 样式美化 ----
        rpt2 = {}
        if args.phase in ('2', 'all'):
            doc_xml, num_xml, sty_xml, rpt2 = phase2_beautify(
                doc_xml, num_xml, sty_xml, args.indent_h4, args.indent_h5)

            with open(doc_path, 'w', encoding='utf-8') as f:
                f.write(doc_xml)
            with open(num_path, 'w', encoding='utf-8') as f:
                f.write(num_xml)
            with open(sty_path, 'w', encoding='utf-8') as f:
                f.write(sty_xml)

        # 验证 & 打包
        final_paras = count_paragraphs(doc_path)
        print(f'\n{"="*55}')
        print(f'  验证:')
        print(f'  段落数: {orig_paras} → {final_paras}', end='')
        if final_paras == orig_paras:
            print(' ✅ 一致')
        else:
            print(f' ⚠️ 差异={final_paras - orig_paras}')

        pack_docx(tmp, args.output)
        out_size = os.path.getsize(args.output)
        in_size = os.path.getsize(args.input)

        print(f'\n{"="*55}')
        print(f'✅ 优化完成!')
        print(f'   📄 输出: {args.output}')
        print(f'   📏 大小: {out_size:,} bytes (原 {in_size:,})')
        print(f'\n   ══════════ 汇总报告 ══════════')

        if args.phase in ('0', 'all'):
            print(f'\n   📋 Phase 0 - H2/H3 编号跟随修复:')
            print(f'      • 章节标题插入numPr: {rpt0.get("num_chapter", 0)} 处')
            print(f'      • 二级标题插入numPr: {rpt0.get("num_h2", 0)} 处')
            print(f'      • 三级标题插入numPr: {rpt0.get("num_h3", 0)} 处')
            print(f'      • 残留编号前缀清除:   {rpt0.get("strip_prefixes", 0)} 处')

        if args.phase in ('1', 'all'):
            print(f'\n   📋 Phase 1 - 编号层级修复:')
            print(f'      • 层级编号修复:   {rpt1.get("fix_count", 0)} 处')
            print(f'      • 多余编号清除:   {rpt1.get("strip_count", 0)} 处')

        if args.phase in ('2', 'all'):
            print(f'\n   📋 Phase 2 - 样式美化:')
            print(f'      • 四级标题套用:   +{rpt2.get("h4_count", 0)} 处')
            print(f'      • 五级标题套用:   +{rpt2.get("h5_count", 0)} 处')
            print(f'      • 重复编号清除:   {rpt2.get("dup_stripped", 0)} 处')
            print(f'      • 正文首行缩进:   +{rpt2.get("body_indent", 0)} 处')
            print(f'      • 编号缩进优化:   {rpt2.get("num_optimizations", 0)} 项')
            print(f'      • 样式属性优化:   {rpt2.get("style_optimizations", 0)} 项')

    finally:
        shutil.rmtree(tmp)


if __name__ == '__main__':
    main()
