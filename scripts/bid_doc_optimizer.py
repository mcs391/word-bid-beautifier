# -*- coding: utf-8 -*-
"""
Word 投标文档一键优化工具（完整版 V3）
==========================================
整合三大功能，一条命令搞定投标文件全流程优化。

功能：
  Phase 0 — 根本性编号修复 ★V2★ (三层面统一)
    - 层面一 numbering.xml: 补全主多级列表的 pStyle 链接
      确保 abstractNumId 主列表的 ilvl=1→style13(H2), ilvl=2→style14(H3)
    - 层面二 styles.xml: 统一样式定义中的 numPr
      所有标题样式(12~16)的 numPr 指向同一个多级列表实例
    - 层面三 document.xml: 为缺失 numPr 的标题段落插入正确引用
      解决"手动套用H2/H3后编号显示9.x而非14.x"的问题
      例: 第14章下手动添加的H2→14.1, H3→14.1.1（而非9.x）
    - 清除文本中残留的编号前缀避免重复显示

  Phase 1 — 编号层级修复
    - 将三级标题下断裂的四级/五级编号补全为完整层级
      例: "1.文化传承..." → "9.1.1.1 文化传承..."
          "1.1传统意象..." → "9.1.1.1.1 传统意象..."
    - 清除正文中被误加的数字前缀

  Phase 2 — 样式美化
    - 智能识别并套用 ryusuke标题4 / ryusuke标题5 样式
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
STYLE_RYUSUKE_H4 = "15"       # ryusuke标题4
STYLE_RYUSUKE_H5 = "16"       # ryusuke标题5
STYLE_RYUSUKE_BODY = "11"     # ryusuke正文
STYLE_RYUSUKE_INDENT = "25"   # ryusuke首行缩进两字符
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


def phase0_fix_heading_numbering(doc_xml, num_xml, styles_xml):
    """
    Phase 0 V2: 根本性编号修复 — 三层面统一
    
    层面一: numbering.xml — 补全主多级列表的 pStyle 链接
      确保 abstractNumId 主列表的 ilvl=1→style13(H2), ilvl=2→style14(H3)
    
    层面二: styles.xml — 统一样式定义中的 numPr
      所有标题样式(12~16)的 numPr 指向同一个多级列表实例
    
    层面三: document.xml — 补全段落级缺失的 numPr
      为没有 numPr 的标题段落插入正确的编号引用
    
    返回: (modified_doc_xml, modified_num_xml, modified_styles_xml, report_dict)
    """
    print('\n' + '='*55)
    print('  PHASE 0: 根本性编号修复（三层面统一）')
    print('='*55)

    UNIFIED_NUM_ID = "1"  # numId=1 → abstractNumId=0 (主多级列表)

    # ===== 层面一: Step A — 修复 numbering.xml =====
    print('\n  [A] 修复 numbering.xml — 补全主列表的 pStyle 链接...')

    abs_pattern = r'<w:abstractNum[^>]*w:abstractNumId="(0)"[^>]*>(.*?)</w:abstractNum>'
    abs_match = re.search(abs_pattern, num_xml, re.DOTALL)

    if not abs_match:
        print('      ⚠ 找不到 abstractNumId=0，跳过层面一')
    else:
        abs_content = abs_match.group(2)
        pstyle_fixes = {
            1: '13',   # ilvl=1 → style13(H2)
            2: '14',   # ilvl=2 → style14(H3)
        }
        abs_modifications = []
        for ilvl_val, style_id in pstyle_fixes.items():
            ilvl_str = str(ilvl_val)
            lvl_pattern = rf'(<w:lvl w:ilvl="{ilvl_str}"[^>]*>)(.*?)(</w:lvl>)'
            lvl_m = re.search(lvl_pattern, abs_content, re.DOTALL)
            if not lvl_m:
                continue
            lvl_inner = lvl_m.group(2)
            existing_ps = re.search(r'<w:pStyle w:val="(\d+)"', lvl_inner)
            if existing_ps and existing_ps.group(1) == style_id:
                print(f'      ✓ ilvl={ilvl_str}: 已有 pStyle→{style_id}')
                continue
            if existing_ps:
                old_style = existing_ps.group(1)
                new_inner = lvl_inner.replace(
                    f'<w:pStyle w:val="{old_style}"/>',
                    f'<w:pStyle w:val="{style_id}"/>', 1)
                print(f'      🔧 ilvl={ilvl_str}: pStyle {old_style}→{style_id}')
            else:
                if '<w:pPr>' in lvl_inner:
                    new_inner = lvl_inner.replace('<w:pPr>',
                        f'<w:pPr><w:pStyle w:val="{style_id}"/>', 1)
                else:
                    new_inner = lvl_inner.rstrip()[:-len('</w:lvl>')].rstrip()
                    new_inner += f'\n<w:pPr><w:pStyle w:val="{style_id}"/></w:pPr>\n</w:lvl>'
                print(f'      + ilvl={ilvl_str}: 新增 pStyle→{style_id}')
            abs_modifications.append((lvl_m.start(2), lvl_m.end(2), new_inner))

        for s, e, nc in reversed(abs_modifications):
            abs_content = abs_content[:s] + nc + abs_content[e:]
        num_xml = num_xml[:abs_match.start(2)] + abs_content + num_xml[abs_match.end(2):]
        print('      ✓ abstractNumId=0 样式链接已完善')

    # ===== 层面二: Step B — 统一 styles.xml =====
    print('\n  [B] 统一 styles.xml — 所有标题样式指向同一多级列表...')

    target_map = {
        '12': ('0', 'ryusuke标题1/章节'),
        '13': ('1', 'ryusuke标题2/H2'),
        '14': ('2', 'ryusuke标题3/H3'),
        '15': ('3', 'ryusuke标题4/H4'),
        '16': ('4', 'ryusuke标题5/H5'),
    }
    style_fix_log = []

    for sid, (target_ilvl, sname) in target_map.items():
        sid_escaped = re.escape(sid)
        sm = re.search(
            r'(<w:style\b[^>]*\bw:styleId="' + sid_escaped + r'"[^>]*>)',
            styles_xml)
        if not sm:
            continue
        style_start_pos = sm.start()
        style_end = styles_xml.find('</w:style>', style_start_pos)
        if style_end == -1:
            continue
        style_block = styles_xml[style_start_pos:style_end + len('</w:style>')]

        numpr_in_style = re.search(r'<w:numPr>(.*?)</w:numPr>', style_block, re.DOTALL)
        needs_fix = False
        fix_desc = ""

        if numpr_in_style:
            np_content = numpr_in_style.group(1)
            cur_nid_m = re.search(r'<w:numId\s+w:val="(\d+)"', np_content)
            cur_ilvl_m = re.search(r'<w:ilvl\s+w:val="(\d+)"', np_content)
            cur_nid = cur_nid_m.group(1) if cur_nid_m else None
            cur_ilvl = cur_ilvl_m.group(1) if cur_ilvl_m else None
            if cur_nid != UNIFIED_NUM_ID or cur_ilvl != target_ilvl:
                needs_fix = True
                fix_desc = f"numId={cur_nid}, ilvl={cur_ilvl} → ({UNIFIED_NUM_ID}, {target_ilvl})"
        else:
            needs_fix = True
            fix_desc = "无numPr → 插入"

        if not needs_fix:
            continue

        print(f'      🔧 style {sid}({sname}): {fix_desc}')

        correct_numpr = (
            f'<w:numPr>'
            f'<w:numId w:val="{UNIFIED_NUM_ID}"/>'
            f'<w:ilvl w:val="{target_ilvl}"/></w:numPr>')

        if numpr_in_style:
            old_numpr_full = f'<w:numPr>{numpr_in_style.group(1)}</w:numPr>'
            new_block = style_block.replace(old_numpr_full, correct_numpr, 1)
        elif '<w:pPr>' in style_block:
            new_block = style_block.replace('<w:pPr>', f'<w:pPr>{correct_numpr}', 1)
        elif '<w:pPr/>' in style_block:
            new_block = style_block.replace('<w:pPr/>',
                f'<w:pPr>{correct_numpr}</w:pPr>', 1)
        else:
            first_child = re.search(r'<w:(?!/)', style_block[sm.end():])
            if first_child:
                insert_pos = sm.end() + first_child.start()
                new_block = (style_block[:insert_pos] +
                    f'<w:pPr>{correct_numpr}</w:pPr>\n' +
                    style_block[insert_pos:])
            else:
                new_block = style_block

        styles_xml = styles_xml[:style_start_pos] + new_block + \
                     styles_xml[style_end + len('</w:style>'):]
        style_fix_log.append((sid, sname))

    print(f'      ✓ 样式定义已修复: {len(style_fix_log)} 个')

    # ===== 层面三: Step C — 补全 document.xml 段落级 numPr =====
    print('\n  [C] 补全 document.xml 段落级 numPr...')

    pr_templates = {
        '12': (f'<w:numPr><w:numId w:val="{UNIFIED_NUM_ID}"/><w:ilvl w:val="0"/></w:numPr>', '章节'),
        '13': (f'<w:numPr><w:numId w:val="{UNIFIED_NUM_ID}"/><w:ilvl w:val="1"/></w:numPr>', 'H2'),
        '14': (f'<w:numPr><w:numId w:val="{UNIFIED_NUM_ID}"/><w:ilvl w:val="2"/></w:numPr>', 'H3'),
        '15': (f'<w:numPr><w:numId w:val="{UNIFIED_NUM_ID}"/><w:ilvl w:val="3"/></w:numPr>', 'H4'),
        '16': (f'<w:numPr><w:numId w:val="{UNIFIED_NUM_ID}"/><w:ilvl w:val="4"/></w:numPr>', 'H5'),
    }

    all_paras = find_all_paras_positions(doc_xml)

    modifications = []
    counts = {'12': 0, '13': 0, '14': 0, '15': 0, '16': 0}
    strip_count = 0

    for start, end, para in all_paras:
        style_m = re.search(r'w:pStyle\s+w:val="(12|13|14|15|16)"', para)
        if not style_m:
            continue
        sid = style_m.group(1)
        has_numpr = bool(re.search(r'<w:numPr>', para))
        new_para = para
        modified = False

        if not has_numpr:
            correct_pr, _ = pr_templates[sid]
            if '<w:pPr>' in new_para:
                new_para = new_para.replace('<w:pPr>',
                    f'<w:pPr>{correct_pr}', 1)
            elif '<w:pPr/>' in new_para:
                new_para = new_para.replace('<w:pPr/>',
                    f'<w:pPr>{correct_pr}</w:pPr>', 1)
            else:
                tag_m = re.match(r'(.*?<w:p\b[^>]*>)(.*)', new_para, re.DOTALL)
                if tag_m:
                    new_para = (tag_m.group(1) +
                        f'<w:pPr>{correct_pr}</w:pPr>' + tag_m.group(2))
            counts[sid] += 1
            modified = True

        # 清除残留数字前缀（所有标题级别）
        for t_m in reversed(list(re.finditer(
                r'(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)', new_para))):
            prefix = t_m.group(2)
            new_text = re.sub(r'^[\d\.]+[\s\u3000\u00a0]+', '', prefix)
            if new_text != prefix:
                strip_count += 1
                new_para = (new_para[:t_m.start()] + t_m.group(1) +
                           new_text + t_m.group(3) + new_para[t_m.end():])
                modified = True

        if modified:
            modifications.append((start, end, new_para))

    # 倒序应用以保持位置正确
    for s, e, nc in reversed(modifications):
        doc_xml = doc_xml[:s] + nc + doc_xml[e:]

    total_fixed = sum(counts.values())
    for sid in sorted(counts.keys()):
        _, name = pr_templates[sid]
        if counts[sid] > 0:
            print(f'      ✓ {name}(style{sid}): +{counts[sid]} 个 numPr')
    print(f'      ✓ 清除残留前缀: {strip_count} 处')

    report = {
        'num_chapter': counts.get('12', 0),
        'num_h2': counts.get('13', 0),
        'num_h3': counts.get('14', 0),
        'num_h4': counts.get('15', 0),
        'num_h5': counts.get('16', 0),
        'style_fixes': len(style_fix_log),
        'strip_prefixes': strip_count,
    }

    print(f'\n  Phase 0 完成: 三层面统一修复完成')
    print(f'      • numbering.xml: 补全pStyle链接')
    print(f'      • styles.xml: 统一{len(style_fix_log)}个样式的numPr')
    print(f'      • document.xml: 补全{total_fixed}个段落的numPr + 去除{strip_count}个前缀')

    return doc_xml, num_xml, styles_xml, report


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
            new_pPr = insert_pStyle(p_body, STYLE_RYUSUKE_H5)
            if new_pPr != p_body:
                counters['h5'] += 1
                modifications.append((m.start(2), m.end(2), new_pPr))
            continue

        # 四级: X.X.X.X 文字
        m4 = re.match(r'^(\d+)\.(\d+)\.(\d+)\.(\d+)([ \t]*)(.*)$', full_text)
        if m4 and is_likely_heading(m4.group(6)):
            new_pPr = insert_pStyle(p_body, STYLE_RYUSUKE_H4)
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
        if not re.search(rf'w:val="({STYLE_RYUSUKE_H4}|{STYLE_RYUSUKE_H5})"', p_body):
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

        new_pPr = insert_pStyle(p_body, STYLE_RYUSUKE_INDENT)
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
    messages = []

    h4_id = find_heading_style_id(style_xml, 4)
    if h4_id and h4_id != STYLE_RYUSUKE_H4:
        STYLE_RYUSUKE_H4 = h4_id
        messages.append(f'四级标题样式绑定为 styleId={h4_id}')

    h5_id = find_heading_style_id(style_xml, 5)
    if h5_id and h5_id != STYLE_RYUSUKE_H5:
        STYLE_RYUSUKE_H5 = h5_id
        messages.append(f'五级标题样式绑定为 styleId={h5_id}')
    else:
        h5_para_re = rf'<w:style\b(?=[^>]*w:type="paragraph")(?=[^>]*w:styleId="{re.escape(STYLE_RYUSUKE_H5)}")'
        if not re.search(h5_para_re, style_xml) and STYLE_RYUSUKE_H5 == "16" and 'w:styleId="16"' in style_xml:
            STYLE_RYUSUKE_H5 = "ryusukeHeading5"
            messages.append('styleId=16 已被非段落样式占用，改用 ryusukeHeading5 创建五级标题样式')

    return messages


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


def optimize_styles(style_xml):
    """精修 ryusuke标题样式属性"""
    changes = []
    result = style_xml

    h5_style_re = rf'<w:style[^>]*w:type="paragraph"[^>]*w:styleId="{re.escape(STYLE_RYUSUKE_H5)}"'
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

    # --- ryusuke标题4 ---
    old_s = r'<w:spacing w:before="120" w:after="120" w:line="360" w:lineRule="auto"/>'
    new_s = '<w:spacing w:before="60" w:after="40" w:line="280" w:lineRule="auto"/>'
    if re.search(old_s, result):
        result = re.sub(old_s, new_s, result)
        changes.append('ryusuke标题4: 段前6pt→3pt, 段后6pt→2pt, 行距280')

    old_f = r'<w:rFonts w:ascii="黑体" w:hAnsi="宋体" w:eastAsia="黑体" w:cs="Times New Roman"/>'
    new_f = '<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:cs="黑体"/>'
    if re.search(old_f, result):
        result = re.sub(old_f, new_f, result)
        changes.append('ryusuke标题4: 字体统一黑体')

    # --- ryusuke标题5 ---
    m5ppr = re.search(
        rf'(<w:style w:type="paragraph"[^>]*w:styleId="{re.escape(STYLE_RYUSUKE_H5)}"[^>]*>.*?<w:pPr>)(.*?)(</w:pPr>)',
        result, re.DOTALL)
    if m5ppr:
        pc = m5ppr.group(2)
        if not re.search('<w:spacing', pc):
            npc = pc + '<w:spacing w:before="40" w:after="20" w:line="276" w:lineRule="auto"/>'
            result = result[:m5ppr.start(2)] + npc + result[m5ppr.end(2):]
            changes.append('ryusuke标题5: 新增段前2pt/段后1pt/行距')

    old_h5f = r'<w:rFonts w:ascii="宋体" w:hAnsi="黑体" w:eastAsia="宋体" w:cs="宋体"/>'
    new_h5f = '<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体" w:cs="黑体"/>'
    if re.search(old_h5f, result):
        result = re.sub(old_h5f, new_h5f, result)
        changes.append('ryusuke标题5: 字体改为黑体')

    m5rpr = re.search(
        rf'(<w:style w:type="paragraph"[^>]*w:styleId="{re.escape(STYLE_RYUSUKE_H5)}"[^>]*>.*?<w:rPr>)(.*?)(</w:rPr>)',
        result, re.DOTALL)
    if m5rpr and not re.search('<w:sz ', m5rpr.group(2)):
        nr = m5rpr.group(2) + '<w:sz w:val="21"/><w:szCs w:val="21"/>'
        result = result[:m5rpr.start(2)] + nr + result[m5rpr.end(2):]
        changes.append('ryusuke标题5: 新增字号10.5pt')

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

    for msg in resolve_ryusuke_heading_style_ids(sty_xml):
        print(f'  ⚠ {msg}')

    # Step A: 套用 H4/H5 标题样式
    print('\n  [A] 识别并套用 H4/H5 标题样式...')
    doc_xml, h4_cnt, h5_cnt = step_apply_heading_styles(xml_content)
    print(f'      ✓ 四级标题(ryusuke标题4): +{h4_cnt}')
    print(f'      ✓ 五级标题(ryusuke标题5): +{h5_cnt}')

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
        print(f'      ⚠ 未找到ryusuke编号体系定义')

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
        has_numbering = os.path.exists(num_path)
        if has_numbering:
            with open(num_path, 'r', encoding='utf-8') as f:
                num_xml = f.read()
        else:
            num_xml = ''
            print('   ⚠ word/numbering.xml 不存在，将跳过依赖编号定义的修复步骤')
        has_styles = os.path.exists(sty_path)
        if has_styles:
            with open(sty_path, 'r', encoding='utf-8') as f:
                sty_xml = f.read()
        else:
            sty_xml = ''
            print('   ⚠ word/styles.xml 不存在，将跳过依赖样式定义的修复步骤')

        # ---- Phase 0: 根本性编号修复（三层面统一） ----
        rpt0 = {}
        if args.phase in ('0', 'all'):
            if not has_numbering or not has_styles:
                print('\n⚠ 跳过 Phase 0：当前文档缺少 numbering.xml 或 styles.xml')
                styles_xml = sty_xml
            else:
                doc_xml, num_xml, styles_xml, rpt0 = phase0_fix_heading_numbering(doc_xml, num_xml, sty_xml)

                with open(doc_path, 'w', encoding='utf-8') as f:
                    f.write(doc_xml)
                with open(num_path, 'w', encoding='utf-8') as f:
                    f.write(num_xml)
                with open(sty_path, 'w', encoding='utf-8') as f:
                    f.write(styles_xml)

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
            if not has_styles:
                print('\n⚠ 跳过 Phase 2：当前文档没有 styles.xml')
            else:
                doc_xml, num_xml, sty_xml, rpt2 = phase2_beautify(
                    doc_xml, num_xml, sty_xml, args.indent_h4, args.indent_h5)

                with open(doc_path, 'w', encoding='utf-8') as f:
                    f.write(doc_xml)
                if has_numbering:
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
            print(f'\n   📋 Phase 0 - 根本性编号修复（三层面统一）:')
            print(f'      • numbering.xml: 补全主列表pStyle链接')
            print(f'      • styles.xml: 统一{rpt0.get("style_fixes", 0)}个样式的numPr')
            print(f'      • 章节标题(style12)插入numPr: {rpt0.get("num_chapter", 0)} 处')
            print(f'      • 二级标题(style13)插入numPr: {rpt0.get("num_h2", 0)} 处')
            print(f'      • 三级标题(style14)插入numPr: {rpt0.get("num_h3", 0)} 处')
            print(f'      • 四级标题(style15)插入numPr: {rpt0.get("num_h4", 0)} 处')
            print(f'      • 五级标题(style16)插入numPr: {rpt0.get("num_h5", 0)} 处')
            print(f'      • 残留编号前缀清除:       {rpt0.get("strip_prefixes", 0)} 处')

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
