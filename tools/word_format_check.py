# -*- coding: utf-8 -*-
"""Word 格式自动检查（用户规矩的验收脚本）—— 每次文档改完必跑
规矩：中文=宋体、英文=Times New Roman；文章题目=二号(44)；正文=五号(21)+首行缩进2字符(420缇或firstLineChars=200)+两端对齐；
标题=五号+加粗+顶格；图表内容与题注=小五(18)+居中+无缩进；图表下的注/Note行=小五+无缩进+左对齐；
全文=单倍行距(line240, lineRule=auto)；公式=居中+编号手打顶右。

用法：python word_format_check.py "文件.docx"    全绿=通过
"""
import sys, zipfile, re
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f'{{{W}}}{t}'
ALLOWED_FONTS = {'Times New Roman', '宋体'}
ALLOWED_SIZES = {'44', '21', '18'}

def check(path):
    issues = []
    z = zipfile.ZipFile(path)
    raw = z.read('word/document.xml').decode('utf-8')
    root = etree.fromstring(z.read('word/document.xml'))
    tree = root.getroottree()

    # 1) 字体
    for m in re.finditer(r'<w:rFonts ([^/>]+)/?>', raw):
        for am in re.finditer(r'w:(\w+)="([^"]+)"', m.group(1)):
            k, v = am.group(1), am.group(2)
            if k == 'hint' or v in ALLOWED_FONTS:
                continue
            issues.append(f'[字体] {k}={v}')

    # 2) 直接字号
    for m in re.finditer(r'<w:sz w:val="(\d+)"/>', raw):
        if m.group(1) not in ALLOWED_SIZES:
            ctx = re.sub(r'<[^>]+>', '', raw[max(0, m.start()-200):m.start()])[-18:]
            issues.append(f'[字号] {m.group(1)} 上文=…{ctx}')

    # 3) 行距
    for m in re.finditer(r'<w:spacing[^>]*w:line="(?!240)\d+"[^>]*/>', raw):
        issues.append('[行距] 非240: ' + m.group(0)[:70])
    for m in re.finditer(r'<w:spacing[^>]*w:lineRule="(?!auto)(\w+)"[^>]*/>', raw):
        issues.append('[行距] 非auto: ' + m.group(0)[:70])

    # 4) 逐段检查（用路径做表格归属判断，稳定可靠）
    tbl_paths = set()
    for tbl in root.iter(q('tbl')):
        for p in tbl.iter(q('p')):
            tbl_paths.add(tree.getpath(p))

    first_seen = False
    prev_note = False
    for p in root.iter(q('p')):
        t = ''.join(x.text or '' for x in p.iter(q('t'))).strip()
        if not t:
            continue
        this_note = t.startswith('注：') or t.startswith('Note:')
        is_note_cont = prev_note and t.startswith('其中')
        prev_note = this_note
        ppr = p.find(q('pPr'))
        st = ind = jc = None
        tabs_center = False
        if ppr is not None:
            e = ppr.find(q('pStyle')); st = e.get(q('val')) if e is not None else None
            e = ppr.find(q('ind'))
            if e is not None:
                ind = {k.split('}')[-1]: v for k, v in e.attrib.items()}
            e = ppr.find(q('jc')); jc = e.get(q('val')) if e is not None else None
            tabs_center = '4153' in etree.tostring(ppr).decode()
        in_tbl = tree.getpath(p) in tbl_paths
        no_ind = bool(ind) and ind.get('firstLineChars') == '0' and ind.get('firstLine') == '0'
        has_body_ind = (bool(ind) and (ind.get('firstLineChars') == '200' or ind.get('firstLine') == '420')) or ind is None or ind == {}
        if not first_seen:
            first_seen = True
            continue  # 第一段=文章题目，主题规则单独管
        is_cap = (t.startswith('Figure ') or t.startswith('Table ')) and not tabs_center
        is_note = this_note or is_note_cont
        is_eq = tabs_center
        is_short_head = len(t) < 40 and bool(re.match(r'^\d+(\.\d+)*[\s　]?[^\d]', t))
        if is_eq or is_short_head:
            continue
        if is_cap:
            if st not in ('aa', 'a3'):
                issues.append(f'[题注] 未用题注样式: {t[:28]!r}')
            if not no_ind and ind:
                issues.append(f'[题注] 有缩进: {t[:28]!r} {ind}')
            if jc not in ('center', None):
                issues.append(f'[题注] 未居中: {t[:28]!r}')
        elif is_note:
            if not no_ind:
                issues.append(f'[注行] 有缩进: {t[:28]!r} {ind}')
            if jc not in ('left', None):
                issues.append(f'[注行] 未左对齐: {t[:28]!r} jc={jc}')
        elif in_tbl:
            for r in p.findall(q('r')):
                e = r.find(q('rPr'))
                if e is not None:
                    s = e.find(q('sz'))
                    if s is not None and s.get(q('val')) != '18':
                        issues.append(f'[表格] 非小五: {s.get(q("val"))}  {t[:18]!r}')
            if jc != 'center':
                issues.append(f'[表格] 未居中: {t[:18]!r}')
        else:  # 正文短语/段
            if len(t) > 40 and no_ind:
                issues.append(f'[正文] 缺首行缩进: {t[:28]!r}')

    # 5) 页眉页脚
    for n in z.namelist():
        if re.search(r'word/(header|footer)\d*\.xml', n):
            x = z.read(n).decode('utf-8')
            if re.sub(r'<[^>]+>', '', x).strip():
                issues.append(f'[页眉页脚] {n} 有内容，需人工核对')

    # 6) 行内公式对象：显示高度必须已按字号缩放（未缩放的旧对象高度=18pt 会顶出格）
    for chunk in re.split(r'(?=<w:p[ >/])', raw):
        if '<w:object' not in chunk:
            continue
        txt = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', chunk)).strip()
        if re.match(r'^[（(]\s*\d+\s*[）)]$', txt):
            continue  # 显示公式行（公式本体允许大）
        for sm in re.finditer(r'<v:shape[^>]*style="[^"]*height:([\d.]+)pt', chunk):
            h = float(sm.group(1))
            if h >= 17.9:
                issues.append(f'[公式对象] 显示高度 {h}pt 偏大（疑似未按字号缩放）')
    z.close()
    return issues

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python word_format_check.py 文件.docx'); sys.exit(2)
    probs = check(sys.argv[1])
    if probs:
        print(f'发现 {len(probs)} 个问题：')
        for x in probs[:80]:
            print('  -', x)
        sys.exit(1)
    print('[全部通过] 文档符合既定格式规矩。')
