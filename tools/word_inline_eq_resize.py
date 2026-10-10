# -*- coding: utf-8 -*-
"""行内公式符号大小统一（用户 Word 标准 · 公式符号修复工具）
用法: PYTHONUTF8=1 python word_inline_eq_resize.py 输入.docx 输出.docx
（输出到新文件，不覆盖原文件；正式使用前请先备份并让用户确认）

原理：本机 MathType 对象内部字形默认 12pt，比注文(9pt)/正文(10.5pt)大 33%/14%，
      会顶端冒头、底部下沉（"偏上偏下"）。本工具按"所在段落字号 ÷ 12"缩放对象的
      显示尺寸（v:shape 宽高 + dxaOrig/dyaOrig 同步），并把 w:position 上抬
      6×(1−缩放比) 磅（换算成半磅取整），使符号与文字基线对齐。
      显示公式段（正文只有 (n) 的段落）不动。

缩放后必须：① 用 Word 导出 PDF；② 用 word_symbol_align_check.py 复核；
③ 个别对象内部基线特殊时（如 Cronbach's α、p(B_S|D)），对个别对象单独
   微调 position（经验：±2~5 半磅）后复测。
"""
import sys
import zipfile
import re


def resize(src, dst):
    z = zipfile.ZipFile(src)
    names = z.namelist()
    items = {n: z.read(n) for n in names}
    z.close()
    raw = items['word/document.xml'].decode('utf-8')

    paras = re.split(r'(?=<w:p[ >/])', raw)
    out = []
    stats = {'scaled': 0, 'skip_disp': 0}
    for chunk in paras:
        if '<w:object' not in chunk:
            out.append(chunk)
            continue
        # 段落文本；显示公式段（只有编号如 (3)）跳过
        txt = ''.join(re.findall(r'<w:t(?: [^>]*)?>([^<]*)</w:t>', chunk)).strip()
        if re.match(r'^[（(]\s*\d+\s*[）)]$', txt):
            stats['skip_disp'] += chunk.count('<w:object')
            out.append(chunk)
            continue
        m = re.search(r'<w:sz w:val="(\d+)"/>', chunk)
        sz = int(m.group(1)) if m else 21  # 半磅；缺省按五号21
        scale = (sz / 2.0) / 12.0
        if abs(scale - 1.0) < 1e-6:
            out.append(chunk)
            continue

        def fix_style(mo):
            style = mo.group(1)
            style = re.sub(r'((?:^|;))width:([\d.]+)pt',
                           lambda mm: '%swidth:%gpt' % (mm.group(1), float(mm.group(2)) * scale), style)
            style = re.sub(r'((?:^|;))height:([\d.]+)pt',
                           lambda mm: '%sheight:%gpt' % (mm.group(1), float(mm.group(2)) * scale), style)
            return 'style="%s"' % style

        chunk = re.sub(r'style="([^"]*width:[^"]*)"', fix_style, chunk)
        corr_hp = round(6.0 * (1.0 - scale) * 2)
        chunk = re.sub(r'<w:position w:val="(-?\d+)"',
                       lambda mm: '<w:position w:val="%d"' % (int(mm.group(1)) + corr_hp), chunk)
        chunk = re.sub(r'<w:object w:dxaOrig="(\d+)" w:dyaOrig="(\d+)">',
                       lambda mm: '<w:object w:dxaOrig="%d" w:dyaOrig="%d">'
                       % (round(int(mm.group(1)) * scale), round(int(mm.group(2)) * scale)), chunk)
        stats['scaled'] += chunk.count('<w:object')
        out.append(chunk)

    raw = ''.join(out)
    items['word/document.xml'] = raw.encode('utf-8')
    zo = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    for n in names:
        zo.writestr(n, items[n])
    zo.close()
    print('已缩放行内公式对象:', stats['scaled'], '个（跳过显示公式对象 %d 个）' % stats['skip_disp'])
    print('输出:', dst)
    print('提示：随后导出 PDF，并用 word_symbol_align_check.py 复核。')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('用法: python word_inline_eq_resize.py 输入.docx 输出.docx')
        sys.exit(2)
    resize(sys.argv[1], sys.argv[2])
