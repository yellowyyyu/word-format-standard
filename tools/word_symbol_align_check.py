# -*- coding: utf-8 -*-
"""行内公式符号对齐检查（用户 Word 标准 · 第 2 步公式专项测量工具）
用法: PYTHONUTF8=1 python word_symbol_align_check.py 文件.pdf

检查内容：每一行文字里，"公式符号字形"（非宋体的字形）相对该行文字基线的
上下偏差 Δy（磅）。主字形（大小≥8磅）应全部落在 ±0.5 磅以内；
下标字形（<8磅）天然低于基线，单列统计不判不合格。

通过标准：主字形 |Δy| ≤ 0.5 磅。
退出码：0=通过，1=有超差，2=用法错误。
"""
import sys
import fitz  # PyMuPDF
from collections import Counter


def is_text_font(f):
    """判断字形是否属于"正文中文"字体（宋体/SimSun，含编码兜底）。"""
    if ('宋体' in f) or ('SimSun' in f):
        return True
    try:
        f2 = f.encode('latin-1', 'ignore').decode('utf-8', 'ignore')
        return ('宋体' in f2) or ('SimSun' in f2)
    except Exception:
        return False


def check(pdf_path):
    doc = fitz.open(pdf_path)
    deltas = []
    for pno in range(len(doc)):
        d = doc[pno].get_text('dict')
        for b in d['blocks']:
            if b.get('type') != 0:
                continue
            for l in b['lines']:
                spans = [s for s in l['spans'] if s['text'].strip()]
                if not spans:
                    continue
                bases = sorted(s['origin'][1] for s in spans if is_text_font(s['font']))
                if not bases:
                    continue  # 没有中文的行（多为显示公式），跳过
                base = bases[len(bases) // 2]
                line_txt = ''.join(s['text'] for s in spans).strip()
                for s in spans:
                    if is_text_font(s['font']):
                        continue
                    if s['size'] > 13.5:
                        continue  # 显示公式的大字形
                    dy = s['origin'][1] - base
                    if abs(dy) > 6:
                        continue  # 显示公式残片噪声
                    cls = 'SUB' if s['size'] < 8.0 else 'MAIN'
                    deltas.append((cls, round(dy, 2), s['size'], s['text'][:10], line_txt[:26], pno + 1))
    doc.close()

    main = [x for x in deltas if x[0] == 'MAIN']
    sub = [x for x in deltas if x[0] == 'SUB']
    print('=== 行内公式符号对齐检查 ===')
    print('文件:', pdf_path)
    print('主字形数:', len(main), '  下标字形数:', len(sub))
    hc = Counter(round(x[1] * 4) / 4 for x in main)
    print('主字形 Δy 分布（0.25 磅一组；负=偏高，正=偏低）:')
    for k in sorted(hc):
        print('   Δy=%+6.2f  n=%d' % (k, hc[k]))

    bad = [x for x in main if abs(x[1]) > 0.5]
    if bad:
        print()
        print('!! 偏差超过 0.5 磅的有 %d 处：' % len(bad))
        for x in bad[:40]:
            print('   第%d页 Δy=%+.2f size=%.1f %r …行:%r' % (x[5], x[1], x[2], x[3], x[4]))
        print('[未通过] 请修复后重新检查。')
        return 1
    print()
    print('[全部通过] 行内公式符号位置检查通过（全部偏差 ≤ 0.5 磅）。')
    return 0


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python word_symbol_align_check.py 文件.pdf')
        sys.exit(2)
    sys.exit(check(sys.argv[1]))
