#!/usr/bin/env python3
"""把 generate_c.py 生成的大图按灰色分隔线切成一组小图。

对每个 (prefix, png) 跑一遍，输出 prefix_1.png / prefix_2.png / ...。

只依赖 Pillow（避免 PEP 668 拦下 numpy 安装）。
"""
from PIL import Image

# (basename without .png, full png path)
TARGETS = [
    ('c_gzh',  'c_gzh.png'),
    ('c_zpxq', 'c_zpxq.png'),
]


def crop_one(src, prefix):
    img = Image.open(src).convert('RGB')
    w, h = img.size
    pixels = img.load()

    # 找全宽 70%+ 都是 (190~210) 灰的行 → 横向分隔线
    sep_rows = []
    for y in range(h):
        gray_count = 0
        for x in range(w):
            r, g, b = pixels[x, y]
            if 190 <= r <= 210 and 190 <= g <= 210 and 190 <= b <= 210:
                gray_count += 1
                # 早退：超过 70% 已经确定是分隔行
                if gray_count * 10 > w * 7:
                    sep_rows.append(y)
                    break

    # 合并相邻行
    merged = []
    for r in sep_rows:
        if not merged or r - merged[-1] > 3:
            merged.append(r)
        else:
            merged[-1] = r

    boundaries = [0] + merged + [h]
    print(f'  {src}: {len(merged)} separators → {len(boundaries) - 1} slices')

    for i in range(len(boundaries) - 1):
        y0, y1 = boundaries[i], boundaries[i + 1]
        out = f'{prefix}_{i + 1}.png'
        img.crop((0, y0, w, y1)).save(out, 'PNG')
        print(f'    → {out}  y={y0}-{y1}  h={y1 - y0}')


def main():
    for prefix, src in TARGETS:
        crop_one(src, prefix)


if __name__ == '__main__':
    main()
