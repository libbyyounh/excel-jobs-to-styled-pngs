#!/usr/bin/env python3
"""Stage 2 — Cut a long rendered PNG into per-card slices along gray separator lines.

Pure Pillow, no numpy (PEP 668 blocks pip install on managed Python).

Usage:
    python3 crop.py                          # auto-discover c_*.png in CWD
    python3 crop.py                          # set TARGETS below to override
"""
import glob
import re
from PIL import Image

# Manual override. Leave as [] to auto-discover every c_*.png in the current directory.
# Format: (prefix_without_underscore_number, full_png_path)
TARGETS = []

# Pattern for an already-sliced file (e.g. c_0714_3.png) — skip these on auto-discovery.
SLICED_PATTERN = re.compile(r'_\d+\.png$')


def auto_discover():
    """Yield (prefix, png_path) for every c_*.png in CWD that isn't already a slice."""
    for path in sorted(glob.glob('c_*.png')):
        if SLICED_PATTERN.search(path):
            continue
        prefix = path[:-4]  # strip ".png"
        yield (prefix, path)


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
    if TARGETS:
        work = list(TARGETS)
    else:
        work = list(auto_discover())
        if not work:
            print('  no c_*.png found in CWD; set TARGETS to override')
            return
    for prefix, src in work:
        crop_one(src, prefix)


if __name__ == '__main__':
    main()
