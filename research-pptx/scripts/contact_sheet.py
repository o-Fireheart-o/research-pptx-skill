# -*- coding: utf-8 -*-
"""Tile rendered slides (sNN.png from ppt_export.ps1) into contact sheets, so a whole deck can be looked over in a
few images instead of one image per slide. Looking at images is the expensive part of QA; use the sheets for the
first pass and open single slides at full size only where a sheet shows a problem.

  python -I contact_sheet.py <render_dir> <out_dir> [--cols 3] [--rows 3] [--width 1800]
      writes <out_dir>/sheet_01.png, sheet_02.png … (cols x rows slides each, slide number in the corner).
      Default 3 x 3 at 1800 px wide: text of a normal content slide stays readable; for very dense slides
      (tables with small type) use --cols 2 --rows 2.
"""
import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont


def font(size):
    for name in ("malgun.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            pass
    return ImageFont.load_default()


def main():
    a = sys.argv[1:]
    opt = {"--cols": 3, "--rows": 3, "--width": 1800}
    for k in list(opt):
        if k in a:
            i = a.index(k)
            opt[k] = int(a[i + 1])
            del a[i:i + 2]
    if len(a) != 2:
        sys.exit(__doc__)
    src, out = a
    files = sorted(f for f in os.listdir(src) if re.match(r"s\d+\.png$", f))
    if not files:
        sys.exit(f"no sNN.png in {src}")
    os.makedirs(out, exist_ok=True)
    cols, rows, width = opt["--cols"], opt["--rows"], opt["--width"]
    gap = 12
    cell_w = (width - gap * (cols + 1)) // cols
    first = Image.open(os.path.join(src, files[0]))
    cell_h = cell_w * first.height // first.width
    per = cols * rows
    label = font(22)
    for page in range(0, len(files), per):
        chunk = files[page:page + per]
        used_rows = (len(chunk) + cols - 1) // cols
        sheet = Image.new("RGB", (width, gap + used_rows * (cell_h + gap)), (90, 90, 90))
        draw = ImageDraw.Draw(sheet)
        for i, name in enumerate(chunk):
            x = gap + (i % cols) * (cell_w + gap)
            y = gap + (i // cols) * (cell_h + gap)
            img = Image.open(os.path.join(src, name)).convert("RGB").resize((cell_w, cell_h), Image.LANCZOS)
            sheet.paste(img, (x, y))
            no = str(int(re.search(r"\d+", name).group()))
            draw.rectangle([x, y, x + 20 + 13 * len(no), y + 30], fill=(200, 0, 0))
            draw.text((x + 8, y + 2), no, fill=(255, 255, 255), font=label)
        path = os.path.join(out, f"sheet_{page // per + 1:02d}.png")
        sheet.save(path)
        print(path)


if __name__ == "__main__":
    main()
