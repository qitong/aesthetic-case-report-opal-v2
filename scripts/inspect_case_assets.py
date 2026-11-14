#!/usr/bin/env python3
"""
Fast Case Asset Inspector & Contact Sheet Generator (Opal v2).

Scans a directory of clinical or anatomical images, prints an instant dimensions/aspect
table, and renders a single numbered Contact Sheet (`_contact_sheet.jpg`).
Instead of calling `view_file` 10+ times on massive DSLR images, the agent views ONE
lightweight contact sheet to classify Frontal / 45-Oblique / 90-Profile / Before / After.

Usage:
  python3 scripts/inspect_case_assets.py ./photos --out ./_contact_sheet.jpg
"""
from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageOps

EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}


def main() -> int:
    ap = argparse.ArgumentParser(description="Inspect clinical photo folder and build a 1-page contact sheet.")
    ap.add_argument("folder", type=Path, help="directory containing clinical photos")
    ap.add_argument("--out", type=Path, default=Path("_contact_sheet.jpg"), help="output contact sheet path")
    ap.add_argument("--cols", type=int, default=4, help="columns in contact sheet (default: 4)")
    args = ap.parse_args()

    if not args.folder.is_dir():
        print(f"[Error] Directory not found: {args.folder}", file=sys.stderr)
        return 1

    files = sorted(
        f for f in args.folder.iterdir()
        if f.is_file() and f.suffix.lower() in EXTS and not f.name.startswith("_")
    )
    if not files:
        print(f"[Warn] No image files found in {args.folder}")
        return 0

    cols = max(1, args.cols)
    rows = math.ceil(len(files) / cols)
    cell_w, cell_h = 340, 400
    sheet = Image.new("RGB", (cols * cell_w, rows * cell_h), (28, 26, 25))
    draw = ImageDraw.Draw(sheet)

    print(f"Found {len(files)} image(s) in {args.folder}:")
    for idx, f in enumerate(files):
        r, c = divmod(idx, cols)
        x0, y0 = c * cell_w, r * cell_h
        with Image.open(f) as raw:
            im = ImageOps.exif_transpose(raw).convert("RGB")
            orig_w, orig_h = im.size
            im.thumbnail((cell_w - 20, cell_h - 56), Image.Resampling.LANCZOS)
            px = x0 + (cell_w - im.width) // 2
            py = y0 + 10 + (cell_h - 56 - im.height) // 2
            sheet.paste(im, (px, py))

        label = f"[{idx+1:02d}] {f.name[:28]} ({orig_w}x{orig_h})"
        draw.rectangle([(x0 + 6, y0 + cell_h - 40), (x0 + cell_w - 6, y0 + cell_h - 6)], fill=(18, 16, 15))
        draw.text((x0 + 12, y0 + cell_h - 30), label, fill=(245, 238, 228))
        print(f"  [{idx+1:02d}] {f.name:36s}  {orig_w}x{orig_h}  (ratio={orig_w/orig_h:.2f})")

    args.out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(args.out, "JPEG", quality=88)
    print(f"[OK] Contact sheet saved -> {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
