"""Render the add-in icons (a check mark in a grid cell) as PNGs with no dependencies.

Run once: python3 scripts/make-icons.py  -> public/assets/icon-{16,32,64,80,128}.png
"""
import math
import struct
import zlib
from pathlib import Path

GREEN = (16, 124, 65)        # Excel brand green #107C41
GREEN_DARK = (11, 92, 48)
WHITE = (255, 255, 255)
OUT = Path(__file__).resolve().parent.parent / "public" / "assets"
SS = 4  # supersampling factor


def seg_dist(px, py, ax, ay, bx, by):
    vx, vy = bx - ax, by - ay
    t = max(0.0, min(1.0, ((px - ax) * vx + (py - ay) * vy) / (vx * vx + vy * vy)))
    cx, cy = ax + t * vx, ay + t * vy
    return math.hypot(px - cx, py - cy)


def in_round_rect(x, y, size, r):
    cx = min(max(x, r), size - r)
    cy = min(max(y, r), size - r)
    return math.hypot(x - cx, y - cy) <= r


def render(size):
    big = size * SS
    r = big * 0.22
    # check mark polyline in unit coordinates
    pts = [(0.27, 0.53), (0.43, 0.69), (0.74, 0.35)]
    pts = [(x * big, y * big) for x, y in pts]
    stroke = big * (0.115 if size <= 16 else 0.095)
    grid_w = max(1.0, big * 0.018)
    pixels = []
    for py in range(size):
        row = []
        for px in range(size):
            acc = [0.0, 0.0, 0.0, 0.0]
            for sy in range(SS):
                for sx in range(SS):
                    x = px * SS + sx + 0.5
                    y = py * SS + sy + 0.5
                    if not in_round_rect(x, y, big, r):
                        continue
                    color = GREEN
                    # faint grid lines (skip on the smallest icon)
                    if size > 16:
                        for g in (big / 3, 2 * big / 3):
                            if abs(x - g) < grid_w or abs(y - g) < grid_w:
                                color = GREEN_DARK
                    d = min(seg_dist(x, y, *pts[0], *pts[1]), seg_dist(x, y, *pts[1], *pts[2]))
                    if d <= stroke / 2:
                        color = WHITE
                    acc[0] += color[0]
                    acc[1] += color[1]
                    acc[2] += color[2]
                    acc[3] += 255
            n = SS * SS
            a = acc[3] / n
            if a > 0:
                row.append((round(acc[0] / (acc[3] / 255)), round(acc[1] / (acc[3] / 255)),
                            round(acc[2] / (acc[3] / 255)), round(a)))
            else:
                row.append((0, 0, 0, 0))
        pixels.append(row)
    return pixels


def write_png(path, pixels):
    h = len(pixels)
    w = len(pixels[0])
    raw = b"".join(b"\x00" + b"".join(struct.pack("BBBB", *p) for p in row) for row in pixels)

    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw, 9))
    png += chunk(b"IEND", b"")
    path.write_bytes(png)


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    for s in (16, 32, 64, 80, 128):
        write_png(OUT / f"icon-{s}.png", render(s))
        print("wrote", OUT / f"icon-{s}.png")
