#!/usr/bin/env python3
from __future__ import annotations

import os
import struct
import zlib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
PNG_PATH = ASSETS / "windowlayouter-native-icon.png"
ICO_PATH = ASSETS / "windowlayouter-native-icon.ico"
PNG_EXPORT_DIR = ASSETS / "png"


def clamp(value: int) -> int:
    return max(0, min(255, value))


def blend(src: tuple[int, int, int, int], dst: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    sr, sg, sb, sa = src
    dr, dg, db, da = dst
    sa_f = sa / 255.0
    da_f = da / 255.0
    out_a = sa_f + da_f * (1.0 - sa_f)
    if out_a <= 0.0:
        return 0, 0, 0, 0
    out_r = int(round((sr * sa_f + dr * da_f * (1.0 - sa_f)) / out_a))
    out_g = int(round((sg * sa_f + dg * da_f * (1.0 - sa_f)) / out_a))
    out_b = int(round((sb * sa_f + db * da_f * (1.0 - sa_f)) / out_a))
    out_alpha = int(round(out_a * 255.0))
    return clamp(out_r), clamp(out_g), clamp(out_b), clamp(out_alpha)


def rgba_to_png_bytes(width: int, height: int, pixels: bytes) -> bytes:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    raw = bytearray()
    stride = width * 4
    for y in range(height):
        raw.append(0)
        start = y * stride
        raw.extend(pixels[start:start + stride])

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
        + chunk(b"IEND", b"")
    )


def rgba_pixels_to_ico_bmp(width: int, pixels: list[tuple[int, int, int, int]]) -> bytes:
    """Encode RGBA pixels as ICO-BMP (BITMAPINFOHEADER + BGRA bottom-up + AND mask)."""
    header = struct.pack("<IiiHHIIiiII",
        40,           # biSize
        width,        # biWidth
        width * 2,    # biHeight (doubled: XOR + AND)
        1,            # biPlanes
        32,           # biBitCount
        0,            # biCompression = BI_RGB
        0,            # biSizeImage
        0, 0,         # biXPelsPerMeter, biYPelsPerMeter
        0, 0,         # biClrUsed, biClrImportant
    )
    pixel_data = bytearray()
    for y in range(width - 1, -1, -1):
        for x in range(width):
            r, g, b, a = pixels[y * width + x]
            pixel_data.extend((b, g, r, a))
    and_row_bytes = (width + 31) // 32 * 4
    and_mask = bytes(and_row_bytes * width)
    return header + bytes(pixel_data) + and_mask


def encode_ico(entries: list[tuple[int, int, bytes]]) -> bytes:
    header = struct.pack("<HHH", 0, 1, len(entries))
    directory = bytearray()
    payload = bytearray()
    offset = 6 + len(entries) * 16

    for width, height, image_data in entries:
        directory.extend(
            struct.pack(
                "<BBBBHHII",
                0 if width == 256 else width,
                0 if height == 256 else height,
                0,
                0,
                1,
                32,
                len(image_data),
                offset,
            )
        )
        payload.extend(image_data)
        offset += len(image_data)

    return bytes(header + directory + payload)


def rounded_rect_mask(size: int, left: int, top: int, right: int, bottom: int, radius: int) -> list[int]:
    mask = [0] * (size * size)
    radius_sq = radius * radius
    for y in range(top, bottom):
        for x in range(left, right):
            dx = 0
            dy = 0
            if x < left + radius:
                dx = left + radius - x - 1
            elif x >= right - radius:
                dx = x - (right - radius)
            if y < top + radius:
                dy = top + radius - y - 1
            elif y >= bottom - radius:
                dy = y - (bottom - radius)

            if dx * dx + dy * dy <= radius_sq:
                mask[y * size + x] = 1
    return mask


def draw_rect(canvas: list[tuple[int, int, int, int]], size: int, left: int, top: int, right: int, bottom: int, color: tuple[int, int, int, int], radius: int = 0) -> None:
    if radius <= 0:
        for y in range(top, bottom):
            for x in range(left, right):
                canvas[y * size + x] = blend(color, canvas[y * size + x])
        return

    mask = rounded_rect_mask(size, left, top, right, bottom, radius)
    for index, active in enumerate(mask):
        if active:
            canvas[index] = blend(color, canvas[index])


def draw_shadow(canvas: list[tuple[int, int, int, int]], size: int, left: int, top: int, right: int, bottom: int, radius: int, spread: int, alpha: int) -> None:
    for y in range(max(0, top - spread), min(size, bottom + spread)):
        for x in range(max(0, left - spread), min(size, right + spread)):
            nearest_x = min(max(x, left), right - 1)
            nearest_y = min(max(y, top), bottom - 1)
            dx = x - nearest_x
            dy = y - nearest_y
            dist = (dx * dx + dy * dy) ** 0.5
            if dist > spread:
                continue
            local_alpha = int(alpha * (1.0 - dist / max(1, spread)))
            if local_alpha <= 0:
                continue
            canvas[y * size + x] = blend((0, 0, 0, local_alpha), canvas[y * size + x])


def decode_raw_pixels(png_data: bytes) -> tuple[int, list[tuple[int, int, int, int]]]:
    """Extract raw RGBA pixels from our simple PNG format (single IDAT, no filtering beyond row-filter 0)."""
    import io
    f = io.BytesIO(png_data)
    f.read(8)  # PNG signature
    chunks: dict[bytes, bytes] = {}
    while True:
        header = f.read(8)
        if len(header) < 8:
            break
        length = struct.unpack(">I", header[:4])[0]
        tag = header[4:8]
        data = f.read(length)
        f.read(4)  # CRC
        chunks[tag] = data
    width = struct.unpack(">I", chunks[b"IHDR"][:4])[0]
    raw = zlib.decompress(chunks[b"IDAT"])
    pixels: list[tuple[int, int, int, int]] = []
    stride = width * 4 + 1
    for y in range(width):
        row_start = y * stride + 1  # skip filter byte
        for x in range(width):
            offset = row_start + x * 4
            r, g, b, a = raw[offset], raw[offset + 1], raw[offset + 2], raw[offset + 3]
            pixels.append((r, g, b, a))
    return width, pixels


def downsample(large_size: int, pixels: list[tuple[int, int, int, int]], target_size: int) -> list[tuple[int, int, int, int]]:
    """Box-filter downsample from large_size to target_size."""
    scale = large_size // target_size
    result: list[tuple[int, int, int, int]] = []
    for ty in range(target_size):
        for tx in range(target_size):
            r_sum = g_sum = b_sum = a_sum = 0
            for dy in range(scale):
                for dx in range(scale):
                    sy = ty * scale + dy
                    sx = tx * scale + dx
                    pr, pg, pb, pa = pixels[sy * large_size + sx]
                    r_sum += pr
                    g_sum += pg
                    b_sum += pb
                    a_sum += pa
            count = scale * scale
            result.append((r_sum // count, g_sum // count, b_sum // count, a_sum // count))
    return result


def make_icon(size: int) -> bytes:
    """Generate icon PNG. Sizes 33-256 use 4x supersampling for smooth edges."""
    if size <= 32 or size > 256:
        if size <= 24:
            return make_small_icon(size)
        return make_large_icon(size)

    ss_factor = 4
    ss_size = size * ss_factor
    raw_png = make_large_icon(ss_size)
    _, large_pixels = decode_raw_pixels(raw_png)
    small_pixels = downsample(ss_size, large_pixels, size)
    pixel_bytes = bytearray()
    for r, g, b, a in small_pixels:
        pixel_bytes.extend((r, g, b, a))
    return rgba_to_png_bytes(size, size, bytes(pixel_bytes))


def make_small_icon(size: int) -> bytes:
    bg = (0, 0, 0, 0)
    canvas = [bg for _ in range(size * size)]

    graphite = (33, 40, 48, 255)
    line = (233, 237, 243, 255)
    amber = (244, 175, 68, 255)
    blue = (90, 117, 194, 255)
    slate = (74, 84, 97, 255)
    pane_dark = (55, 64, 75, 255)

    frame_left = max(1, size // 8)
    frame_top = max(1, size // 8)
    frame_right = size - frame_left
    frame_bottom = size - frame_top
    frame_radius = max(3, size // 5)

    draw_shadow(canvas, size, frame_left, frame_top, frame_right, frame_bottom, frame_radius, max(1, size // 10), 64)
    draw_rect(canvas, size, frame_left, frame_top, frame_right, frame_bottom, graphite, frame_radius)

    inset = max(1, size // 10)
    inner_left = frame_left + inset
    inner_top = frame_top + inset
    inner_right = frame_right - inset
    inner_bottom = frame_bottom - inset
    inner_width = inner_right - inner_left
    inner_height = inner_bottom - inner_top
    split_x = inner_left + inner_width // 2
    split_y = inner_top + inner_height // 2
    gap = max(1, size // 16)
    thickness = max(1, size // 12)
    pane_radius = max(2, size // 8)

    panes = [
        (inner_left, inner_top, split_x - gap, split_y - gap, amber),
        (split_x + gap, inner_top, inner_right, split_y - gap, blue),
        (inner_left, split_y + gap, split_x - gap, inner_bottom, pane_dark),
        (split_x + gap, split_y + gap, inner_right, inner_bottom, slate),
    ]
    for left, top, right, bottom, color in panes:
        draw_rect(canvas, size, left, top, right, bottom, color, pane_radius)

    draw_rect(canvas, size, split_x - thickness // 2, inner_top, split_x + (thickness + 1) // 2, inner_bottom, line)
    draw_rect(canvas, size, inner_left, split_y - thickness // 2, inner_right, split_y + (thickness + 1) // 2, line)

    pixels = bytearray()
    for r, g, b, a in canvas:
        pixels.extend((r, g, b, a))
    return rgba_to_png_bytes(size, size, bytes(pixels))


def make_large_icon(size: int) -> bytes:
    bg = (0, 0, 0, 0)
    canvas = [bg for _ in range(size * size)]

    shell_left = int(size * 0.10)
    shell_top = int(size * 0.10)
    shell_right = int(size * 0.90)
    shell_bottom = int(size * 0.90)
    shell_radius = max(3, size // 8)

    screen_left = int(size * 0.18)
    screen_top = int(size * 0.17)
    screen_right = int(size * 0.82)
    screen_bottom = int(size * 0.73)
    screen_radius = max(3, size // 14)

    stand_width = int(size * 0.18)
    stand_height = max(2, int(size * 0.06))
    stand_left = (size - stand_width) // 2
    stand_top = int(size * 0.78)
    stand_right = stand_left + stand_width
    stand_bottom = stand_top + stand_height

    base_width = int(size * 0.34)
    base_height = max(2, int(size * 0.05))
    base_left = (size - base_width) // 2
    base_top = int(size * 0.86)
    base_right = base_left + base_width
    base_bottom = base_top + base_height

    graphite = (34, 40, 49, 255)
    graphite_high = (55, 63, 74, 255)
    slate = (87, 98, 113, 255)
    line = (207, 214, 224, 255)
    amber = (242, 172, 66, 255)
    amber_soft = (255, 199, 109, 255)
    pane = (72, 82, 95, 255)
    pane_dark = (58, 67, 79, 255)

    draw_shadow(canvas, size, shell_left, shell_top, shell_right, shell_bottom, shell_radius, max(2, size // 18), 80)
    draw_rect(canvas, size, shell_left, shell_top, shell_right, shell_bottom, graphite, shell_radius)
    draw_rect(canvas, size, screen_left, screen_top, screen_right, screen_bottom, graphite_high, screen_radius)

    inset = max(2, size // 36)
    pane_gap = max(2, size // 32)
    inner_left = screen_left + inset
    inner_top = screen_top + inset
    inner_right = screen_right - inset
    inner_bottom = screen_bottom - inset
    inner_width = inner_right - inner_left
    inner_height = inner_bottom - inner_top
    split_x = inner_left + inner_width // 2
    split_y = inner_top + inner_height // 2

    panes = [
        (inner_left, inner_top, split_x - pane_gap // 2, split_y - pane_gap // 2, amber),
        (split_x + pane_gap // 2, inner_top, inner_right, split_y - pane_gap // 2, pane),
        (inner_left, split_y + pane_gap // 2, split_x - pane_gap // 2, inner_bottom, pane_dark),
        (split_x + pane_gap // 2, split_y + pane_gap // 2, inner_right, inner_bottom, slate),
    ]
    pane_radius = max(2, size // 22)
    for left, top, right, bottom, color in panes:
        draw_rect(canvas, size, left, top, right, bottom, color, pane_radius)

    line_thickness = max(1, size // 48)
    draw_rect(canvas, size, split_x - line_thickness, inner_top, split_x + line_thickness, inner_bottom, line)
    draw_rect(canvas, size, inner_left, split_y - line_thickness, inner_right, split_y + line_thickness, line)

    glow = max(1, size // 40)
    draw_rect(canvas, size, inner_left + glow, inner_top + glow, split_x - pane_gap // 2 - glow, split_y - pane_gap // 2 - glow, amber_soft, max(1, pane_radius - 1))

    draw_rect(canvas, size, stand_left, stand_top, stand_right, stand_bottom, graphite_high, max(1, stand_height // 2))
    draw_rect(canvas, size, base_left, base_top, base_right, base_bottom, slate, max(1, base_height // 2))

    pixels = bytearray()
    for r, g, b, a in canvas:
        pixels.extend((r, g, b, a))
    return rgba_to_png_bytes(size, size, bytes(pixels))


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    PNG_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256]
    png_256 = make_icon(256)
    PNG_PATH.write_bytes(png_256)

    entries: list[tuple[int, int, bytes]] = []
    for size in sizes:
        png_data = make_icon(size)
        if size <= 128:
            _, pixels = decode_raw_pixels(png_data)
            bmp_data = rgba_pixels_to_ico_bmp(size, pixels)
            entries.append((size, size, bmp_data))
        else:
            entries.append((size, size, png_data))
    ICO_PATH.write_bytes(encode_ico(entries))

    export_sizes = [16, 32, 48, 64, 128, 256, 512, 1024]
    for size in export_sizes:
        export_path = PNG_EXPORT_DIR / f"windowlayouter-native-icon-{size}.png"
        export_path.write_bytes(make_icon(size))

    print(PNG_PATH)
    print(ICO_PATH)


if __name__ == "__main__":
    main()
