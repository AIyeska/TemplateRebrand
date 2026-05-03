"""
Generates SoftwareOne logo as PNG bytes using Pillow.
White variant for dark backgrounds, dark variant for light backgrounds.
"""
import io
import math
from PIL import Image, ImageDraw, ImageFont


def _get_font(size, bold=False):
    candidates = [
        "arialbd.ttf" if bold else "arial.ttf",
        "Arial Bold.ttf" if bold else "Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except (IOError, OSError):
            continue
    return ImageFont.load_default()


def make_logo_png(white_on_dark=True, width=360) -> bytes:
    """
    Returns PNG bytes of a SoftwareOne-style logo.
    white_on_dark=True  → white logo, transparent background (for dark Office slides/headers)
    white_on_dark=False → dark-blue logo, transparent background (for light Word docs)
    """
    height = int(width * 0.28)
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    fg = (255, 255, 255, 255) if white_on_dark else (0, 48, 135, 255)

    # --- Circle badge ---
    r = int(height * 0.42)
    cx = r + int(height * 0.06)
    cy = height // 2
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fg)

    # "S" inside circle
    s_color = (0, 48, 135, 255) if white_on_dark else (255, 255, 255, 255)
    s_font = _get_font(int(r * 1.2), bold=True)
    draw.text((cx, cy), "S", font=s_font, fill=s_color, anchor="mm")

    # --- Wordmark: "software" small + "one" large ---
    text_x = cx + r + int(width * 0.04)

    small_font = _get_font(int(height * 0.22))
    big_font   = _get_font(int(height * 0.52), bold=True)

    small_h = int(height * 0.22)
    big_h   = int(height * 0.52)
    total_h = small_h + big_h - int(height * 0.06)
    top_y   = (height - total_h) // 2

    draw.text((text_x, top_y), "software", font=small_font, fill=fg)
    draw.text((text_x, top_y + small_h - int(height * 0.06)), "one", font=big_font, fill=fg)

    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


# Cached instances
_logo_white: bytes | None = None
_logo_dark:  bytes | None = None


def get_logo_white(width=360) -> bytes:
    global _logo_white
    if _logo_white is None:
        _logo_white = make_logo_png(white_on_dark=True, width=width)
    return _logo_white


def get_logo_dark(width=360) -> bytes:
    global _logo_dark
    if _logo_dark is None:
        _logo_dark = make_logo_png(white_on_dark=False, width=width)
    return _logo_dark
