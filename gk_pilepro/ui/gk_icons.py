# -*- coding: utf-8 -*-

import ctypes

from PIL import Image

try:
    from ctypes import wintypes
except Exception:
    wintypes = None


def rounded_icon_image(image, size=(128, 128), radius_ratio=0.18):

    image = image.convert("RGBA").resize(size, Image.Resampling.LANCZOS)

    mask = Image.new("L", size, 0)

    from PIL import ImageDraw

    radius = max(1, int(min(size) * radius_ratio))

    ImageDraw.Draw(mask).rounded_rectangle(

        (0, 0, size[0] - 1, size[1] - 1),

        radius=radius,

        fill=255,

    )

    alpha = image.getchannel("A")

    image.putalpha(Image.composite(alpha, Image.new("L", size, 0), mask))

    return image



def sharp_icon_image(image, size):

    """

    Resize an icon for Windows/taskbar use without soft masking.

    """

    img = image.convert("RGBA").resize(size, Image.Resampling.LANCZOS)

    try:

        from PIL import ImageEnhance

        img = ImageEnhance.Sharpness(img).enhance(1.35)

        img = ImageEnhance.Contrast(img).enhance(1.06)

    except Exception:

        pass

    return img



def sharp_icon_image_small(image, size):

    """

    Resize an icon aggressively for very small taskbar/window sizes.

    """

    img = image.convert("RGBA").resize(size, Image.Resampling.LANCZOS)

    try:

        from PIL import ImageEnhance, ImageFilter

        img = ImageEnhance.Sharpness(img).enhance(1.9)

        img = ImageEnhance.Contrast(img).enhance(1.14)

        img = img.filter(ImageFilter.UnsharpMask(radius=0.35, percent=245, threshold=0))

    except Exception:

        pass

    return img


def get_windows_work_area():
    try:
        if not hasattr(ctypes, "windll") or wintypes is None:
            return None
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        class RECT(ctypes.Structure):
            _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

        class POINT(ctypes.Structure):
            _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

        class MONITORINFO(ctypes.Structure):
            _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", RECT), ("rcWork", RECT), ("dwFlags", wintypes.DWORD)]

        pt = POINT()
        try:
            user32.GetCursorPos(ctypes.byref(pt))
        except Exception:
            pt.x = 0
            pt.y = 0

        MONITOR_DEFAULTTONEAREST = 2
        hmon = user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
        if not hmon:
            return None
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(MONITORINFO)
        if not user32.GetMonitorInfoW(hmon, ctypes.byref(mi)):
            return None
        return {
            "left": int(mi.rcWork.left),
            "top": int(mi.rcWork.top),
            "width": int(mi.rcWork.right - mi.rcWork.left),
            "height": int(mi.rcWork.bottom - mi.rcWork.top),
        }
    except Exception:
        return None


def get_windows_dpi(hwnd=None):
    try:
        if not hasattr(ctypes, "windll"):
            return 96
        user32 = ctypes.windll.user32
        if hwnd and hasattr(user32, "GetDpiForWindow"):
            dpi = int(user32.GetDpiForWindow(hwnd))
            if dpi > 0:
                return dpi
        if hasattr(user32, "GetDpiForSystem"):
            dpi = int(user32.GetDpiForSystem())
            if dpi > 0:
                return dpi
    except Exception:
        pass
    return 96



def build_simplified_taskbar_icon(size=256):

    """

    Tạo app-mark phẳng, dễ đọc ở kích thước nhỏ cho taskbar/shortcut.

    """

    from PIL import ImageDraw, ImageFont, ImageEnhance



    img = Image.new("RGBA", (size, size), (255, 255, 255, 0))

    draw = ImageDraw.Draw(img)



    bg = "#0f2b4d"

    bg2 = "#153d6c"

    gold = "#d7a21f"

    white = "#f5f7fb"

    accent = "#7fb2ff"



    # A compact mark that stays readable at 16 px without a dark box.

    try:

        font = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", max(82, size // 2))

    except Exception:

        font = ImageFont.load_default()



    # Center the letters manually using measured bounds.

    g_x = int(size * 0.11)

    k_x = int(size * 0.50)

    y = int(size * 0.18)

    draw.text((g_x, y), "G", font=font, fill=white)

    draw.text((k_x, y), "K", font=font, fill=gold)



    # Small base bar to hint the pilepress brand without clutter.

    bar_y = int(size * 0.78)

    draw.rounded_rectangle((int(size * 0.14), bar_y, int(size * 0.86), bar_y + max(9, size // 18)), radius=max(4, size // 28), fill="#0d2748")

    draw.rectangle((int(size * 0.43), int(size * 0.14), int(size * 0.49), int(size * 0.75)), fill="#1d3557")

    draw.rectangle((int(size * 0.49), int(size * 0.14), int(size * 0.52), int(size * 0.75)), fill=accent)



    img = ImageEnhance.Sharpness(img).enhance(1.12)

    img = ImageEnhance.Contrast(img).enhance(1.06)

    return img



def build_detailed_app_icon(source_path, size=256):

    """

    Build the app/desktop icon from the original detailed logo.

    """

    img = Image.open(source_path).convert("RGBA")

    w, h = img.size

    # Keep the emblem only, matching the reference icon style.

    img = img.crop((0, 0, w, int(h * 0.76)))

    bbox = img.getchannel("A").getbbox()

    if bbox:

        img = img.crop(bbox)

    canvas = Image.new("RGBA", (size, size), (255, 255, 255, 0))

    fit_w = int(size * 1.08)

    fit_h = int(size * 1.08)

    scale = min(fit_w / img.size[0], fit_h / img.size[1])

    new_size = (max(1, int(img.size[0] * scale)), max(1, int(img.size[1] * scale)))

    img = img.resize(new_size, Image.Resampling.LANCZOS)

    canvas.alpha_composite(img, ((size - new_size[0]) // 2, 0))

    img = canvas

    try:

        from PIL import ImageEnhance, ImageFilter

        img = ImageEnhance.Sharpness(img).enhance(1.6)

        img = ImageEnhance.Contrast(img).enhance(1.12)

        img = img.filter(ImageFilter.UnsharpMask(radius=0.4, percent=230, threshold=0))

    except Exception:

        pass

    return img



def build_icon_variant(source_path, size):

    """

    Build a size-specific transparent app icon variant from the logo source.

    """

    img = Image.open(source_path).convert("RGBA")

    w, h = img.size

    img = img.crop((0, 0, w, int(h * 0.76)))

    bbox = img.getchannel("A").getbbox()

    if bbox:

        img = img.crop(bbox)

    params = {
        16: (1.18, 1.95, 1.16, 0.34, 250),
        32: (1.12, 1.80, 1.14, 0.36, 235),
        48: (1.08, 1.65, 1.12, 0.40, 220),
        64: (1.06, 1.50, 1.10, 0.42, 200),
        128: (1.04, 1.34, 1.08, 0.46, 175),
        256: (1.02, 1.22, 1.06, 0.50, 155),
    }

    fit_ratio, sharpness, contrast, radius, percent = params.get(int(size), (1.00, 1.45, 1.08, 0.45, 180))

    canvas = Image.new("RGBA", (size, size), (255, 255, 255, 0))

    fit_w = int(size * fit_ratio)

    fit_h = int(size * fit_ratio)

    scale = min(fit_w / img.size[0], fit_h / img.size[1])

    new_size = (max(1, int(img.size[0] * scale)), max(1, int(img.size[1] * scale)))

    img = img.resize(new_size, Image.Resampling.LANCZOS)

    canvas.alpha_composite(img, ((size - new_size[0]) // 2, 0))

    try:

        from PIL import ImageEnhance, ImageFilter

        canvas = ImageEnhance.Sharpness(canvas).enhance(sharpness)

        canvas = ImageEnhance.Contrast(canvas).enhance(contrast)

        canvas = canvas.filter(ImageFilter.UnsharpMask(radius=radius, percent=percent, threshold=0))

    except Exception:

        pass

    return canvas
