"""Capture a scrollable editor window into stitched screenshots (Windows).

Two modes:

    # list top-level windows, find the exact title to pass
    python capture_notebook.py --list-windows --filter ipynb

    # capture: focus the window, scroll from top to bottom, save frames
    python capture_notebook.py --title "test-1.ipynb" --out _cap/test1 \
        --crop-top 70 --crop-bottom 34 --step 2

    # stitch already-captured frames (also runs automatically after capture)
    python capture_notebook.py --stitch _cap/test1 -o c1_tall.png

Design notes that matter (see references/screenshot-evidence.md):

* Real foreground is required: Chromium ignores posted WM_MOUSEWHEEL, so this
  script moves the cursor and emits real wheel events.
* Wheel delta is POSITIVE for scrolling up, NEGATIVE for scrolling down.
* DPI awareness must be set before any window API call, otherwise window
  rectangles come back in scaled logical pixels and the crops are blurry.
* Frames are cropped to the scrolling area (fixed chrome removed) so the
  stitched image does not repeat the tab bar.

Verified on: Windows 10, VS Code 1.9x + Jupyter, 150% scaling, 2560x1600.
crop-top / crop-bottom are DPI-dependent: always check the preview.
"""

from __future__ import annotations

import argparse
import hashlib
import sys
import time
from pathlib import Path

from PIL import Image, ImageGrab

MOUSEEVENTF_WHEEL = 0x0800
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_WHEEL_DELTA = 120
SW_MAXIMIZE = 3


def _u32():
    import ctypes

    if sys.platform != "win32":
        raise SystemExit("此脚本仅支持 Windows(需要 Win32 窗口 API)")
    return ctypes.windll.user32


def dpi_aware() -> None:
    """Declare per-monitor DPI awareness before touching any window API."""
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def list_windows(keyword: str = "") -> list[dict]:
    import ctypes
    from ctypes import wintypes

    u32 = _u32()
    results: list[dict] = []
    proc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

    def _cb(hwnd, _lparam):
        if not u32.IsWindowVisible(hwnd):
            return True
        length = u32.GetWindowTextLengthW(hwnd)
        if not length:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        u32.GetWindowTextW(hwnd, buf, length + 1)
        title = buf.value
        if keyword and keyword.lower() not in title.lower():
            return True
        rect = wintypes.RECT()
        u32.GetWindowRect(hwnd, ctypes.byref(rect))
        results.append({
            "hwnd": hwnd,
            "title": title,
            "rect": (rect.left, rect.top, rect.right, rect.bottom),
            "size": (rect.right - rect.left, rect.bottom - rect.top),
        })
        return True

    u32.EnumWindows(proc(_cb), 0)
    return results


def find_window(title: str) -> int:
    for win in list_windows(title):
        if win["size"][0] > 400 and win["size"][1] > 300:
            return win["hwnd"]
    raise SystemExit(f"没找到标题包含「{title}」的窗口,先用 --list-windows 确认")


def focus(hwnd: int) -> None:
    """Bring the window forward, then click once so the webview owns the queue."""
    import ctypes
    from ctypes import wintypes

    u32 = _u32()
    u32.ShowWindow(hwnd, SW_MAXIMIZE)
    time.sleep(0.35)
    u32.SetForegroundWindow(hwnd)
    time.sleep(0.15)

    rect = wintypes.RECT()
    u32.GetWindowRect(hwnd, ctypes.byref(rect))
    cx = (rect.left + rect.right) // 2
    cy = (rect.top + rect.bottom) // 2
    u32.SetCursorPos(cx, cy)
    u32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    u32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.12)
    # clicking inside a notebook cell can enter edit mode; ESC backs out
    u32.keybd_event(0x1B, 0, 0, 0)
    u32.keybd_event(0x1B, 0, 2, 0)
    time.sleep(0.12)


def wheel(notches: int) -> None:
    """Positive = scroll up (content moves down), negative = scroll down."""
    _u32().mouse_event(MOUSEEVENTF_WHEEL, 0, 0, MOUSEEVENTF_WHEEL_DELTA * notches, 0)


def scroll_to_top(rounds: int = 30) -> None:
    """Scroll up until two consecutive screen hashes stop changing."""
    last = None
    for _ in range(rounds):
        wheel(6)
        time.sleep(0.15)
        cur = _screen_hash()
        if cur == last:
            return
        last = cur


def _screen_hash() -> str:
    img = ImageGrab.grab()
    return hashlib.md5(img.tobytes()).hexdigest()


def grab(hwnd: int, crop_top: int, crop_bottom: int) -> Image.Image:
    import ctypes
    from ctypes import wintypes

    rect = wintypes.RECT()
    _u32().GetWindowRect(hwnd, ctypes.byref(rect))
    box = (rect.left, rect.top, rect.right, rect.bottom)
    try:
        img = ImageGrab.grab(bbox=box)          # Pillow >= 10 spans all screens
    except TypeError:
        img = ImageGrab.grab(bbox=box, all_screens=True)
    w, h = img.size
    top = max(0, min(crop_top, h // 3))
    bottom = max(0, min(crop_bottom, h // 3))
    return img.crop((0, top, w, h - bottom))


def _gray(img: Image.Image, div: int = 4) -> Image.Image:
    return img.convert("L").resize(
        (max(1, img.width // div), max(1, img.height // div)), Image.BOX)


def _edge(img: Image.Image, div: int = 4) -> Image.Image:
    """Edge view: suppresses smooth gradients that make blank bands self-similar."""
    from PIL import ImageFilter

    g = img.convert("L").filter(ImageFilter.FIND_EDGES)
    return g.resize((max(1, g.width // div), max(1, g.height // div)), Image.BOX)


def _mad(a: Image.Image, b: Image.Image) -> float:
    from PIL import ImageChops, ImageStat

    return ImageStat.Stat(ImageChops.difference(a, b)).mean[0]


def _scan(p: Image.Image, c: Image.Image, lo: int, hi: int) -> tuple[int, float]:
    W, H = p.size
    best, best_cost = lo, None
    for d in range(max(1, lo), max(lo + 1, hi)):
        overlap = H - d
        if overlap < 8:
            break
        cost = _mad(p.crop((0, d, W, H)), c.crop((0, 0, W, overlap)))
        if best_cost is None or cost < best_cost:
            best_cost, best = cost, d
    return best, (best_cost if best_cost is not None else 1e9)


def estimate_offset(prev: Image.Image, cur: Image.Image, div: int = 4,
                    pf: Image.Image | None = None,
                    cf: Image.Image | None = None) -> tuple[int, float]:
    """Scroll distance between two frames, plus the overlap ratio.

    A 2-D block match on the edge view beats a 1-D row profile: notebook pages
    repeat identical blank bands, and row profiles happily lock onto the wrong
    one. Edges keep the match anchored on text and plot borders.
    """
    p_s = pf if pf is not None else _edge(prev, div)
    c_s = cf if cf is not None else _edge(cur, div)
    W, H = p_s.size

    best, _ = _scan(p_s, c_s, int(H * 0.06), int(H * 0.96))
    coarse = best * div

    # refine on plain gray at quarter resolution, then at full resolution
    p_f, c_f = _gray(prev, 2), _gray(cur, 2)
    W_f, H_f = p_f.size
    half = div
    d2, _ = _scan(p_f, c_f, max(1, coarse // 2 - half), min(H_f - 8, coarse // 2 + half))
    coarse = d2 * 2

    p_g, c_g = prev.convert("L"), cur.convert("L")
    lo = max(1, coarse - 6)
    hi = min(p_g.height - 8, coarse + 6)
    best_f, _ = _scan(p_g, c_g, lo, max(lo + 1, hi))

    overlap = 1.0 - (best_f / prev.height) if prev.height else 0.0
    return max(0, best_f), max(0.0, min(1.0, overlap))


def stitch(frames: list[Path], out: Path, quiet: bool = False,
           crop: tuple[int, int] = (0, 0)) -> tuple[int, list[float]]:
    top, bottom = crop
    imgs = []
    for f in frames:
        im = Image.open(f).convert("RGB")
        if top or bottom:
            im = im.crop((0, top, im.width, max(top + 1, im.height - bottom)))
        imgs.append(im)
    if not imgs:
        raise SystemExit("没有可拼接的帧")
    sizes = {im.size for im in imgs}
    if len(sizes) > 1:
        detail = ", ".join(f"{f.name}={Image.open(f).size}" for f in frames[:6])
        raise SystemExit(
            f"帧尺寸不一致,不能拼接: {detail}\n"
            f"提示:采集帧请放独立目录(自动拼接只认 frame_*.png),"
            f"不要把长图、预览图混进帧目录")

    feats = [_edge(im) for im in imgs]
    offsets = [0]
    overlaps: list[float] = []
    for i in range(1, len(imgs)):
        offset, overlap = estimate_offset(
            imgs[i - 1], imgs[i], pf=feats[i - 1], cf=feats[i])
        if offset <= 0:
            offset = max(1, imgs[i].height // 2)
            overlap = 0.5
        offsets.append(offset)
        overlaps.append(overlap)

    total = imgs[0].height + sum(offsets[1:])
    canvas = Image.new("RGB", (imgs[0].width, max(total, 1)), (18, 19, 20))
    canvas.paste(imgs[0], (0, 0))
    y = 0
    for i in range(1, len(imgs)):
        y += offsets[i]
        canvas.paste(imgs[i], (0, y))
    out.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out)

    # self-check: a wrong offset shows up as a large residual in the overlap
    suspect = []
    gray = [_gray(im) for im in imgs]
    for i in range(1, len(imgs)):
        h, w = gray[i - 1].size[1], gray[i - 1].size[0]
        d4 = offsets[i] // 4
        if d4 < 4 or h - d4 < 8:
            continue
        residual = _mad(gray[i - 1].crop((0, d4, w, h)), gray[i].crop((0, 0, w, h - d4)))
        if residual > 12:
            suspect.append((i, round(residual, 1)))

    if not quiet:
        print(f"拼接 {len(imgs)} 帧 -> {out} ({canvas.width}x{canvas.height})")
        if overlaps:
            print(f"  重叠率 min {min(overlaps):.2f} / mean {sum(overlaps)/len(overlaps):.2f}"
                  f"(低于 0.35 请减小 --step)")
        if suspect:
            print(f"  WARN 第 {', '.join(str(i) for i, _ in suspect)} 帧对齐可疑"
                  f"(残差 {', '.join(str(r) for _, r in suspect)});"
                  f"请检查这一段的拼接图是否有重影或跳行")
    return canvas.height, overlaps


def preview(hwnd: int, crop_top: int, crop_bottom: int, out: Path) -> None:
    from PIL import ImageDraw

    img = grab(hwnd, 0, 0).convert("RGB")
    draw = ImageDraw.Draw(img)
    for y, color in ((crop_top, (255, 0, 0)), (img.height - crop_bottom, (255, 0, 0))):
        draw.line((0, y, img.width, y), fill=color, width=3)
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out)
    print(f"预览 -> {out};红线之外是要裁掉的固定 UI")


def capture(title: str, out_dir: Path, crop_top: int, crop_bottom: int, step: int,
            wait_ms: int, max_frames: int, quiet: bool) -> list[Path]:
    hwnd = find_window(title)
    focus(hwnd)
    scroll_to_top()
    time.sleep(0.25)

    out_dir.mkdir(parents=True, exist_ok=True)
    frames: list[Path] = []
    last_hash = ""
    for i in range(max_frames):
        img = grab(hwnd, crop_top, crop_bottom)
        digest = hashlib.md5(img.tobytes()).hexdigest()
        if digest == last_hash:
            if not quiet:
                print(f"  第 {i} 帧与上一帧相同,已到底")
            break
        path = out_dir / f"frame_{i:03d}.png"
        img.save(path)
        frames.append(path)
        last_hash = digest
        wheel(-step)
        time.sleep(wait_ms / 1000)
        if not quiet and i % 10 == 0:
            print(f"  帧 {i:03d} {img.width}x{img.height}")
    if not quiet:
        print(f"共 {len(frames)} 帧 -> {out_dir}")
    return frames


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="采集滚动窗口为拼接长图(Windows)")
    ap.add_argument("--list-windows", action="store_true", help="列出可见窗口后退出")
    ap.add_argument("--filter", default="", help="配合 --list-windows 过滤标题")
    ap.add_argument("--title", help="目标窗口标题(包含匹配)")
    ap.add_argument("--out", "-o", help="capture: 帧目录;--stitch: 长图路径")
    ap.add_argument("--stitch", metavar="DIR", help="只拼接已有帧目录")
    ap.add_argument("--crop-top", type=int, default=None,
                    help="裁掉顶部固定 UI 的像素(capture 默认 175,stitch 默认 0);"
                         "150%% 缩放下标题栏+菜单+标签栏+工具栏约 175px,用 --preview 校准")
    ap.add_argument("--crop-bottom", type=int, default=None,
                    help="裁掉底部固定 UI 的像素(capture 默认 55,stitch 默认 0)")
    ap.add_argument("--step", type=int, default=2, help="每帧滚动的滚轮格数(默认 2)")
    ap.add_argument("--wait", type=int, default=260, help="每帧等待毫秒(默认 260)")
    ap.add_argument("--max-frames", type=int, default=400)
    ap.add_argument("--preview", metavar="PNG", help="只抓一帧并画出裁切线")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    if args.list_windows:
        dpi_aware()
        wins = list_windows(args.filter)
        wins.sort(key=lambda w: -w["size"][0] * w["size"][1])
        for w in wins[:40]:
            print(f"{w['hwnd']:#010x}  {w['size'][0]}x{w['size'][1]}  {w['title'][:90]}")
        print(f"共 {len(wins)} 个窗口")
        return 0

    if args.stitch:
        src = Path(args.stitch)
        out = Path(args.out) if args.out else src.with_suffix("_tall.png")
        frames = sorted(src.glob("frame_*.png"))
        if not frames:
            frames = [
                f for f in sorted(src.glob("*.png"))
                if not f.name.startswith("_")
                and "_tall" not in f.stem
                and f.name != out.name
            ]
        stitch(frames, out, args.quiet, (args.crop_top or 0, args.crop_bottom or 0))
        return 0

    if not args.title:
        ap.error("capture 需要 --title(或用 --list-windows 先看)")

    dpi_aware()
    crop_top = 175 if args.crop_top is None else args.crop_top
    crop_bottom = 55 if args.crop_bottom is None else args.crop_bottom
    if args.preview:
        preview(find_window(args.title), crop_top, crop_bottom, Path(args.preview))
        return 0

    out_dir = Path(args.out) if args.out else Path("_cap") / "frames"
    frames = capture(args.title, out_dir, crop_top, crop_bottom,
                     args.step, args.wait, args.max_frames, args.quiet)
    if frames:
        stitch(frames, out_dir.with_name(out_dir.name + "_tall.png"), args.quiet)
    return 0


if __name__ == "__main__":
    sys.exit(main())
