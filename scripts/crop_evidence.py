"""Crop a raw screenshot with an explicit coordinate transform and audit sidecar.

Default coordinates are physical screenshot pixels. For CSS/logical coordinates,
pass the measured --scale and screenshot-pixel --origin. Neither DPR nor viewport
size is inferred: desktop captures can include pixels outside the viewport.
This validates geometry/provenance, not which code cell the crop contains.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import sys
from pathlib import Path

from PIL import Image

sys.path.insert(0, str(Path(__file__).parent))
from _docx_utils import setup_stdout  # noqa: E402


def pixel_box(
    box: tuple[float, float, float, float], size: tuple[int, int],
    scale: float = 1.0, origin: tuple[float, float] = (0.0, 0.0),
) -> tuple[int, int, int, int]:
    """Map x/y/width/height to physical LTRB; reject rather than clip overflow."""
    x, y, width, height = box
    values = (*box, scale, *origin)
    if not all(math.isfinite(v) for v in values):
        raise ValueError("Coordinates and scale must be finite")
    if scale <= 0 or width <= 0 or height <= 0:
        raise ValueError("Scale, width and height must be positive")
    left = origin[0] + x * scale
    top = origin[1] + y * scale
    right = origin[0] + (x + width) * scale
    bottom = origin[1] + (y + height) * scale
    if left < 0 or top < 0 or right > size[0] or bottom > size[1]:
        raise ValueError("Crop exceeds raw screenshot; remeasure this capture")
    result = (math.floor(left), math.floor(top), math.ceil(right), math.ceil(bottom))
    return result


def crop_evidence(
    source: str | Path, output: str | Path, box: tuple[float, float, float, float], *,
    scale: float = 1.0, origin: tuple[float, float] = (0.0, 0.0),
    identity: str | None = None,
) -> dict:
    """Write a lossless crop and source/transform metadata without replacing raw evidence."""
    source, output = Path(source), Path(output)
    sidecar = output.with_suffix(output.suffix + ".json")
    if source.resolve() == output.resolve():
        raise ValueError("Preserve the raw screenshot; choose a different output")
    if output.exists() or sidecar.exists():
        raise FileExistsError("Use a new output filename for each crop revision")
    if output.suffix.lower() != ".png":
        raise ValueError("Use PNG to preserve screenshot pixels")
    with Image.open(source) as raw:
        bounds = pixel_box(box, raw.size, scale, origin)
        result = raw.crop(bounds)
        record = {
            "source": str(source.resolve()),
            "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "source_size_px": list(raw.size),
            "box_xywh": list(box),
            "scale": scale,
            "origin_px": list(origin),
            "crop_ltrb_px": list(bounds),
            "output": str(output.resolve()),
            "output_size_px": list(result.size),
            "identity": identity,
            "semantic_review": "pending: check cell identity, first/last line and output",
        }
        output.parent.mkdir(parents=True, exist_ok=True)
        result.save(output, format="PNG")
    record["output_sha256"] = hashlib.sha256(output.read_bytes()).hexdigest()
    sidecar.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
    return record


def main(argv: list[str] | None = None) -> int:
    setup_stdout()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--box", type=float, nargs=4, required=True,
                        metavar=("X", "Y", "WIDTH", "HEIGHT"))
    parser.add_argument("--scale", type=float, default=1.0,
                        help="Measured coordinate-to-screenshot scale (default: physical pixels)")
    parser.add_argument("--origin", type=float, nargs=2, default=(0.0, 0.0),
                        help="Screenshot-pixel offset of the coordinate origin")
    parser.add_argument("--identity", help="Task/notebook/cell and code/output identity")
    args = parser.parse_args(argv)
    try:
        record = crop_evidence(args.source, args.output, args.box, scale=args.scale,
                               origin=args.origin, identity=args.identity)
    except (OSError, ValueError) as exc:
        parser.exit(2, f"crop failed: {exc}\n")
    print(json.dumps(record, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
