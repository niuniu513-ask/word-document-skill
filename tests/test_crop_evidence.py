"""Check pixel preservation and stale/out-of-bounds coordinate rejection."""

import hashlib
import json
import sys
from pathlib import Path

import pytest
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
from crop_evidence import crop_evidence, pixel_box  # noqa: E402


def test_scaled_crop_preserves_pixels_and_provenance(tmp_path):
    raw = Image.new("RGB", (80, 60))
    raw.putdata([(i % 256, i // 256, 10) for i in range(80 * 60)])
    source, output = tmp_path / "raw.png", tmp_path / "crop.png"
    raw.save(source)
    before = source.read_bytes()
    record = crop_evidence(source, output, (4, 3, 10, 8), scale=1.25,
                           origin=(2, 4), identity="numpy/cell5/code")
    assert record["crop_ltrb_px"] == [7, 7, 20, 18]
    with Image.open(output) as result:
        assert result.tobytes() == raw.crop((7, 7, 20, 18)).tobytes()
    assert source.read_bytes() == before
    assert record["source_sha256"] == hashlib.sha256(before).hexdigest()
    assert json.loads(output.with_suffix(".png.json").read_text("utf-8")) == record


@pytest.mark.parametrize("box,scale", [
    ((0, 0, 81, 2), 1), ((-1, 0, 5, 5), 1), ((0, 0, 1, 1), 0),
    ((0, 0, 0, 3), 1), ((0, float("nan"), 3, 3), 1),
    ((0, 0, 60, 40), 1.5),
])
def test_invalid_transform_rejected(box, scale):
    with pytest.raises(ValueError):
        pixel_box(box, (80, 60), scale)


def test_existing_crop_and_raw_never_overwritten(tmp_path):
    source, output = tmp_path / "raw.png", tmp_path / "crop.png"
    Image.new("RGB", (20, 20)).save(source)
    with pytest.raises(ValueError):
        crop_evidence(source, source, (0, 0, 10, 10))
    crop_evidence(source, output, (0, 0, 10, 10))
    before = output.read_bytes()
    with pytest.raises(FileExistsError):
        crop_evidence(source, output, (0, 0, 5, 5))
    assert output.read_bytes() == before
