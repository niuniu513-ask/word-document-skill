"""Bridge to the pure-MoonBit document engine that ships in ``engine/``.

The skill has two implementations of the same contract:

* the Python scripts (``analyze_docx.py``, ``validate_docx.py``, ``repair_docx.py``
  — the ones an agent calls through SKILL.md), and
* the MoonBit engine under ``engine/``, which implements the analysis,
  editing and format-delta checks from the OOXML layer up.

This module lets the Python side hand work to the MoonBit engine when the
toolchain is installed, and stay silent (``EngineUnavailable``) when it is not,
so the skill keeps working on a plain Python install.

Usage:
    python mbt_engine.py doctor                      # is the engine usable?
    python mbt_engine.py analyze report.docx -o spec.json
    python mbt_engine.py validate before.docx after.docx
    python mbt_engine.py fill report.docx --set "【待补充】=值" --out filled.docx
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_ENGINE_DIR = REPO_ROOT / "engine"
DEFAULT_TIMEOUT = 600  # seconds; the first run compiles the engine


class EngineUnavailable(RuntimeError):
    """Raised when the MoonBit engine cannot be reached on this machine."""


def engine_dir() -> Path:
    """Directory that holds ``moon.mod`` for the engine."""
    override = os.environ.get("DOCX_MBT_DIR")
    return Path(override) if override else DEFAULT_ENGINE_DIR


def moon_binary() -> str | None:
    """Locate the ``moon`` executable (PATH, then $MOON_HOME/bin)."""
    found = shutil.which("moon")
    if found:
        return found
    home = os.environ.get("MOON_HOME")
    if home:
        for name in ("moon.exe", "moon"):
            candidate = Path(home) / "bin" / name
            if candidate.exists():
                return str(candidate)
    return None


def available() -> bool:
    """True when both the toolchain and the engine sources are present."""
    return moon_binary() is not None and (engine_dir() / "moon.mod").exists()


def _require() -> tuple[str, Path]:
    moon = moon_binary()
    if moon is None:
        raise EngineUnavailable(
            "the MoonBit toolchain was not found: install it from "
            "https://www.moonbitlang.com/download or set MOON_HOME"
        )
    root = engine_dir()
    if not (root / "moon.mod").exists():
        raise EngineUnavailable(f"engine sources not found under {root}")
    return moon, root


def _run(*args: str, timeout: int = DEFAULT_TIMEOUT) -> str:
    """Run one ``cmd/docx`` subcommand and return its stdout."""
    moon, root = _require()
    cmd = [moon, "run", "cmd/docx", "--target", "js", "--", *args]
    try:
        proc = subprocess.run(
            cmd,
            cwd=str(root),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:  # pragma: no cover - slow machines
        raise EngineUnavailable(f"engine timed out after {timeout}s") from exc
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()
        raise EngineUnavailable(f"engine exited with {proc.returncode}: {detail[:400]}")
    return proc.stdout


def analyze(path: str | Path, timeout: int = DEFAULT_TIMEOUT) -> dict:
    """Full Document Specification for one .docx, computed by the MoonBit engine."""
    path = Path(path).resolve()
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "spec.json"
        _run("analyze", str(path), "--out", str(out), timeout=timeout)
        if not out.exists():  # pragma: no cover - defensive
            raise EngineUnavailable("engine did not write the specification")
        return json.loads(out.read_text(encoding="utf-8"))


def validate(before: str | Path, after: str | Path, timeout: int = DEFAULT_TIMEOUT) -> dict:
    """Formatting delta between two documents, as a JSON object."""
    text = _run(
        "validate",
        str(Path(before).resolve()),
        str(Path(after).resolve()),
        "--json",
        timeout=timeout,
    )
    start = text.find("{")
    if start < 0:
        raise EngineUnavailable(f"engine returned no JSON: {text[:200]}")
    return json.loads(text[start:])


def fill(
    path: str | Path,
    values: dict[str, str],
    out: str | Path,
    timeout: int = DEFAULT_TIMEOUT,
) -> dict:
    """Replace placeholder literals in place and save the result."""
    if not values:
        raise ValueError("fill() needs at least one placeholder")
    args: list[str] = ["fill", str(Path(path).resolve())]
    for key, value in values.items():
        args += ["--set", f"{key}={value}"]
    args += ["--out", str(Path(out).resolve())]
    text = _run(*args, timeout=timeout)
    match = re.search(r"replaced (\d+) occurrence", text)
    return {
        "replaced": int(match.group(1)) if match else 0,
        "output": str(Path(out).resolve()),
        "raw": text.strip(),
    }


def analyze_or_none(path: str | Path) -> dict | None:
    """Convenience for callers that want a silent fallback to the Python path."""
    if not available():
        return None
    try:
        return analyze(path)
    except EngineUnavailable:
        return None


def doctor() -> dict:
    moon = moon_binary()
    root = engine_dir()
    report = {
        "engine_dir": str(root),
        "engine_sources": (root / "moon.mod").exists(),
        "moon_binary": moon,
        "available": available(),
    }
    if report["available"]:
        try:
            out = _run("info", str(REPO_ROOT / "engine" / "tests" / "fixtures" / "mini.docx"))
            report["smoke"] = out.strip().splitlines()[0] if out.strip() else ""
            report["ok"] = True
        except EngineUnavailable as exc:
            report["smoke"] = str(exc)
            report["ok"] = False
    else:
        report["ok"] = False
        report["hint"] = "install the MoonBit toolchain, or set MOON_HOME / DOCX_MBT_DIR"
    return report


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in {"-h", "--help", "help"}:
        print(__doc__)
        return 0
    command, rest = argv[0], argv[1:]
    try:
        if command == "doctor":
            print(json.dumps(doctor(), ensure_ascii=False, indent=2))
            return 0
        if command == "analyze":
            path = rest[0]
            spec = analyze(path)
            out = None
            if "-o" in rest:
                out = rest[rest.index("-o") + 1]
            elif "--out" in rest:
                out = rest[rest.index("--out") + 1]
            text = json.dumps(spec, ensure_ascii=False, indent=2)
            if out:
                Path(out).write_text(text, encoding="utf-8")
                print(f"wrote {out}")
            else:
                print(text)
            return 0
        if command == "validate":
            delta = validate(rest[0], rest[1])
            print(json.dumps(delta, ensure_ascii=False, indent=2))
            return 0
        if command == "fill":
            values: dict[str, str] = {}
            out = ""
            i = 1
            while i < len(rest):
                if rest[i] == "--set" and i + 1 < len(rest):
                    key, _, value = rest[i + 1].partition("=")
                    values[key] = value
                    i += 2
                elif rest[i] == "--out" and i + 1 < len(rest):
                    out = rest[i + 1]
                    i += 2
                else:
                    i += 1
            if not out:
                print("error: fill needs --out")
                return 2
            result = fill(rest[0], values, out)
            print(f"replaced {result['replaced']} occurrence(s) -> {result['output']}")
            return 0
        print(f"unknown command: {command}")
        return 2
    except EngineUnavailable as exc:
        print(f"engine unavailable: {exc}")
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
