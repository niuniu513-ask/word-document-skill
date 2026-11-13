#!/usr/bin/env python
"""docx_agent.py - one entry point for every Word-document task in this skill.

A single router so an agent (or a person) does not have to memorise twelve script
names: ``python scripts/docx_agent.py <task> ...``.  The task table lives in
``capabilities.json`` at the repository root, so the manifest a reader sees and
the commands that actually run can never drift apart (``tests/test_docx_agent.py``
keeps them honest).

Examples
    python scripts/docx_agent.py list
    python scripts/docx_agent.py analyze report.docx -o spec.json
    python scripts/docx_agent.py review report.docx --profile railway-plan --html r.html
    python scripts/docx_agent.py --capabilities
    python scripts/docx_agent.py --doctor
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _docx_utils import setup_stdout  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
MANIFEST = ROOT / "capabilities.json"

ALIASES = {
    "分析": "analyze", "结构": "analyze",
    "取内容": "extract", "提取": "extract",
    "样式": "styles",
    "生成": "generate", "填充": "generate", "插入": "generate",
    "验证": "validate", "格式差异": "validate",
    "体检": "review", "审查": "review", "合规": "review",
    "修复": "repair", "排版修复": "repair",
    "渲染": "render", "出图": "render", "导出pdf": "render",
    "证据": "evidence", "对账": "evidence",
    "裁图": "crop",
    "采集": "capture", "截图": "capture",
    "notebook截图": "notebook-html",
    "实验报告": "lab-report",
    "打包": "package",
}


def load_manifest() -> dict:
    with open(MANIFEST, encoding="utf-8") as fh:
        return json.load(fh)


def tasks_by_id(manifest: dict) -> dict:
    return {t["id"]: t for t in manifest["tasks"]}


def print_usage(manifest: dict) -> None:
    print(f"{manifest['name']} v{manifest['version']} - {manifest['positioning']}\n")
    print("用法: python scripts/docx_agent.py <task> [参数...]")
    print("      python scripts/docx_agent.py list | --capabilities | --doctor | --guide\n")
    width = max(len(t["id"]) for t in manifest["tasks"])
    for t in manifest["tasks"]:
        print(f"  {t['id']:<{width}}  {t['title']}")
    print("\n不知道该用哪个? 先跑 list, 再看 AGENT-GUIDE.md 的任务路由表。")


def print_list(manifest: dict) -> None:
    width = max(len(t["id"]) for t in manifest["tasks"])
    print(f"{manifest['name']} v{manifest['version']}: {len(manifest['tasks'])} tasks\n")
    for t in manifest["tasks"]:
        print(f"{t['id']:<{width}}  {t['title']}")
        print(f"{'':<{width}}  何时用: {' / '.join(t['when'])}")
        print(f"{'':<{width}}  命令:   python {t['script']} {t['args']}")
        if t.get("workflow"):
            print(f"{'':<{width}}  工作流: {t['workflow']}")
        print()


def print_guide() -> None:
    guide = ROOT / "AGENT-GUIDE.md"
    if not guide.exists():
        print("AGENT-GUIDE.md 不存在", file=sys.stderr)
        return
    print(f"指南: {guide}\n")
    for line in guide.read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            print("  " + line[3:])
    print("\n直接读全文: less AGENT-GUIDE.md")


def prepare_outputs(rest: list[str], root: Path) -> list[str]:
    """Create parent directories for -o/--output/--out values (agents hate ENOENT)."""
    made: list[str] = []
    for i, token in enumerate(rest):
        if token not in ("-o", "--output", "--out") or i + 1 >= len(rest):
            continue
        target = Path(rest[i + 1])
        if not target.is_absolute():
            target = root / target
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.suffix == "":
                target.mkdir(parents=True, exist_ok=True)
            made.append(str(target))
        except OSError:
            pass
    return made


def run_doctor(manifest: dict) -> int:
    problems = 0
    print(f"python  {sys.version.split()[0]}")
    for mod in ("docx", "lxml"):
        try:
            __import__(mod)
            print(f"  ✓ {mod}")
        except Exception as exc:  # pragma: no cover - environment dependent
            print(f"  ✗ {mod}: {exc}")
            problems += 1
    for t in manifest["tasks"]:
        target = ROOT / t["script"]
        mark = "✓" if target.exists() else "✗"
        if not target.exists():
            problems += 1
        print(f"  {mark} {t['id']:<14} {t['script']}")
    for rule in sorted((ROOT / "rules").glob("*.json")):
        try:
            json.loads(rule.read_text(encoding="utf-8"))
            print(f"  ✓ rules/{rule.name}")
        except Exception as exc:
            print(f"  ✗ rules/{rule.name}: {exc}")
            problems += 1
    for doc_key, rel in manifest.get("docs", {}).items():
        mark = "✓" if (ROOT / rel).exists() else "✗"
        if not (ROOT / rel).exists():
            problems += 1
        print(f"  {mark} {doc_key:<7} {rel}")
    print("\n结论: " + ("一切就绪" if problems == 0 else f"{problems} 处需要处理"))
    return 0 if problems == 0 else 1


def main(argv: list[str]) -> int:
    setup_stdout()
    manifest = load_manifest()
    if not argv or argv[0] in ("-h", "--help", "help"):
        print_usage(manifest)
        return 0
    cmd = argv[0]
    if cmd in ("list", "--list", "-l"):
        print_list(manifest)
        return 0
    if cmd in ("--capabilities", "capabilities", "manifest"):
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
        return 0
    if cmd in ("--doctor", "doctor"):
        return run_doctor(manifest)
    if cmd in ("--guide", "guide"):
        print_guide()
        return 0

    task_id = ALIASES.get(cmd, cmd)
    tasks = tasks_by_id(manifest)
    if task_id not in tasks:
        print(f"不认识的任务: {cmd}", file=sys.stderr)
        print("可用任务: " + ", ".join(tasks), file=sys.stderr)
        print("提示: python scripts/docx_agent.py list", file=sys.stderr)
        return 2

    task = tasks[task_id]
    rest = argv[1:]
    dry = "--dry-run" in rest
    rest = [a for a in rest if a != "--dry-run"]
    script = ROOT / task["script"]
    prepare_outputs(rest, ROOT)
    command = [sys.executable, str(script), *rest]
    if dry:
        shown = " ".join(f'"{c}"' if " " in c else c for c in command[1:])
        print(f"[dry-run] python {shown}")
        print(f"[task]    {task['id']} - {task['title']}")
        print(f"[expect]  输入: {', '.join(task['inputs'])}  ->  产出: {', '.join(task['outputs'])}")
        return 0
    if not rest and task["id"] != "lab-report":
        print(f"[task] {task['id']} - {task['title']}", file=sys.stderr)
        print(f"[example] python {task['script']} {task['args']}", file=sys.stderr)
    return subprocess.call(command)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
