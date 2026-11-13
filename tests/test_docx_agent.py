"""docx_agent.py（唯一入口）与 capabilities.json（任务清单）的测试。

两组断言：
1. 清单自洽——清单里写的脚本 / 工作流 / 规则包 / 文档都必须真的存在，
   否则这份「指南书」就是在骗读者。
2. 路由正确——任务能找到、未知任务退出码 2、--dry-run 不执行副作用。
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
AGENT = ROOT / "scripts" / "docx_agent.py"
MANIFEST = ROOT / "capabilities.json"


def run(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(AGENT), *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=ROOT,
    )


@pytest.fixture(scope="module")
def manifest() -> dict:
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def test_manifest_shape(manifest):
    assert manifest["name"] == "word-document-skill"
    assert manifest["tasks"], "任务清单不能为空"
    ids = [t["id"] for t in manifest["tasks"]]
    assert len(ids) == len(set(ids)), "任务 id 必须唯一"
    assert len(manifest["iron_rules"]) == 3


def test_every_referenced_file_exists(manifest):
    missing = []
    for task in manifest["tasks"]:
        for key in ("script", "workflow"):
            rel = task.get(key)
            if rel and not (ROOT / rel).exists():
                missing.append(f"{task['id']}.{key}: {rel}")
    for rel in manifest["docs"].values():
        if not (ROOT / rel).exists():
            missing.append(f"docs: {rel}")
    assert not missing, "清单引用了不存在的文件: " + "; ".join(missing)


def test_every_task_documents_its_contract(manifest):
    for task in manifest["tasks"]:
        assert task["title"].strip(), task["id"]
        assert task["when"] and all(w.strip() for w in task["when"]), task["id"]
        assert task["args"].strip(), task["id"]
        assert task["inputs"] and task["outputs"], task["id"]


def test_script_covered_tasks_are_real_scripts(manifest):
    """清单里出现的每个 scripts/*.py 都必须存在且不是目录。"""
    for task in manifest["tasks"]:
        script = ROOT / task["script"]
        assert script.is_file(), script
        assert script.suffix == ".py", script


def test_capabilities_flag_prints_valid_json(manifest):
    out = run("--capabilities")
    assert out.returncode == 0, out.stderr
    parsed = json.loads(out.stdout)
    assert parsed["version"] == manifest["version"]
    assert [t["id"] for t in parsed["tasks"]] == [t["id"] for t in manifest["tasks"]]


def test_list_and_help_exit_zero():
    for flag in ("list", "help", "--guide"):
        out = run(flag)
        assert out.returncode == 0, (flag, out.stderr)
    assert "review" in run("list").stdout


def test_unknown_task_exits_2_with_hint():
    out = run("frobnicate")
    assert out.returncode == 2
    assert "不认识的任务" in out.stderr
    assert "review" in out.stderr


def test_chinese_alias_resolves(manifest):
    out = run("体检", "--dry-run", "a.docx")
    assert out.returncode == 0
    assert "review" in out.stdout
    assert "review.py" in out.stdout


def test_dry_run_does_not_touch_disk(tmp_path):
    target = tmp_path / "spec.json"
    out = run("analyze", "--dry-run", "missing.docx", "-o", str(target))
    assert out.returncode == 0, out.stderr
    assert "[dry-run]" in out.stdout
    assert "analyze_docx.py" in out.stdout
    assert not target.exists()


def test_prepare_outputs_creates_missing_dirs(tmp_path):
    import docx_agent

    nested = tmp_path / "a" / "b" / "spec.json"
    docx_agent.prepare_outputs(["-o", str(nested)], docx_agent.ROOT)
    assert nested.parent.is_dir()
    assert not nested.exists(), "只应建目录,不应创建文件"


def test_prepare_outputs_creates_directory_target(tmp_path):
    import docx_agent

    pages = tmp_path / "pages"
    docx_agent.prepare_outputs(["--out", str(pages)], docx_agent.ROOT)
    assert pages.is_dir()


def test_doctor_reports_ready():
    out = run("--doctor")
    assert "结论" in out.stdout
    # 依赖齐全时必须一切就绪；环境缺依赖时应给出明确提示而不是崩溃
    assert out.returncode == 0, out.stdout + out.stderr
