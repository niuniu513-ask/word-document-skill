# Contributing to word-document-skill

Thanks for your interest in contributing! This document covers how to set up
your development environment, run tests, and submit changes.

## Development Setup

```bash
# 1. Clone
git clone https://github.com/<your-org>/word-document-skill.git
cd word-document-skill

# 2. Create virtual environment
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements-dev.txt
```

## Project Structure

```
word-document-skill/
├── SKILL.md              # Agent entry point (six-phase workflow + rules)
├── README.md             # Human-facing overview
├── scripts/              # 8 Python CLI tools (+ shared _docx_utils.py)
├── schemas/              # JSON Schema definitions
├── workflows/            # Task-type-specific workflow guides
├── references/           # Deep-dive rule documents
├── evals/                # Integration test cases
├── tests/                # Unit tests (pytest)
└── examples/             # Example content.json files
```

## Code Style

- **Python**: target 3.10+, use `from __future__ import annotations`.
- **Formatting**: run `ruff check scripts/ tests/` before committing.
- **Type hints**: all public functions should have type annotations.
- **Docstrings**: every script starts with a module docstring showing usage.
- **Chinese comments**: the project uses Chinese for user-facing messages;
  code comments may be English or Chinese, but be consistent within a file.

## Running Tests

```bash
# Unit tests
pytest tests/ -v

# With coverage
pytest tests/ --cov=scripts --cov-report=term-missing

# Smoke test all scripts (requires fixtures)
python scripts/analyze_docx.py fixtures/实验报告模板.docx -o /tmp/spec.json
```

## Adding a New Script

1. Place it in `scripts/` with a clear CLI interface (`argparse`).
2. Import shared helpers from `_docx_utils.py` — don't duplicate XML parsing.
3. Call `setup_stdout()` at the start of `main()` for Windows UTF-8 safety.
4. Add a row to the script quick-reference table in `SKILL.md` §8.
5. Add unit tests in `tests/test_<script_name>.py`.

## Adding a New Workflow

1. Create `workflows/<task-type>.md`.
2. Register it in the Phase 0 routing table in `SKILL.md` §2.
3. Include: when to use, exact commands, common pitfalls, validation steps.

## Pull Request Checklist

- [ ] `ruff check scripts/ tests/` passes
- [ ] `pytest tests/` passes
- [ ] New functionality has tests
- [ ] `SKILL.md` and `README.md` updated if behavior changed
- [ ] `CHANGELOG.md` updated with an entry under `## [Unreleased]`
- [ ] No hardcoded absolute paths (use `Path(__file__).parent` for skill-relative)

## Reporting Issues

When filing a bug, include:
- The exact command run
- The input docx (or a minimal reproduction)
- Expected vs actual output
- Python version and OS (`python --version`, `uname -a` or `ver`)

## Release Process

1. Update version in `pyproject.toml` and `CHANGELOG.md`.
2. Tag: `git tag v1.x.x && git push --tags`.
3. Build: `python -m build`.
4. The `.skill` package is created by `scripts/package_skill.py`.
