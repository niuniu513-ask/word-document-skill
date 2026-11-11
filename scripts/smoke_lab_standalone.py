"""Self-contained smoke test. TEST graphics are fixtures, never IDE evidence."""

import argparse
import subprocess
import sys
from pathlib import Path

import nbformat
from docx import Document
from lab_report import read, write
from PIL import Image, ImageDraw


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", required=True)
    root = Path(parser.parse_args().out).resolve()
    root.mkdir(parents=True, exist_ok=True)
    scripts = Path(__file__).resolve().parent

    def run(*args):
        subprocess.run(
            [sys.executable, "-X", "utf8", str(scripts / "lab_report.py"), *map(str, args)],
            check=True,
            cwd=root,
        )

    nb = nbformat.v4.new_notebook(cells=[nbformat.v4.new_code_cell("print(2+3)")])
    source = root / "input.ipynb"
    nbformat.write(nb, source)
    run("prepare", source, "--execute", "--out", root / "execution")
    plan = read(root / "execution/capture-plan.json")
    image = root / "TEST-fixture.png"
    im = Image.new("RGB", (600, 120), "white")
    draw = ImageDraw.Draw(im)
    draw.text((30, 30), "TEST FIXTURE - NOT A SCREENSHOT", fill="black")
    draw.rectangle((30, 65, 560, 95), fill="blue")
    im.save(image)
    for cell in plan["cells"]:
        cell.update(
            env={"w": 600, "h": 120, "d": 1},
            codeFile=str(image),
            outputFile=str(image),
            codeBox={"x": 0, "y": 0, "width": 600, "height": 120},
            outputBox={"x": 0, "y": 0, "width": 600, "height": 120},
        )
    write(root / "capture.json", plan)
    run("crop", root / "capture.json", "--out", root / "crops")
    d = Document()
    p = d.add_paragraph("Class: ")
    p.add_run("___").underline = True
    cell = d.add_table(rows=1, cols=1).cell(0, 0)
    cell.text = "FIXED TEMPLATE TEXT"
    for text in ["heading", "caption", "analysis", "section", "question"]:
        cell.add_paragraph(text)
    template = root / "template.docx"
    d.save(template)
    config = {
        "template": str(template),
        "notebook": str(root / "execution/executed.ipynb"),
        "output": str(root / "report.docx"),
        "container": {"kind": "cell", "table": 0, "row": 0, "column": 0},
        "keep_prefix": 1,
        "cover_edits": [{"paragraph": 0, "run": 1, "expected": "___", "text": "TEST"}],
        "prototypes": {"heading": 1, "caption": 2, "analysis": 3, "section": 4, "question": 5},
        "screenshots": str(root / "crops/screenshots-cropped.json"),
        "max_width_cm": 14,
        "groups": [
            {"title": "1. Test", "items": [{"cell": 0, "analysis": "Actual executed output is 5."}]}
        ],
        "questions": [{"section": "Questions", "text": "What is the result?", "answer": "5"}],
    }
    write(root / "config.json", config)
    run("build", root / "config.json")
    run("audit", root / "config.json", "--out", root / "audit.json")
    subprocess.run(
        [
            sys.executable,
            str(scripts / "verify_evidence.py"),
            str(root / "report.docx"),
            "--expect",
            "代码截图=1",
            "--expect",
            "实验结果截图=1",
            "--strict",
        ],
        check=True,
        cwd=root,
    )
    assert Document(root / "report.docx").paragraphs[0].runs[1].underline
    assert nbformat.read(root / "execution/executed.ipynb", 4).cells[0].outputs[0].text == "5\n"
    print("PASS: standalone execution, crop, template fill, underlines, hashes and evidence counts")


if __name__ == "__main__":
    main()
