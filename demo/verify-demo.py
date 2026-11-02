"""Smoke check the packaged demo worker against the local NL books."""

import hashlib
import json
import subprocess
from pathlib import Path

from pptx import Presentation

from watermark import NOTICE, WATERMARK_ANGLE, WATERMARK_COUNT


ROOT = Path(__file__).resolve().parent.parent
DEMO = ROOT / "demo"
WORKER = DEMO / "dist" / "papagun-demo-engine.exe"
OUTPUT = DEMO / ".build" / "demo-qa.pptx"
BOOKS = {
    "quantity": ROOT / "Исходники" / "NL_Кол. Бюджет.xlsx",
    "financial": ROOT / "Исходники" / "NL_Фин. Бюджет.xlsx",
}


def call(action, payload):
    result = subprocess.run(
        [str(WORKER), action], input=json.dumps(payload, ensure_ascii=False),
        text=True, encoding="utf-8", capture_output=True, check=False,
    )
    response = json.loads(result.stdout)
    if result.returncode:
        raise RuntimeError(response.get("error", "Demo worker failed"))
    return response


def main():
    assert WORKER.is_file()
    before = {name: hashlib.sha256(path.read_bytes()).digest() for name, path in BOOKS.items()}
    request = {name: str(path) for name, path in BOOKS.items()}
    catalog = call("inspect", request)
    assert len(catalog["availableMonths"]) >= 2 and catalog["availableBranches"]
    months = catalog["availableMonths"][:2]
    article = next(item["id"] for item in catalog["articles"]
                   if item["group"] == "Лекарственные препараты" and item["code"] != "211")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    response = call("generate", {
        **request,
        "months": months,
        "branches": catalog["availableBranches"][:1],
        "articles": [article], "groups": [], "output": str(OUTPUT),
    })
    presentation = Presentation(OUTPUT)
    assert len(presentation.slides) == response["slides"] == 12
    for slide in list(presentation.slides)[-4:]:
        charts = [shape.chart for shape in slide.shapes if shape.has_chart]
        assert len(charts) == 3, "Expected three expense shares per new slide"
        for chart in charts:
            values = list(chart.series[0].values)
            assert values == sorted(values, reverse=True), "Expense shares are not sorted"
    for slide in (presentation.slides[0], presentation.slides[1]):
        title = next(shape.text for shape in slide.shapes if shape.has_text_frame
                     and ("АППАРАТЫ:" in shape.text or "ПАЦИЕНТЫ:" in shape.text))
        assert all(month in title for month in months), "First slides ignore the selected period"
    for slide in presentation.slides:
        matches = [shape for shape in slide.shapes if shape.has_text_frame and shape.text == NOTICE]
        assert len(matches) == WATERMARK_COUNT, "Missing demo notices"
        assert matches == list(slide.shapes)[-WATERMARK_COUNT:], "Demo notices are not the top shapes"
        assert all(shape.rotation == WATERMARK_ANGLE for shape in matches)
        assert len({(shape.left, shape.top) for shape in matches}) == WATERMARK_COUNT
        assert all("\n" not in shape.text for shape in matches)
    from pptx.oxml.ns import qn
    verifier = presentation._element.find(qn("p:modifyVerifier"))
    assert verifier is not None and verifier.get("spinCount") == "100000"
    assert all(hashlib.sha256(path.read_bytes()).digest() == before[name] for name, path in BOOKS.items())
    print(f"PASS: {len(presentation.slides)} slides, twelve sorted expense charts, five diagonal notices per slide; write protection set; source books unchanged")


if __name__ == "__main__":
    main()
