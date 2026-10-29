"""Local report worker. The calculation and slide functions come from excel_loading.ipynb."""

import json
import os
import sys
import tempfile
import re
from pathlib import Path


def asset(name):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / name


NOTEBOOK = asset("excel_loading.ipynb")
CENTERS = asset("Соответствие названий центров в презентации и эксель Nephrolife.xlsx")
LOGO = asset("nephrolife_logo.jpg")
if not CENTERS.is_file():
    CENTERS = asset("Исходники/Соответствие названий центров в презентации и эксель Nephrolife.xlsx")
if not LOGO.is_file():
    LOGO = asset("resources/nephrolife_logo.jpg")
MONTHS = ("январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь")
BRANCHES = (
    "Актобе", "Актобе АМЦ", "Астана Биржан Сал", "Астана Сарыарка", "Атырау", "Атырау Нурсая",
    "Атырау Нурсая 2", "Караганда ЦАД", "Караганда Строителей", "Кентау", "Костанай", "Лисаковск",
    "Рудный", "Павлодар Дюсенова", "Павлодар Павлова", "Экибастуз", "Петропавловск",
    "Усть-Каменогорск (NL UK)", "Шымкент", "Шымкент Абая",
)


def checked_files(payload):
    quantity = Path(payload["quantity"])
    financial = Path(payload["financial"])
    if quantity.resolve() == financial.resolve() or any(p.suffix.lower() != ".xlsx" or not p.is_file() for p in (quantity, financial)):
        raise ValueError("Выберите две разные существующие книги .xlsx.")
    return {"quantity": quantity, "financial": financial}


def group_for(code):
    try:
        n = float(code)
    except (TypeError, ValueError):
        return "Прочие статьи"
    if n <= 200:
        return "Основные показатели"
    prefix = str(code).strip().split(".")[0]
    if prefix == "211": return "Лекарственные препараты"
    if prefix == "212": return "Изделия медицинского назначения"
    if prefix == "213": return "Расходные медицинские материалы"
    if 214 <= int(prefix) <= 219: return "Переменные косвенные затраты"
    if int(prefix) >= 220: return "Постоянные затраты"
    return "Прочие статьи"


def inspect(payload):
    from openpyxl import load_workbook

    files = checked_files(payload)
    items = {}
    available = None
    for book, path in files.items():
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            sheets = [name for name in BRANCHES if name in wb.sheetnames]
            if not sheets:
                raise ValueError(f"В книге {book} не найдены листы филиалов NL.")
            available = set(sheets) if available is None else available & set(sheets)
            limit = 141 if book == "quantity" else 193
            for branch in sheets:
                ws = wb[branch]
                if str(ws["B6"].value).strip().casefold() != "код статей":
                    raise ValueError(f"Не найден заголовок кодов: {book}, {branch}, B6.")
                for row in ws.iter_rows(min_row=9, max_row=limit, min_col=2, max_col=3, values_only=True):
                    code, name = row
                    if code is None or name is None or not str(code).strip() or not str(name).strip():
                        continue
                    code, name = str(code).strip(), str(name).strip()
                    if not re.fullmatch(r"\d+(?:\.\d+)*\.?", code):
                        continue
                    key = (book, code, name)
                    items[key] = {"id": "\u001f".join(key), "book": book, "code": code, "name": name, "group": group_for(code)}
        finally:
            wb.close()
    ordered = sorted(items.values(), key=lambda x: (x["group"], x["code"], x["name"], x["book"]))
    return {"months": MONTHS, "branches": [b for b in BRANCHES if b in available], "articles": ordered}


def notebook_namespace(files, branches, months):
    import pandas as pd

    cells = json.loads(NOTEBOOK.read_text(encoding="utf-8"))["cells"]
    ns = {"__name__": "papagun_notebook"}
    for index in (3, 5, 12):
        source = "".join(cells[index]["source"])
        if index == 5:
            source = source.replace("ROOT = Path.cwd()", "ROOT = Path.cwd()")
            source = source.replace("if not (ROOT / 'excel_loading.ipynb').is_file():", "if False:")
        exec(compile(source, str(NOTEBOOK), "exec"), ns)
    ns["CENTERS"] = CENTERS
    ns["LOGO"] = LOGO
    ns["BRANCH_SHEETS"] = list(branches)
    ns["excel_files"] = files
    ns["workbooks"] = {
        key: pd.read_excel(path, sheet_name=list(branches), header=None, dtype=object,
                           keep_default_na=False, engine="openpyxl")
        for key, path in files.items()
    }
    ns["REPORT_PERIODS"] = tuple(months)
    source = "".join(cells[13]["source"])
    source = source[:source.index("quantity_df, quantity_provenance")]
    exec(compile(source, str(NOTEBOOK), "exec"), ns)
    for index in (17, 19, 21):
        exec(compile("".join(cells[index]["source"]), str(NOTEBOOK), "exec"), ns)
    return ns


def generate(payload):
    import hashlib
    from io import BytesIO
    from zipfile import ZipFile
    from pptx import Presentation

    files = checked_files(payload)
    months = tuple(m for m in MONTHS if m in payload["months"])
    branches = tuple(b for b in BRANCHES if b in payload["branches"])
    articles = payload["articles"]
    if not months or not branches or not articles:
        raise ValueError("Выберите хотя бы один месяц, филиал и статью.")
    if len(months) != len(set(payload["months"])) or len(branches) != len(set(payload["branches"])):
        raise ValueError("Список месяцев или филиалов содержит неизвестные или повторяющиеся значения.")
    output = Path(payload["output"])
    if output.suffix.lower() != ".pptx" or output.resolve() in (p.resolve() for p in files.values()):
        raise ValueError("Укажите выходной файл .pptx, отличный от исходных книг.")
    output.parent.mkdir(parents=True, exist_ok=True)
    before = {key: hashlib.sha256(path.read_bytes()).digest() for key, path in files.items()}
    ns = notebook_namespace(files, branches, months)
    frames = {}
    for book in files:
        frames[book] = ns["build_branches_dataframe"](book, periods=months)[:2]
    selected = {"quantity": [], "financial": []}
    for raw in articles:
        book, code, name = raw.split("\u001f")
        if book not in selected or not code or not name:
            raise ValueError("Некорректная статья в выборе.")
        pair = (code, name)
        if pair in selected[book]:
            raise ValueError("Статья выбрана повторно.")
        selected[book].append(pair)
    payloads = []
    for book, choices in selected.items():
        frame, provenance = frames[book]
        for start in range(0, len(choices), 3):
            group = choices[start:start + 3]
            models = ns["_drug_models_by_key"](
                frame, provenance, list(branches), months=months, book=book,
                first_metric="Утверждено" if book == "quantity" else "Факт по норме",
                divisor=ns["Decimal"](1 if book == "quantity" else 1000), articles=group,
            )
            heading = ("КОЛИЧЕСТВЕННЫЙ БЮДЖЕТ" if book == "quantity" else "ФИНАНСОВЫЙ БЮДЖЕТ") + ": 2026 г."
            slide_bytes = ns["render_drug_models"](
                models, heading=heading,
                series_names=("План", "Факт", "отк:") if book == "quantity" else ("Факт по норме", "Факт по 1С", "отк:"),
                financial=book == "financial", articles=group,
            )
            deck = Presentation(BytesIO(slide_bytes))
            ns["text_box"](deck.slides[0], "Месяцы: " + ", ".join(months), .25, .53, 11.7, .16, size=8, color="475467")
            payloads.append(ns["presentation_bytes"](deck))
    result = ns["combine_created_slides"](payloads)
    with ZipFile(BytesIO(result)) as archive:
        if archive.testzip() is not None:
            raise ValueError("Созданный PPTX повреждён.")
    if len(Presentation(BytesIO(result)).slides) != len(payloads):
        raise ValueError("Количество слайдов не совпало с выбранными статьями.")
    for key, path in files.items():
        if hashlib.sha256(path.read_bytes()).digest() != before[key]:
            raise ValueError("Исходная книга изменилась во время обработки.")
    fd, temp_path = tempfile.mkstemp(suffix=".pptx", dir=output.parent)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(result)
        os.replace(temp_path, output)
    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)
    return {"output": str(output), "slides": len(payloads), "articles": len(articles)}


if __name__ == "__main__":
    sys.stdin.reconfigure(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        request = json.load(sys.stdin)
        response = inspect(request) if sys.argv[1] == "inspect" else generate(request)
        print(json.dumps(response, ensure_ascii=False))
    except Exception as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False))
        sys.exit(1)
