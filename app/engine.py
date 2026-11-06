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
LOGO = asset("nephrolife_logo.jpg")
if not LOGO.is_file():
    LOGO = asset("resources/nephrolife_logo.jpg")
MONTHS = ("январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь")
def checked_files(payload):
    quantity = Path(payload["quantity"])
    financial = Path(payload["financial"])
    if quantity.resolve() == financial.resolve() or any(p.suffix.lower() != ".xlsx" or not p.is_file() for p in (quantity, financial)):
        raise ValueError("Выберите две разные существующие книги .xlsx.")
    return {"quantity": quantity, "financial": financial}


def group_for(code):
    try:
        prefix = int(str(code).strip().split(".")[0])
    except (TypeError, ValueError):
        return "Прочие статьи"
    if prefix < 200:
        return "Статьи до 200"
    if prefix == 211: return "Лекарственные препараты"
    if prefix == 212: return "Изделия медицинского назначения"
    if prefix == 213: return "Расходные медицинские материалы"
    if prefix >= 214: return "Статьи от 214 и выше"
    return "Прочие статьи"


APP_CELL_IDS = (
    "papagun-001", "4c50e724", "papagun-007", "papagun-010", "papagun-011",
    "f3520855", "e9a9c483", "papagun-assembly",
)
REQUIRED_NAMES = (
    "read_branch_index", "build_branches_dataframe", "_drug_models_by_key", "render_drug_models",
    "combine_created_slides", "quarter_execution_info", "drug_execution_text", "_center_key",
    "numeric_chart", "new_presentation", "presentation_bytes", "text_box", "set_plot_layout",
    "OxmlElement", "STYLE", "RGBColor", "_number", "_procedure_number", "_counted",
    "add_header", "add_summary",
)
ROOT_CHECK = "if not (ROOT / 'excel_loading.ipynb').is_file():"
SELECTION_ERROR = "Список месяцев или филиалов содержит неизвестные или повторяющиеся значения."


def notebook_sources():
    cells = json.loads(NOTEBOOK.read_text(encoding="utf-8"))["cells"]
    found = {}
    for cell in cells:
        cid = cell.get("id")
        if cid not in APP_CELL_IDS:
            continue
        if cid in found:
            raise ValueError(f"Повторяющийся идентификатор ячейки {cid}.")
        found[cid] = "".join(cell["source"])
    missing = [cid for cid in APP_CELL_IDS if cid not in found]
    if missing:
        raise ValueError("В notebook нет ячеек: " + ", ".join(missing))
    return found


def exec_notebook_cell(ns, cid, source):
    if cid == "4c50e724":
        if source.count(ROOT_CHECK) != 1:
            raise ValueError("Не найдена проверка корня notebook.")
        source = source.replace(ROOT_CHECK, "if False:", 1)
    exec(compile(source, str(NOTEBOOK), "exec"), ns)


def require_names(ns):
    missing = [name for name in REQUIRED_NAMES if name not in ns]
    if missing:
        raise ValueError("Notebook не определил: " + ", ".join(missing))


def bind_center_map(ns, branches):
    def current_center_map():
        mapping = {}
        for branch in branches:
            key = ns["_center_key"](branch)
            if key in mapping and mapping[key] != branch:
                raise ValueError(f"Названия филиалов не различаются после нормализации: {mapping[key]} и {branch}.")
            mapping[key] = branch
        return mapping

    ns["slide5_center_map"] = current_center_map


def bind_execution_text(ns):
    original_drug_execution_text = ns["drug_execution_text"]
    original_quarter_execution_info = ns["quarter_execution_info"]

    def app_drug_execution_text(norm, actual):
        return "0%" if norm == 0 else original_drug_execution_text(norm, actual)

    def app_quarter_execution_info(norm, actual):
        return ("0%", "000000") if norm == 0 else original_quarter_execution_info(norm, actual)

    ns["drug_execution_text"] = app_drug_execution_text
    ns["quarter_execution_info"] = app_quarter_execution_info


def prepare_namespace():
    """Load notebook functions and stop before opening Excel when a name is missing."""
    sources = notebook_sources()
    ns = {"__name__": "papagun_notebook", "branch_maps": {"quantity": {}}}
    for cid in APP_CELL_IDS:
        exec_notebook_cell(ns, cid, sources[cid])
    require_names(ns)
    ns["LOGO"] = LOGO
    bind_execution_text(ns)
    return ns


def common_branches(ns, files):
    branch_maps = {key: ns["read_branch_index"](path) for key, path in files.items()}
    branches = [name for name in branch_maps["quantity"] if name in branch_maps["financial"]]
    if not branches:
        raise ValueError("В обеих книгах нет общих филиалов из листа Филиалы.")
    ns["branch_maps"] = branch_maps
    return branches


def read_frames(ns, files, branches):
    import pandas as pd

    ns["BRANCH_SHEETS"] = list(branches)
    ns["excel_files"] = files
    ns["workbooks"] = {
        key: pd.read_excel(path, sheet_name=[ns["branch_maps"][key][branch] for branch in branches],
                           header=None, dtype=object, keep_default_na=False, engine="openpyxl")
        for key, path in files.items()
    }
    ns["REPORT_PERIODS"] = MONTHS
    bind_center_map(ns, branches)
    return {book: ns["build_branches_dataframe"](book, periods=MONTHS)[:2] for book in files}


def narrow_frames(frame, provenance, branches, months):
    chosen = frame.index[frame["Наименование филиала"].isin(branches) & frame["Месяц"].isin(months)]
    return frame.loc[chosen], provenance.loc[chosen]


def catalog_from(branches, frames):
    import pandas as pd

    quantity_df = frames["quantity"][0]
    financial_df = frames["financial"][0]
    month_sets = []
    for frame in (quantity_df, financial_df):
        facts = frame.loc[frame["Факт по 1С"].map(
            lambda value: isinstance(value, (int, float)) and not isinstance(value, bool))]
        month_sets.append(set(facts["Месяц"]))

    finance_names_by_code = {}
    for code_value, name_value in financial_df[["Код статей", "Наименование статей"]].itertuples(index=False, name=None):
        if isinstance(code_value, str) and not pd.isna(name_value):
            finance_names_by_code.setdefault(code_value.strip(), set()).add(name_value)

    items = {}
    for code_value, name_value in quantity_df[["Код статей", "Наименование статей"]].drop_duplicates().itertuples(index=False, name=None):
        if code_value is None or name_value is None:
            continue
        code, name = str(code_value).strip(), str(name_value).strip()
        if not code or not name or not re.fullmatch(r"\d+(?:\.\d+)*\.?", code):
            continue
        group = group_for(code)
        if group not in ("Лекарственные препараты", "Изделия медицинского назначения",
                         "Расходные медицинские материалы"):
            continue
        finance_names = finance_names_by_code.get(code, set())
        if not finance_names or (name not in finance_names and len(finance_names) != 1):
            continue
        key = (code, name)
        items[key] = {"id": "\u001f".join(key), "code": code, "name": name, "group": group}

    ordered = sorted(items.values(), key=lambda item: (item["group"], item["code"], item["name"]))
    report_branches = [name for name in branches if name not in ("Атырау Нурсая2", "Атырау Нурсая 2")]
    return {"months": MONTHS, "availableMonths": [month for month in MONTHS
            if all(month in month_set for month_set in month_sets)],
            "branches": branches, "availableBranches": report_branches, "articles": ordered}


def has_default_components(frame, branches, months, codes):
    selected = frame.loc[frame["Наименование филиала"].isin(branches) & frame["Месяц"].isin(months)]
    found = {value.strip() if isinstance(value, str) else str(value)
             for value in selected["Код статей"].dropna()}
    return bool(found.intersection(codes))


def inspect(payload):
    files = checked_files(payload)
    ns = prepare_namespace()
    branches = common_branches(ns, files)
    return catalog_from(branches, read_frames(ns, files, branches))


def financial_selector(frame, code, quantity_name, branches, months):
    """Use the quantity key; accept a different finance label only for an unambiguous code."""
    codes = frame["Код статей"].map(lambda value: value.strip() if isinstance(value, str) else value)
    matches = frame.loc[codes.eq(code)
                        & frame["Наименование филиала"].isin(branches)
                        & frame["Месяц"].isin(months)]
    names = set(matches["Наименование статей"].dropna())
    if not names:
        raise ValueError(f"Код {code} отсутствует в финансовой книге.")
    if quantity_name in names:
        return (code, quantity_name)
    if len(names) == 1:
        return code
    raise ValueError(f"Код {code} имеет несколько финансовых наименований; сопоставление с «{quantity_name}» неоднозначно.")


def article_groups(items):
    """Keep categories independent; use a three-item last page for odd counts."""
    count = len(items)
    blocks = []
    index = 0
    while count - index > 3:
        blocks.append(items[index:index + 2])
        index += 2
    if index < count:
        blocks.append(items[index:])
    return blocks


def period_caption(months):
    indexes = [MONTHS.index(month) for month in months]
    if len(indexes) == 12 and indexes == list(range(12)):
        return "за год"
    if len(indexes) == 6 and indexes in (list(range(6)), list(range(6, 12))):
        return "за полугодие"
    if len(indexes) == 3 and indexes[0] % 3 == 0 and indexes == list(range(indexes[0], indexes[0] + 3)):
        return "за квартал"
    if len(months) == 2:
        return " и ".join(months)
    if indexes == list(range(indexes[0], indexes[-1] + 1)):
        return months[0] if len(months) == 1 else f"{months[0]}–{months[-1]}"
    return ", ".join(months)


def total_period_caption(months):
    indexes = [MONTHS.index(month) for month in months]
    if len(indexes) == 1:
        return months[0]
    if indexes == list(range(12)):
        return "год"
    if len(indexes) == 6 and indexes in (list(range(6)), list(range(6, 12))):
        return f"{1 if indexes[0] == 0 else 2} полугодие"
    if len(indexes) == 3 and indexes[0] % 3 == 0 and indexes == list(range(indexes[0], indexes[0] + 3)):
        return f"{indexes[0] // 3 + 1} квартал"
    count = len(indexes)
    form = "месяца" if count in (2, 3, 4) else "месяцев"
    return f"{count} {form}" if indexes == list(range(count)) else f"{count} выбранных {form}"


def combined_model(ns, frame, provenance, branches, months, book, item):
    from decimal import Decimal

    components = item["components"]
    selectors = item["selectors"][book]
    raw = ns["_drug_models_by_key"](
        frame, provenance, list(branches), months=months, book=book,
        first_metric="Утверждено" if book == "quantity" else "Факт по норме",
        divisor=Decimal(1 if book == "quantity" else 1000), articles=selectors,
        missing_as_zero=True,
    )
    if len(components) == 1:
        return raw[0]
    result = []
    for members in zip(*raw):
        if len({(row["label"], row["branch"]) for row in members}) != 1:
            raise ValueError("Компоненты группы относятся к разным филиалам.")
        row = dict(label=members[0]["label"], branch=members[0]["branch"],
                   article=item["title"], article_code=item["key"],
                   component_keys=tuple((r["article_code"], r["article"]) for r in members),
                   source_indices=tuple(i for r in members for i in r["source_indices"]))
        for metric in ("plan", "actual", "norm"):
            row[metric] = sum((r[metric] for r in members), Decimal(0))
        row["difference"] = row["norm"] - row["actual"]
        result.append(row)
    return result


def quarter_models(ns, frame, provenance, branches, months, book, items):
    from decimal import Decimal

    # Group the chosen months by calendar quarter; never create a zero for an unchosen quarter.
    selected_months = set(months)
    periods = [(f"{q + 1} квартал", tuple(m for m in MONTHS[q * 3:q * 3 + 3] if m in selected_months))
               for q in range(4)]
    periods = [(label, part) for label, part in periods if part]
    result = []
    for item in items:
        rows = []
        for label, period_months in periods:
            centers = combined_model(ns, frame, provenance, branches, period_months, book, item)
            norm = sum((r["norm"] for r in centers), Decimal(0))
            actual = sum((r["actual"] for r in centers), Decimal(0))
            rows.append(dict(label=label, article=item["title"], article_code=item["key"],
                             plan=norm, norm=norm, actual=actual, difference=norm - actual))
        norm = sum((r["norm"] for r in rows), Decimal(0))
        actual = sum((r["actual"] for r in rows), Decimal(0))
        total_label = "9 месяцев" if tuple(months) == MONTHS[:9] else ("12 месяцев" if tuple(months) == MONTHS else "Итого")
        rows.append(dict(label=total_label, article=item["title"], article_code=item["key"],
                         plan=norm, norm=norm, actual=actual, difference=norm - actual))
        result.append(rows)
    return result


def render_quarters(ns, models, title):
    deck = ns["new_presentation"]()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    ns["add_header"](slide, title, width=11.8)
    count = len(models)
    for index, rows in enumerate(models):
        quarter_count = len(rows) - 1
        if not 1 <= quarter_count <= 4:
            raise ValueError("Квартальный слайд требует от одного до четырёх кварталов и итог.")
        slots = [None] * (4 - quarter_count) + rows[:-1] + [rows[-1]]
        if count == 4:
            top, height = .65 + index * 1.67, 1.63
        elif count == 3:
            top, height = .65 + index * 2.16, 2.04
        else:
            top, height = ((.65, 3.10) if index == 0 else (3.88, 3.12))
        series = [(name, [r[key] if r is not None else None for r in slots], color) for name, key, color in
                  zip(("Факт по норме", "Факт по 1С", "отк:"), ("norm", "actual", "difference"),
                      (ns["STYLE"]["teal"], ns["STYLE"]["gray"], ns["STYLE"]["gray"]))]
        chart = ns["numeric_chart"](slide, [r["label"] if r is not None else "" for r in slots], series,
                                      .15, top, 12.99, height, fmt="#,##0", compact_base=5,
                                     drug=True, total=True)
        if count == 4:
            ns["set_plot_layout"](chart.chart, x=.01, y=.30, w=.98, h=.55)
        elif count == 3:
            ns["set_plot_layout"](chart.chart, x=.01, y=.38, w=.98, h=.48)
        else:
            ns["set_plot_layout"](chart.chart, x=.01, y=.28, w=.98, h=.62)
        ns["text_box"](slide, rows[0]["article"], .25, top + .025, 12.79, .22,
                       size=11, bold=True, color="0070C0", name="quarter_article")
        for j, row in enumerate(rows):
            slot = 4 - quarter_count + j if j < quarter_count else 4
            center = .15 + 12.99 * (.01 + (slot + .5) * .98 / 5)
            percent, color = ns["quarter_execution_info"](row["norm"], row["actual"])
            panel = ns["text_box"](slide, percent, center - .325, top + .30,
                                   .65, .18, size=10, bold=True, color=color, fill=ns["STYLE"]["gray"])
            panel.line.color.rgb = ns["RGBColor"](0, 0, 0)
    return ns["presentation_bytes"](deck)


def statistic_slide(ns, frame, provenance, branches, labels, months, row_number, code, title, forms):
    from decimal import Decimal
    import pandas as pd

    def number_or_zero(value, number):
        if pd.isna(value) or isinstance(value, (str, bool)):
            return Decimal(0)
        return number(value)

    period = tuple(months)
    model = []
    for branch, label in zip(branches, labels):
        selected = frame.loc[frame["Наименование филиала"].eq(branch)
                             & frame["Месяц"].isin(period)
                             & provenance["Строка Excel"].eq(row_number)]
        if (len(selected) != len(period) or set(selected["Месяц"]) != set(period)
                or selected["Месяц"].duplicated().any()
                or not selected["Код статей"].map(lambda value: str(value).strip()).eq(code).all()
                or not provenance.loc[selected.index, "Лист"].eq(ns["branch_maps"]["quantity"][branch]).all()):
            raise ValueError(f"Нет однозначных данных показателя «{title}» для филиала «{branch}».")
        number = ns["_procedure_number"] if row_number == 15 else ns["_number"]
        plan = sum((number_or_zero(value, number) for value in selected["Утверждено"]), Decimal(0))
        actual = sum((number_or_zero(value, number) for value in selected["Факт по 1С"]), Decimal(0))
        model.append(dict(label=label, plan=plan, actual=actual))
    deck = ns["new_presentation"]()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    ns["add_header"](slide, f"{title}: {period_caption(period)} 2026 г.", width=11.8)
    plan = sum((r["plan"] for r in model), Decimal(0))
    actual = sum((r["actual"] for r in model), Decimal(0))
    ns["add_summary"](slide, [(f"{len(model)} центров", "000000"),
                               ("План: " + ns["_counted"](plan, forms), "000000"),
                               ("Факт: " + ns["_counted"](actual, forms), "000000"),
                               ("Отклонение: " + ns["_counted"](actual - plan, forms), "000000")])
    ns["numeric_chart"](slide, [r["label"] for r in model],
                         [("План", [r["plan"] for r in model], ns["STYLE"]["teal"]),
                          ("Факт", [r["actual"] for r in model], ns["STYLE"]["gray"])],
                         0, 1.85, 13 + 1 / 3, 5.5, fmt="General", compact_base=len(model))
    return ns["presentation_bytes"](deck)


SHARE_SPECS = (
    ("Лекарственные препараты", "211."),
    ("Изделия медицинского назначения", "212."),
    ("Расходные медицинские материалы", "213."),
    ("Диагностирование пациентов", ("214.00",)),
    ("Утилизация медицинских отходов", ("219.00",)),
    ("Затраты на питание пациентов", ("222.00",)),
    ("Транспортные расходы", ("223.01", "223.02", "223.03", "223.04")),
    ("Командировочные расходы", ("224.01", "224.02", "224.03")),
    ("Коммунальные услуги", ("228.01", "228.02", "228.03", "228.04", "228.05")),
    ("Аренда помещения", ("232.00",)),
    ("Общий фонд оплаты труда", ("235.00",)),
    ("Оплата труда", ("235.01",)),
)


def financial_share_models(ns, frame, provenance, branches, months, *, average_branches=None):
    """Use the full reference set for means and selected branches for columns."""
    from decimal import Decimal, InvalidOperation
    import pandas as pd

    average_branches = tuple(branches if average_branches is None else average_branches)
    if not average_branches:
        raise ValueError("Для средних Плана и Факта нужен хотя бы один филиал.")
    calculation_branches = tuple(dict.fromkeys((*branches, *average_branches)))
    month_set = set(months)
    branch_set = set(calculation_branches)
    selected = frame.loc[frame["Наименование филиала"].isin(branch_set)
                         & frame["Месяц"].isin(month_set)]
    by_period = {(branch, month): [] for branch in calculation_branches for month in months}
    expected_sheets = ns["branch_maps"]["financial"]
    for index, row in selected.iterrows():
        branch, month = row["Наименование филиала"], row["Месяц"]
        if provenance.at[index, "Лист"] != expected_sheets[branch]:
            raise ValueError(f"Неверный лист финансового показателя для филиала «{branch}».")
        code = str(row["Код статей"]).strip() if pd.notna(row["Код статей"]) else ""
        by_period[(branch, month)].append((code, str(row["Наименование статей"]).strip(),
                                           row["Утверждено"], row["Факт по 1С"]))

    def amount(value, branch, month, title, metric):
        if pd.isna(value) or (isinstance(value, str) and not value.strip()):
            return None
        try:
            number = Decimal(str(value))
        except InvalidOperation as error:
            raise ValueError(f"Нечисловое значение «{metric}» статьи «{title}» для филиала «{branch}», месяц «{month}».") from error
        if not number.is_finite():
            raise ValueError(f"Неконечное значение «{metric}» статьи «{title}» для филиала «{branch}», месяц «{month}».")
        return number

    models = []
    has_realization = any(name == "ВСЕГО РЕАЛИЗАЦИЯ"
                          for entries in by_period.values() for _, name, _, _ in entries)
    if not has_realization:
        return models
    for title, selector in SHARE_SPECS:
        has_components = any(
            code.startswith(selector) if isinstance(selector, str) else code in selector
            for branch in branches for month in months
            for code, _, _, _ in by_period[(branch, month)]
        )
        if not has_components:
            continue
        rows = []
        for branch in calculation_branches:
            numerators = {"plan": Decimal(0), "fact": Decimal(0)}
            denominators = {"plan": Decimal(0), "fact": Decimal(0)}
            for month in months:
                entries = by_period[(branch, month)]
                for _, name, plan_value, fact_value in entries:
                    if name != "ВСЕГО РЕАЛИЗАЦИЯ":
                        continue
                    for metric, value in (("plan", plan_value), ("fact", fact_value)):
                        denominators[metric] += amount(value, branch, month, "ВСЕГО РЕАЛИЗАЦИЯ",
                                                       "Утверждено" if metric == "plan" else "Факт по 1С") or Decimal(0)
                if isinstance(selector, str):
                    values = [entry for entry in entries if entry[0].startswith(selector)]
                else:
                    values = [entry for entry in entries if entry[0] in selector]
                for _, _, plan_value, fact_value in values:
                    for metric, value in (("plan", plan_value), ("fact", fact_value)):
                        numerators[metric] += amount(value, branch, month, title,
                                                      "Утверждено" if metric == "plan" else "Факт по 1С") or Decimal(0)
            plan_percent = numerators["plan"] / denominators["plan"] * 100 if denominators["plan"] else Decimal(0)
            fact_percent = numerators["fact"] / denominators["fact"] * 100 if denominators["fact"] else Decimal(0)
            rows.append({"branch": branch, "label": branch, "numerator": numerators["fact"],
                         "realization": denominators["fact"], "percent": fact_percent,
                         "plan_percent": plan_percent, "plan_numerator": numerators["plan"],
                         "plan_realization": denominators["plan"]})
        rows_by_branch = {row["branch"]: row for row in rows}
        plan_numerator = sum((rows_by_branch[branch]["plan_numerator"] for branch in average_branches), Decimal(0))
        plan_realization = sum((rows_by_branch[branch]["plan_realization"] for branch in average_branches), Decimal(0))
        fact_numerator = sum((rows_by_branch[branch]["numerator"] for branch in average_branches), Decimal(0))
        fact_realization = sum((rows_by_branch[branch]["realization"] for branch in average_branches), Decimal(0))
        plan_reference = plan_numerator / plan_realization * 100 if plan_realization else Decimal(0)
        fact_reference = fact_numerator / fact_realization * 100 if fact_realization else Decimal(0)
        displayed_rows = sorted((rows_by_branch[branch] for branch in branches),
                                key=lambda row: (-row["percent"], row["branch"]))
        models.append((title, displayed_rows,
                       plan_reference, fact_reference))
    return models


def financial_share_slides(ns, models):
    from pptx.enum.chart import XL_DATA_LABEL_POSITION
    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    result = []
    for first in range(0, len(models), 3):
        deck = ns["new_presentation"]()
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        for position, (title, rows, plan_reference, fact_reference) in enumerate(models[first:first + 3]):
            top = .40 + position * 2.33
            categories = ["План", "Факт ср."] + [row["label"] for row in rows]
            values = [plan_reference, fact_reference] + [row["percent"] for row in rows]
            chart = ns["numeric_chart"](
                slide, categories,
                [("Доля от всего реализация", values, ns["STYLE"]["teal"])],
                .12, top + .24, 13.05, 1.84, fmt='0.0', compact_base=21, compact_labels=True,
            )
            fraction = .96 * min(1, max(.25, len(rows) / 19))
            ns["set_plot_layout"](chart.chart, x=(1 - fraction) / 2, y=.18, w=fraction, h=.34)
            chart.chart.category_axis.tick_labels.font.size = Pt(6)
            axis = chart.chart.category_axis._element
            body = axis.get_or_add_txPr().xpath('./a:bodyPr')[0]
            body.set('rot', '-2700000')
            body.set('wrap', 'none')
            tick_skip = axis.xpath('./c:tickLblSkip')
            if not tick_skip:
                tick_skip = [ns["OxmlElement"]('c:tickLblSkip')]
                axis.append(tick_skip[0])
            tick_skip[0].set('val', '1')
            series = chart.chart.series[0]
            series.data_labels.position = XL_DATA_LABEL_POSITION.CENTER
            series.data_labels.font.size = Pt(11)
            series.data_labels.font.bold = True
            plot_height_pt = 1.84 * .34 * 72
            axis_maximum = chart.chart.value_axis.maximum_scale
            axis_minimum = chart.chart.value_axis.minimum_scale
            for index, point in enumerate(series.points):
                label = point.data_label
                label.position = XL_DATA_LABEL_POSITION.CENTER
                label.font.name = ns["STYLE"]["font"]
                label.font.size = Pt(11)
                label.font.bold = True
                color = "FFFF00" if index < 2 else ("FFFFFF" if rows[index - 2]["percent"] > fact_reference else "000000")
                # Keep every value visible when its column cannot hold an 11 pt label.
                short_column = abs(float(values[index])) / (axis_maximum - axis_minimum) * plot_height_pt < 14
                if short_column:
                    label.position = XL_DATA_LABEL_POSITION.OUTSIDE_END
                    if color == "FFFFFF":
                        color = "666666"
                label.font.color.rgb = RGBColor.from_string(color)
                for label_body in series._element.xpath(f'./c:dLbls/c:dLbl[c:idx/@val="{index}"]/c:txPr/a:bodyPr'):
                    label_body.set('rot', '0')
                    label_body.set('wrap', 'none')
            ns["text_box"](slide, title, .36, top + .02, 3.95, .29,
                           size=13, bold=True, color="000000", fill="ADADAD", name="Показатель доли")
        result.append(ns["presentation_bytes"](deck))
    return result


def save_pptx(data, output):
    """Save atomically; keep an existing report when Windows denies its replacement."""
    fd, temp_path = tempfile.mkstemp(suffix=".pptx", dir=output.parent)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
        try:
            os.replace(temp_path, output)
            return output
        except PermissionError:
            if not output.is_file():
                raise
            for number in range(2, 1002):
                alternative = output.with_name(f"{output.stem} ({number}){output.suffix}")
                try:
                    os.rename(temp_path, alternative)
                    return alternative
                except FileExistsError:
                    continue
            raise FileExistsError("Не найдено свободное имя для нового отчёта в выбранной папке.")
    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)


def generate(payload):
    import hashlib
    from io import BytesIO
    from zipfile import ZipFile
    from pptx import Presentation

    files = checked_files(payload)
    requested_months = payload["months"]
    requested_branches = payload["branches"]
    articles = payload.get("articles", [])
    custom_groups = payload.get("groups", [])
    months = tuple(month for month in MONTHS if month in requested_months)
    if not months or not requested_branches:
        raise ValueError("Выберите хотя бы один месяц и филиал.")
    if len(requested_months) != len(set(requested_months)) or len(months) != len(set(requested_months)):
        raise ValueError(SELECTION_ERROR)
    output = Path(payload["output"])
    if output.suffix.lower() != ".pptx" or output.resolve() in (p.resolve() for p in files.values()):
        raise ValueError("Укажите выходной файл .pptx, отличный от исходных книг.")
    before = {key: hashlib.sha256(path.read_bytes()).digest() for key, path in files.items()}
    ns = prepare_namespace()
    common = common_branches(ns, files)
    branches = tuple(branch for branch in ns["branch_maps"]["quantity"] if branch in requested_branches)
    if not branches:
        raise ValueError("Выберите хотя бы один месяц и филиал.")
    if (len(requested_branches) != len(set(requested_branches))
            or len(branches) != len(set(requested_branches))):
        raise ValueError(SELECTION_ERROR)
    full_frames = read_frames(ns, files, common)
    available_catalog = catalog_from(common, full_frames)
    if any(month not in available_catalog["availableMonths"] for month in months) or any(
            branch not in available_catalog["availableBranches"] for branch in branches):
        raise ValueError("Выбран месяц или филиал без полного набора данных для этого отчёта.")
    bind_center_map(ns, branches)
    frames = {book: narrow_frames(*full_frames[book], branches, months) for book in files}
    catalog = {item["id"]: item for item in available_catalog["articles"]}
    used = set()

    def parse(raw):
        if raw not in catalog:
            raise ValueError("Выбрана статья, которой нет в количественной книге.")
        parts = raw.split("\u001f")
        if len(parts) != 2 or not all(parts):
            raise ValueError("Некорректная статья в выборе.")
        if raw in used:
            raise ValueError("Статья используется повторно в выборе или группе.")
        used.add(raw)
        if catalog[raw]["group"] not in ("Лекарственные препараты", "Изделия медицинского назначения", "Расходные медицинские материалы"):
            raise ValueError("Отдельные слайды строятся только для лекарств, медицинских изделий и расходных материалов.")
        if catalog[raw]["code"] in ("211", "212", "213"):
            raise ValueError("Выберите исходные статьи, а не итоговую строку раздела.")
        return tuple(parts), catalog[raw]["group"]

    items = []
    for raw in articles:
        pair, category = parse(raw)
        items.append(dict(title=pair[1], key=pair[0], category=category, components=[pair]))
    for entry in custom_groups:
        title = str(entry.get("title", "")).strip()
        members = entry.get("articles", [])
        if not title or len(members) < 2 or len(members) != len(set(members)):
            raise ValueError("Группе нужны заголовок и хотя бы две разные статьи.")
        parsed = [parse(raw) for raw in members]
        categories = {category for _, category in parsed}
        if len(categories) != 1:
            raise ValueError("В одной группе можно объединять статьи только из одного раздела.")
        components = [pair for pair, _ in parsed]
        items.append(dict(title=title, key=" + ".join(code for code, _ in components),
                          category=categories.pop(), components=components))
    financial_used = []
    for item in items:
        item["selectors"] = {"quantity": item["components"],
                             "financial": [financial_selector(frames["financial"][0], code, name, branches, months)
                                           for code, name in item["components"]]}
        financial_used.extend(item["selectors"]["financial"])
    if len(set(financial_used)) != len(financial_used):
        raise ValueError("Несколько выбранных количественных позиций соответствуют одной финансовой статье. Сопоставление неоднозначно.")
    payloads = []
    quantity, quantity_provenance = frames["quantity"]
    labels = branches
    for row, code, title, forms in ((13, "2", "АППАРАТЫ", ("аппарат", "аппарата", "аппаратов")),
                                    (14, "3", "ПАЦИЕНТЫ", ("пациент", "пациента", "пациентов")),
                                    (15, "4", "ПРОЦЕДУРЫ", ("процедура", "процедуры", "процедур"))):
        payloads.append(statistic_slide(ns, quantity, quantity_provenance, branches, labels, months, row, code, title, forms))
    category_titles = {"Лекарственные препараты": "ЛС", "Изделия медицинского назначения": "ИМН",
                       "Расходные медицинские материалы": "РАСХОДНЫЕ МАТЕРИАЛЫ"}
    for category in category_titles:
        for block in article_groups([item for item in items if item["category"] == category]):
            for book in ("quantity", "financial"):
                frame, provenance = frames[book]
                models = [combined_model(ns, frame, provenance, labels, months, book, item) for item in block]
                selectors = [(item["key"], rows[0]["article"]) for item, rows in zip(block, models)]
                heading = ("РАСХОДЫ " + category_titles[category] + ": " + period_caption(months) + " 2026 г."
                           + (" (количество)" if book == "quantity" else " (тыс. тенге)"))
                payloads.append(ns["render_drug_models"](
                    models, heading=heading, series_names=("План", "Факт", "отк:") if book == "quantity"
                    else ("Факт по норме", "Факт по 1С", "отк:"), financial=book == "financial",
                    articles=selectors, sort_by_difference=True, total_period=total_period_caption(months)))
            models = quarter_models(ns, quantity, quantity_provenance, labels, months, "quantity", block)
            payloads.append(render_quarters(ns, models, "РАСХОДЫ " + category_titles[category] + ": по кварталам, " + period_caption(months) + " 2026 г. (количество)"))

    financial, financial_provenance = frames["financial"]
    expense_items = [dict(title=label, key=" + ".join(codes), components=[(code, "") for code in codes],
                          selectors={"financial": list(codes)}) for label, codes in (
                              ("Транспортные расходы", ("223.01", "223.02", "223.03", "223.04")),
                              ("Расходы на питание", ("222.00", "222.01")))
                     if has_default_components(financial, branches, months, codes)]
    expense_models = [combined_model(ns, financial, financial_provenance, labels, months, "financial", item)
                      for item in expense_items]
    if expense_models:
        payloads.append(ns["render_drug_models"](
            expense_models, heading="РАСХОДЫ: " + period_caption(months) + " 2026 г. (тыс. тенге)",
            series_names=("Факт по норме", "Факт по 1С", "отк:"), financial=True,
            articles=[(item["key"], item["title"]) for item in expense_items], sort_by_difference=True,
            total_period=total_period_caption(months)))
    quarter_items = [dict(title=label, key=" + ".join(codes), components=[(code, "") for code in codes],
                          selectors={"financial": list(codes)}) for label, codes in (
                              ("Расходы на транспорт", ("223.01", "223.02")),
                              ("Возмещение транспортных расходов пациентам", ("223.04",)),
                              ("Расходы на питание", ("222.00",)),
                              ("Возмещение расходов на питание", ("222.01",)))
                     if has_default_components(financial, branches, months, codes)]
    if quarter_items:
        payloads.append(render_quarters(ns, quarter_models(ns, financial, financial_provenance, labels, months,
                                                           "financial", quarter_items),
                                        "ФИНАНСОВЫЕ ПОКАЗАТЕЛИ: по кварталам, " + period_caption(months) + " 2026 г. (тыс. тенге)"))
    payloads.extend(financial_share_slides(
        ns, financial_share_models(ns, *full_frames["financial"], branches, months,
                                   average_branches=common)))
    result = ns["combine_created_slides"](payloads)
    with ZipFile(BytesIO(result)) as archive:
        if archive.testzip() is not None:
            raise ValueError("Созданный PPTX повреждён.")
    if len(Presentation(BytesIO(result)).slides) != len(payloads):
        raise ValueError("Количество слайдов не совпало с выбранными статьями.")
    for key, path in files.items():
        if hashlib.sha256(path.read_bytes()).digest() != before[key]:
            raise ValueError("Исходная книга изменилась во время обработки.")
    output.parent.mkdir(parents=True, exist_ok=True)
    saved_output = save_pptx(result, output)
    return {"output": str(saved_output), "alternateOutput": saved_output != output,
            "slides": len(payloads), "articles": len(used)}


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
