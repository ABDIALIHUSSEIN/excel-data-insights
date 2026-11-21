"""Spreadsheet ingestion, validation, and exports."""
from io import BytesIO
from datetime import date, datetime
import re
from zipfile import ZipFile

import pandas as pd

MAX_ROWS = 100_000
MAX_COLUMNS = 250
MAX_CELLS = 1_000_000


def inspect_workbook(content):
    with ZipFile(BytesIO(content)) as archive:
        if sum(part.file_size for part in archive.infolist()) > 120 * 1024**2:
            raise ValueError("Workbook exceeds the expanded size limit.")
    with pd.ExcelFile(BytesIO(content), engine="openpyxl") as workbook:
        return workbook.sheet_names


def read_grid(content, sheet=None, delimiter=","):
    if sheet is None:
        grid = pd.read_csv(BytesIO(content), header=None, dtype=str,
                           keep_default_na=False, skip_blank_lines=False,
                           sep=delimiter, encoding="utf-8-sig", nrows=MAX_ROWS + 1)
    else:
        grid = pd.read_excel(BytesIO(content), sheet_name=sheet, header=None,
                             dtype=object, keep_default_na=False,
                             engine="openpyxl", nrows=MAX_ROWS + 1)
    if grid.empty:
        raise ValueError("The file or worksheet has no data.")
    if len(grid) > MAX_ROWS or len(grid.columns) > MAX_COLUMNS or grid.size > MAX_CELLS:
        raise ValueError("Limit exceeded: 100,000 rows, 250 columns, or 1 million cells.")
    return grid.replace(r"^\s*$", pd.NA, regex=True)


def use_header(grid, row):
    names = []
    used = set()
    for index, value in enumerate(grid.iloc[row - 1], 1):
        base = f"Column {index}" if pd.isna(value) else str(value).strip()
        name = base
        suffix = 2
        while name in used:
            name = f"{base} ({suffix})"
            suffix += 1
        names.append(name)
        used.add(name)
    result = grid.iloc[row:].copy().reset_index(drop=True)
    result.columns = names
    return result


def guess_type(series):
    values = series.dropna()
    if values.empty:
        return "Text"
    if values.map(lambda value: isinstance(value, (datetime, date))).all():
        return "Date"
    if values.astype(str).str.match(r"^[+-]?0\d+$").any():
        return "Text"
    if parse_number(values).notna().all():
        return "Number"
    if values.astype(str).str.fullmatch(r"\d{4}-\d{2}-\d{2}").all():
        return "Date"
    return "Text"


# Optional sign, currency before or after, or a trailing percent: -$1,200 · 1,200 USD · 12%
# ponytail: any 3 capital letters count as a currency code; use an allow-list if that misfires.
AMOUNT = (r"^(?P<sign>[+-]?)\s*(?P<pre>[$€£¥₹]|[A-Z]{3})?\s*"
          r"(?P<num>[+-]?(?:\d[\d,]*)?\.?\d+(?:[eE][+-]?\d+)?)\s*"
          r"(?P<post>[$€£¥₹]|[A-Z]{3}|%)?$")


def parse_number(series):
    values = series.astype("string").str.strip()
    parts = values.str.extract(AMOUNT)
    number = parts["num"]
    # Recognize comma thousands separators while rejecting ambiguous commas.
    grouped = number.str.fullmatch(r"[+-]?\d{1,3}(,\d{3})+(\.\d+)?", na=False)
    number = number.where(~grouped, number.str.replace(",", "", regex=False))
    percent = parts["post"].eq("%").fillna(False)
    currency = parts["pre"].fillna(parts["post"].where(~percent))
    bad = ((parts["sign"].ne("") & number.str.match(r"[+-]"))           # -$-5
           | (parts["pre"].notna() & parts["post"].notna())             # $5 USD, $5%
           | (currency.notna() & (currency.nunique() > 1))).fillna(False)  # $ and € mixed
    text = (parts["sign"] + number).where(number.notna(), values)
    result = pd.to_numeric(text.where(~bad), errors="coerce").astype("Float64")
    result = result.where(~percent, result / 100)
    return result.replace([float("inf"), -float("inf")], pd.NA)


def prepare(frame, settings, normalize=False, clear_invalid=False, deduplicate=False):
    cleaned = pd.DataFrame(index=frame.index)
    invalid = {}
    report = []
    for column, (kind, date_format, unit) in settings.items():
        raw = frame[column]
        text = raw.astype("string")
        variants = 0
        outliers = 0
        if kind == "Number":
            parsed = parse_number(raw)
            measured = parsed.dropna().astype(float)
            if len(measured) >= 4:
                low, high = measured.quantile([0.25, 0.75])
                spread = high - low
                outliers = int(((measured < low - 1.5 * spread) |
                                (measured > high + 1.5 * spread)).sum())
        elif kind == "Date":
            existing = raw.map(lambda value: isinstance(value, (date, datetime)))
            parsed_text = pd.to_datetime(text.where(~existing).str.strip(),
                                         format=date_format, errors="coerce")
            parsed_dates = pd.to_datetime(raw.where(existing), errors="coerce")
            parsed = parsed_dates.combine_first(parsed_text)
        else:
            trimmed = text.str.strip()
            variants = int(text.nunique() - trimmed.str.casefold().nunique())
            parsed = trimmed if normalize else text
            if normalize:
                # Most common spelling wins; ties go to the one seen first.
                counts = trimmed.dropna().value_counts()
                best = counts.groupby(counts.index.str.casefold(), sort=False).head(1).index
                canonical = dict(zip(best.str.casefold(), best))
                parsed = trimmed.str.casefold().map(canonical).astype("string")
        bad = raw.notna() & parsed.isna()
        invalid[column] = raw[bad]
        cleaned[column] = parsed if clear_invalid or not bad.any() else raw
        report.append({"Column": column, "Type": kind, "Unit": unit,
                       "Missing": int(raw.isna().sum()), "Invalid": int(bad.sum()),
                       "Category variants": variants, "Unusual values": outliers})
    if deduplicate:
        cleaned = cleaned.drop_duplicates()
    return cleaned, pd.DataFrame(report), invalid


def csv_bytes(frame):
    def safe_cell(value):
        if isinstance(value, str) and re.match(r"^[\s\u200b\ufeff]*[=+@-]", value):
            return "'" + value
        return value
    escaped = frame.map(safe_cell)
    escaped.columns = [safe_cell(name) for name in frame.columns]
    return escaped.to_csv(index=False).encode("utf-8-sig")


def excel_bytes(frame):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter", engine_kwargs={
        "options": {"strings_to_formulas": False, "strings_to_urls": False}
    }) as writer:
        frame.to_excel(writer, index=False, sheet_name="Data")
    return buffer.getvalue()
