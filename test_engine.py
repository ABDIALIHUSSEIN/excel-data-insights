"""Run: .venv/bin/python test_engine.py"""
from pathlib import Path

import pandas as pd

from engine import csv_bytes, guess_type, parse_number, prepare, read_grid, use_header


def clean(values, normalize=True):
    frame = pd.DataFrame({"R": pd.Series(values, dtype=object)})
    return prepare(frame, {"R": ("Text", "", "")}, normalize=normalize)[0]["R"].tolist()


# Category cleanup keeps the most common spelling, ties go to the first seen.
assert clean(["Baidoa", "baidoa", "baidoa", " BAIDOA "]) == ["baidoa"] * 4
assert clean(["Baidoa", "baidoa"]) == ["Baidoa", "Baidoa"]
assert clean(["a", None, "A"])[1] is pd.NA

# Sample file end to end.
grid = read_grid(Path(__file__).with_name("sample.csv").read_bytes())
frame = use_header(grid, 1)
types = {name: guess_type(frame[name]) for name in frame}
assert types == {"Date": "Date", "Region": "Text", "Sales": "Number", "Orders": "Number"}
cleaned, report, _ = prepare(frame, {n: (t, "%Y-%m-%d", "") for n, t in types.items()},
                             normalize=True, clear_invalid=True, deduplicate=True)
assert len(cleaned) == 7 and set(cleaned["Region"]) == {"Mogadishu", "Baidoa", "Kismayo"}
assert cleaned["Sales"].sum() == 14900

# Numbers, IDs, and spreadsheet formula escaping.
assert parse_number(pd.Series(["1,234", "1,23", "inf"])).tolist()[0] == 1234
assert parse_number(pd.Series(["1,23", "inf"])).isna().all()
assert guess_type(pd.Series(["01234", "02345"])) == "Text"

# Currency and percentages.
dollars = parse_number(pd.Series(["$1,200", "-$5.50", "$-5", " 7 $", "1e3", None]))
assert dollars[:5].tolist() == [1200, -5.5, -5, 7, 1000] and dollars.isna()[5]
assert parse_number(pd.Series(["1,200 USD", "USD 3"])).tolist() == [1200, 3]
assert parse_number(pd.Series(["$1", "€7", "2"])).isna().tolist() == [True, True, False]  # mixed
assert parse_number(pd.Series(["12%", "-0.5%", "100 %"])).round(6).tolist() == [0.12, -0.005, 1]
assert parse_number(pd.Series(["$5 USD", "$5%", "-$-5", "1,23 USD", "abc", "$"])).isna().all()
assert parse_number(pd.Series(["£2", "3"])).tolist() == [2, 3]  # one currency plus bare numbers is fine
assert guess_type(pd.Series(["$1,200", "$950"])) == "Number"
assert guess_type(pd.Series(["5%", "12.5%"])) == "Number"
assert guess_type(pd.Series(["ID 12", "ID 13"])) == "Text"
assert "'=1+1" in csv_bytes(pd.DataFrame({"x": ["=1+1"]})).decode("utf-8-sig")
print("engine ok")
