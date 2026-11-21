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

# Dashboard period comparison: latest month vs previous, per group, partial months.
from agent import compare_periods  # noqa: E402

sales = pd.DataFrame({
    "Date": pd.to_datetime(["2026-08-03", "2026-08-20", "2026-09-04", "2026-09-30", "2026-09-12"]),
    "Region": ["West", "East", "West", "East", "South"],
    "Revenue": pd.array([100, 50, 60, 80, 10], "Float64"),
    "Margin": pd.array([0.20, 0.10, 0.15, 0.10, 0.30], "Float64")})
c = compare_periods(sales, "Revenue", "Date", "Region")
assert (str(c["latest"]), c["current"], c["prior"], c["change"]) == ("2026-09", 150, 150, 0)
assert c["pct"] == 0 and not c["partial"]
assert c["table"].index[0] == "West" and c["table"].loc["West", "Change"] == -40
assert c["table"]["Latest"].sum() == c["current"]  # groups reconcile to the total
assert c["table"].loc["South", "Previous"] == 0  # no rows last month: nothing recorded
vanished = pd.DataFrame({"Date": pd.to_datetime(["2026-08-05", "2026-08-06", "2026-09-30"]),
                         "Region": ["West", "East", "East"], "Revenue": pd.array([90, 10, 30], "Float64")})
gone = compare_periods(vanished, "Revenue", "Date", "Region")["table"]
assert gone.index[0] == "West" and gone.loc["West", "Latest"] == 0  # vanished group is the top decline
assert pd.isna(compare_periods(sales, "Margin", "Date", "Region", rate=True)["table"].loc["South", "Previous"])
r = compare_periods(sales, "Margin", "Date", rate=True)
assert round(r["change"], 4) == round(0.55 / 3 - 0.15, 4) and r["pct"] is None  # pp, not %
assert compare_periods(sales[sales.Date.dt.month == 9], "Revenue", "Date")["change"] is None
zero = sales.assign(Revenue=pd.array([0, 0, 5, 0, 0], "Float64"))
assert compare_periods(zero, "Revenue", "Date")["pct"] is None  # prior 0 -> n/a
assert compare_periods(sales.iloc[:3], "Revenue", "Date")["partial"]  # data ends 4 Sep

# Year over year: same month last year, skipping the months in between.
yearly = pd.DataFrame({"Date": pd.to_datetime(["2025-09-10", "2026-08-10", "2026-09-10"]),
                       "Revenue": pd.array([200, 999, 250], "Float64")})
y = compare_periods(yearly, "Revenue", "Date", lag=12)
assert (str(y["previous"]), y["prior"], y["change"], y["pct"]) == ("2025-09", 200, 50, 0.25)
assert compare_periods(yearly.iloc[1:], "Revenue", "Date", lag=12)["prior"] is None  # no data a year ago
print("engine ok")

# Reading: ragged rows, Windows and UTF-16 encodings, empty sheets.
from engine import DataError, guess_date_format, read_grid as rg  # noqa: E402

assert rg(b"Sales,Region\n,,\n,North\n").shape == (3, 3)  # extra trailing comma no longer rejects the file
assert rg("Région\nÉté\n".encode("cp1252")).iloc[1, 0] == "Été"
assert rg("Région\tVille\nÉté\tMogadishu\n".encode("utf-16"), delimiter="\t").iloc[1, 1] == "Mogadishu"
for empty in (b"", b",,\n,,\n"):
    try:
        rg(empty)
        raise AssertionError("empty file accepted")
    except DataError as exc:
        assert "no data" in str(exc)

# Date formats are detected, day-first when ambiguous.
assert guess_date_format(pd.Series(["15/03/2026", "01/02/2026"])) == "%d/%m/%Y"
assert guess_date_format(pd.Series(["03/15/2026"])) == "%m/%d/%Y"
assert guess_date_format(pd.Series(["2026-01-01 10:30"])) == "%Y-%m-%d %H:%M"
assert guess_date_format(pd.Series(["26-March-2025", "1-Apr-2025"])) is None  # mixed month styles
assert guess_date_format(pd.Series(["26-March-2025"])) == "%d-%B-%Y"
assert guess_type(pd.Series(["15/03/2026", "16/03/2026"])) == "Date"
assert guess_type(pd.Series(["Mogadishu", "Baidoa"])) == "Text"
print("reading ok")

# Mostly-numeric columns stay numbers; the odd cell is reported as invalid.
assert guess_type(pd.Series(["$1"] * 9 + ["n/a"])) == "Number"
assert guess_type(pd.Series(["1", "2", "n/a"])) == "Text"  # 67% is not enough
assert guess_date_format(pd.Series(["2026-01-0%d" % d for d in range(1, 10)] + ["soon"])) == "%Y-%m-%d"
print("detection ok")

# Row limit counts data rows, not the header.
from engine import MAX_ROWS  # noqa: E402
header_plus_max = ("n\n" + "1\n" * MAX_ROWS).encode()
assert len(rg(header_plus_max)) == MAX_ROWS + 1
try:
    rg(("n\n" + "1\n" * (MAX_ROWS + 1)).encode())
    raise AssertionError("over-limit file accepted")
except DataError:
    pass
print("limits ok")
