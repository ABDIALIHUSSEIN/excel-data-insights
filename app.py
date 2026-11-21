"""Launch with: python -m streamlit run app.py"""
from datetime import datetime
from hashlib import sha256
import math
import os
from pathlib import Path

import pandas as pd
import streamlit as st

from agent import analyze, compare_periods, plan
from engine import (DataError, csv_bytes, excel_bytes, guess_date_format, guess_type, inspect_workbook,
                    prepare, read_grid, use_header)

st.set_page_config(page_title="Excel Data Insights", page_icon="📊", layout="wide")
title_area, language_area = st.columns([5, 1], vertical_alignment="center")
language = language_area.segmented_control("Language / Luqadda", ["English", "Soomaali"], default="English",
                                           key="language", label_visibility="collapsed") or "English"

SYMBOLS = "$€£¥₹"
FORMATS = {"%": "percent", "$": "dollar", "€": "euro", "¥": "yen"}


def tr(english, somali):
    return somali if language == "Soomaali" else english


def show(value, unit="", spec=None):
    """Readable number: whole numbers without decimals, % with one decimal, currency symbol in front."""
    if value is None or pd.isna(value) or not math.isfinite(float(value)):
        return "—"
    value = float(value) * (100 if unit == "%" else 1)
    if spec is None:
        spec = ".3g" if abs(value) >= 1e15 else ",.1f" if unit == "%" else ",.0f" if value.is_integer() else ",.2f"
    text = f"{abs(value) if unit in SYMBOLS and unit else value:{spec}}"
    if unit == "%":
        return text + "%"
    if unit and unit in SYMBOLS:
        return ("-" if value < 0 else "") + unit + text
    return f"{text} {unit}".strip()


def detect_unit(series):
    # Preset the Unit when every value carries the same % sign or currency symbol.
    values = series.dropna().astype("string").str.strip()
    if values.empty:
        return ""
    # 90% rather than all, so a few "n/a" cells don't hide the unit.
    if values.str.endswith("%").mean() >= 0.9:
        return "%"
    for symbol in SYMBOLS:
        if (values.str.lstrip("+-").str.startswith(symbol) | values.str.endswith(symbol)).mean() >= 0.9:
            return symbol
    return ""


def percent_columns(table):
    # Number columns with a % or currency unit, formatted in tables (0.125 → 12.50%, 1200 → $1,200.00).
    return {name: st.column_config.NumberColumn(
                format=FORMATS.get(settings[name][2], "localized"),
                label=f"{name} ({settings[name][2]})" if settings[name][2] not in FORMATS else None)
            for name in table.columns
            if name in settings and settings[name][2] and pd.api.types.is_numeric_dtype(table[name])}


def chart_ready(table):
    # Altair reads "name:type" in field names, so a ":" in a column or index name crashes charts.
    # Swap it for the look-alike "∶" (U+2236) in chart labels only.
    safe = table.rename(columns=lambda name: str(name).replace(":", "∶"))
    safe.index.name = None if table.index.name is None else str(table.index.name).replace(":", "∶")
    return safe


def is_rate(name):
    return name in settings and settings[name][2] == "%"


def percent_chart(table):
    # Charts cannot format their axis, so scale % columns to 0-100 and label the unit in the series name.
    scaled = table.apply(lambda column: column * 100 if is_rate(column.name) else column)
    return scaled.rename(columns=lambda name: f"{name} ({settings[name][2]})"
                         if name in settings and settings[name][2] else name)


class Sample:
    """The bundled sample, used like an uploaded file."""
    name = "sample_finance.csv"

    def getvalue(self):
        return Path(__file__).with_name(self.name).read_bytes()


title_area.title("📊 Excel Data Insights")
title_area.caption(tr("Upload a spreadsheet, check its quality, and see what changed — in English or Somali.",
                      "Soo geli fayl, hubi tayada xogta, oo arag waxa isbeddelay — Ingiriis ama Soomaali."))

# Once a file is loaded, the upload box and file settings fold into one line so the data comes first.
loaded = st.session_state.get("upload") is not None or st.session_state.get("use_sample")
if loaded:
    current = st.session_state["upload"].name if st.session_state.get("upload") is not None else Sample.name
    file_panel = st.expander(f"📄 {current} — " + tr("change file or file settings", "beddel faylka ama dejinta"))
else:
    file_panel = st.container()
uploaded = file_panel.file_uploader(tr("Upload Excel (.xlsx) or CSV, up to 20 MB", "Soo geli Excel (.xlsx) ama CSV, ilaa 20 MB"),
                                    type=["xlsx", "csv"], key="upload")
if uploaded is not None:
    st.session_state.pop("use_sample", None)
elif st.session_state.get("use_sample"):
    uploaded = Sample()
    file_panel.caption(tr("🧪 Showing the sample file. Upload your own file above to replace it.",
                          "🧪 Waxaa muuqda faylka tusaalaha. Fayl kaaga ah kor ku soo geli si aad u beddesho."))
else:
    steps = st.columns(3)
    for column, (icon, title, text) in zip(steps, [
            ("📤", tr("1. Upload", "1. Soo geli"), tr("An Excel workbook or a CSV file.", "Buug Excel ama fayl CSV.")),
            ("🔍", tr("2. Review", "2. Hubi"), tr("Fix column types, blanks, duplicates, and spelling variants.",
                                                  "Sax noocyada tiirarka, bannaanada iyo kuwa soo noqnoqda.")),
            ("📈", tr("3. Explore", "3. Sahami"), tr("KPIs vs last month or last year, trends, and what to check next.",
                                                    "Tilmaamayaasha, isbeddelka iyo waxa xiga ee la hubiyo."))]):
        with column.container(border=True):
            st.markdown(f"#### {icon} {title}")
            st.caption(text)
    if st.button(tr("Try it with sample sales data", "Ku tijaabi xog iib tusaale ah"), type="primary", icon="🧪"):
        st.session_state["use_sample"] = True
        st.rerun()
    st.stop()

content = uploaded.getvalue()
if len(content) > 20 * 1024**2:
    st.error(tr("Choose a file under 20 MB.", "Dooro fayl ka yar 20 MB."))
    st.stop()

is_excel = uploaded.name.lower().endswith(".xlsx")
file_key = sha256(content).hexdigest()
sheet = None
delimiter = ","
settings_row = file_panel.columns(2)
try:
    if is_excel:
        sheet = settings_row[0].selectbox(tr("Worksheet", "Xaashida Excel"), inspect_workbook(content),
                                          key=f"sheet-{file_key}")
    else:
        delimiters = {tr("Comma", "Hakad"): ",", tr("Semicolon", "Hakad-dhibic"): ";",
                      "Tab": "\t", "Pipe |": "|"}
        # Preselect the separator that appears most in the first line.
        first_line = content[:4096].split(b"\n")[0]
        likely = max(delimiters.values(), key=lambda mark: first_line.count(mark.encode()))
        delimiter = delimiters[settings_row[0].selectbox(tr("CSV separator", "Kala-soocaha CSV"), list(delimiters),
                                                         index=list(delimiters.values()).index(likely),
                                                         key=f"separator-{file_key}")]
    grid = read_grid(content, sheet, delimiter)
except DataError as exc:
    st.error(tr(str(exc), {"The file or worksheet has no data.": "Faylka ama xaashidan xog kuma jirto."}.get(
        str(exc), "Faylku wuxuu dhaafay xadka: 100,000 saf, 250 tiir ama 1 milyan unug.")), icon="⚠️")
    st.stop()
except Exception as exc:
    st.error(tr("Could not read the file. Check that it is a real .xlsx or CSV file and pick the right separator.",
                "Faylka lama akhrin karin. Hubi inuu yahay .xlsx ama CSV dhab ah oo dooro kala-soocaha saxda ah."),
             icon="⚠️")
    st.caption(type(exc).__name__)
    st.stop()

source_key = f"{file_key}:{sheet}:{delimiter}"
header = settings_row[1].number_input(tr("Header row", "Safka cinwaannada"), min_value=1, max_value=len(grid),
                                      value=1, key=f"header-{source_key}",
                                      help=tr("The row that holds the column names. Use it to skip title rows above the table.",
                                              "Safka magacyada tiirarka."))
frame = use_header(grid, header)
key = f"{source_key}:{header}"
dashboard, review, assistant, exports = st.tabs([
    "📈 " + tr("Dashboard", "Shaashadda xogta"), "🔍 " + tr("Review data", "Hubi xogta"),
    "💬 " + tr("Ask the agent", "Weydii kaaliyaha"), "⬇️ " + tr("Downloads", "Soo dejinta")
])

with review:
    # Settings first (they drive cleaning), then the summary cards are filled in above them.
    summary_cards = st.container()

    with st.container(border=True):
        st.markdown("##### " + tr("1 · Column types", "1 · Noocyada tiirarka"))
        st.caption(tr("Detected automatically. Change a type, date format (e.g. %d/%m/%Y), or unit if it looks wrong. "
                      "Unit % shows 0.125 as 12.5%; $ € £ ¥ ₹ show as currency.",
                      "Si toos ah ayaa loo ogaaday. Beddel nooca, qaabka taariikhda ama halbeegga haddii uu khaldan yahay."))
        kinds = {name: guess_type(frame[name]) for name in frame}
        schema = pd.DataFrame([{"Column": name, "Type": kinds[name],
                                "Date format": (guess_date_format(frame[name]) or "%Y-%m-%d") if kinds[name] == "Date" else "",
                                "Unit": detect_unit(frame[name])} for name in frame])
        edited = st.data_editor(schema, key=f"schema-{key}", hide_index=True, disabled=["Column"], width="stretch",
                                column_config={"Type": st.column_config.SelectboxColumn(
                                    options=["Text", "Number", "Date"], required=True)})
    settings = {row["Column"]: (row["Type"], row["Date format"] or "%Y-%m-%d", row["Unit"] or "")
                for row in edited.to_dict("records")}

    with st.container(border=True):
        st.markdown("##### " + tr("2 · Cleaning", "2 · Hagaajinta"))
        options = st.columns(3)
        normalize = options[0].checkbox(tr("Apply category cleanup", "Midee qaybaha isku magaca ah"),
                                        help=tr("Trim spaces and merge case variants such as Baidoa / baidoa.",
                                                "Ka saar bannaanada oo midee faraha kala duwan."), key=f"normalize-{key}")
        clear_invalid = options[1].checkbox(tr("Replace invalid values with blanks", "Qiyamka khaldan ka dhig bannaan"),
                                            help=tr("Needed to calculate a column that has some invalid cells.",
                                                    "Waa lagama maarmaan si tiir leh qiyam khaldan loo xisaabiyo."),
                                            key=f"invalid-{key}")
        deduplicate = options[2].checkbox(tr("Remove duplicate rows", "Ka saar safafka soo noqnoqda"),
                                          key=f"duplicates-{key}")
    cleaned, report, invalid = prepare(frame, settings, normalize, clear_invalid, deduplicate)

    with st.container(border=True):
        st.markdown("##### " + tr("3 · Quality report", "3 · Warbixinta tayada"))
        problems = report[["Missing", "Invalid", "Category variants", "Unusual values"]].sum(axis=1)
        report.insert(0, tr("Status", "Xaalad"), [
            "✅ OK" if count == 0 else f"⚠️ {count:,} " + tr("to check", "la hubiyo") for count in problems])
        st.dataframe(report, hide_index=True, width="stretch")
        st.caption(tr("Unusual values use the 1.5 × IQR rule. Category variants are spellings that differ only in case or spaces.",
                      "Qiyamka aan caadiga ahayn waxaa lagu calaamadeeyaa 1.5 × IQR."))
        issues = [column for column in invalid if not invalid[column].empty]
        if issues:
            column = st.selectbox(tr("Inspect invalid values", "Eeg qiyamka khaldan"), issues, key=f"issues-{key}")
            st.dataframe(invalid[column].head(30).astype("string").rename(column), width="stretch")
            if not clear_invalid:
                st.caption(tr("While invalid values are kept, the column stays in its original format and can't be calculated.",
                              "Inta qiyamka khaldan la hayo, tiirka lama xisaabin karo."))

    with st.container(border=True):
        st.markdown("##### " + tr("4 · Preview (first 30 rows, as in the file)", "4 · Muuqaal (30 saf ee ugu horreeya)"))
        st.dataframe(frame.head(30).astype("string"), hide_index=True, width="stretch")

    with summary_cards:
        cells = max(frame.size, 1)
        quality = st.columns(5)
        quality[0].metric(tr("Rows", "Safaf"), f"{len(frame):,}", border=True)
        quality[1].metric(tr("Columns", "Tiirar"), f"{len(frame.columns):,}", border=True)
        quality[2].metric(tr("Blank cells", "Unugyo bannaan"), f"{int(frame.isna().sum().sum()):,}",
                          f"{frame.isna().sum().sum() / cells:.1%} " + tr("of cells", "unugyada"),
                          delta_color="off", delta_arrow="off", border=True)
        quality[3].metric(tr("Invalid values", "Qiyam khaldan"), f"{int(report['Invalid'].sum()):,}", border=True)
        quality[4].metric(tr("Duplicate rows", "Safaf soo noqnoqday"), f"{int(frame.duplicated().sum()):,}",
                          tr(f"{len(frame) - len(cleaned):,} removed", f"{len(frame) - len(cleaned):,} la saaray"),
                          delta_color="off", delta_arrow="off", border=True)

numeric = list(cleaned.select_dtypes(include="number").columns)
dates = list(cleaned.select_dtypes(include="datetime").columns)
categories = [name for name in cleaned if name not in numeric + dates]
with dashboard:
    # Reading order: purpose and period, KPIs, trend, segment comparison, what to check, detail, notes.
    blocked = {column: values for column, values in invalid.items()
               if not values.empty and settings[column][0] in ("Number", "Date") and not clear_invalid}
    variants = int(report["Category variants"].sum())
    if blocked:
        examples = ", ".join(f"{column} ({len(values):,}: “{values.astype('string').iloc[0]}”)" for column, values in blocked.items())
        fix, act = st.columns([5, 2], vertical_alignment="center")
        fix.warning(tr(f"Not in the dashboard yet because of invalid cells: {examples}.",
                       f"Weli kuma jiraan shaashadda sababo la xiriira unugyo khaldan: {examples}."), icon="⚠️")
        act.button(tr("Blank the invalid cells", "Bannaanee unugyada khaldan"), type="primary", width="stretch",
                   on_click=lambda: st.session_state.update({f"invalid-{key}": True}))
    if variants and not normalize:
        fix, act = st.columns([5, 2], vertical_alignment="center")
        fix.info(tr(f"{variants:,} spelling variant(s) such as Baidoa / baidoa are counted separately.",
                    f"{variants:,} qoraal oo kala duwan sida Baidoa / baidoa ayaa si gooni ah loo tiriyay."), icon="💡")
        act.button(tr("Merge spelling variants", "Midee qoraallada"), width="stretch",
                   on_click=lambda: st.session_state.update({f"normalize-{key}": True}))

    header = st.container()
    controls = st.container(border=True)
    picks = controls.columns(4)
    measure = picks[0].selectbox(tr("Measure", "Cabbirka"), numeric, key=f"measure-{key}") if numeric else None
    date_column = picks[1].selectbox(tr("Date", "Taariikhda"), dates, key=f"date-{key}") if dates else None
    category = picks[2].selectbox(tr("Compare by", "Isbarbardhig"), categories, key=f"category-{key}") if categories else None
    unit = settings[measure][2] if measure else ""
    rate = unit == "%"
    lag = 1
    if date_column and measure:
        lag = picks[3].segmented_control(tr("Compare with", "Is barbar dhig"), [1, 12], default=1, key=f"lag-{key}",
                                         format_func=lambda months: tr("Last month", "Bishii hore") if months == 1
                                         else tr("Last year", "Sannadkii hore")) or 1
    filtered = cleaned
    active = tr("All rows", "Dhammaan safafka")
    filter_key = f"filter-{key}"
    pick_filter, pick_values, pick_target, pick_reset = controls.columns([2, 3, 2, 1], vertical_alignment="bottom")
    target = 0.0
    if measure and not rate:
        target = pick_target.number_input(tr("Monthly target (optional)", "Hadafka bishii (ikhtiyaari)"),
                                          min_value=0.0, value=0.0, key=f"target-{key}")
    filter_column = pick_filter.selectbox(tr("Filter", "Shaandhayn"), [None, *cleaned.columns], key=filter_key,
                                          format_func=lambda name: name if name is not None else tr("All rows", "Dhammaan safafka"))
    pick_reset.button(tr("Reset", "Dib u celi"), on_click=lambda: st.session_state.update({filter_key: None}),
                      disabled=filter_column is None, width="stretch")
    if filter_column is not None:
        values = cleaned[filter_column]
        with pick_values:
            if filter_column in dates and values.notna().any():
                start, end = values.min().date(), values.max().date()
                interval = st.date_input(tr("Date range", "Muddada"), (start, end), min_value=start, max_value=end,
                                         key=f"range-{key}-{filter_column}-{start}-{end}")
                if len(interval) == 2:
                    filtered = cleaned[values.dt.date.between(*interval)]
                    active = f"{filter_column} {interval[0]:%d %b %Y} – {interval[1]:%d %b %Y}"
            elif filter_column in numeric and values.notna().any():
                low, high = float(values.min()), float(values.max())
                if low < high:
                    bounds = st.slider(tr("Value range", "Xadka qiimaha"), low, high, (low, high),
                                       key=f"bounds-{key}-{filter_column}-{low}-{high}")
                    filtered = cleaned[values.between(*bounds)]
                    active = f"{filter_column} {show(bounds[0], settings[filter_column][2])} – {show(bounds[1], settings[filter_column][2])}"
            elif filter_column in categories:
                labels = values.astype("string")
                if labels.nunique() > 500:
                    search = st.text_input(tr("Contains", "Ka kooban"), key=f"contains-{key}-{filter_column}")
                    if search:
                        filtered = cleaned[labels.str.contains(search, case=False, regex=False, na=False)]
                        active = f"{filter_column} ∋ “{search}”"
                else:
                    chosen = st.multiselect(tr("Keep these values", "Hay qiyamkan"), sorted(labels.dropna().unique()),
                                            key=f"choices-{key}-{filter_column}")
                    if chosen:
                        filtered = cleaned[labels.isin(chosen)]
                        active = f"{filter_column}: {', '.join(chosen)}"

    period = ""
    if date_column and filtered[date_column].notna().any():
        period = f"{filtered[date_column].min():%d %b %Y} – {filtered[date_column].max():%d %b %Y}"
    with header:
        st.subheader(tr(f"{measure} overview", f"Dulmarka {measure}") if measure else tr("Data overview", "Dulmarka xogta"),
                     anchor=False)
        st.caption(" · ".join(part for part in [
            period and tr("Period: ", "Muddada: ") + period,
            tr("Filter: ", "Shaandhayn: ") + active,
            tr(f"{len(filtered):,} rows", f"{len(filtered):,} saf")] if part))

    numbers = filtered[measure].dropna() if measure else pd.Series(dtype="Float64")
    if filtered.empty:
        st.info(tr("No rows match this filter. Press Reset to see all rows.",
                   "Shaandhayntan saf uma soo bixin. Riix Dib u celi si aad u aragto dhammaan safafka."))
    else:
        compared = compare_periods(filtered, measure, date_column, category, rate, lag) \
            if measure and date_column and not numbers.empty else None
        cards = st.columns(4)
        if compared and compared["current"] is not None:
            latest, previous = compared["latest"].strftime("%b %Y"), compared["previous"].strftime("%b %Y")
            partial = tr(" (partial)", " (qayb)") if compared["partial"] else ""
            delta = None
            if compared["change"] is not None:
                delta = f"{compared['change'] * 100:+.1f} pp" if rate else \
                    f"{compared['pct']:+.1%}" if compared["pct"] is not None else "n/a"
            cards[0].metric(f"{measure}, {latest}{partial}", show(compared["current"], unit), delta,
                            delta_description=tr(f"vs {previous}", f"marka loo eego {previous}"), border=True)
            cards[1].metric(f"{measure}, {previous}",
                            show(compared["prior"], unit) if compared["prior"] is not None else "—",
                            help=tr("No rows in this month." if compared["prior"] is None else None, "Bishan saf kuma jiro."
                                    if compared["prior"] is None else None), border=True)
            if target > 0:
                remaining = target - compared["current"]
                cards[2].metric(tr(f"Of {show(target, unit)} target", f"Hadafka {show(target, unit)}"),
                                f"{compared['current'] / target:.1%}",
                                tr(f"{show(abs(remaining), unit)} {'to go' if remaining > 0 else 'over'}",
                                   f"{show(abs(remaining), unit)} {'ka haray' if remaining > 0 else 'ka badan'}"),
                                delta_color="off", delta_arrow="off", border=True)
            else:
                cards[2].metric(tr("Average per row", "Celcelis saf kasta"), show(numbers.mean(), unit), border=True)
            if compared["partial"]:
                st.caption(tr(f"⚠ {latest} is partial: data runs to {filtered[date_column].max():%d %b %Y}. "
                              "Compare it with a full month with care.",
                              f"⚠ {latest} waa qayb: xogtu waxay ku egtahay {filtered[date_column].max():%d %b %Y}."))
        elif measure and not numbers.empty:
            cards[0].metric(tr("Average", "Celcelis") if rate else tr("Total", "Wadar"),
                            show(numbers.mean() if rate else numbers.sum(), unit), border=True)
            cards[1].metric(tr("Average per row", "Celcelis saf kasta"), show(numbers.mean(), unit), border=True)
            cards[2].metric(tr("Blank cells", "Unugyo bannaan"), f"{filtered.isna().sum().sum():,}", border=True)
        else:
            cards[0].metric(tr("Columns", "Tiirar"), len(filtered.columns), border=True)
            cards[1].metric(tr("Blank cells", "Unugyo bannaan"), f"{filtered.isna().sum().sum():,}", border=True)
        cards[3].metric(tr("Rows in view", "Safafka muuqda"), f"{len(filtered):,}", border=True)

        left, right = st.columns(2)
        if date_column and not numbers.empty:
            with left.container(border=True):
                st.markdown("**" + tr(f"Monthly {measure}", f"{measure} bil kasta") + "**")
                months = filtered[date_column].dt.to_period("M")
                monthly = filtered[measure].groupby(months)
                trend = (monthly.mean() if rate else monthly.sum(min_count=1)).dropna()
                if not trend.empty:
                    every_month = pd.period_range(trend.index.min(), trend.index.max(), freq="M")
                    trend = trend.reindex(every_month).tail(36)
                    trend.index = trend.index.strftime("%Y-%m")
                    st.line_chart(chart_ready(percent_chart(trend.to_frame(measure))), height=300)
        if category:
            with (right if date_column and not numbers.empty else left).container(border=True):
                labels = filtered[category].astype("string").fillna(tr("(Blank)", "(Bannaan)"))
                if numbers.empty:
                    st.markdown("**" + tr(f"Rows by {category}", f"Safafka {category} kasta") + "**")
                    bars = labels.value_counts().head(15).rename(tr("Rows", "Safaf")).to_frame()
                else:
                    st.markdown("**" + tr(f"{measure} by {category}", f"{measure} {category} kasta") + "**")
                    grouped = filtered[measure].groupby(labels)
                    bars = percent_chart((grouped.mean() if rate else grouped.sum(min_count=1)).dropna()
                                         .nlargest(15).to_frame(measure))
                # Sorted horizontal bars, largest first; bars start at zero.
                bars = chart_ready(bars)
                st.bar_chart(bars, horizontal=True, sort="-" + bars.columns[0], height=300)

        table = compared["table"] if compared else None
        if table is not None and table["Change"].notna().any():
            with st.container(border=True):
                st.markdown("**" + tr(f"{measure} by {category}: {previous} vs {latest}",
                                      f"{measure} {category} kasta: {previous} iyo {latest}") + "**")
                worst, best = table["Change"].idxmin(), table["Change"].idxmax()
                move = table.loc[worst, "Change"]
                name, size = (worst, move) if move < 0 else (best, table.loc[best, "Change"])
                amount = f"{abs(size) * 100:.1f} pp" if rate else show(abs(size), unit)
                st.info(tr(f"**What to check next:** {name} had the largest {'decline' if size < 0 else 'increase'} in {measure}, "
                           f"{amount} from {previous} to {latest}. Filter to {category} = {name} to see which rows drove it.",
                           f"**Waxa xiga ee la hubiyo:** {name} ayaa ugu {'hoos u dhac' if size < 0 else 'kororka'} badnaa {measure}, "
                           f"{amount} laga bilaabo {previous} ilaa {latest}. Shaandhee {category} = {name} si aad u aragto safafka."))
                shown = table.copy()
                shown.loc[tr("Total", "Wadar")] = [compared["prior"], compared["current"], compared["change"]]
                if rate:
                    shown["Change"] = shown["Change"] * 100
                else:
                    shown["% change"] = (shown["Change"] / shown["Previous"].abs()).where(shown["Previous"] != 0)
                shown = shown.rename(columns={"Previous": previous, "Latest": latest,
                                              "Change": tr("Change (pp)", "Isbeddel (pp)") if rate else tr("Change", "Isbeddel")})
                change_column = shown.columns[2]
                formats = {previous: lambda v: show(v, unit), latest: lambda v: show(v, unit),
                           change_column: (lambda v: f"{v:+.1f} pp") if rate else (lambda v: ("+" if v > 0 else "") + show(v, unit)),
                           "% change": lambda v: f"{v:+.1%}"}
                text = shown.astype(object).apply(lambda column: column.map(
                    lambda v: "—" if pd.isna(v) else formats.get(column.name, str)(v)))
                st.dataframe(text, width="stretch", column_config={
                    name: st.column_config.TextColumn(alignment="right") for name in text.columns})

        with st.expander(tr(f"Detail: first 100 of {len(filtered):,} rows", f"Faahfaahin: 100 ka mid ah {len(filtered):,} saf"),
                         icon="📋"):
            st.dataframe(filtered.head(100), hide_index=True, width="stretch",
                         column_config=percent_columns(filtered))

    applied = [label for label, on in [(tr("category cleanup", "midaynta qaybaha"), normalize),
                                       (tr("invalid values blanked", "qiyamka khaldan la bannaanay"), clear_invalid),
                                       (tr("duplicates removed", "safafka soo noqnoqday la saaray"), deduplicate)] if on]
    st.caption(" · ".join(part for part in [
        tr("Source: ", "Isha: ") + uploaded.name + (f" › {sheet}" if sheet else ""),
        tr(f"{len(cleaned):,} of {len(frame):,} rows after cleaning", f"{len(cleaned):,} ka mid ah {len(frame):,} saf"),
        tr("Cleaning: ", "Hagaajin: ") + (", ".join(applied) or tr("none", "midna")),
        tr("Months are calendar months; the latest month is the most recent date in view.",
           "Bilaha waa bilaha kalandarka; bisha ugu dambeysa waa taariikhda ugu dambeysa."),
        tr("% measures are averages of row values.", "Cabbirada % waa celcelis qiyamka safafka.") if rate else "",
        tr("Generated ", "La sameeyey ") + f"{datetime.now().astimezone():%d %b %Y %H:%M %Z}"] if part))

with assistant:
    ai_enabled = False
    if os.getenv("OPENAI_API_KEY"):
        ai_enabled = st.toggle(tr("Use OpenAI to understand free-form questions", "Adeegso OpenAI si loo fahmo su’aalaha"),
                               help=tr("Sends only your question and column names and types. Calculations happen here.",
                                       "Waxaa loo diraa su’aasha iyo magacyada tiirarka oo keliya."))
    # One-click questions built from this file's own columns.
    # Prefer an amount over a % column: a total of percentages means nothing.
    first_number = next((name for name in numeric if not is_rate(name)), numeric[0] if numeric else None)
    examples = [tr("Overview", "Dulmar")]
    if first_number and categories:
        examples.append(f"Total {first_number} by {categories[0]}")
    if first_number:
        examples.append(f"Average {first_number}")
    if first_number and dates:
        examples.append(f"Monthly trend {first_number} by {dates[0]}")
    if len(numeric) > 1:
        examples.append(f"Correlation {numeric[0]} and {numeric[1]}")
    if categories:
        examples.append(f"Count by {categories[0]}")
    question_key, example_key, run_key = f"question-{key}", f"example-{key}", f"ask-{key}"

    def use_example():
        # Copy the picked example into the box, clear the pick, and answer it once.
        st.session_state[question_key] = st.session_state[example_key]
        st.session_state[example_key] = None
        st.session_state[run_key] = True

    st.pills(tr("Try a question", "Isku day su’aal"), examples, key=example_key, on_change=use_example)
    question = st.text_input(tr("Or type your own, using column names", "Ama qor su’aashaada, adoo isticmaalaya magacyada tiirarka"),
                             max_chars=1000, key=question_key,
                             placeholder=tr("e.g. Total Revenue by Region", "tusaale: Wadarta Revenue by Region"))
    st.caption(tr(f"Answers use the dashboard view: {active} · {len(filtered):,} rows.",
                  f"Jawaabaha waxay adeegsadaan shaandhaynta: {active} · {len(filtered):,} saf."))
    asked = st.button(tr("Answer", "Jawaab"), type="primary", disabled=not question.strip())
    if question.strip() and (asked or st.session_state.pop(run_key, False)):
        try:
            operation, selected_measure, grouping = plan(question, filtered, ai_enabled)
            summary, detail = analyze(filtered, operation, selected_measure, grouping)
        except ValueError:
            st.warning(tr("Try a supported question using the exact column names. For conditions, apply a dashboard filter first.",
                          "Isku day su’aal tusaalayaasha la mid ah oo leh magacyada saxda ah ee tiirarka. Marka hore shaandhee haddii ay jirto shuruud."))
        except Exception:
            st.error(tr("The assistant could not complete the request. Check the API setup or switch to local mode.",
                        "Kaaliyuhu ma dhammeystirin codsiga. Hubi dejinta API-ga ama u beddel habka gudaha."))
        else:
            phrases = {
                "overview": ("Data overview", "Dulmarka xogta"), "count": ("Row count", "Tirada safafka"),
                "missing": ("Missing values", "Qiyamka maqan"), "sum": ("Total", "Wadar"),
                "mean": ("Average", "Celcelis"), "min": ("Minimum", "Ugu yar"),
                "max": ("Maximum", "Ugu badan"), "median": ("Median", "Qiimaha dhexe"),
                "group_sum": ("Totals by category, up to 25", "Wadarta qayb kasta, ilaa 25"),
                "group_count": ("Rows by category, up to 25", "Safafka qayb kasta, ilaa 25"),
                "trend": ("Monthly totals, up to 36 months", "Wadarta bil kasta, ilaa 36 bilood"),
                "profile": ("Numeric statistics", "Tirakoobka tiirka"),
                "correlation": ("Pearson correlation", "Xiriirka Pearson")
            }
            st.subheader(tr(*phrases[operation]))
            if "rows" in summary:
                st.write(tr(f"This view has {summary['rows']:,} rows, {summary['columns']} columns, and {summary['missing']:,} blank cells.",
                            f"Xogtan waxay leedahay {summary['rows']:,} saf, {summary['columns']} tiir iyo {summary['missing']:,} unug oo bannaan."))
            if "value" in summary:
                fields = " / ".join([summary["measure"], *([summary["group"]] if "group" in summary else [])])
                st.write(fields)
                unit = settings[summary["measure"]][2] if operation != "correlation" else ""
                st.metric(tr("Result", "Natiijada"), show(summary["value"], unit, ",.6g"))
                st.caption(tr(f"Observations used: {summary['n']:,}", f"Tirada la xisaabiyey: {summary['n']:,}"))
            if detail is not None:
                if operation == "profile" and any(map(is_rate, detail.columns)):
                    # The count row is a number of rows, not a percentage.
                    detail = percent_chart(detail.drop(index="count")).round(4)
                    st.caption(tr(f"Observations used: {summary['n']:,}", f"Tirada la xisaabiyey: {summary['n']:,}"))
                if operation in {"group_sum", "group_count"} and len(detail) > 1:
                    bars = chart_ready(percent_chart(detail.rename(index=lambda v: "(Blank)" if pd.isna(v) else str(v))))
                    st.bar_chart(bars, horizontal=True, sort="-" + bars.columns[0], height=min(60 + 32 * len(bars), 400))
                elif operation == "trend" and len(detail) > 1:
                    st.line_chart(chart_ready(percent_chart(detail)), height=280)
                st.dataframe(detail, column_config=percent_columns(detail))

with exports:
    export_format = st.segmented_control(tr("Format", "Nooca faylka"), ["Excel (.xlsx)", "CSV"], default="Excel (.xlsx)",
                                         key="export-format") or "Excel (.xlsx)"
    as_excel = export_format.startswith("Excel")
    export = excel_bytes if as_excel else csv_bytes
    extension = "xlsx" if as_excel else "csv"
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if as_excel else "text/csv"
    stem = Path(uploaded.name).stem
    cards = st.columns(3)
    # Files are built only when a button is clicked, so large data doesn't slow every page update.
    for column, (title, text, data, name) in zip(cards, [
            (tr("Cleaned data", "Xogta la hagaajiyey"),
             tr(f"{len(cleaned):,} rows with your cleaning applied.", f"{len(cleaned):,} saf oo la hagaajiyey."),
             lambda: export(cleaned), f"{stem}-cleaned.{extension}"),
            (tr("Dashboard view", "Xogta la shaandheeyey"),
             tr(f"{len(filtered):,} rows · {active}", f"{len(filtered):,} saf · {active}"),
             lambda: export(filtered), f"{stem}-filtered.{extension}"),
            (tr("Original file", "Faylka asalka ah"), tr("Exactly as uploaded.", "Sidii loo soo geliyey."),
             content, uploaded.name)]):
        with column.container(border=True):
            st.markdown(f"**{title}**")
            st.caption(text)
            st.download_button(tr("Download", "Soo dejiso"), data, name,
                               mime if data is not content else "application/octet-stream",
                               icon="⬇️", width="stretch", on_click="ignore", key=f"download-{name}")
    st.caption(tr("Excel keeps text as text. CSV adds a leading ' to text a spreadsheet could run as a formula.",
                  "Excel wuxuu qoraalka u hayaa qoraal. CSV wuxuu ' ku daraa qoraalka u eg qaacido."))
