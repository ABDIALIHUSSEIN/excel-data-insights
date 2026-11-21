"""Launch with: python -m streamlit run app.py"""
from hashlib import sha256
import os
from pathlib import Path

import pandas as pd
import streamlit as st

from agent import analyze, plan
from engine import csv_bytes, excel_bytes, guess_type, inspect_workbook, prepare, read_grid, use_header

st.set_page_config(page_title="Excel Data Insights", page_icon="📊", layout="wide")
st.title("Excel Data Insights")
language = st.sidebar.radio("Language / Luqadda", ["English", "Soomaali"], horizontal=True)


def tr(english, somali):
    return somali if language == "Soomaali" else english


def show(value, unit, spec=",.2f"):
    # A "%" unit means values are stored as fractions, so 0.125 shows as 12.5%.
    if unit == "%":
        return f"{value * 100:{spec}}%"
    return f"{value:{spec}} {unit}".strip()


def is_percent(series):
    values = series.dropna().astype("string").str.strip()
    return not values.empty and values.str.endswith("%").all()


def percent_columns(table):
    # Table columns whose unit is %, shown by Streamlit as percentages.
    return {name: st.column_config.NumberColumn(format="percent") for name in table.columns
            if name in settings and settings[name][2] == "%"}


def percent_chart(table):
    # Charts cannot format their axis, so scale to 0-100 and say so in the series name.
    return table.apply(lambda column: column * 100 if column.name in percent_columns(table) else column) \
        .rename(columns=lambda name: f"{name} (%)" if name in percent_columns(table) else name)


st.caption(tr("Upload a spreadsheet, review its quality, and explore what it shows.",
              "Soo geli fayl, hubi tayada xogta, kadibna sahami natiijooyinka."))
uploaded = st.file_uploader(tr("Upload Excel or CSV", "Soo geli Excel ama CSV"), type=["xlsx", "csv"])
if uploaded is None:
    st.info(tr("Drop a file above, or try the sample below.", "Fayl kor ku rid ama tijaabi tusaalaha hoose."))
    st.download_button(tr("Download sample", "Soo dejiso tusaale"), (Path(__file__).with_name("sample.csv")).read_bytes(),
                       "sample.csv", "text/csv")
    st.stop()

content = uploaded.getvalue()
if len(content) > 20 * 1024**2:
    st.error(tr("Choose a file under 20 MB.", "Dooro fayl ka yar 20 MB."))
    st.stop()

is_excel = uploaded.name.lower().endswith(".xlsx")
file_key = sha256(content).hexdigest()
sheet = None
delimiter = ","
try:
    if is_excel:
        sheet = st.selectbox(tr("Worksheet", "Xaashida Excel"), inspect_workbook(content), key=f"sheet-{file_key}")
    else:
        delimiters = {tr("Comma", "Hakad"): ",", tr("Semicolon", "Hakad-dhibic"): ";",
                      "Tab": "\t", "Pipe |": "|"}
        delimiter = delimiters[st.selectbox(tr("CSV separator", "Kala-soocaha CSV"), list(delimiters))]
    grid = read_grid(content, sheet, delimiter)
except Exception as exc:
    st.error(tr("Could not read the file. Check its format and separator.",
                "Faylka lama akhrin karin. Hubi nooca faylka iyo kala-soocaha."))
    st.caption(type(exc).__name__)
    st.stop()

source_key = f"{file_key}:{sheet}:{delimiter}"
header = st.number_input(tr("Header row", "Safka cinwaannada"), min_value=1, max_value=len(grid),
                         value=1, key=f"header-{source_key}")
frame = use_header(grid, header)
key = f"{source_key}:{header}"
review, dashboard, assistant, exports = st.tabs([
    tr("Review data", "Hubi xogta"), tr("Dashboard", "Shaashadda xogta"),
    tr("Ask the agent", "Weydii kaaliyaha"), tr("Downloads", "Soo dejinta")
])

with review:
    st.caption(tr(f"{len(frame):,} data rows · {len(frame.columns)} columns",
                  f"{len(frame):,} saf oo xog ah · {len(frame.columns)} tiir"))
    st.dataframe(frame.head(30).astype("string"), hide_index=True, width="stretch")
    schema = pd.DataFrame([{"Column": name, "Type": guess_type(frame[name]),
                            "Date format": "%Y-%m-%d", "Unit": "%" if is_percent(frame[name]) else ""}
                           for name in frame])
    edited = st.data_editor(schema, key=f"schema-{key}", hide_index=True, disabled=["Column"],
                            column_config={"Type": st.column_config.SelectboxColumn(
                                options=["Text", "Number", "Date"], required=True)})
    st.caption(tr("Set a date format such as %d/%m/%Y. Units are labels; the unit % shows 0.125 as 12.5%.",
                  "Dooro qaabka taariikhda sida %d/%m/%Y. Halbeegyadu waa calaamado; halbeegga % wuxuu 0.125 u muujiyaa 12.5%."))
    settings = {row["Column"]: (row["Type"], row["Date format"] or "%Y-%m-%d", row["Unit"] or "")
                for row in edited.to_dict("records")}
    normalize = st.checkbox(tr("Apply category cleanup", "Midee qaybaha isku magaca ah"),
                            help=tr("Trim spaces and merge case variants.", "Ka saar bannaanada oo midee faraha kala duwan."), key=f"normalize-{key}")
    clear_invalid = st.checkbox(tr("Replace invalid values with blanks", "Qiyamka khaldan ka dhig bannaan"), key=f"invalid-{key}")
    deduplicate = st.checkbox(tr("Remove duplicate rows", "Ka saar safafka soo noqnoqda"), key=f"duplicates-{key}")
    cleaned, report, invalid = prepare(frame, settings, normalize, clear_invalid, deduplicate)
    st.dataframe(report, hide_index=True, width="stretch")
    st.caption(tr(f"Original duplicate rows: {frame.duplicated().sum():,}. Rows removed: {len(frame) - len(cleaned):,}. Unusual numbers are flagged using the 1.5 × IQR rule.",
                  f"Safafka soo noqnoqday: {frame.duplicated().sum():,}. Safafka laga saaray: {len(frame) - len(cleaned):,}. Qiyamka aan caadiga ahayn waxaa lagu calaamadeeyaa 1.5 × IQR."))
    issues = [column for column in invalid if not invalid[column].empty]
    if issues:
        column = st.selectbox(tr("Inspect invalid values", "Eeg qiyamka khaldan"), issues, key=f"issues-{key}")
        st.dataframe(invalid[column].head(30).astype("string").rename(column))
        st.caption(tr("If invalid values are kept, the affected column remains in its original format.",
                      "Haddii qiyamka khaldan la hayo, tiirka wuxuu sii ahaanayaa qaabkiisii hore."))

numeric = list(cleaned.select_dtypes(include="number").columns)
dates = list(cleaned.select_dtypes(include="datetime").columns)
categories = [name for name in cleaned if name not in numeric + dates]
with dashboard:
    filtered = cleaned
    filter_column = st.selectbox(tr("Filter column", "Tiirka shaandhaynta"), [None, *cleaned.columns],
                                 format_func=lambda name: name if name is not None else tr("All rows", "Dhammaan safafka"), key=f"filter-{key}")
    if filter_column is not None:
        values = cleaned[filter_column]
        if filter_column in dates and values.notna().any():
            start, end = values.min().date(), values.max().date()
            interval = st.date_input(tr("Date range", "Muddada"), (start, end), min_value=start, max_value=end,
                                     key=f"range-{key}-{filter_column}-{start}-{end}")
            if len(interval) == 2:
                filtered = cleaned[values.dt.date.between(*interval)]
        elif filter_column in numeric and values.notna().any():
            low, high = float(values.min()), float(values.max())
            if low < high:
                bounds = st.slider(tr("Value range", "Xadka qiimaha"), low, high, (low, high),
                                   key=f"bounds-{key}-{filter_column}-{low}-{high}")
                filtered = cleaned[values.between(*bounds)]
        elif filter_column in categories:
            labels = values.astype("string")
            if labels.nunique() > 500:
                search = st.text_input(tr("Contains", "Ka kooban"), key=f"contains-{key}-{filter_column}")
                if search:
                    filtered = cleaned[labels.str.contains(search, case=False, regex=False, na=False)]
            else:
                chosen = st.multiselect(tr("Keep these values", "Hay qiyamkan"), sorted(labels.dropna().unique()),
                                        key=f"choices-{key}-{filter_column}")
                if chosen:
                    filtered = cleaned[labels.isin(chosen)]

    metrics = st.columns(3)
    metrics[0].metric(tr("Rows in view", "Safafka muuqda"), f"{len(filtered):,}")
    metrics[1].metric(tr("Columns", "Tiirar"), len(filtered.columns))
    metrics[2].metric(tr("Blank cells", "Unugyo bannaan"), f"{filtered.isna().sum().sum():,}")
    if filtered.empty:
        st.info(tr("No rows match this view.", "Shaandhayntan saf uma soo bixin."))
    else:
        measure = None
        if numeric:
            measure = st.selectbox(tr("Numeric measure", "Cabbirka tirada"), numeric, key=f"measure-{key}")
            numbers = filtered[measure].dropna()
            unit = settings[measure][2]
            if not numbers.empty:
                totals = st.columns(2)
                totals[0].metric(tr("Total", "Wadar"), show(numbers.sum(), unit))
                totals[1].metric(tr("Average", "Celcelis"), show(numbers.mean(), unit))
        if categories:
            category = st.selectbox(tr("Compare by category", "Isbarbardhig qaybaha"), categories, key=f"category-{key}")
            _, chart = analyze(filtered, "group_sum" if measure is not None and filtered[measure].notna().any() else "group_count",
                               measure, category)
            if not chart.empty:
                labels = chart.index.to_series().astype("string").fillna(tr("(Blank)", "(Bannaan)"))
                chart.index = labels.to_list()
                st.bar_chart(percent_chart(chart.head(15)))
        if dates and measure is not None and filtered[measure].notna().any():
            date_column = st.selectbox(tr("Date for monthly trend", "Taariikhda isbeddelka bil kasta"), dates, key=f"date-{key}")
            _, trend = analyze(filtered, "trend", measure, date_column)
            if not trend.empty:
                st.line_chart(percent_chart(trend))
        st.dataframe(filtered.head(100), hide_index=True, width="stretch",
                     column_config=percent_columns(filtered))

with assistant:
    ai_enabled = st.checkbox(tr("Use OpenAI to interpret my question", "Adeegso OpenAI si loo fahmo su’aashayda"),
                             disabled=not bool(os.getenv("OPENAI_API_KEY")))
    st.caption(tr("Optional AI sends your question and column names and types. Calculations happen locally. Without an API key, use the example question patterns below.",
                  "OpenAI waxaa loo diraa su’aasha, magacyada tiirarka iyo noocyadooda. Xisaabtu waxay ka dhacdaa kombiyuutarkaaga. Furaha API la’aantiis, raac tusaalooyinkan."))
    st.caption(tr("Examples: Overview · Total Sales by Region · Average Sales · Monthly trend Sales by Date",
                  "Tusaale: Dulmar · Wadarta Sales by Region · Celcelis Sales · Bil kasta Sales by Date"))
    question = st.text_input(tr("Ask about the filtered data", "Weydii wax ku saabsan xogta la shaandheeyay"), max_chars=1000)
    if st.button(tr("Analyze question", "Falanqee su’aasha"), type="primary", disabled=not question.strip()):
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
                if operation == "profile" and percent_columns(detail):
                    # The count row is a number of rows, not a percentage.
                    detail = percent_chart(detail.drop(index="count")).round(4)
                    st.caption(tr(f"Observations used: {summary['n']:,}", f"Tirada la xisaabiyey: {summary['n']:,}"))
                st.dataframe(detail, column_config=percent_columns(detail))

with exports:
    st.download_button(tr("Download original file", "Soo dejiso faylka asalka ah"), content,
                       uploaded.name, "application/octet-stream")
    export_format = st.radio(tr("Export format", "Nooca faylka"), ["Excel (.xlsx)", "CSV"], horizontal=True)
    as_excel = export_format.startswith("Excel")
    export = excel_bytes if as_excel else csv_bytes
    extension = "xlsx" if as_excel else "csv"
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if as_excel else "text/csv"
    if st.button(tr("Prepare downloads", "Diyaari faylasha")):
        st.download_button(tr("Download cleaned data", "Soo dejiso xogta la hagaajiyey"), export(cleaned),
                           f"cleaned.{extension}", mime, on_click="ignore")
        st.download_button(tr("Download filtered data", "Soo dejiso xogta la shaandheeyey"), export(filtered),
                           f"filtered.{extension}", mime, on_click="ignore")
