# Excel Data Insights

Open this folder in VS Code or Visual Studio with Python support. Use Python 3.10 or newer and open a terminal in the folder containing app.py.

## Windows PowerShell

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py --server.address 127.0.0.1
```

## macOS or Linux

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py --server.address 127.0.0.1
```

Open http://localhost:8501 if your browser does not open automatically.

## Try the portal

1. Upload sample.csv, a UTF-8 CSV, or an .xlsx workbook. Select a worksheet or CSV separator.
2. In **Review data**, choose the header row and review detected column types. Select cleaning options to apply them. If a numeric or date column contains invalid cells, approve replacing invalid cells with blanks to enable calculations on its valid values.
3. Open **Dashboard** to filter the active data and view totals, category comparisons, and monthly trends.
4. Try **Ask the agent** with `Overview`, `Total Sales by Region`, `Average Sales`, `Monthly trend Sales by Date`, or `Correlation Sales and Orders`. Use the dashboard to apply conditions before asking a question. In Somali, try `Dulmar`, `Wadarta Sales`, or `Celcelis Sales`.
5. Open **Downloads** and choose the original, cleaned, or filtered data. Excel exports preserve strings as text. CSV exports add a leading apostrophe to text that a spreadsheet could interpret as a formula.

## Optional OpenAI mode

Set an OpenAI API key in the same terminal before starting the portal.

```powershell
$env:OPENAI_API_KEY = "your-api-key"
```

```bash
export OPENAI_API_KEY="your-api-key"
```

Enable **Use OpenAI to interpret my question** in the agent tab. The question, column names and their types are sent to OpenAI for operation selection. Avoid entering confidential values in questions or using this mode with sensitive column names. Worksheet values and calculated results stay local. The default model is gpt-4.1-mini. Set OPENAI_MODEL to use another Responses API model. Local mode uses keyword matching, and both modes support a fixed set of analyses calculated in Python.

The portal runs locally without a database. Uploads are limited to 20 MB, 100,000 worksheet or CSV rows, 250 columns, and 1 million cells. The date parser defaults to ISO dates. For other date strings, set the input type to Date and provide an explicit format. The language selector translates navigation, guidance, and agent answers. Data and column names remain as uploaded.
