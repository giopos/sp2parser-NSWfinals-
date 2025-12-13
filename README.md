# Swim Meet PDF → Heats XLSX (Streamlit)

This repo contains:
- `pdf_to_heats_xlsx.py`: the parser + Excel workbook generator
- `app.py`: a Streamlit web UI for uploading a PDF and downloading outputs

## Features
- Upload a meet program PDF
- Automatically parses events/heats + alternates
- Download:
  - XLSX workbook (same as the CLI script output)
  - Heats CSV
  - Alternates CSV
- Preview tables in the browser
- **Copy all** (TSV) button for easy paste into Excel / Google Sheets

## Run locally

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL Streamlit prints (usually http://localhost:8501).

## Deploy on Streamlit Community Cloud
1. Push this repo to GitHub.
2. Go to https://streamlit.io/cloud and create a new app.
3. Select your repo and set:
   - **Main file path**: `app.py`
4. This repo includes `runtime.txt` with `python-3.11` so Streamlit Cloud uses Python 3.11.

## CLI usage (optional)
You can still run the original script directly:

```bash
python pdf_to_heats_xlsx.py "/path/to/program.pdf" "/path/to/output.xlsx"
```

(or run it without arguments in a folder containing PDFs to select interactively).
