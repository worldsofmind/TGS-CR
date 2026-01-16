# CR Dashboard Prototype (Streamlit)

This repo is a **prototype** Streamlit dashboard for internal stakeholders.
It recreates the **CR Overview** page first (KPIs + pies + Top 10 bar charts), and is designed to expand into additional pages/visualisations later.

## Files
- `app.py` – Streamlit app (prototype)
- `requirements.txt` – Python dependencies

## Data handling (prototype-safe)
✅ **Do not commit internal Excel data to GitHub.**

Instead, the app uses:
- **File upload** (`Upload Excel (.xlsx)`) as the default for prototypes

The uploaded Excel should contain:
- Sheet: **CR Master** (required)
- Sheet: **Look Up** (optional)

## Prototype features
- Sidebar filters:
  - Delivery Timeline (Year)
  - Cleaned Timeline (derived if missing in Excel)
  - Division
  - PO
- KPIs:
  - No of Change Requests (CR)
  - Estimated Effort (Man-days)
- Charts:
  - MoSCoW (pie)
  - CR Prep Status (pie)
  - Delivery Status (pie)
  - Change Requests by Divisions (**default Top 10 + Others (n=XX)**, with **Show all** toggle)
  - Change Requests by PO (**default Top 10 + Others (n=XX)**, with **Show all** toggle)

## Run locally

### 1) Create a virtual environment (recommended)
```bash
python -m venv .venv
# Windows
.\.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate
```

### 2) Install dependencies
```bash
pip install -r requirements.txt
```

### 3) Run Streamlit
```bash
streamlit run app.py
```

## Deploy (prototype) to Streamlit Community Cloud
1. Push this repo to GitHub
2. Create a Streamlit Cloud app
3. Select:
   - **Repository**: your repo
   - **Branch**: main
   - **Main file**: `app.py`
4. Deploy

> Note: confirm your organisation’s policies before uploading any sensitive data via the uploader.

## Next steps (when ready)
- Add more pages: Timeline / Division / Details / Data Quality
- Add an admin-only diagnostics panel:
  - “Values changed by standardisation”
  - “Possible duplicates detected”
- Add production data source options:
  - fixed internal file path
  - scheduled extract to a database
