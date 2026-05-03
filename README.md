# 2xFR4 Curve Filler

A Streamlit web app that automatically fills raw optical transceiver test data (PC & WRP) into the `TEC_2xFR4_Curve` Excel template.

## Features

- **PC Raw Data** — pairs Operational + Maximum CSVs per SN → one Excel per SN
- **WRP Raw Data** — splits by `TESTNUMBER` → separate Excel per test group
  - Group 1 → `SN_WRP.xlsx`
  - Group 2 → `SN_WRP_1.xlsx`
  - Group 3 → `SN_WRP_2.xlsx`
- Accepts **CSV** files or **ZIP** archives
- Auto-extends formula rows in `Operational`, `Maximum`, `Curve` sheets
- Download individual files or all as a single ZIP

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Cloud

1. Fork / push this repo to your GitHub account
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → `app.py`
4. Click **Deploy**

## File Naming Convention

| Input | Output |
|-------|--------|
| `P204261000290_Operational Current_*.csv` + `P204261000290_Maximum Current_*.csv` | `TEC_2xFR4_Curve_P204261000290_PC.xlsx` |
| `P204261000290.csv` (1 TESTNUMBER) | `TEC_2xFR4_Curve_P204261000290_WRP.xlsx` |
| `P204261000290.csv` (2 TESTNUMBERs) | `TEC_2xFR4_Curve_P204261000290_WRP.xlsx` + `..._WRP_1.xlsx` |

## Requirements

- Python 3.9+
- streamlit
- pandas
- openpyxl
