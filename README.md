# File Splitter

Streamlit application that ingests any consolidated workbook (`.xlsx`) and produces downloadable Excel files—one per Team Lead or Mentor—plus a zipped bundle for convenience. The app automatically skips sheets that lack the selected header, normalizes common aliases, and lets you prepend a custom prefix to each exported filename. A companion **File Consolidator** page lets you merge multiple workbooks back together, generating fresh exports per Team Lead or Mentor.

## Run Locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run filesplit.py
```

## Deploy to Streamlit Cloud

1. Commit the project to a public Git repository containing `filesplit.py`, the `pages/` directory, `requirements.txt`, and optionally the `data/` directory for sample files.
2. Visit [https://streamlit.io/cloud](https://streamlit.io/cloud) and click **Deploy an app**.
3. Select the repository, branch, and set the main file to `filesplit.py` (or `app.py` if you rename it).
4. Confirm the Python version if needed, then click **Deploy**. Streamlit Cloud installs dependencies from `requirements.txt` automatically.

## Folder Structure

```
.
├── filesplit.py               # Main Streamlit page (File Splitter)
├── pages/
│   └── 1_Consolidate.py       # Additional page for merging exports
├── requirements.txt
├── README.md
└── data/
    └── sample_timesheet.xlsx  (optional example workbook)
```
