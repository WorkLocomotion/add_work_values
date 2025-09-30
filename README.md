# Add Work Values

[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/WorkLocomotion/add_work_values/blob/main/Enrich_Work_Values.ipynb)

Enrich a list of mapped company job titles with O*NET **Work Values** and export a consolidated Excel file.

## What this does
- Reads **Company Job Titles – Mapped.xlsx** with titles and SOC codes.
- Uses the hosted **Work Values.xlsx** (or a custom GitHub RAW URL/local file if you prefer).
- Merges six Work Values by SOC code: **Achievement, Independence, Recognition, Relationships, Support, Working Conditions**.
- Writes **Company Job Titles – Mapped.with_work_values.xlsx**.

## Repository layout
- `enrich_work_values.py` — CLI script (root-level).
- `Enrich_Work_Values.ipynb` — Colab notebook (one-click).
- `templates/sample_input.xlsx` — example input template.
- `requirements.txt`, `LICENSE`, `CHANGELOG.md`.

## Expected input columns
Single sheet; headers are case-insensitive. Accepted:
- `Job Title` **or** `Job Titles`
- `HeadCount` *(or `Headcount`; optional, defaults to 1 if missing)*
- SOC code column (any one of): `O*NET-SOC Code`, `O*NET SOC Code`, `SOC Code`, `SOC`, `Code`
- `Occupational Title` *(optional; falls back to Job Title if absent)*

> The script normalizes SOC variants and the **output field is `SOC Code`**.

## O*NET Work Values workbook
Default: uses the copy hosted in this repo (`Work Values.xlsx`).  
Advanced: pass your own source via `--onet_values_url` (local path or GitHub RAW).

## Run in Google Colab (recommended)
Click the badge above and **Run all**.  
When prompted, upload your `Company Job Titles – Mapped.xlsx`. The enriched file will download automatically.

## Run locally (Windows PowerShell)
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

python .\enrich_work_values.py `
  --input_excel ".\Company Job Titles - Mapped.xlsx" `
  --onet_values_url "https://raw.githubusercontent.com/WorkLocomotion/add_work_values/main/Work%20Values.xlsx" `
  --output_excel ".\Company Job Titles - Mapped.with_work_values.xlsx"

---

WORK LOCOMOTION: Make Potential Actual
