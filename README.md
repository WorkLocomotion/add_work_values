# Work Values Enrichment
[![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/WorkLocomotion/enrich_work_values/blob/main/Enrich_Work_Values.ipynb)

Enrich a matched list of company job titles with O*NET Work Values scores and export a consolidated Excel file.

## What this does
- Reads **Company Job Titles - Matched.xlsx** (or similarly named file) containing your mapped titles and SOC codes.
- Downloads/reads **Work Values.xlsx** (from a GitHub raw URL or local file).
- Merges the six O*NET Work Values (Achievement, Independence, Recognition, Relationships, Support, Working Conditions) by **SOC Code**.
- Writes **Company Job Titles - Mapped.with_work_values.xlsx**.

## Files
- `enrich_work_values.py` — core script.
- `Enrich_Work_Values.ipynb` — Colab-ready notebook.
- `sample_input.xlsx` — input template.
- `requirements.txt` — dependencies.

## Expected columns in your input Excel
One sheet with headers (case-insensitive accepted):
- `Job Title` or `Job Titles`
- `HeadCount` (optional; defaults to 1 if missing)
- `O*NET SOC Code` (required)
- `Occupational Title` (optional; falls back to Job Title if absent)

## O*NET Work Values workbook
Provide a GitHub **raw** URL to your `ONET Work Values.xlsx`. The sheet must include:
- A column for SOC code (e.g., `SOC Code`, `SOC`, `Code`, etc.)
- Six columns named (or close variants of):  
  `Achievement`, `Independence`, `Recognition`, `Relationships`, `Support`, `Working Conditions`.
The script auto-detects reasonable variants (underscores/no-spaces).

## Run locally
## Run locally (Windows PowerShell)
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

python .\enrich_work_values.py `
  --input_excel ".\Company Job Titles - Mapped.xlsx" `
  --onet_values_url ".\Work Values.xlsx" `
  --output_excel ".\Company Job Titles - Mapped.with_work_values.xlsx"
```

## Run in Google Colab
1. Open `notebooks/Enrich_Work_Values.ipynb` in Colab.
2. Upload your input Excel when prompted.
3. Set `ONET_VALUES_URL` to your GitHub raw link.
4. Run all cells; a download will start for the enriched Excel.

## Notes
- SOC codes are normalized best-effort to O*NET style.
- If an output filename is locked, the script auto-increments the name.
- All input columns are preserved after the main fields and values.

---

WORK LOCOMOTION: Make Potential Actual
