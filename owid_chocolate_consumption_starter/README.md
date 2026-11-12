# OWID Cocoa Supply – Applied Statistics Project Starter

## Files expected
Place the supplied `chocolate-consumption-per-person.csv` in `data/raw/`.

## Run in Codespaces
```bash
python -m pip install --upgrade pip
pip install pandas numpy matplotlib openpyxl python-docx
python scripts/build_analysis.py
```

## Generated outputs
- `data/processed/chocolate_consumption_clean.csv`
- `excel/chocolate_consumption_analysis.xlsx`
- `report/CIE3_Statistical_Analysis_Report.docx`
- `charts/*.png`
- `build_summary.json`

## Important analytical notes
The indicator is per-capita food supply of cocoa and specified cocoa products, not measured actual intake. The full dataset is a country/region-year panel. The script calculates a pooled descriptive Pearson correlation for country/territory rows but explicitly warns that repeated observations are not independent and that the pooled result is not causal. Review and improve the analysis before submission.

The script is a starting point, not a guarantee of rubric compliance. Check all workbook formulas in Excel/LibreOffice and inspect the DOCX formatting and charts.
