
# Purpose Built - Fulfillment Extractor

Streamlit app to pull a specific **tab** from a shared Google Sheet (orders), merge with **store_master_data**,
map to the exact schema of **fulfillment_sheet_proto_V1**, and export a ShipStation-ready CSV named
`Purpose_Built_Fulfillment_Sheet_YYYY-MM-DD.csv`.

## Quickstart
```bash
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Flow
1. Paste Google Sheets share link → Fetch tabs → Pick your *new order* tab.
2. Upload `store_master_data.xls/.xlsx` (optional if you don't need enrichment).
3. Upload `fulfillment_sheet_proto_V1.xls/.xlsx` to capture the exact output column order (or type columns manually).
4. Map **source** columns to **target** columns via the UI. Provide defaults as needed.
5. (Optional) Join order tab with store master on keys you select.
6. Export CSV: `Purpose_Built_Fulfillment_Sheet_YYYY-MM-DD.csv`.

## Private Sheets
- For private Google Sheets, ensure your link allows access or we can add Google OAuth/Service Account.
  (Next step if you need it.)

## Notes
- Legacy `.xls` reading requires `xlrd`. If reading fails locally, run: `pip install xlrd`.
