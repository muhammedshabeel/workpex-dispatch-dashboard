# Workpex Dispatch Dashboard

A Streamlit dashboard that converts a Workpex order export into courier upload files using the supplied `UploadFormat.xlsx` template.

## Current dispatch actions

- **UAE Dispatch:** exports all UAE orders.
- **Oud Al Salam Dispatch:** exports UAE orders for Al Huda, Premium Edition/Collection, and Luminex/Luminux.
- **KSA, Bahrain, Qatar:** visible placeholders, intentionally disabled for the next phase.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repository to GitHub.
2. Open Streamlit Community Cloud and create a new app.
3. Select this repository, branch `main`, and entry file `app.py`.
4. Deploy. No secrets are required.

## Export logic

The output retains the exact 13-column template structure. UAE phone numbers are converted to local 9-digit format. Clearly prepaid orders receive a COD amount of zero; COD orders retain the numeric amount from Workpex.

## Files

- `app.py` — dashboard UI
- `dispatch_core.py` — parsing, filtering, mapping, and Excel generation
- `assets/UploadFormat.xlsx` — courier template
- `tests/test_dispatch.py` — core validation
