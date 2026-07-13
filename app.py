from __future__ import annotations

import io
from pathlib import Path

import streamlit as st

from dispatch_core import DispatchError, build_export, read_workpex, transform

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "assets" / "UploadFormat.xlsx"

st.set_page_config(page_title="Workpex Dispatch Dashboard", page_icon="📦", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 2rem; padding-bottom: 3rem; max-width: 1450px;}
.hero {padding: 1.45rem 1.6rem; border: 1px solid #e5e7eb; border-radius: 18px; background: linear-gradient(135deg,#ffffff,#f8fafc); margin-bottom: 1rem;}
.hero h1 {margin:0; font-size:2rem;}
.hero p {margin:.45rem 0 0; color:#475569;}
.status-card {padding:1rem 1.1rem; border:1px solid #e5e7eb; border-radius:14px; background:#fff; min-height:112px;}
.status-card h4 {margin:0 0 .4rem;}
.status-card p {color:#64748b; margin:0; font-size:.9rem;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>Workpex Dispatch Dashboard</h1>
  <p>Upload the Workpex Excel export, verify the filtered orders, and download courier-ready Excel files.</p>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Upload Workpex order export",
    type=["xlsx", "xls"],
    help="Expected format: workpexexports.xlsx",
)

if not uploaded:
    st.info("Upload a Workpex export to activate dispatch downloads.")
    cols = st.columns(5)
    names = ["KSA Dispatch", "Oud Al Salam Dispatch", "UAE Dispatch", "Bahrain Dispatch", "Qatar Dispatch"]
    for col, name in zip(cols, names):
        with col:
            st.markdown(
                f'<div class="status-card"><h4>{name}</h4><p>Waiting for an uploaded file.</p></div>',
                unsafe_allow_html=True,
            )
    st.stop()

raw_bytes = uploaded.getvalue()
try:
    source_df = read_workpex(io.BytesIO(raw_bytes))
except DispatchError as exc:
    st.error(str(exc))
    st.stop()

uae_df = transform(source_df, "uae")
oas_df = transform(source_df, "oud_al_salam")
bahrain_df = transform(source_df, "bahrain")
qatar_df = transform(source_df, "qatar")

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Uploaded orders", f"{len(source_df):,}")
m2.metric("UAE dispatch", f"{len(uae_df):,}")
m3.metric("Oud Al Salam", f"{len(oas_df):,}")
m4.metric("Bahrain", f"{len(bahrain_df):,}")
m5.metric("Qatar", f"{len(qatar_df):,}")

st.subheader("Dispatch exports")
buttons = st.columns(5)

with open(TEMPLATE_PATH, "rb") as template:
    uae_result = build_export(io.BytesIO(raw_bytes), template, "uae")
with open(TEMPLATE_PATH, "rb") as template:
    oas_result = build_export(io.BytesIO(raw_bytes), template, "oud_al_salam")
with open(TEMPLATE_PATH, "rb") as template:
    bahrain_result = build_export(io.BytesIO(raw_bytes), template, "bahrain")
with open(TEMPLATE_PATH, "rb") as template:
    qatar_result = build_export(io.BytesIO(raw_bytes), template, "qatar")

with buttons[0]:
    st.button("KSA Dispatch", disabled=True, width="stretch", help="Reserved for the next phase")
    st.caption("Coming later")
with buttons[1]:
    st.download_button(
        "Oud Al Salam Dispatch",
        data=oas_result.workbook_bytes,
        file_name="Oud_Al_Salam_Dispatch.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        disabled=oas_result.exported_rows == 0,
    )
    st.caption("UAE + Al Huda / Premium / Luminex")
with buttons[2]:
    st.download_button(
        "UAE Dispatch",
        data=uae_result.workbook_bytes,
        file_name="UAE_Dispatch.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        disabled=uae_result.exported_rows == 0,
    )
    st.caption("UAE excluding Oud Al Salam products")
with buttons[3]:
    st.download_button(
        "Bahrain Dispatch",
        data=bahrain_result.workbook_bytes,
        file_name="Bahrain_Dispatch.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        disabled=bahrain_result.exported_rows == 0,
    )
    st.caption("Country = Bahrain")
with buttons[4]:
    st.download_button(
        "Qatar Dispatch",
        data=qatar_result.workbook_bytes,
        file_name="Qatar_Dispatch.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        disabled=qatar_result.exported_rows == 0,
    )
    st.caption("Country = Qatar")

st.divider()
preview_options = {
    "UAE Dispatch": uae_df,
    "Oud Al Salam Dispatch": oas_df,
    "Bahrain Dispatch": bahrain_df,
    "Qatar Dispatch": qatar_df,
}
preview_type = st.radio("Preview", list(preview_options), horizontal=True)
preview_df = preview_options[preview_type]
if preview_df.empty:
    st.warning("No matching rows were found for this dispatch type.")
else:
    st.dataframe(preview_df, width="stretch", hide_index=True, height=480)

with st.expander("Current transformation rules"):
    st.markdown("""
- **UAE Dispatch:** UAE orders excluding Al Huda, Premium Edition/Collection, and Luminex/Luminux.
- **Oud Al Salam Dispatch:** UAE orders containing Al Huda, Premium Edition/Collection, or Luminex/Luminux.
- **Bahrain Dispatch:** only rows where Country is Bahrain, aligned to the uploaded Bahrain 16-column format.
- **Qatar Dispatch:** only rows where Country is Qatar, aligned to the uploaded Qatar 16-column format.
- UAE phone numbers export as 9-digit local numbers. Qatar and Bahrain phone numbers export as 8-digit local numbers.
- Any repeated final exported phone number is treated as a duplicate and every matching row is highlighted light red.
- KSA remains reserved for the next phase.
""")
