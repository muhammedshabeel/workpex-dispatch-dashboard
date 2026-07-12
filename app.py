from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
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

uploaded = st.file_uploader("Upload Workpex order export", type=["xlsx", "xls"], help="Expected format: workpexexports.xlsx")

if not uploaded:
    st.info("Upload a Workpex export to activate UAE and Oud Al Salam dispatch downloads.")
    cols = st.columns(5)
    names = ["KSA Dispatch", "Oud Al Salam Dispatch", "UAE Dispatch", "Bahrain Dispatch", "Qatar Dispatch"]
    for col, name in zip(cols, names):
        with col:
            st.markdown(f'<div class="status-card"><h4>{name}</h4><p>Waiting for an uploaded file.</p></div>', unsafe_allow_html=True)
    st.stop()

raw_bytes = uploaded.getvalue()
try:
    source_df = read_workpex(io.BytesIO(raw_bytes))
except DispatchError as exc:
    st.error(str(exc))
    st.stop()

uae_df = transform(source_df, "uae")
oas_df = transform(source_df, "oud_al_salam")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Uploaded orders", f"{len(source_df):,}")
m2.metric("UAE orders", f"{len(uae_df):,}")
m3.metric("Oud Al Salam orders", f"{len(oas_df):,}")
m4.metric("Other-country orders", f"{len(source_df) - len(uae_df):,}")

st.subheader("Dispatch exports")
buttons = st.columns(5)

with open(TEMPLATE_PATH, "rb") as template:
    uae_result = build_export(io.BytesIO(raw_bytes), template, "uae")
with open(TEMPLATE_PATH, "rb") as template:
    oas_result = build_export(io.BytesIO(raw_bytes), template, "oud_al_salam")

with buttons[0]:
    st.button("KSA Dispatch", disabled=True, use_container_width=True, help="Reserved for the next phase")
    st.caption("Coming later")
with buttons[1]:
    st.download_button(
        "Oud Al Salam Dispatch",
        data=oas_result.workbook_bytes,
        file_name="Oud_Al_Salam_Dispatch.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=oas_result.exported_rows == 0,
    )
    st.caption("UAE + Al Huda / Premium / Luminex")
with buttons[2]:
    st.download_button(
        "UAE Dispatch",
        data=uae_result.workbook_bytes,
        file_name="UAE_Dispatch.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=uae_result.exported_rows == 0,
    )
    st.caption("All UAE orders")
with buttons[3]:
    st.button("Bahrain Dispatch", disabled=True, use_container_width=True, help="Reserved for the next phase")
    st.caption("Coming later")
with buttons[4]:
    st.button("Qatar Dispatch", disabled=True, use_container_width=True, help="Reserved for the next phase")
    st.caption("Coming later")

st.divider()
preview_type = st.radio("Preview", ["UAE Dispatch", "Oud Al Salam Dispatch"], horizontal=True)
preview_df = uae_df if preview_type == "UAE Dispatch" else oas_df
if preview_df.empty:
    st.warning("No matching rows were found for this dispatch type.")
else:
    st.dataframe(preview_df, use_container_width=True, hide_index=True, height=480)

with st.expander("Current transformation rules"):
    st.markdown("""
- **UAE Dispatch:** every row where Country is United Arab Emirates/UAE.
- **Oud Al Salam Dispatch:** UAE rows whose Product contains Al Huda, Premium Edition/Collection, or Luminex/Luminux.
- Phone 1 is used only when it is a valid UAE mobile. If Phone 1 is foreign or invalid, Phone 2 is used. The export contains only 9-digit UAE-local mobile numbers without `971` or the leading `0`.
- COD amount keeps the order value only when Payment Method is COD/Cash on Delivery. Every other payment method exports `0`.
- KSA, Bahrain, and Qatar buttons are placeholders for the next phase.
""")
