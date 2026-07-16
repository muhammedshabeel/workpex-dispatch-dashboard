from __future__ import annotations

import io
from pathlib import Path

import streamlit as st

from dispatch_core import DispatchError, build_export, read_workpex, transform
from gcc_portal import build_gcc_export, transform_gcc
from ksa_portal import build_ksa_exports

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "assets" / "UploadFormat.xlsx"

st.set_page_config(page_title="Workpex Dispatch Dashboard