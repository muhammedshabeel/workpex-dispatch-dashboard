from __future__ import annotations

import io
import re
from dataclasses import dataclass

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADERS = [
    "client_order_ref", "customer_name", "partner_id", "whatsapp_no", "source_id",
