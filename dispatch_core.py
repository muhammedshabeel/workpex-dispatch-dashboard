from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

TEMPLATE_HEADERS = [
    "CUSTOMER_NAME", "MOBILE_NO", "LANDLINE_NO", "ADDRESS_1", "ADDRESS_2",
    "ADDRESS_3", "FLAT/VILLA NO", "DELIVERY_CITY", "COD_AMOUNT", "REMARKS",
    "REFERENCE_NO", "OTHER_REMARKS", "COLLECT_RETURN (YES/NO)",
]

UAE_NAMES = {"uae", "united arab emirates", "united arab emirate", "emirates"}
OAS_PRODUCT_PATTERNS = (
    r"\bal\s*huda\b",
    r"\bpremium\s*(edition|collection)?\b",
   