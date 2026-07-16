from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

UAE_HEADERS = [
    "CUSTOMER_NAME", "MOBILE_NO", "LANDLINE_NO", "ADDRESS_1", "ADDRESS_2",
    "ADDRESS_3", "FLAT/VILLA NO", "