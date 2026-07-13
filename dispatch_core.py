from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook

TEMPLATE_HEADERS = [
    "CUSTOMER_NAME", "MOBILE_NO", "LANDLINE_NO", "ADDRESS_1", "ADDRESS_2",
    "ADDRESS_3", "FLAT/VILLA NO", "DELIVERY_CITY", "COD_AMOUNT", "REMARKS",
    "REFERENCE_NO", "OTHER_REMARKS", "COLLECT_RETURN (YES/NO)",
]

UAE_NAMES = {"uae", "united arab emirates", "united arab emirate", "emirates"}
OAS_PRODUCT_PATTERNS = (
    r"\bal\s*huda\b",
    r"\bpremium\s*(edition|collection)?\b",
    r"\blumin(e|u)x\b",
)

COLUMN_ALIASES = {
    "name": ["Lead Name", "Customer Name", "Name", "First Name"],
    "primary_phone": ["Primary Phone", "Phone 1", "Phone1", "Phone", "Mobile", "Mobile No"],
    "secondary_phone": ["Secondary Phone", "Phone 2", "Phone2", "Alternate Phone", "WhatsApp Number"],
    "country": ["Country"],
    "state": ["State", "Emirate", "Province"],
    "street": ["Street", "Address", "Address 1"],
    "city": ["CITY", "City", "Delivery City"],
    "amount": ["Actual Amount", "Forecasted Amount", "Amount", "COD Amount"],
    "payment": ["Payment Method", "Payment"],
    "product": ["Product", "Product Name"],
    "qty": ["QTY", "Quantity"],
    "product2": ["PRODUCT 2", "Product 2"],
    "qty2": ["QTY OF PRODUCT 2", "Quantity of Product 2"],
    "description": ["Lead Description", "Remarks", "Notes"],
    "national_code": ["National Code", "Reference No", "Order ID"],
}


class DispatchError(ValueError):
    pass


@dataclass(frozen=True)
class ExportResult:
    dataframe: pd.DataFrame
    workbook_bytes: bytes
    source_rows: int
    exported_rows: int


def _clean(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalized(value) -> str:
    return re.sub(r"\s+", " ", _clean(value).lower()).strip()


def _find_column(df: pd.DataFrame, aliases: Iterable[str], required: bool = False) -> str | None:
    exact = {_normalized(c): c for c in df.columns}
    for alias in aliases:
        key = _normalized(alias)
        if key in exact:
            return exact[key]
    if required:
        raise DispatchError(f"Required column missing. Expected one of: {', '.join(aliases)}")
    return None


def read_workpex(file) -> pd.DataFrame:
    try:
        df = pd.read_excel(file, dtype=object)
    except Exception as exc:
        raise DispatchError(f"Could not read the uploaded Excel file: {exc}") from exc
    if df.empty:
        raise DispatchError("The uploaded Excel file contains no order rows.")
    df.columns = [str(c).strip() for c in df.columns]
    _find_column(df, COLUMN_ALIASES["country"], required=True)
    _find_column(df, COLUMN_ALIASES["product"], required=True)
    _find_column(df, COLUMN_ALIASES["name"], required=True)
    return df


def _is_uae(value) -> bool:
    return _normalized(value) in UAE_NAMES


def _is_oas_product(value) -> bool:
    text = _normalized(value).replace("_", " ").replace("-", " ")
    return any(re.search(pattern, text, flags=re.I) for pattern in OAS_PRODUCT_PATTERNS)


def _digits(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = _clean(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\D", "", text)


def _as_uae_local(value) -> str:
    number = _digits(value)
    if number.startswith("00971"):
        number = number[5:]
    elif number.startswith("971"):
        number = number[3:]
    elif number.startswith("0"):
        number = number[1:]
    return number if len(number) == 9 and number.startswith("5") else ""


def _uae_mobile(primary, secondary) -> str:
    return _as_uae_local(primary) or _as_uae_local(secondary)


def _money(value) -> float:
    text = _clean(value).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group()) if match else 0.0


def _cod_amount(amount, payment) -> float:
    payment_text = _normalized(payment)
    cod_methods = {"cod", "cash on delivery", "cash-on-delivery"}
    return _money(amount) if payment_text in cod_methods else 0.0


def _join_nonempty(*values, sep=" | ") -> str:
    parts = [_clean(v) for v in values if _clean(v)]
    return sep.join(parts)


def transform(df: pd.DataFrame, dispatch_type: str) -> pd.DataFrame:
    cols = {key: _find_column(df, aliases, required=key in {"name", "country", "product"})
            for key, aliases in COLUMN_ALIASES.items()}

    country_series = df[cols["country"]].map(_is_uae)
    product_series = df[cols["product"]].map(_is_oas_product)

    if dispatch_type == "uae":
        # UAE Dispatch must exclude Oud Al Salam products to prevent duplicate dispatching.
        filtered = df[country_series & ~product_series].copy()
    elif dispatch_type == "oud_al_salam":
        filtered = df[country_series & product_series].copy()
    else:
        raise DispatchError(f"Unsupported dispatch type: {dispatch_type}")

    def get(row, key):
        col = cols.get(key)
        return row[col] if col else ""

    records = []
    for index, row in filtered.iterrows():
        product = get(row, "product")
        qty = get(row, "qty")
        product2 = get(row, "product2")
        qty2 = get(row, "qty2")
        payment = get(row, "payment")
        description = get(row, "description")
        state = get(row, "state")
        city = get(row, "city")
        street = get(row, "street")
        reference = _clean(get(row, "national_code")) or f"WPX-{index + 2}"

        product_note = _join_nonempty(
            f"{_clean(product)} x {_clean(qty) or '1'}" if _clean(product) else "",
            f"{_clean(product2)} x {_clean(qty2) or '1'}" if _clean(product2) else "",
        )
        remarks = _join_nonempty(product_note, f"Payment: {_clean(payment)}" if _clean(payment) else "")

        records.append({
            "CUSTOMER_NAME": _clean(get(row, "name")),
            "MOBILE_NO": _uae_mobile(get(row, "primary_phone"), get(row, "secondary_phone")),
            "LANDLINE_NO": "",
            "ADDRESS_1": _clean(street),
            "ADDRESS_2": _clean(state),
            "ADDRESS_3": "",
            "FLAT/VILLA NO": "",
            "DELIVERY_CITY": _clean(city) or _clean(state),
            "COD_AMOUNT": _cod_amount(get(row, "amount"), payment),
            "REMARKS": remarks,
            "REFERENCE_NO": reference,
            "OTHER_REMARKS": _clean(description),
            "COLLECT_RETURN (YES/NO)": "NO",
        })

    return pd.DataFrame(records, columns=TEMPLATE_HEADERS)


def create_workbook(export_df: pd.DataFrame, template_file) -> bytes:
    template_bytes = template_file.read() if hasattr(template_file, "read") else bytes(template_file)
    if hasattr(template_file, "seek"):
        template_file.seek(0)
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb["Upload Format"] if "Upload Format" in wb.sheetnames else wb.active

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for col_index, header in enumerate(TEMPLATE_HEADERS, start=1):
        ws.cell(row=1, column=col_index, value=header)

    for row_index, values in enumerate(export_df.itertuples(index=False, name=None), start=2):
        for col_index, value in enumerate(values, start=1):
            ws.cell(row=row_index, column=col_index, value=value)

    for row_index in range(2, ws.max_row + 1):
        ws.cell(row=row_index, column=2).number_format = "@"
        ws.cell(row=row_index, column=11).number_format = "@"
        ws.cell(row=row_index, column=9).number_format = "0.00"

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def build_export(source_file, template_file, dispatch_type: str) -> ExportResult:
    source_df = read_workpex(source_file)
    export_df = transform(source_df, dispatch_type)
    workbook = create_workbook(export_df, template_file)
    return ExportResult(export_df, workbook, len(source_df), len(export_df))
