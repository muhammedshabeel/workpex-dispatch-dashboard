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
    "Pricelist Name", "street_no", "building_no", "zone_id (governarate)",
    "wilayat_id", "city_id", "order_line/product_id", "order_line/product_uom",
    "order_line/price_unit", "order_line/product_uom_qty", "remarks", "Agent Name", "Source",
]

SETTINGS = {
    "qatar": {
        "countries": {"qatar", "qa"},
        "code": "974",
        "prefix": "EMQ",
        "source_id": "SCENT PASSION - QATAR",
        "pricelist": "Public Pricelist",
    },
    "bahrain": {
        "countries": {"bahrain", "bh"},
        "code": "973",
        "prefix": "EMB",
        "source_id": "SCENT PASSION - BAHRAIN",
        "pricelist": "Default BHD pricelist",
    },
}

ALIASES = {
    "name": ["Lead Name", "Customer Name", "Name", "First Name"],
    "phone1": ["Primary Phone", "Phone 1", "Phone1", "Phone", "Mobile", "Mobile No"],
    "phone2": ["Secondary Phone", "Phone 2", "Phone2", "Alternate Phone", "WhatsApp Number"],
    "country": ["Country"],
    "street": ["Street", "Address", "Address 1"],
    "building": ["Building No", "Building Number", "Building", "Flat/Villa No"],
    "zone": ["Zone", "Zone ID", "Zone Id", "Governorate ID", "Governarate ID"],
    "state": ["State", "Province", "Governorate", "Governarate"],
    "city": ["CITY", "City", "Delivery City", "Wilayat", "Wilayat Name"],
    "product": ["Product", "Product Name"],
    "qty": ["QTY", "Quantity"],
    "amount": ["Actual Amount", "Forecasted Amount", "Amount", "COD Amount"],
    "remarks": ["Lead Description", "Remarks", "Notes"],
    "reference": ["National Code", "Reference No", "Order ID"],
    "agent_name": ["Assigned", "Assigned User"],
    "source": ["Source"],
}


@dataclass(frozen=True)
class GCCResult:
    dataframe: pd.DataFrame
    workbook_bytes: bytes
    exported_rows: int


def _clean(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _norm(value) -> str:
    return re.sub(r"\s+", " ", _clean(value).lower()).strip()


def _column(df: pd.DataFrame, key: str, required: bool = False) -> str | None:
    columns = {_norm(col): col for col in df.columns}
    for alias in ALIASES[key]:
        if _norm(alias) in columns:
            return columns[_norm(alias)]
    if required:
        raise ValueError(f"Required column missing for portal export: {ALIASES[key][0]}")
    return None


def _digits(value) -> str:
    text = _clean(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\D", "", text)


def _phone(value, country: str) -> str:
    number = _digits(value)
    code = SETTINGS[country]["code"]
    if number.startswith("00" + code):
        number = number[len(code) + 2:]
    elif number.startswith(code):
        number = number[len(code):]
    elif number.startswith("0"):
        number = number[1:]
    return number if len(number) == 8 else ""


def _number(value) -> float:
    text = _clean(value).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group()) if match else 0.0


def _quantity(value):
    number = _number(value)
    if number <= 0:
        return 1
    return int(number) if number.is_integer() else number


def _zone(value):
    text = _clean(value)
    return int(text) if re.fullmatch(r"\d+", text) else None


def transform_gcc(source_df: pd.DataFrame, country: str) -> pd.DataFrame:
    if country not in SETTINGS:
        raise ValueError(f"Unsupported GCC country: {country}")

    cols = {
        key: _column(source_df, key, required=key in {"name", "country", "product"})
        for key in ALIASES
    }
    mask = source_df[cols["country"]].map(lambda value: _norm(value) in SETTINGS[country]["countries"])
    filtered = source_df[mask].copy()
    config = SETTINGS[country]
    records = []

    def get(row, key):
        col = cols.get(key)
        return row[col] if col else ""

    for index, row in filtered.iterrows():
        phone = _phone(get(row, "phone1"), country) or _phone(get(row, "phone2"), country)
        zone = _zone(get(row, "zone")) or _zone(get(row, "state"))
        wilayat = _clean(get(row, "city"))
        if not wilayat and _zone(get(row, "state")) is None:
            wilayat = _clean(get(row, "state"))

        product = _clean(get(row, "product"))
        reference = _clean(get(row, "reference")) or f"{config['prefix']}{index + 2}"

        records.append({
            "client_order_ref": reference,
            "customer_name": _clean(get(row, "name")),
            "partner_id": int(phone) if phone else None,
            "whatsapp_no": int(phone) if phone else None,
            "source_id": config["source_id"],
            "Pricelist Name": config["pricelist"],
            "street_no": _clean(get(row, "street")),
            "building_no": _clean(get(row, "building")) or None,
            "zone_id (governarate)": zone,
            "wilayat_id": wilayat or None,
            "city_id": None,
            "order_line/product_id": product,
            "order_line/product_uom": "Units",
            "order_line/price_unit": _number(get(row, "amount")),
            "order_line/product_uom_qty": _quantity(get(row, "qty")),
            "remarks": _clean(get(row, "remarks")) or None,
            "Agent Name": _clean(get(row, "agent_name")) or None,
            "Source": _clean(get(row, "source")) or None,
        })

    return pd.DataFrame(records, columns=HEADERS)


def _workbook_bytes(export_df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    duplicate_fill = PatternFill(fill_type="solid", fgColor="FFC7CE")

    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(1, col, header)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    phones = export_df["partner_id"].fillna("").astype(str) if not export_df.empty else pd.Series(dtype=str)
    duplicate_mask = phones.ne("") & phones.duplicated(keep=False)

    for row_no, values in enumerate(export_df.itertuples(index=False, name=None), start=2):
        is_duplicate = bool(duplicate_mask.iloc[row_no - 2])
        for col_no, value in enumerate(values, start=1):
            cell = ws.cell(row_no, col_no, value)
            cell.font = Font(name="Calibri", size=11)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = border
            if is_duplicate:
                cell.fill = duplicate_fill

    widths = [18, 24, 15, 15, 28, 24, 62, 18, 24, 24, 16, 34, 25, 24, 28, 34, 22, 22]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width

    for row_no in range(2, ws.max_row + 1):
        ws.cell(row_no, 1).number_format = "@"
        ws.cell(row_no, 3).number_format = "0"
        ws.cell(row_no, 4).number_format = "0"
        ws.cell(row_no, 9).number_format = "0"
        ws.cell(row_no, 14).number_format = "0.00"
        ws.cell(row_no, 15).number_format = "0.##"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:R{max(1, ws.max_row)}"
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def build_gcc_export(source_df: pd.DataFrame, country: str) -> GCCResult:
    export_df = transform_gcc(source_df, country)
    return GCCResult(export_df, _workbook_bytes(export_df), len(export_df))
