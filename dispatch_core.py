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
    "ADDRESS_3", "FLAT/VILLA NO", "DELIVERY_CITY", "COD_AMOUNT", "REMARKS",
    "REFERENCE_NO", "OTHER_REMARKS", "COLLECT_RETURN (YES/NO)", "Agent Name", "Source",
]

GCC_HEADERS = [
    "client_order_ref", "customer_name", "partner_id", "whatsapp_no", "source_id",
    "Pricelist Name", "street_no", "building_no", "zone_id (governarate)",
    "wilayat_id", "city_id", "order_line/product_id", "order_line/product_uom",
    "order_line/price_unit", "order_line/product_uom_qty", "remarks", "Agent Name", "Source",
]

COUNTRY_NAMES = {
    "uae": {"uae", "united arab emirates", "united arab emirate", "emirates"},
    "qatar": {"qatar", "qa"},
    "bahrain": {"bahrain", "bh"},
}

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
    "state": ["State", "Emirate", "Province", "Zone"],
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
    "agent_name": ["Agent Name", "Agent", "Assigned Agent", "Sales Agent", "Owner Name"],
    "source": ["Source", "Lead Source", "Order Source", "Campaign Source"],
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
    for key in ("country", "product", "name"):
        _find_column(df, COLUMN_ALIASES[key], required=True)
    return df


def _is_country(value, country: str) -> bool:
    return _normalized(value) in COUNTRY_NAMES[country]


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


def _local_phone(value, country: str) -> str:
    number = _digits(value)
    code = {"uae": "971", "qatar": "974", "bahrain": "973"}[country]
    if number.startswith("00" + code):
        number = number[len(code) + 2:]
    elif number.startswith(code):
        number = number[len(code):]
    elif number.startswith("0"):
        number = number[1:]

    if country == "uae":
        return number if len(number) == 9 and number.startswith("5") else ""
    return number if len(number) == 8 else ""


def _best_phone(primary, secondary, country: str) -> str:
    return _local_phone(primary, country) or _local_phone(secondary, country)


def _money(value) -> float:
    text = _clean(value).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group()) if match else 0.0


def _cod_amount(amount, payment) -> float:
    payment_text = _normalized(payment)
    return _money(amount) if payment_text in {"cod", "cash on delivery", "cash-on-delivery"} else 0.0


def _join_nonempty(*values, sep=" | ") -> str:
    return sep.join(_clean(v) for v in values if _clean(v))


def _columns(df: pd.DataFrame) -> dict[str, str | None]:
    return {
        key: _find_column(df, aliases, required=key in {"name", "country", "product"})
        for key, aliases in COLUMN_ALIASES.items()
    }


def transform(df: pd.DataFrame, dispatch_type: str) -> pd.DataFrame:
    cols = _columns(df)
    country_col = cols["country"]
    product_col = cols["product"]

    if dispatch_type in {"uae", "oud_al_salam"}:
        country_mask = df[country_col].map(lambda value: _is_country(value, "uae"))
        product_mask = df[product_col].map(_is_oas_product)
        filtered = df[country_mask & (~product_mask if dispatch_type == "uae" else product_mask)].copy()
        return _transform_uae(filtered, cols)

    if dispatch_type in {"qatar", "bahrain"}:
        country_mask = df[country_col].map(lambda value: _is_country(value, dispatch_type))
        return _transform_gcc(df[country_mask].copy(), cols, dispatch_type)

    raise DispatchError(f"Unsupported dispatch type: {dispatch_type}")


def _transform_uae(filtered: pd.DataFrame, cols: dict[str, str | None]) -> pd.DataFrame:
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
        reference = _clean(get(row, "national_code")) or f"WPX-{index + 2}"
        product_note = _join_nonempty(
            f"{_clean(product)} x {_clean(qty) or '1'}" if _clean(product) else "",
            f"{_clean(product2)} x {_clean(qty2) or '1'}" if _clean(product2) else "",
        )
        records.append({
            "CUSTOMER_NAME": _clean(get(row, "name")),
            "MOBILE_NO": _best_phone(get(row, "primary_phone"), get(row, "secondary_phone"), "uae"),
            "LANDLINE_NO": "",
            "ADDRESS_1": _clean(get(row, "street")),
            "ADDRESS_2": _clean(get(row, "state")),
            "ADDRESS_3": "",
            "FLAT/VILLA NO": "",
            "DELIVERY_CITY": _clean(get(row, "city")) or _clean(get(row, "state")),
            "COD_AMOUNT": _cod_amount(get(row, "amount"), payment),
            "REMARKS": _join_nonempty(product_note, f"Payment: {_clean(payment)}" if _clean(payment) else ""),
            "REFERENCE_NO": reference,
            "OTHER_REMARKS": _clean(get(row, "description")),
            "COLLECT_RETURN (YES/NO)": "NO",
            "Agent Name": _clean(get(row, "agent_name")),
            "Source": _clean(get(row, "source")),
        })
    return pd.DataFrame(records, columns=UAE_HEADERS)


def _transform_gcc(filtered: pd.DataFrame, cols: dict[str, str | None], country: str) -> pd.DataFrame:
    def get(row, key):
        col = cols.get(key)
        return row[col] if col else ""

    settings = {
        "qatar": ("EMQ", "SCENT PASSION - QATAR", "Public Pricelist"),
        "bahrain": ("EMB", "SCENT PASSION - BAHRAIN", "Default BHD pricelist"),
    }
    prefix, source_id, pricelist = settings[country]
    records = []

    for index, row in filtered.iterrows():
        phone = _best_phone(get(row, "primary_phone"), get(row, "secondary_phone"), country)
        reference = _clean(get(row, "national_code")) or f"{prefix}{index + 2}"
        state = _clean(get(row, "state"))
        city = _clean(get(row, "city"))
        zone = state if re.fullmatch(r"\d+", state) else ""
        records.append({
            "client_order_ref": reference,
            "customer_name": _clean(get(row, "name")),
            "partner_id": phone,
            "whatsapp_no": phone,
            "source_id": source_id,
            "Pricelist Name": pricelist,
            "street_no": _clean(get(row, "street")),
            "building_no": "",
            "zone_id (governarate)": zone,
            "wilayat_id": city or state,
            "city_id": "",
            "order_line/product_id": _clean(get(row, "product")),
            "order_line/product_uom": "Units",
            "order_line/price_unit": _money(get(row, "amount")),
            "order_line/product_uom_qty": _money(get(row, "qty")) or 1,
            "remarks": _clean(get(row, "description")),
            "Agent Name": _clean(get(row, "agent_name")),
            "Source": _clean(get(row, "source")),
        })
    return pd.DataFrame(records, columns=GCC_HEADERS)


def _duplicate_mask(export_df: pd.DataFrame) -> pd.Series:
    if export_df.empty:
        return pd.Series(dtype=bool)
    phone_col = "MOBILE_NO" if "MOBILE_NO" in export_df.columns else "partner_id"
    phones = export_df[phone_col].map(_digits)
    return phones.ne("") & phones.duplicated(keep=False)


def _apply_rows(ws, export_df: pd.DataFrame, headers: list[str]) -> None:
    duplicate_mask = _duplicate_mask(export_df)
    duplicate_fill = PatternFill(fill_type="solid", fgColor="FFC7CE")
    for col_index, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col_index, value=header)
    for row_index, values in enumerate(export_df.itertuples(index=False, name=None), start=2):
        is_duplicate = bool(duplicate_mask.iloc[row_index - 2])
        for col_index, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            if is_duplicate:
                cell.fill = duplicate_fill


def create_workbook(export_df: pd.DataFrame, template_file, dispatch_type: str) -> bytes:
    if dispatch_type in {"uae", "oud_al_salam"}:
        template_bytes = template_file.read() if hasattr(template_file, "read") else bytes(template_file)
        if hasattr(template_file, "seek"):
            template_file.seek(0)
        wb = load_workbook(io.BytesIO(template_bytes))
        ws = wb["Upload Format"] if "Upload Format" in wb.sheetnames else wb.active
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        _apply_rows(ws, export_df, UAE_HEADERS)
        for row_index in range(2, ws.max_row + 1):
            ws.cell(row=row_index, column=2).number_format = "@"
            ws.cell(row=row_index, column=11).number_format = "@"
            ws.cell(row=row_index, column=9).number_format = "0.00"
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        _apply_rows(ws, export_df, GCC_HEADERS)
        for row_index in range(2, ws.max_row + 1):
            ws.cell(row=row_index, column=1).number_format = "@"
            ws.cell(row=row_index, column=3).number_format = "@"
            ws.cell(row=row_index, column=4).number_format = "@"
            ws.cell(row=row_index, column=14).number_format = "0.00"

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def build_export(source_file, template_file, dispatch_type: str) -> ExportResult:
    source_df = read_workpex(source_file)
    export_df = transform(source_df, dispatch_type)
    workbook = create_workbook(export_df, template_file, dispatch_type)
    return ExportResult(export_df, workbook, len(source_df), len(export_df))