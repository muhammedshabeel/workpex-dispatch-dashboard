from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAQEL_PRODUCTS = {
    "ABSOLUTE MOUNTAIN AVENUE": ("ABSOLUTE MOUNTAIN AVENUE_OUD AL SALAM", 155.0),
    "ARBE PURO COMBO": ("ARBE PURO COMBO_LPG", 160.0),
    "LEON": ("LEON_LPG", 140.0),
    "OUD LOVERS": ("OUD LOVERS_LPG", 180.0),
    "PREMIUM EDITION": ("PREMIUM COLLECTION_OUD AL SALAM", 180.0),
    "PREMIUM COLLECTION": ("PREMIUM COLLECTION_OUD AL SALAM", 180.0),
}

JNT_PRODUCTS = {
    "ARCHER COMBO": ("THE ARCHER COMBO", 160.0),
    "HECTOR": ("HECTOR COMBO", 170.0),
    "MIRAMAR": ("MIRAMAR", 180.0),
    "ASEEL COMBO": ("ASEEL COMBO", 165.0),
    "SHADOW FLAME": ("SHADOW FLAME", 210.0),
    "VOLGA COMBO": ("VOLGA EDITION PERFUME COMBO", 170.0),
    "VOLGA EDITION": ("VOLGA EDITION PERFUME COMBO", 170.0),
    "COLLECTION OF MOOD": ("COLLECTION OF MOOD", 190.0),
}

NAQEL_CITY_CODES = {
    "riyadh": "RUH", "jeddah": "JED", "makkah": "MAC", "mecca": "MAC",
    "taif": "TIF", "dammam": "DMM", "jubail": "QJB", "aljubail": "QJB",
    "khobar": "DMM", "alkhobar": "DMM", "dhahran": "DMM", "abqaiq": "DMM",
    "hofuf": "HOF", "alhofuf": "HOF", "hafaralbatin": "HBT", "buraydah": "ELQ",
    "buraidah": "ELQ", "hail": "HAS", "madinah": "MED", "medinah": "MED",
    "tabuk": "TUU", "yanbu": "YNB", "abha": "AHB", "jizan": "GIZ",
    "gizan": "GIZ", "najran": "EAM", "albaha": "ABT", "arar": "RAE",
    "alula": "ULH", "alwajh": "EJH", "alqurayyat": "URY", "sakaka": "AJF",
    "aljouf": "AJF", "wadialdawasir": "WAE", "khamismushait": "AHB",
    "muhayil": "AHB", "sabya": "GIZ", "bisha": "BHH", "alkhurma": "TIF",
}

PROVINCE_MAP = {
    "riyadh": "Riyadh Province", "makkah": "Makkah Province", "mecca": "Makkah Province",
    "jeddah": "Makkah Province", "taif": "Makkah Province", "eastern": "Eastern Province",
    "dammam": "Eastern Province", "khobar": "Eastern Province", "alkhobar": "Eastern Province",
    "jubail": "Eastern Province", "dhahran": "Eastern Province", "hofuf": "Eastern Province",
    "madinah": "Madinah Province", "medinah": "Madinah Province", "tabuk": "Tabuk Province",
    "qassim": "Qassim Province", "buraidah": "Qassim Province", "buraydah": "Qassim Province",
    "hail": "Hail Province", "aseer": "Aseer Province", "asir": "Aseer Province",
    "abha": "Aseer Province", "khamismushait": "Aseer Province", "jizan": "Gizan Province",
    "gizan": "Gizan Province", "najran": "Najran Province", "albaha": "Al Baha Province",
    "aljouf": "Al Jouf Province", "sakaka": "Al Jouf Province", "northern": "Northern Borders Province",
}

NAQEL_HEADERS = [
    "RefNo", "Origin", "Destination", "Name", "Email", "PhoneNo", "MobileNo", "Address",
    "Location", "NationalAddress", "BuildingNo", "POBox", "Date", "Peices", "Weight", "Width",
    "Length", "Height", "Amount", "DeliveryInstruction", "PODType", "DeclaredValue", "Currency",
    "Incoterm", "GoodDesc", "ConsigneeNationalID", "ConsigneeNationalIdExpiry", "ConsigneeBirthDate",
    "Latitude", "Longitude", "OriginCountryCode", "DestinationCountryCode", "Agent Name", "Source",
    "Payment Method", "Unit Price",
]

JNT_HEADERS = [
    "Customer Order Number", "*Receiver name", "*Receiver phone number", "Receiver Backup NO.",
    "Receiver province", "*Receiver city", "Receiver district", "Receiver street", "*Receiver address",
    "Receiver Short Address", "Receiver Building Number", "Receiver Additional Number", "Sender Short Address",
    "Receiver email", "Receiver company name", "*Product type", "Payment type", "Package Number",
    "*Item type", "*Item weight (kg)", "*Item name", "*compensation ceiling or not?",
    "shipment value subject to service", "Platform name", "Customer account", "COD amount",
    "*Customer unpacking inspection", "Notes", "Agent Name", "Source", "Payment Method", "Quantity",
    "Unit Price",
]

ALIASES = {
    "name": ["Lead Name", "Customer Name", "Name", "First Name"],
    "phone1": ["Primary Phone", "Phone 1", "Phone1", "Phone", "Mobile", "Mobile No"],
    "phone2": ["Secondary Phone", "Phone 2", "Phone2", "Alternate Phone", "WhatsApp Number"],
    "country": ["Country"], "state": ["State", "Province", "Region"],
    "street": ["Street", "Address", "Address 1"], "city": ["CITY", "City", "Delivery City"],
    "reference": ["National Code", "Reference No", "Order ID"],
    "product1": ["Product", "Product Name"], "qty1": ["QTY", "Quantity"],
    "product2": ["PRODUCT 2", "Product 2"],
    "qty2": ["QTY OF PRODUCT 2", "Quantity of Product 2", "QTY 2", "Qty 2"],
    "amount": ["Actual Amount", "Forecasted Amount", "Amount", "COD Amount"],
    "payment": ["Payment Method", "Payment"], "remarks": ["Lead Description", "Remarks", "Notes"],
    "agent": ["Assigned", "Assigned User"], "source": ["Source"],
}


@dataclass(frozen=True)
class KSAResult:
    naqel_df: pd.DataFrame
    jnt_df: pd.DataFrame
    naqel_bytes: bytes
    jnt_bytes: bytes
    naqel_orders: int
    jnt_orders: int


def _clean(value) -> str:
    return "" if pd.isna(value) else str(value).strip()


def _norm(value) -> str:
    return re.sub(r"\s+", " ", _clean(value).lower()).strip()


def _key(value) -> str:
    text = unicodedata.normalize("NFKD", _clean(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _column(df: pd.DataFrame, key: str, required: bool = False) -> str | None:
    columns = {_norm(col): col for col in df.columns}
    for alias in ALIASES[key]:
        if _norm(alias) in columns:
            return columns[_norm(alias)]
    if required:
        raise ValueError(f"Required Workpex column missing: {ALIASES[key][0]}")
    return None


def _city_from_row(row: pd.Series) -> str:
    for column in row.index:
        name = _norm(column)
        if name.startswith("city") or "delivery city" in name:
            value = _clean(row[column])
            if value:
                return value
    return ""


def _digits(value) -> str:
    text = _clean(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\D", "", text)


def _phone(value) -> str:
    number = _digits(value)
    if number.startswith("00966"):
        number = number[5:]
    elif number.startswith("966"):
        number = number[3:]
    if number.startswith("0"):
        number = number[1:]
    return number if len(number) == 9 and number.startswith("5") else ""


def _number(value) -> float:
    match = re.search(r"-?\d+(?:\.\d+)?", _clean(value).replace(",", ""))
    return float(match.group()) if match else 0.0


def _qty(value, has_product: bool) -> float:
    if not has_product:
        return 0
    number = _number(value)
    return number if number > 0 else 1


def _product(raw: str) -> tuple[str, float, bool]:
    key = _key(raw)
    for name, (portal_name, price) in NAQEL_PRODUCTS.items():
        if _key(name) in key:
            return portal_name, price, True
    for name, (portal_name, price) in JNT_PRODUCTS.items():
        if _key(name) in key:
            return portal_name, price, False
    return _clean(raw).upper(), 0.0, False


def _province(state: str, city: str) -> str:
    for text in (state, city):
        key = _key(text)
        for token, province in PROVINCE_MAP.items():
            if token in key:
                return province
    text = _clean(state)
    if text:
        return text if text.lower().endswith("province") else f"{text} Province"
    return ""


def _destination(city: str) -> str:
    key = _key(city)
    if key in NAQEL_CITY_CODES:
        return NAQEL_CITY_CODES[key]
    for token, code in NAQEL_CITY_CODES.items():
        if token and token in key:
            return code
    return ""


def _same(a: str, b: str) -> bool:
    return bool(a and b and _key(a) == _key(b))


def _style_sheet(ws, headers: list[str], phone_column: int):
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col, header)
        cell.font = Font(name="Calibri", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        cell.fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions[get_column_letter(9)].width = 55
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, ws.max_row)}"
    duplicate_fill = PatternFill(fill_type="solid", fgColor="FFC7CE")
    values = [str(ws.cell(row, phone_column).value or "") for row in range(2, ws.max_row + 1)]
    duplicates = {value for value in values if value and values.count(value) > 1}
    for row in range(2, ws.max_row + 1):
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row, col)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = border
            if str(ws.cell(row, phone_column).value or "") in duplicates:
                cell.fill = duplicate_fill


def _workbook(sheet_name: str, headers: list[str], rows: list[list], phone_column: int) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    _style_sheet(sheet, headers, phone_column)
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def build_ksa_exports(source_df: pd.DataFrame) -> KSAResult:
    cols = {
        key: _column(source_df, key, required=key in {"name", "country", "product1"})
        for key in ALIASES
    }

    def get(row, key):
        column = cols.get(key)
        return row[column] if column else ""

    mask = source_df[cols["country"]].map(
        lambda value: _norm(value) in {"saudi arabia", "ksa", "saudi"}
    )
    filtered = source_df[mask].copy()
    naqel_records, jnt_records = [], []
    naqel_rows, jnt_rows = [], []
    naqel_orders = 0
    jnt_orders = 0

    for index, row in filtered.iterrows():
        raw1 = _clean(get(row, "product1"))
        raw2 = _clean(get(row, "product2"))
        if not raw1:
            continue

        product1, price1, is_naqel1 = _product(raw1)
        product2, price2, is_naqel2 = _product(raw2) if raw2 else ("", 0.0, True)
        qty1 = _qty(get(row, "qty1"), bool(raw1))
        qty2 = _qty(get(row, "qty2"), bool(raw2))
        total = _number(get(row, "amount"))
        total_qty = qty1 + qty2
        fallback = total / total_qty if total_qty else total
        unit1 = price1 or fallback
        unit2 = price2 or fallback
        city = _city_from_row(row) or _clean(get(row, "state"))
        state = _clean(get(row, "state"))
        street = _clean(get(row, "street"))
        destination = _destination(city)
        route = "naqel" if is_naqel1 and is_naqel2 and destination else "jnt"
        phone = _phone(get(row, "phone1")) or _phone(get(row, "phone2"))
        backup = _phone(get(row, "phone2"))
        reference = _clean(get(row, "reference")) or f"EMK{index + 2}"
        payment = _clean(get(row, "payment"))
        remarks = _clean(get(row, "remarks"))
        agent = _clean(get(row, "agent"))
        source = _clean(get(row, "source"))

        lines = [(product1, qty1, unit1)]
        if product2:
            if _same(product1, product2):
                lines = [(product1, qty1 + qty2, unit1)]
            else:
                lines.append((product2, qty2, unit2))

        if route == "naqel":
            naqel_orders += 1
            for product, quantity, unit_price in lines:
                amount = unit_price * quantity
                naqel_records.append({
                    "Reference": reference,
                    "City": city,
                    "Destination": destination,
                    "Product": product,
                    "Quantity": quantity,
                    "Unit Price": unit_price,
                    "Agent Name": agent,
                    "Source": source,
                    "Payment Method": payment,
                })
                naqel_rows.append([
                    reference, "RUH", destination, _clean(get(row, "name")), None,
                    phone, phone, city, street, None, None, None, None, quantity, 1.5,
                    10, 10, 10, amount, remarks or "NIL", "NIL", amount, "SAR",
                    "OTHER/UNKNOWN", product, None, None, None, None, None, "KSA", "KSA",
                    agent, source, payment, unit_price,
                ])
        else:
            jnt_orders += 1
            province = _province(state, city)
            for product, quantity, unit_price in lines:
                amount = unit_price * quantity
                jnt_records.append({
                    "Reference": reference,
                    "Province": province,
                    "City": city,
                    "Product": product,
                    "Quantity": quantity,
                    "Unit Price": unit_price,
                    "Agent Name": agent,
                    "Source": source,
                    "Payment Method": payment,
                })
                jnt_rows.append([
                    reference, _clean(get(row, "name")), int(phone) if phone else None,
                    int(backup) if backup else None, province, city, None, None, street,
                    None, None, None, "HOSA5275", None, None, "STANDARD",
                    "Cash" if _norm(payment) in {"cod", "cash on delivery"} else payment,
                    None, "Others", 1, product, "No", None, None, None, amount, "NO",
                    remarks, agent, source, payment, quantity, unit_price,
                ])

    return KSAResult(
        pd.DataFrame(naqel_records),
        pd.DataFrame(jnt_records),
        _workbook("GenerateWaybills", NAQEL_HEADERS, naqel_rows, 6),
        _workbook("Template", JNT_HEADERS, jnt_rows, 3),
        naqel_orders,
        jnt_orders,
    )
