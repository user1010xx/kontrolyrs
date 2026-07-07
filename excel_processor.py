from __future__ import annotations

from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd

SKIP_SHEETS = {"TOPLAM", "MANUEL EKLENENLER"}
DATE_FORMATS = ("%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%d/%m/%y")
DATETIME_FORMATS = ("%d.%m.%Y %H:%M", "%d.%m.%y %H:%M", "%d/%m/%Y %H:%M", "%d/%m/%y %H:%M")


def parse_date_input(text: str) -> tuple[date, date]:
    parts = [p.strip() for p in text.replace(";", ",").split(",") if p.strip()]
    if len(parts) != 2:
        raise ValueError("Tarih formatı: GG.AA.YYYY,GG.AA.YYYY")

    start = _parse_date(parts[0])
    end = _parse_date(parts[1])
    if start > end:
        raise ValueError("Başlangıç tarihi bitiş tarihinden büyük olamaz.")
    return start, end


def _parse_date(value: str) -> date:
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Geçersiz tarih: {value}")


def _parse_datetime(value) -> datetime | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())

    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None

    for fmt in DATETIME_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _yatirim_sayisi(kayit: datetime, yatirim: datetime | None) -> int:
    if not yatirim:
        return 0
    deadline = (kayit.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)).replace(
        hour=10, minute=0, second=0, microsecond=0
    )
    return 1 if yatirim <= deadline else 0


def _process_sheet(df: pd.DataFrame, start: date, end: date) -> tuple[int, int]:
    uye = 0
    yat = 0
    kayit_col = "KAYIT TARİHİ"
    yatirim_col = "İLK YATIRIM TARİHİ"

    if kayit_col not in df.columns:
        return 0, 0

    for _, row in df.iterrows():
        kayit = _parse_datetime(row.get(kayit_col))
        if not kayit:
            continue
        if not (start <= kayit.date() <= end):
            continue
        uye += 1
        yatirim = _parse_datetime(row.get(yatirim_col)) if yatirim_col in df.columns else None
        yat += _yatirim_sayisi(kayit, yatirim)
    return uye, yat


def process_excel(source: str | Path | BytesIO, start: date, end: date) -> BytesIO:
    xl = pd.ExcelFile(source)
    rows: list[dict[str, object]] = []

    for sheet in xl.sheet_names:
        if sheet in SKIP_SHEETS:
            continue
        df = pd.read_excel(source, sheet_name=sheet)
        uye, yat = _process_sheet(df, start, end)
        rows.append({"Personel Adı": sheet, "Üye Adedi": uye, "Yatırım Adedi": yat})

    rows.sort(key=lambda r: (-int(r["Üye Adedi"]), str(r["Personel Adı"])))
    result = pd.DataFrame(rows, columns=["Personel Adı", "Üye Adedi", "Yatırım Adedi"])

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Rapor")
    output.seek(0)
    return output