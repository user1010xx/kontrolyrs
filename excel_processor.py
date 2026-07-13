from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd

SKIP_SHEETS = frozenset({"TOPLAM", "MANUEL EKLENENLER"})

DATE_FORMATS = (
    "%d.%m.%Y",
    "%d.%m.%y",
    "%d/%m/%Y",
    "%d/%m/%y",
    "%Y-%m-%d",
)

DATETIME_FORMATS = (
    "%d.%m.%Y %H:%M:%S",
    "%d.%m.%Y %H:%M",
    "%d.%m.%y %H:%M:%S",
    "%d.%m.%y %H:%M",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%y %H:%M:%S",
    "%d/%m/%y %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%SZ",
)

# Excel serial date origin (with the 1900 leap-year bug convention used by Excel)
_EXCEL_EPOCH = datetime(1899, 12, 30)
_EXCEL_SERIAL_MIN = 20000  # ~1954 — avoid treating small numbers as dates
_EXCEL_SERIAL_MAX = 60000  # ~2064

_TR_MAP = str.maketrans(
    {
        "İ": "I",
        "I": "I",
        "ı": "I",
        "i": "I",
        "Ş": "S",
        "ş": "S",
        "Ğ": "G",
        "ğ": "G",
        "Ü": "U",
        "ü": "U",
        "Ö": "O",
        "ö": "O",
        "Ç": "C",
        "ç": "C",
    }
)


@dataclass
class ProcessResult:
    output: BytesIO
    warnings: list[str] = field(default_factory=list)
    total_uye: int = 0
    total_yat: int = 0
    personel_count: int = 0
    skipped_empty: int = 0


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


def _norm_text(value: str) -> str:
    return " ".join(str(value).translate(_TR_MAP).upper().split())


def _norm_col(name: Any) -> str:
    return _norm_text(str(name).strip())


_SKIP_SHEETS_NORM = frozenset(_norm_text(s) for s in SKIP_SHEETS)


def _is_skip_sheet(name: str) -> bool:
    return _norm_text(name) in _SKIP_SHEETS_NORM


def _find_columns(df: pd.DataFrame) -> tuple[str | None, str | None]:
    """Locate kayıt / ilk yatırım columns with flexible header matching."""
    kayit_col: str | None = None
    yatirim_col: str | None = None

    for col in df.columns:
        n = _norm_col(col)
        if not n or n.startswith("UNNAMED"):
            continue
        if kayit_col is None and "KAYIT" in n and "TARIH" in n:
            kayit_col = str(col)
            continue
        if yatirim_col is None and "YATIRIM" in n and "TARIH" in n:
            yatirim_col = str(col)
            continue

    # Fallback: exact classic names after normalize
    if kayit_col is None:
        wanted = _norm_text("KAYIT TARİHİ")
        for col in df.columns:
            if _norm_col(col) == wanted:
                kayit_col = str(col)
                break
    if yatirim_col is None:
        wanted = _norm_text("İLK YATIRIM TARİHİ")
        for col in df.columns:
            if _norm_col(col) == wanted:
                yatirim_col = str(col)
                break

    return kayit_col, yatirim_col


def _from_excel_serial(value: float) -> datetime | None:
    if value < _EXCEL_SERIAL_MIN or value > _EXCEL_SERIAL_MAX:
        return None
    days = int(value)
    frac = float(value) - days
    seconds = int(round(frac * 86400))
    if seconds >= 86400:
        days += 1
        seconds -= 86400
    return _EXCEL_EPOCH + timedelta(days=days, seconds=seconds)


def _to_naive_datetime(value: datetime) -> datetime:
    if getattr(value, "tzinfo", None) is not None:
        return value.replace(tzinfo=None)
    return value


def _parse_datetime(value: Any) -> datetime | None:
    if value is None:
        return None

    # NaT / NA must be checked before isinstance(..., datetime)
    # because pandas.NaT reports as datetime subclass on some versions.
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        return _to_naive_datetime(value.to_pydatetime())

    if isinstance(value, datetime):
        return _to_naive_datetime(value)

    if isinstance(value, date) and not isinstance(value, datetime):
        return datetime.combine(value, datetime.min.time())

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        parsed = _from_excel_serial(float(value))
        if parsed is not None:
            return parsed

    text = str(value).strip()
    if not text or text.lower() in {"nan", "nat", "none", "null", "-"}:
        return None

    # ISO / fromisoformat (handles "2026-07-04 10:57:00" and with T)
    iso_candidate = text.replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(iso_candidate)
        return _to_naive_datetime(dt)
    except ValueError:
        pass

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
    if yatirim is None:
        return 0
    # Kayıt takvim gününden önceki yatırımlar sayılmaz.
    # Aynı güne ait saatsiz (00:00) Excel tarihleri elenmez.
    if yatirim.date() < kayit.date():
        return 0
    deadline = (
        kayit.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
    ).replace(hour=10, minute=0, second=0, microsecond=0)
    return 1 if yatirim <= deadline else 0


def _process_sheet(
    df: pd.DataFrame, start: date, end: date
) -> tuple[int, int, str | None]:
    """
    Returns (uye, yat, error_or_none).
    error is set when the sheet has no usable kayıt column.
    """
    if df is None or df.empty:
        return 0, 0, None

    kayit_col, yatirim_col = _find_columns(df)
    if not kayit_col:
        return 0, 0, "kayit_sutunu_yok"

    uye = 0
    yat = 0

    kayit_series = df[kayit_col]
    yatirim_series = df[yatirim_col] if yatirim_col else None

    for idx in range(len(df)):
        kayit = _parse_datetime(kayit_series.iloc[idx])
        if kayit is None:
            continue
        kayit_day = kayit.date()
        if kayit_day < start or kayit_day > end:
            continue
        uye += 1
        yatirim = _parse_datetime(yatirim_series.iloc[idx]) if yatirim_series is not None else None
        yat += _yatirim_sayisi(kayit, yatirim)

    return uye, yat, None if yatirim_col else "yatirim_sutunu_yok"


def process_excel(
    source: str | Path | BytesIO,
    start: date,
    end: date,
    *,
    include_zero_rows: bool = True,
    max_rows_per_sheet: int = 100_000,
) -> ProcessResult:
    warnings: list[str] = []
    rows: list[dict[str, object]] = []
    missing_kayit: list[str] = []
    missing_yatirim: list[str] = []
    truncated: list[str] = []
    skipped_empty = 0

    # Ensure BytesIO is readable from the start
    if isinstance(source, BytesIO):
        source.seek(0)

    with pd.ExcelFile(source) as xl:
        sheet_names = list(xl.sheet_names)
        for sheet in sheet_names:
            if _is_skip_sheet(sheet):
                continue

            df = xl.parse(sheet_name=sheet)
            if max_rows_per_sheet and len(df) > max_rows_per_sheet:
                truncated.append(sheet)
                df = df.iloc[:max_rows_per_sheet].copy()

            uye, yat, issue = _process_sheet(df, start, end)
            if issue == "kayit_sutunu_yok":
                missing_kayit.append(sheet)
                continue
            if issue == "yatirim_sutunu_yok":
                missing_yatirim.append(sheet)

            if uye == 0 and yat == 0 and not include_zero_rows:
                skipped_empty += 1
                continue

            rows.append(
                {
                    "Personel Adı": sheet,
                    "Üye Adedi": uye,
                    "Yatırım Adedi": yat,
                }
            )

    if missing_kayit:
        sample = ", ".join(missing_kayit[:5])
        extra = f" (+{len(missing_kayit) - 5})" if len(missing_kayit) > 5 else ""
        warnings.append(
            f"Kayıt tarihi sütunu bulunamayan sheet'ler atlandı ({len(missing_kayit)}): "
            f"{sample}{extra}"
        )
    if missing_yatirim:
        sample = ", ".join(missing_yatirim[:5])
        extra = f" (+{len(missing_yatirim) - 5})" if len(missing_yatirim) > 5 else ""
        warnings.append(
            f"İlk yatırım sütunu olmayan sheet'lerde yatırım=0 sayıldı ({len(missing_yatirim)}): "
            f"{sample}{extra}"
        )
    if truncated:
        warnings.append(
            f"Satır limiti ({max_rows_per_sheet}) aşıldığı için kısaltılan sheet: "
            f"{', '.join(truncated[:5])}"
        )
    if not rows and not missing_kayit:
        warnings.append("İşlenecek personel sheet'i bulunamadı.")
    if not rows and missing_kayit:
        warnings.append(
            "Hiçbir sheet'te 'KAYIT TARİHİ' (veya benzeri) sütun bulunamadı. "
            "Başlık satırını kontrol edin."
        )

    rows.sort(key=lambda r: (-int(r["Üye Adedi"]), str(r["Personel Adı"])))
    result = pd.DataFrame(rows, columns=["Personel Adı", "Üye Adedi", "Yatırım Adedi"])

    total_uye = int(result["Üye Adedi"].fillna(0).sum()) if len(result) else 0
    total_yat = int(result["Yatırım Adedi"].fillna(0).sum()) if len(result) else 0

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Rapor")
    output.seek(0)

    return ProcessResult(
        output=output,
        warnings=warnings,
        total_uye=total_uye,
        total_yat=total_yat,
        personel_count=len(result),
        skipped_empty=skipped_empty,
    )
