"""Unit tests for excel_processor and bot helpers."""

from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from bot import _cleanup, _cleanup_dir  # noqa: E402
from excel_processor import (  # noqa: E402
    _find_columns,
    _parse_datetime,
    _process_sheet,
    _yatirim_sayisi,
    parse_date_input,
    process_excel,
)


def _xlsx_bytes(sheets: dict[str, pd.DataFrame]) -> BytesIO:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name)
    bio.seek(0)
    return bio


class ParseDateInputTests(unittest.TestCase):
    def test_ok(self):
        s, e = parse_date_input("02.07.2026,04.07.2026")
        self.assertEqual(s, date(2026, 7, 2))
        self.assertEqual(e, date(2026, 7, 4))

    def test_semicolon(self):
        s, e = parse_date_input("02.07.2026;04.07.2026")
        self.assertEqual((s, e), (date(2026, 7, 2), date(2026, 7, 4)))

    def test_reversed(self):
        with self.assertRaises(ValueError):
            parse_date_input("04.07.2026,02.07.2026")

    def test_single(self):
        with self.assertRaises(ValueError):
            parse_date_input("tek tarih")


class ParseDatetimeTests(unittest.TestCase):
    def test_nat(self):
        self.assertIsNone(_parse_datetime(pd.NaT))
        self.assertIsNone(_parse_datetime(None))
        self.assertIsNone(_parse_datetime(float("nan")))

    def test_timestamp(self):
        ts = pd.Timestamp("2026-07-04 10:57:00")
        self.assertEqual(_parse_datetime(ts), datetime(2026, 7, 4, 10, 57))

    def test_string_hm(self):
        self.assertEqual(
            _parse_datetime("04.07.26 10:57"), datetime(2026, 7, 4, 10, 57)
        )

    def test_string_seconds(self):
        self.assertEqual(
            _parse_datetime("04.07.2026 10:57:30"),
            datetime(2026, 7, 4, 10, 57, 30),
        )

    def test_iso(self):
        self.assertEqual(
            _parse_datetime("2026-07-04 10:57:00"),
            datetime(2026, 7, 4, 10, 57),
        )
        self.assertEqual(
            _parse_datetime("2026-07-04T10:57:00"),
            datetime(2026, 7, 4, 10, 57),
        )

    def test_excel_serial(self):
        # 2026-07-04 00:00 roughly
        serial = (datetime(2026, 7, 4) - datetime(1899, 12, 30)).days
        parsed = _parse_datetime(float(serial))
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.date(), date(2026, 7, 4))

    def test_small_float_not_date(self):
        self.assertIsNone(_parse_datetime(12.5))


class YatirimRuleTests(unittest.TestCase):
    def setUp(self):
        self.kayit = datetime(2026, 7, 5, 13, 41)

    def test_before_deadline(self):
        self.assertEqual(
            _yatirim_sayisi(self.kayit, datetime(2026, 7, 6, 9, 55)), 1
        )

    def test_at_deadline(self):
        self.assertEqual(
            _yatirim_sayisi(self.kayit, datetime(2026, 7, 6, 10, 0)), 1
        )

    def test_after_deadline(self):
        self.assertEqual(
            _yatirim_sayisi(self.kayit, datetime(2026, 7, 6, 10, 1)), 0
        )

    def test_none(self):
        self.assertEqual(_yatirim_sayisi(self.kayit, None), 0)

    def test_same_day(self):
        self.assertEqual(
            _yatirim_sayisi(self.kayit, datetime(2026, 7, 5, 18, 0)), 1
        )

    def test_before_kayit_day(self):
        self.assertEqual(
            _yatirim_sayisi(self.kayit, datetime(2026, 7, 4, 18, 0)), 0
        )

    def test_same_day_earlier_clock_still_counts(self):
        # Saatsiz / erken saat Excel tarihleri aynı günde sayılır
        self.assertEqual(
            _yatirim_sayisi(self.kayit, datetime(2026, 7, 5, 0, 0)), 1
        )


class ColumnAndSheetTests(unittest.TestCase):
    def test_whitespace_columns(self):
        df = pd.DataFrame(
            {
                "KAYIT TARİHİ ": [datetime(2026, 7, 4, 10, 0)],
                "İLK YATIRIM TARİHİ": [datetime(2026, 7, 4, 12, 0)],
            }
        )
        kayit, yat = _find_columns(df)
        self.assertIsNotNone(kayit)
        self.assertIsNotNone(yat)
        uye, y, err = _process_sheet(df, date(2026, 7, 4), date(2026, 7, 4))
        self.assertEqual((uye, y, err), (1, 1, None))

    def test_missing_kayit(self):
        df = pd.DataFrame({"A": [1], "B": [2]})
        uye, y, err = _process_sheet(df, date(2026, 7, 4), date(2026, 7, 4))
        self.assertEqual(uye, 0)
        self.assertEqual(err, "kayit_sutunu_yok")

    def test_nat_rows_do_not_crash(self):
        df = pd.DataFrame(
            {
                "KAYIT TARİHİ": [pd.NaT, datetime(2026, 7, 4, 10, 0)],
                "İLK YATIRIM TARİHİ": [pd.NaT, datetime(2026, 7, 4, 12, 0)],
            }
        )
        uye, y, err = _process_sheet(df, date(2026, 7, 4), date(2026, 7, 4))
        self.assertEqual((uye, y, err), (1, 1, None))

    def test_skip_sheet_case_insensitive(self):
        sheets = {
            "toplam": pd.DataFrame({"X": [1]}),
            "Manuel Eklenenler": pd.DataFrame({"X": [1]}),
            "Ali": pd.DataFrame(
                {
                    "KAYIT TARİHİ": [datetime(2026, 7, 4, 10, 0)],
                    "İLK YATIRIM TARİHİ": [datetime(2026, 7, 4, 11, 0)],
                }
            ),
        }
        result = process_excel(_xlsx_bytes(sheets), date(2026, 7, 4), date(2026, 7, 4))
        summary = pd.read_excel(result.output)
        self.assertEqual(list(summary["Personel Adı"]), ["Ali"])
        self.assertEqual(int(summary.iloc[0]["Üye Adedi"]), 1)


class ProcessExcelTests(unittest.TestCase):
    def test_bytesio_and_empty_cells(self):
        df = pd.DataFrame(
            {
                "KAYIT TARİHİ": [
                    datetime(2026, 7, 4, 10, 0),
                    None,
                    datetime(2026, 7, 5, 11, 0),
                ],
                "İLK YATIRIM TARİHİ": [
                    datetime(2026, 7, 4, 12, 0),
                    None,
                    datetime(2026, 7, 6, 11, 0),  # after next-day 10:00
                ],
            }
        )
        bio = _xlsx_bytes({"Ayşe": df, "TOPLAM": pd.DataFrame({"A": [1]})})
        result = process_excel(bio, date(2026, 7, 4), date(2026, 7, 5))
        summary = pd.read_excel(result.output)
        self.assertEqual(len(summary), 1)
        self.assertEqual(int(summary.iloc[0]["Üye Adedi"]), 2)
        self.assertEqual(int(summary.iloc[0]["Yatırım Adedi"]), 1)
        self.assertEqual(result.total_uye, 2)
        self.assertEqual(result.total_yat, 1)

    def test_file_closed_after_process(self):
        df = pd.DataFrame(
            {
                "KAYIT TARİHİ": [datetime(2026, 7, 4, 10, 0)],
                "İLK YATIRIM TARİHİ": [datetime(2026, 7, 4, 12, 0)],
            }
        )
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "sample.xlsx"
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                df.to_excel(w, index=False, sheet_name="Veli")
            result = process_excel(path, date(2026, 7, 4), date(2026, 7, 4))
            self.assertEqual(result.total_uye, 1)
            # Windows: file should be releasable
            path.unlink()
            self.assertFalse(path.exists())

    def test_missing_column_warning(self):
        bio = _xlsx_bytes({"X": pd.DataFrame({"A": [1]})})
        result = process_excel(bio, date(2026, 7, 4), date(2026, 7, 4))
        self.assertEqual(result.personel_count, 0)
        self.assertTrue(any("Kayıt" in w or "kayit" in w.lower() for w in result.warnings))

    def test_missing_yatirim_column_warns_zero(self):
        df = pd.DataFrame({"KAYIT TARİHİ": [datetime(2026, 7, 4, 10, 0)]})
        result = process_excel(_xlsx_bytes({"Zeynep": df}), date(2026, 7, 4), date(2026, 7, 4))
        self.assertEqual(result.total_uye, 1)
        self.assertEqual(result.total_yat, 0)
        self.assertTrue(any("yatırım" in w.lower() or "Yatırım" in w for w in result.warnings))

    def test_xlsx_magic(self):
        df = pd.DataFrame(
            {
                "KAYIT TARİHİ": [datetime(2026, 7, 4, 10, 0)],
                "İLK YATIRIM TARİHİ": [datetime(2026, 7, 4, 12, 0)],
            }
        )
        result = process_excel(_xlsx_bytes({"A": df}), date(2026, 7, 4), date(2026, 7, 4))
        self.assertEqual(result.output.getvalue()[:2], b"PK")


class CleanupTests(unittest.TestCase):
    def test_nested_cleanup(self):
        with tempfile.TemporaryDirectory() as outer:
            # create our own dir that cleanup will remove
            d = Path(outer) / "nested_tmp"
            d.mkdir()
            (d / "a.txt").write_text("x", encoding="utf-8")
            sub = d / "sub"
            sub.mkdir()
            (sub / "b.txt").write_text("y", encoding="utf-8")
            _cleanup_dir(d)
            self.assertFalse(d.exists())

    def test_yukle_style_cleanup_not_orphan(self):
        ctx = MagicMock()
        ctx.user_data = {}
        d = Path(tempfile.mkdtemp(prefix="leak_"))
        fp = d / "x.xlsx"
        fp.write_bytes(b"dummy")
        ctx.user_data["excel_path"] = str(fp)
        ctx.user_data["temp_dir"] = str(d)
        _cleanup(ctx)
        self.assertEqual(ctx.user_data, {})
        self.assertFalse(d.exists())


class BotImportTests(unittest.TestCase):
    def test_import_bot(self):
        import bot  # noqa: F401

        self.assertTrue(hasattr(bot, "main"))


if __name__ == "__main__":
    unittest.main()
