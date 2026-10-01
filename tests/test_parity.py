import io
import re
import sys
import os
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd

from barcode_tools import core, ops, presets
from barcode_tools import io as btio


def orig_deep_clean(series):
    return (
        series.astype(str)
        .str.replace(r"(?i)_x001D_", "", regex=True)
        .str.replace(r'[\s\x00-\x1F\x7F"]+', "", regex=True)
        .str.replace(r"^\]d2", "", regex=True)
        .str.upper()
        .str[:31]
    )


def orig_strip_match(series):
    return (
        series.astype(str)
        .str.replace(r"_x001[dD]_", "", regex=True)
        .str.replace("\x1d", "")
        .str.strip()
        .str.upper()
    )


class CoreParityTests(unittest.TestCase):
    def setUp(self):
        self.raw_gs = "\x1d"
        self.samples = pd.Series(
            [
                "0108606018940011215LT+mbL8'urit_x001D_91EE12_x001D_92abc=",
                "0108606018940011215abc\u001d91EE12\u001d92xyz=",
                "]d2010468067960195921500ABC_x001D_93ZZ",
                "  spaced _x001d_ value  ",
                "0108606018940011215\u001d93TAIL",
                "",
            ]
        )

    def test_core_id_matches_original(self):
        expected = orig_deep_clean(self.samples)
        got = self.samples.map(lambda v: core.core_id(v))
        self.assertListEqual(list(got), list(expected))

    def test_match_key_full_matches_original(self):
        expected = orig_strip_match(self.samples)
        got = self.samples.map(lambda v: core.match_key(v, "full"))
        self.assertListEqual(list(got), list(expected))

    def test_gs_detection(self):
        self.assertTrue(core.has_gs("abc\u001ddef"))
        self.assertTrue(core.has_gs("abc_x001D_def"))
        self.assertTrue(core.has_gs("abc_x001d_def"))
        self.assertFalse(core.has_gs("abcdef"))

    def test_gs_before_and_surround(self):
        val = "x_x001D_91EE12_x001D_92yyy"
        self.assertTrue(core.has_gs_surrounding(val, "91EE12"))
        self.assertTrue(core.has_gs_before("abc\u001d93zzz", "93"))
        self.assertFalse(core.has_gs_before("abc93zzz", "93"))

    def test_to_excel_safe(self):
        self.assertEqual(core.to_excel_safe("a\x1db"), "a_x001D_b")
        self.assertEqual(core.to_excel_safe("a\x00b"), "a_x0000_b")

    def test_visualize(self):
        self.assertEqual(core.visualize("a\x1db", "[GS]"), "a[GS]b")
        self.assertEqual(core.visualize("a_x001D_b", "[GS]"), "a[GS]b")
        self.assertEqual(core.visualize("plain"), "plain")
        self.assertTrue(core.has_invisible("a\x1db"))
        self.assertTrue(core.has_invisible("a_x001d_b"))
        self.assertFalse(core.has_invisible("plain"))


class OperationTests(unittest.TestCase):
    def make_df(self):
        return pd.DataFrame(
            {
                "QR Code": [
                    "0108606018940011215AAAA_x001D_91EE12_x001D_92BBB=",
                    "0108606018940011215AAAA_x001D_91EE12_x001D_92BBB=",
                    "0108606018940011215CCCC_x001D_91EE12_x001D_92DDD=",
                    "9999999999BADPREFIX_x001D_91EE12_x001D_92EEE=",
                ]
            }
        )

    def test_deduplicate(self):
        df = self.make_df()
        res = ops.deduplicate(df, "QR Code")
        self.assertEqual(len(res.result), 3)
        self.assertEqual(res.summary["Duplicates removed"], 1)

    def test_prefix_offenders(self):
        df = self.make_df()
        res = ops.check_prefix(df, "QR Code", "0108606018940011215")
        self.assertEqual(res.summary["Non-matching prefix"], 1)

    def test_gs_check(self):
        df = self.make_df()
        res = ops.check_gs(df, "QR Code", "strict_surround", "91EE12", "93")
        self.assertEqual(res.summary["Valid"], 4)

    def test_length_audit(self):
        df = pd.DataFrame({"QR Code": ["a" * 31, "b" * 30, "c" * 31]})
        res = ops.check_length(df, "QR Code", 31)
        self.assertEqual(res.summary["Mismatched length"], 1)

    def test_compare(self):
        a = pd.DataFrame({"C": ["AAA", "BBB", "CCC"]})
        b = pd.DataFrame({"C": ["BBB", "CCC", "DDD"]})
        res = ops.compare_files(a, b, "in_a_not_b")
        self.assertEqual(list(res.result["C"]), ["AAA"])

    def test_compare_dedupe_non_duplicates(self):
        a = pd.DataFrame({"C": ["AAA", "BBB", "CCC"]})
        b = pd.DataFrame({"C": ["BBB", "CCC", "DDD", "DDD", "EEE"]})
        res = ops.compare_files(a, b, "in_b_not_a", dedupe=True)
        self.assertEqual(sorted(res.result["C"]), ["DDD", "EEE"])
        self.assertEqual(res.summary["Internal duplicates removed"], 1)
        raw = ops.compare_files(a, b, "in_b_not_a", dedupe=False)
        self.assertEqual(len(raw.result), 3)

    def test_filter_split(self):
        prefix = "0108606018940011215"
        jammed = prefix + "A" * 80 + prefix + "B" * 80
        df = pd.DataFrame({"QR Code": [jammed]})
        res = ops.filter_clean(df, "QR Code", min_length=80, split_prefix=prefix)
        self.assertEqual(len(res.result), 2)

    def test_excel_export_sanitizes_gs_and_controls(self):
        df = pd.DataFrame({"QR Code": ["a\x1db\x00c", 123, None]})
        buf = btio.to_excel_bytes(df)
        back = pd.read_excel(io.BytesIO(buf), dtype=str)
        self.assertIn("_x001D_", back.iloc[0, 0])
        self.assertIn("_x0000_", back.iloc[0, 0])
        self.assertNotIn("\x1d", back.iloc[0, 0])

    def test_duplicate_check_variation(self):
        prefix = "0108606018940011215"
        df = pd.DataFrame(
            {
                "QR Code": [
                    prefix + "AAAA\x1d92BBB=",
                    prefix + "AAAA92BBB=",
                    prefix + "CCCC\x1d92DDD=",
                ]
            }
        )
        res = ops.duplicate_check(df, "QR Code", key_mode="core_id")
        self.assertEqual(res.summary["Duplicate rows"], 2)
        self.assertEqual(res.summary["Corrupted / variation groups"], 1)
        self.assertEqual(res.summary["Exact groups"], 0)

    def test_duplicate_check_remove_corrupted(self):
        prefix = "0108606018940011215"
        df = pd.DataFrame(
            {
                "QR Code": [
                    prefix + "AAAA\x1d92BBB=",
                    prefix + "AAAA92BBB=",
                    prefix + "CCCC\x1d92DDD=",
                ]
            }
        )
        res = ops.duplicate_check(
            df, "QR Code", key_mode="core_id", gs_prefix="92", action="remove_corrupted"
        )
        self.assertEqual(res.summary["Rows removed"], 1)
        self.assertEqual(len(res.result), 2)
        self.assertTrue(all("\x1d" in str(v) for v in res.result["QR Code"]))

    def test_guess_gs_tokens_and_surround(self):
        prefix = "0108606018940011215"
        series = pd.Series(
            [prefix + "AAAA\x1d93x", prefix + "BBBB\x1d93y", prefix + "CCCC\x1d92z"]
        )
        self.assertEqual(btio.guess_gs_tokens(series)[0], "93")
        self.assertEqual(btio.guess_common_prefix(series), prefix)
        surround = pd.Series([prefix + "AAAA\x1d91EE12\x1d92z"])
        self.assertEqual(btio.guess_surround_marker(surround)[0], "91EE12")

    def test_presets_roundtrip(self):
        loaded = presets.load_presets()
        self.assertEqual(loaded[0]["name"], presets.BLANK["name"])
        self.assertTrue(any("Milk" in n for n in presets.names(loaded)))
        blob = presets.to_json_bytes(loaded)
        parsed = presets.parse_presets(blob)
        self.assertTrue(any("Nelly" in p["name"] for p in parsed))
        merged = presets.upsert(parsed, {"name": "Custom", "prefix": "123"})
        self.assertEqual(presets.get(merged, "Custom")["prefix"], "123")

    def test_full_audit_length_skipped_and_verdict(self):
        prefix = "0108606018940011215"
        df = pd.DataFrame(
            {"QR Code": [prefix + "AAAA\x1d91EE12\x1d92BBB="]}
        )
        res = ops.full_audit(
            df, "QR Code", prefix=prefix, gs_mode="strict_surround", marker="91EE12"
        )
        self.assertEqual(res.summary["Length check"], "skipped")
        self.assertTrue(res.verdict["passed"])
        bad = pd.DataFrame({"QR Code": ["nonsense"]})
        res2 = ops.full_audit(bad, "QR Code", prefix=prefix, gs_mode="presence")
        self.assertFalse(res2.verdict["passed"])
        self.assertTrue(res2.verdict["issues"])

    def test_full_audit_cleaned_export(self):
        prefix = "0108606018940011215"
        df = pd.DataFrame(
            {
                "QR Code": [
                    prefix + "AAAA\x1d91EE12\x1d92BBB=",
                    prefix + "AAAA\x1d91EE12\x1d92BBB=",
                    prefix + "CCCC\x1d91EE12\x1d92DDD=",
                    "9999BAD\x1d91EE12\x1d92EEE=",
                    prefix + "FFFF\x1d92GGG=",
                ]
            }
        )
        keep = ops.full_audit(
            df,
            "QR Code",
            prefix=prefix,
            gs_mode="strict_surround",
            marker="91EE12",
            clean_duplicates="keep_first",
        )
        self.assertEqual(len(keep.result), 2)
        self.assertEqual(keep.summary["Rows after cleaning"], 2)
        remove_all = ops.full_audit(
            df,
            "QR Code",
            prefix=prefix,
            gs_mode="strict_surround",
            marker="91EE12",
            clean_duplicates="remove_all",
        )
        self.assertEqual(len(remove_all.result), 1)

    def test_full_audit_gs_before_mode(self):
        prefix = "0104680679602796215"
        df = pd.DataFrame(
            {
                "QR Code": [
                    prefix + "loFHU\x1d93yCxf",
                    prefix + "|pG6u\x1d93KFzI",
                    prefix + "p'-f\x1d93gqEJ",
                ]
            }
        )
        res = ops.full_audit(df, "QR Code", gs_mode="gs_before", gs_prefix="93")
        self.assertEqual(res.summary["GS missing"], 0)
        self.assertEqual(res.summary["GS placement issues (directly before 93)"], 0)
        self.assertIn("GS Placement (directly before 93)", res.sheets)
        strict = ops.full_audit(
            df, "QR Code", gs_mode="strict_surround", marker="91EE12"
        )
        self.assertEqual(
            strict.summary["GS placement issues (both sides of 91EE12)"], 3
        )
        presence = ops.full_audit(df, "QR Code", gs_mode="presence")
        self.assertFalse(
            any("placement" in key.lower() for key in presence.summary)
        )
        self.assertFalse(
            any("placement" in key.lower() for key in presence.sheets)
        )

    def test_guess_code_column(self):
        df = pd.DataFrame(
            {
                "Date": ["2026-01-01", "2026-01-02"],
                "QR Code": ["0" * 31, "1" * 31],
            }
        )
        self.assertEqual(btio.guess_code_column(df), "QR Code")

    def test_full_audit_sheets(self):
        prefix = "0108606018940011215"
        df = pd.DataFrame(
            {
                "QR Code": [
                    prefix + "AAAA\x1d91EE12_x001D_92BBB=",
                    prefix + "AAAA91EE1292BBB=",
                    "9999BAD\x1d91EE12\x1d92CCC=",
                    prefix + "DDDD\x1d91EE12\x1d92EEE=",
                    prefix + "DDDD\x1d91EE12\x1d92EEE=",
                ]
            }
        )
        res = ops.full_audit(
            df,
            "QR Code",
            prefix=prefix,
            gs_mode="strict_surround",
            marker="91EE12",
            gs_prefix="93",
        )
        self.assertTrue(res.has_sheets)
        for key in [
            "Summary",
            "GS Missing",
            "GS Placement (both sides of 91EE12)",
            "Prefix Offenders",
            "Length Offenders",
            "Duplicates",
        ]:
            self.assertIn(key, res.sheets)
        self.assertEqual(res.summary["GS missing"], 1)
        self.assertEqual(
            res.summary["GS placement issues (both sides of 91EE12)"], 1
        )
        self.assertEqual(res.summary["Prefix mismatches"], 1)
        self.assertEqual(res.summary["Duplicate rows"], 4)

    def test_multi_sheet_export(self):
        sheets = {
            "Sheet[1]": pd.DataFrame({"a": [1]}),
            "Sheet:2": pd.DataFrame({"b": [2]}),
            "AVeryLongSheetNameThatExceedsTheThirtyOneCharLimit": pd.DataFrame({"c": [3]}),
        }
        buf = btio.to_excel_bytes_sheets(sheets)
        book = pd.ExcelFile(io.BytesIO(buf))
        self.assertEqual(len(book.sheet_names), 3)
        for name in book.sheet_names:
            self.assertLessEqual(len(name), 31)
            self.assertNotIn("[", name)
            self.assertNotIn(":", name)

    def test_zip_csv_export(self):
        import zipfile

        sheets = {
            "A": pd.DataFrame({"x": [1, 2]}),
            "B": pd.DataFrame({"y": ["a\x1db"]}),
        }
        blob = btio.to_zip_csv_bytes(sheets, sep=",", ext="csv")
        with zipfile.ZipFile(io.BytesIO(blob)) as archive:
            names = archive.namelist()
        self.assertIn("A.csv", names)
        self.assertEqual(len(names), 2)

    def test_insert_dummies_actually_inserts(self):
        df = pd.DataFrame({"QR Code": ["ORIG" + str(i) for i in range(50)]})
        res = ops.insert_dummies(df, "QR Code", 10, seed=1)
        self.assertEqual(len(res.result), 60)
        self.assertEqual(res.result["QR Code"].duplicated().sum(), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
