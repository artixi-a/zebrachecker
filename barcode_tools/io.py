import importlib.util
import io as _io
import os
import re
import zipfile
from collections import Counter

import pandas as pd

from .core import normalize, to_excel_safe, unxml

try:
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
except Exception:
    ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")

_HAS_CALAMINE = importlib.util.find_spec("python_calamine") is not None

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".xlsb", ".ods")
CSV_EXTENSIONS = (".csv", ".tsv", ".txt")


def _rewind(source):
    if hasattr(source, "seek"):
        try:
            source.seek(0)
        except Exception:
            pass
    return source


def _excel_engine(force=None):
    if force:
        return force
    return "calamine" if _HAS_CALAMINE else "openpyxl"


def detect_format(source, fmt=None):
    if fmt:
        return fmt
    name = ""
    if isinstance(source, str):
        name = source
    else:
        name = getattr(source, "name", "") or ""
    ext = os.path.splitext(name)[1].lower()
    if ext in CSV_EXTENSIONS:
        return "csv"
    if ext in EXCEL_EXTENSIONS:
        return "excel"
    return "excel"


def read_excel(source, sheet_name=0, header="infer", usecols=None, dtype=None, engine=None):
    resolved_header = 0 if header == "infer" else header
    return pd.read_excel(
        _rewind(source),
        sheet_name=sheet_name,
        header=resolved_header,
        usecols=usecols,
        dtype=dtype,
        engine=_excel_engine(engine),
    )


def read_csv_raw(source, skip_header_keywords=("QR Code", "QR")):
    _rewind(source)
    if isinstance(source, str) and os.path.exists(source):
        handle = open(source, "r", encoding="utf-8", errors="ignore")
        close = True
    else:
        raw = _rewind(source).read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8", errors="ignore")
        handle = _io.StringIO(raw)
        close = False
    lines = []
    try:
        for line in handle:
            cleaned = line.rstrip("\r\n").strip().strip('"')
            if cleaned:
                cleaned = cleaned.rstrip(";")
                lines.append(cleaned)
    finally:
        if close:
            handle.close()
    if lines and any(k in lines[0] for k in skip_header_keywords):
        lines = lines[1:]
    return pd.DataFrame({"Raw_Code": lines})


def read_table(source, fmt=None, sheet_name=0, header="infer", usecols=None):
    fmt = detect_format(source, fmt)
    if fmt == "csv":
        return read_csv_raw(source)
    return read_excel(source, sheet_name=sheet_name, header=header, usecols=usecols)


def sheet_names(source, engine=None):
    try:
        book = pd.read_excel(_rewind(source), sheet_name=None, nrows=0, engine=_excel_engine(engine))
        return list(book.keys())
    except Exception:
        return []


def columns_of(source, fmt=None, sheet_name=0):
    df = read_table(source, fmt=fmt, sheet_name=sheet_name, header="infer", usecols=None)
    return list(df.columns)


def first_column(df):
    return df.columns[0]


_CODE_HEADER_HINTS = ("qr", "code", "barcode", "bar code", "serial", "zebra", "scan")


def guess_code_column(df):
    columns = list(df.columns)
    if not columns:
        return None
    best = columns[0]
    best_score = float("-inf")
    row_count = max(len(df), 1)
    for pos, col in enumerate(columns):
        score = 0.0
        name = str(col).lower()
        for hint in _CODE_HEADER_HINTS:
            if hint in name:
                score += 5
        strings = df[col].dropna().astype(str)
        if len(strings):
            score += (len(strings) / row_count) * 2
            avg_len = strings.str.len().mean()
            if 25 <= avg_len <= 100:
                score += 2
            if strings.str.contains(r"\x1d|_x001[dD]_", regex=True).any():
                score += 3
            if strings.str.startswith("01").mean() > 0.5:
                score += 2
        score += (len(columns) - pos) * 0.01
        if score > best_score:
            best_score = score
            best = col
    return best


_GS_TOKEN_RE = re.compile(r"(?:_x001[dD]_|\x1d|!)([0-9]{2,4})")
_GS_ALNUM_RE = re.compile(r"(?:_x001[dD]_|\x1d|!)([0-9A-Za-z]{2,6})")
_GS_SURROUND_RE = re.compile(
    r"(?:_x001[dD]_|\x1d)([0-9A-Za-z]{2,8})(?:_x001[dD]_|\x1d)"
)


def _sample_strings(series, limit):
    return series.dropna().astype(str).head(limit)


def guess_gs_tokens(series, limit=5000, top=5):
    counts = Counter()
    for value in _sample_strings(series, limit):
        for match in _GS_TOKEN_RE.finditer(value):
            counts[match.group(1)] += 1
    if not counts:
        for value in _sample_strings(series, limit):
            for match in _GS_ALNUM_RE.finditer(value):
                counts[match.group(1)] += 1
    return [token for token, _ in counts.most_common(top)]


def guess_surround_marker(series, limit=5000, top=5):
    counts = Counter()
    for value in _sample_strings(series, limit):
        for match in _GS_SURROUND_RE.finditer(value):
            counts[match.group(1)] += 1
    return [token for token, _ in counts.most_common(top)]


def guess_common_prefix(series, length=19, limit=5000, min_share=0.5):
    values = [normalize(v, upper=False) for v in _sample_strings(series, limit)]
    values = [v for v in values if v]
    if not values:
        return ""
    counts = Counter(v[:length] for v in values)
    prefix, count = counts.most_common(1)[0]
    if count / len(values) < min_share:
        return ""
    return prefix


def code_lengths(series):
    return series.dropna().astype(str).map(lambda v: len(unxml(v)))


def guess_code_length(series):
    lengths = code_lengths(series)
    if lengths.empty:
        return None
    return int(lengths.value_counts().idxmax())


def length_distribution(series):
    lengths = code_lengths(series)
    counts = lengths.value_counts().sort_index(ascending=False)
    total = int(counts.sum()) or 1
    return pd.DataFrame(
        {
            "Length": counts.index.astype(int),
            "Count": counts.values,
            "Percent": (counts.values / total * 100).round(2),
        }
    )


def _excel_safe_cell(value):
    if isinstance(value, str):
        return to_excel_safe(value)
    return value


def safe_dataframe(df):
    out = df.copy()
    for col in out.columns:
        dtype = out[col].dtype
        if dtype == object or pd.api.types.is_string_dtype(dtype):
            out[col] = out[col].map(_excel_safe_cell)
    return out


def _strip_all_illegal(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def _write_excel(df, target, index, header, sheet_name):
    try:
        safe_dataframe(df).to_excel(
            target, index=index, header=header, sheet_name=sheet_name, engine="openpyxl"
        )
        return
    except Exception:
        if hasattr(target, "seek"):
            target.seek(0)
            target.truncate(0)
    fallback = df.copy()
    for col in fallback.columns:
        fallback[col] = fallback[col].map(_strip_all_illegal)
    fallback.to_excel(
        target, index=index, header=header, sheet_name=sheet_name, engine="openpyxl"
    )


def to_excel_bytes(df, index=False, header=True, sheet_name="Sheet1"):
    buf = _io.BytesIO()
    _write_excel(df, buf, index, header, sheet_name)
    buf.seek(0)
    return buf.getvalue()


def to_csv_bytes(df, index=False, header=True, sep=","):
    return df.to_csv(index=index, header=header, sep=sep).encode("utf-8")


def _safe_sheet_name(name, used):
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", str(name)).strip() or "Sheet"
    cleaned = cleaned[:31]
    if cleaned not in used:
        used.add(cleaned)
        return cleaned
    base = cleaned
    counter = 1
    while cleaned in used:
        suffix = f"_{counter}"
        cleaned = base[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(cleaned)
    return cleaned


def to_excel_bytes_sheets(sheets, index=False, header=True):
    buf = _io.BytesIO()
    used = set()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            sheet_name = _safe_sheet_name(name, used)
            try:
                safe_dataframe(frame).to_excel(
                    writer, index=index, header=header, sheet_name=sheet_name
                )
            except Exception:
                fallback = frame.copy()
                for col in fallback.columns:
                    fallback[col] = fallback[col].map(_strip_all_illegal)
                fallback.to_excel(
                    writer, index=index, header=header, sheet_name=sheet_name
                )
    buf.seek(0)
    return buf.getvalue()


def to_zip_csv_bytes(sheets, sep=",", ext="csv"):
    buf = _io.BytesIO()
    used = set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, frame in sheets.items():
            file_name = _safe_sheet_name(name, used) + f".{ext}"
            archive.writestr(file_name, frame.to_csv(index=False, sep=sep))
    buf.seek(0)
    return buf.getvalue()


def write_excel(df, path, index=False, header=True, sheet_name="Sheet1"):
    _write_excel(df, path, index, header, sheet_name)
    return path
