import io
import os
import re

import pandas as pd
import streamlit as st

from barcode_tools import core, io as bt_io, ops, presets, reports

st.set_page_config(page_title="Barcode / GS Report Toolkit", layout="wide")

EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
UPLOAD_TYPES = ["xlsx", "xlsm", "xls", "xlsb", "ods", "csv", "tsv", "txt"]


def _user():
    return getattr(st, "user", None) or getattr(st, "experimental_user", None)


def require_auth():
    try:
        auth_cfg = st.secrets["auth"]
    except Exception:
        return None
    user = _user()
    if user is None or not getattr(user, "is_logged_in", False):
        st.title("Sign in")
        st.write("Sign in with Google to use the report toolkit.")
        try:
            st.login()
        except Exception as exc:
            st.error(
                "Login is unavailable. Make sure 'Authlib' is installed (it is in "
                "requirements.txt) and the [auth] secrets are configured. "
                f"Details: {exc}"
            )
        st.stop()
    email = getattr(user, "email", None)
    allowed = list(auth_cfg.get("allowed_emails", []))
    if allowed and email not in allowed:
        st.error(f"Access denied for {email}.")
        st.logout()
        st.stop()
    return email


def sidebar_footer(email):
    with st.sidebar:
        st.divider()
        if email:
            st.caption(f"Signed in as {email}")
            if st.button("Log out"):
                st.logout()
        else:
            st.caption("Local mode (no auth configured)")


@st.cache_data(show_spinner="Reading file...")
def _read_cached(data: bytes, name: str, header_mode: str, sheet: str | int):
    fmt = bt_io.detect_format(name)
    buf = io.BytesIO(data)
    if fmt == "csv":
        return bt_io.read_csv_raw(buf)
    header = 0 if header_mode == "header" else None
    return bt_io.read_excel(buf, sheet_name=sheet, header=header)


def load_uploaded(uploaded, key):
    if uploaded is None:
        return None, None, None
    data = uploaded.getvalue()
    fmt = bt_io.detect_format(uploaded.name)
    sheets = ["__none__"]
    if fmt == "excel":
        sheets = bt_io.sheet_names(io.BytesIO(data))
    header_mode = st.radio(
        "Header row",
        ["header", "no_header"],
        format_func=lambda x: "First row is header" if x == "header" else "No header (column numbers)",
        horizontal=True,
        key=f"hdr_{key}",
    )
    sheet = 0
    if fmt == "excel" and len(sheets) > 1:
        sheet = st.selectbox("Sheet", sheets, key=f"sheet_{key}")
    df = _read_cached(data, uploaded.name, header_mode, sheet)
    return df, header_mode, uploaded.name


def pick_column(df, key):
    cols = list(df.columns)
    guessed = bt_io.guess_code_column(df)
    default_index = cols.index(guessed) if guessed in cols else 0
    idx = st.selectbox(
        "Column with the codes",
        cols,
        index=default_index,
        format_func=str,
        key=f"col_{key}",
    )
    if guessed in cols:
        st.caption(f"Auto-detected column: {guessed}")
    return idx


def _show_invis():
    return st.session_state.get("show_invis", True)


def _gs_token():
    return st.session_state.get("gs_token", "[GS]") or "[GS]"


def current_preset():
    return st.session_state.get("preset", {}) or {}


def pv(field, fallback):
    value = current_preset().get(field)
    if value in (None, "", []):
        return fallback
    return value


def ptag():
    return st.session_state.get("ptag", "none")


def _display(df):
    if _show_invis():
        return reports.display_frame(df, _gs_token())
    return df


def _download_buttons(result, base_name, header=True):
    stem = os.path.splitext(base_name)[0]
    c1, c2, c3 = st.columns(3)
    c1.download_button(
        "Download Excel (.xlsx)",
        data=bt_io.to_excel_bytes(result, header=header),
        file_name=f"{stem}.xlsx",
        mime=EXCEL_MIME,
        key=f"dl_xlsx_{base_name}",
    )
    c2.download_button(
        "Download CSV",
        data=bt_io.to_csv_bytes(result, header=header),
        file_name=f"{stem}.csv",
        mime="text/csv",
        key=f"dl_csv_{base_name}",
    )
    c3.download_button(
        "Download TSV",
        data=bt_io.to_csv_bytes(result, header=header, sep="\t"),
        file_name=f"{stem}.tsv",
        mime="text/tab-separated-values",
        key=f"dl_tsv_{base_name}",
    )


def _download_sheets(sheets, base_name):
    stem = os.path.splitext(base_name)[0]
    c1, c2, c3 = st.columns(3)
    c1.download_button(
        "Download Excel (.xlsx, all sheets)",
        data=bt_io.to_excel_bytes_sheets(sheets),
        file_name=f"{stem}.xlsx",
        mime=EXCEL_MIME,
        key=f"dl_sheets_xlsx_{base_name}",
    )
    c2.download_button(
        "Download CSV (zip)",
        data=bt_io.to_zip_csv_bytes(sheets, sep=",", ext="csv"),
        file_name=f"{stem}_csv.zip",
        mime="application/zip",
        key=f"dl_sheets_csv_{base_name}",
    )
    c3.download_button(
        "Download TSV (zip)",
        data=bt_io.to_zip_csv_bytes(sheets, sep="\t", ext="tsv"),
        file_name=f"{stem}_tsv.zip",
        mime="application/zip",
        key=f"dl_sheets_tsv_{base_name}",
    )


def _colored_preview(table):
    if not _show_invis() or table is None or "Value" not in table.columns:
        return
    values = table["Value"].head(20)
    if not any(core.has_invisible(v) for v in values):
        return
    st.caption(f"First {len(values)} rows with invisible characters highlighted")
    st.markdown(
        reports.highlight_html(values, _gs_token(), limit=20),
        unsafe_allow_html=True,
    )


def show_result(res):
    if callable(res):
        with st.spinner("Working..."):
            try:
                res = res()
            except Exception as exc:
                st.error(f"Operation failed: {exc}")
                return
    st.subheader(res.title)
    if res.verdict is not None:
        if res.verdict.get("passed"):
            st.success("ALL CHECKS PASSED - no issues found.")
        else:
            issues = res.verdict.get("issues", [])
            st.error("ISSUES FOUND:\n\n" + "\n".join(f"- {issue}" for issue in issues))
    cols = st.columns(min(max(len(res.summary), 1), 4))
    for i, (k, v) in enumerate(res.summary.items()):
        cols[i % len(cols)].metric(k, f"{v:,}" if isinstance(v, int) else str(v))
    if res.table is not None and len(res.table):
        if res.table_caption:
            st.caption(res.table_caption)
        st.dataframe(_display(res.table).head(200), use_container_width=True, height=360)
        _colored_preview(res.table)
    if res.has_sheets:
        if res.result is not None:
            st.markdown(
                "**Cleaned file** - every failing row removed, duplicates collapsed"
            )
            _download_buttons(res.result, res.output_name, res.output_header)
            with st.expander("Preview cleaned file (first 100 rows)"):
                st.dataframe(
                    _display(res.result).head(100),
                    use_container_width=True,
                )
        st.caption("Audit sheets: " + ", ".join(res.sheets.keys()))
        for name, sheet_df in res.sheets.items():
            with st.expander(f"{name} ({len(sheet_df):,} rows)"):
                st.dataframe(
                    _display(sheet_df).head(200),
                    use_container_width=True,
                    height=300,
                )
        _download_sheets(res.sheets, res.output_name)
    elif res.result is not None:
        _download_buttons(res.result, res.output_name, res.output_header)
        with st.expander("Preview output (first 100 rows)"):
            st.dataframe(_display(res.result).head(100), use_container_width=True)


def single_file_section(title, uploaded, key):
    df, header_mode, name = load_uploaded(uploaded, key)
    if df is None:
        st.info("Upload a file to begin.")
        return
    st.caption(f"Loaded {len(df):,} rows x {len(df.columns)} columns from {name}")
    op = st.selectbox("Operation", list(title.keys()), key=f"op_{key}")
    st.divider()
    title[op](df, key)


def op_gs(df, key):
    column = pick_column(df, key)
    mode_options = ["presence", "strict_surround", "gs_before"]
    mode_default = pv("gs_mode", "presence")
    if mode_default not in mode_options:
        mode_default = "presence"
    mode = st.radio(
        "Check mode",
        mode_options,
        format_func=lambda m: {
            "presence": "GS separator present anywhere",
            "strict_surround": "GS on both sides of marker",
            "gs_before": "GS directly before prefix",
        }[m],
        index=mode_options.index(mode_default),
        key=f"gsmode_{key}_{ptag()}",
    )
    values = df[column]
    tokens = bt_io.guess_gs_tokens(values)
    surround = bt_io.guess_surround_marker(values)
    marker = st.text_input(
        "Marker (strict mode)",
        pv("marker", surround[0] if surround else core.DEFAULT_MARKER),
        key=f"gsmark_{key}_{ptag()}",
    )
    gs_prefix = st.text_input(
        "Prefix after GS",
        pv("gs_prefix", tokens[0] if tokens else core.DEFAULT_GS_PREFIX),
        key=f"gspfx_{key}_{ptag()}",
    )
    if st.button("Run", type="primary", key=f"run_gs_{key}"):
        show_result(lambda: ops.check_gs(df, column, mode, marker, gs_prefix))


def op_prefix(df, key):
    column = pick_column(df, key)
    detected = bt_io.guess_common_prefix(df[column])
    prefix = st.text_input(
        "Required prefix",
        pv("prefix", detected if len(detected) >= 6 else core.COMMON_PREFIXES[0]),
        key=f"pfx_{key}_{ptag()}",
    )
    action = st.radio(
        "Action",
        ["report", "remove"],
        format_func=lambda a: "Report offenders only" if a == "report" else "Remove offenders (keep valid rows)",
        horizontal=True,
        key=f"pfxact_{key}",
    )
    if st.button("Run", type="primary", key=f"run_pfx_{key}"):
        show_result(lambda: ops.check_prefix(df, column, prefix, action=action))


def op_length(df, key):
    column = pick_column(df, key)
    target = st.number_input(
        "Target cleaned length",
        min_value=1,
        value=int(pv("expected_length", 0)) or core.DEFAULT_CORE_LENGTH,
        key=f"len_{key}_{ptag()}",
    )
    if st.button("Run", type="primary", key=f"run_len_{key}"):
        show_result(lambda: ops.check_length(df, column, int(target)))


def _key_mode_default():
    default = pv("key_mode", "core_id")
    return default if default in ("core_id", "full") else "core_id"


def op_dedupe(df, key):
    column = pick_column(df, key)
    key_mode = st.radio(
        "Match basis",
        ["core_id", "full"],
        format_func=lambda m: "Core ID (first 31 chars)" if m == "core_id" else "Full normalized string",
        index=["core_id", "full"].index(_key_mode_default()),
        horizontal=True,
        key=f"dedupemode_{key}_{ptag()}",
    )
    length = st.number_input("Core ID length", min_value=1, value=core.DEFAULT_CORE_LENGTH, key=f"dedupelen_{key}_{ptag()}")
    if st.button("Run", type="primary", key=f"run_dedupe_{key}"):
        show_result(lambda: ops.deduplicate(df, column, key_mode, int(length)))


def op_duplicate(df, key):
    column = pick_column(df, key)
    key_mode = st.radio(
        "Match basis",
        ["core_id", "full"],
        format_func=lambda m: "Core ID (first 31 chars)" if m == "core_id" else "Full normalized string",
        index=["core_id", "full"].index(_key_mode_default()),
        horizontal=True,
        key=f"dupkey_{key}_{ptag()}",
    )
    c1, c2 = st.columns(2)
    view = c1.selectbox(
        "Show",
        ["all", "exact", "variation"],
        format_func=lambda v: {
            "all": "All duplicate groups",
            "exact": "Exact matches only",
            "variation": "Corrupted / variation only",
        }[v],
        key=f"dupview_{key}",
    )
    tokens = bt_io.guess_gs_tokens(df[column])
    gs_prefix = c2.text_input(
        "GS must precede (blank = off)",
        pv("gs_prefix", tokens[0] if tokens else core.DEFAULT_GS_PREFIX),
        key=f"dupgs_{key}_{ptag()}",
    )
    length = st.number_input("Core ID length", min_value=1, value=core.DEFAULT_CORE_LENGTH, key=f"duplen_{key}_{ptag()}")
    action = st.radio(
        "Action",
        ["report", "remove_all", "remove_corrupted"],
        format_func=lambda a: {
            "report": "Report only (download the duplicate rows)",
            "remove_all": "Remove all duplicates (keep first occurrence)",
            "remove_corrupted": "Remove corrupted (keep row where GS precedes prefix)",
        }[a],
        key=f"dupact_{key}",
    )
    if st.button("Run", type="primary", key=f"run_dup_{key}"):
        show_result(
            lambda: ops.duplicate_check(
                df,
                column,
                key_mode=key_mode,
                length=int(length),
                gs_prefix=gs_prefix or None,
                action=action,
                view=view,
            )
        )


def op_random(df, key):
    column = pick_column(df, key)
    action = st.selectbox(
        "Random action",
        ["sample", "shuffle", "shuffle_middle", "scatter_last", "insert_dummies"],
        format_func=lambda a: {
            "sample": "Randomly sample N rows",
            "shuffle": "Shuffle everything",
            "shuffle_middle": "Keep edges, shuffle the middle",
            "scatter_last": "Scatter the last N rows",
            "insert_dummies": "Insert N dummy codes",
        }[a],
        key=f"rand_{key}",
    )
    seed = st.number_input("Random seed (-1 = random)", value=-1, step=1, key=f"seed_{key}")
    seed_val = None if int(seed) < 0 else int(seed)
    if action == "sample":
        n = st.number_input("Sample size", min_value=1, value=min(1000, len(df)), key=f"samp_{key}")
        run = lambda: ops.sample_rows(df, int(n), seed_val)
    elif action == "shuffle":
        run = lambda: ops.shuffle_rows(df, seed_val)
    elif action == "shuffle_middle":
        first = st.number_input("Keep first N", min_value=0, value=1000, key=f"sm1_{key}")
        last = st.number_input("Keep last N", min_value=0, value=1000, key=f"sm2_{key}")
        run = lambda: ops.shuffle_middle(df, int(first), int(last), seed_val)
    elif action == "scatter_last":
        n = st.number_input("Scatter last N", min_value=1, value=50, key=f"scat_{key}")
        run = lambda: ops.scatter_last(df, int(n), seed_val)
    else:
        n = st.number_input("Number of dummies", min_value=1, value=100, key=f"dum_{key}")
        prefix1 = st.text_input("Dummy prefix", core.COMMON_PREFIXES[0], key=f"dumpre_{key}")
        infix = st.text_input("Dummy infix", "91EE1292", key=f"duminf_{key}")
        run = lambda: ops.insert_dummies(df, column, int(n), prefix1, infix, seed=seed_val)
    if st.button("Run", type="primary", key=f"run_rand_{key}"):
        show_result(lambda: run())


def op_filter(df, key):
    column = pick_column(df, key)
    min_length = st.number_input(
        "Minimum cleaned length (0 = off)",
        min_value=0,
        value=int(pv("min_length", 0)) or core.DEFAULT_MIN_LENGTH,
        key=f"minlen_{key}_{ptag()}",
    )
    detected = bt_io.guess_common_prefix(df[column])
    default_prefix = pv("prefix", detected if len(detected) >= 6 else "")
    split_prefix = st.text_input("Split jammed codes on prefix (blank = off)", default_prefix, key=f"split_{key}_{ptag()}")
    remove_prefix = st.text_input("Remove rows starting with (blank = off)", "", key=f"rempfx_{key}_{ptag()}")
    dedupe = st.checkbox("Deduplicate afterwards", value=True, key=f"fdedupe_{key}_{ptag()}")
    if st.button("Run", type="primary", key=f"run_filter_{key}"):
        show_result(
            lambda: ops.filter_clean(
                df,
                column,
                min_length=int(min_length) or None,
                split_prefix=split_prefix or None,
                remove_prefix=remove_prefix or None,
                dedupe=dedupe,
            )
        )


def op_full_audit(df, key):
    column = pick_column(df, key)
    values = df[column]
    tokens = bt_io.guess_gs_tokens(values)
    surround = bt_io.guess_surround_marker(values)
    detected_prefix = bt_io.guess_common_prefix(values)
    auto_mode = "strict_surround" if surround else "gs_before" if tokens else "presence"
    default_marker = pv("marker", surround[0] if surround else core.DEFAULT_MARKER)
    default_token = pv("gs_prefix", tokens[0] if tokens else core.DEFAULT_GS_PREFIX)
    default_prefix = pv(
        "prefix", detected_prefix if len(detected_prefix) >= 6 else core.COMMON_PREFIXES[0]
    )
    default_mode = pv("gs_mode", auto_mode)
    mode_options = ["presence", "strict_surround", "gs_before"]
    if default_mode not in mode_options:
        default_mode = auto_mode
    mode = st.radio(
        "GS check location",
        mode_options,
        format_func=lambda m: {
            "presence": "Anywhere (presence)",
            "strict_surround": "GS on both sides of a marker",
            "gs_before": "GS directly before a token",
        }[m],
        index=mode_options.index(default_mode),
        key=f"fa_mode_{key}_{ptag()}",
    )
    marker = default_marker
    gs_prefix = default_token
    if mode == "strict_surround":
        marker = st.text_input("Marker for strict placement", default_marker, key=f"fa_mark_{key}_{ptag()}")
    elif mode == "gs_before":
        gs_prefix = st.text_input("GS must directly precede (token)", default_token, key=f"fa_gs_{key}_{ptag()}")
    c1, c2 = st.columns(2)
    prefix = c1.text_input("Required prefix (blank = skip)", default_prefix, key=f"fa_pfx_{key}_{ptag()}")
    expected = c2.number_input(
        "Expected length (0 = skip)",
        min_value=0,
        value=int(pv("expected_length", 0)),
        key=f"fa_len_{key}_{ptag()}",
    )
    dedupe_options = ["core_id", "full"]
    dedupe_default = pv("dedupe_mode", "core_id")
    if dedupe_default not in dedupe_options:
        dedupe_default = "core_id"
    dedupe_mode = st.radio(
        "Duplicate basis",
        dedupe_options,
        format_func=lambda m: "Core ID (first 31 chars)" if m == "core_id" else "Full normalized string",
        index=dedupe_options.index(dedupe_default),
        horizontal=True,
        key=f"fa_dup_{key}_{ptag()}",
    )
    clean_dup = st.radio(
        "In the cleaned file, duplicates are:",
        ["keep_first", "remove_all", "off"],
        format_func=lambda c: {
            "keep_first": "Keep first occurrence",
            "remove_all": "Remove all copies",
            "off": "Leave duplicates alone",
        }[c],
        horizontal=True,
        key=f"fa_cleandup_{key}_{ptag()}",
    )
    include_passed = st.checkbox("Include a 'Passed' sheet", value=False, key=f"fa_pass_{key}_{ptag()}")
    if st.button("Run full audit", type="primary", key=f"run_fa_{key}"):
        show_result(
            lambda: ops.full_audit(
                df,
                column,
                prefix=prefix or None,
                expected_length=int(expected) or None,
                gs_mode=mode,
                marker=marker,
                gs_prefix=gs_prefix or None,
                dedupe_mode=dedupe_mode,
                include_passed=include_passed,
                clean_duplicates=clean_dup,
            )
        )
    with st.expander("Save these settings as a preset"):
        new_name = st.text_input("Preset name", key=f"fa_savename_{key}")
        if st.button("Save preset", key=f"fa_savebtn_{key}"):
            if new_name.strip():
                new_preset = {
                    "name": new_name.strip(),
                    "prefix": prefix or "",
                    "gs_mode": mode,
                    "marker": marker or "",
                    "gs_prefix": gs_prefix or "",
                    "expected_length": int(expected),
                    "dedupe_mode": dedupe_mode,
                }
                current = st.session_state.get("presets") or presets.load_presets()
                st.session_state["presets"] = presets.upsert(current, new_preset)
                st.success(f"Saved '{new_name.strip()}'. Pick it in the sidebar.")
            else:
                st.warning("Enter a preset name first.")


def two_file_section(uploaded_a, uploaded_b, key):
    df_a, _, name_a = load_uploaded(uploaded_a, f"{key}a")
    df_b, _, name_b = load_uploaded(uploaded_b, f"{key}b")
    if df_a is None or df_b is None:
        st.info("Upload both files to begin.")
        return
    st.caption(f"File 1: {name_a} ({len(df_a):,} rows) | File 2: {name_b} ({len(df_b):,} rows)")
    op = st.selectbox(
        "Operation",
        ["compare", "merge", "combine_sample"],
        format_func=lambda o: {
            "compare": "Compare / find missing or extra codes",
            "merge": "Merge validated codes from file 2 into file 1",
            "combine_sample": "Combine + randomly sample from file 2",
        }[o],
        key=f"twoop_{key}",
    )
    st.divider()
    if op == "compare":
        mode = st.selectbox(
            "Result",
            ["in_b_not_a", "in_a_not_b", "symmetric", "common"],
            format_func=lambda m: {
                "in_b_not_a": "Only in file 2 (not in file 1) - non-duplicates",
                "in_a_not_b": "Only in file 1 (not in file 2) - non-duplicates",
                "symmetric": "Unique to either file (both directions)",
                "common": "Present in both",
            }[m],
            key=f"cmpmode_{key}",
        )
        c1, c2 = st.columns(2)
        key_mode = c1.radio(
            "Match basis",
            ["core_id", "full"],
            index=["core_id", "full"].index(_key_mode_default()),
            horizontal=True,
            key=f"cmpkey_{key}_{ptag()}",
        )
        keep = c2.radio(
            "When a code repeats in the result",
            ["first", "last"],
            format_func=lambda k: "Keep first" if k == "first" else "Keep last",
            horizontal=True,
            key=f"cmpkeep_{key}",
        )
        dedupe = st.checkbox(
            "Export only unique codes (remove duplicates within the result)",
            value=True,
            key=f"cmpdd_{key}",
        )
        if st.button("Run", type="primary", key=f"run_cmp_{key}"):
            show_result(
                lambda: ops.compare_files(
                    df_a, df_b, mode, key_mode, dedupe=dedupe, keep=keep
                )
            )
    elif op == "merge":
        column = pick_column(df_a, f"{key}col")
        target = st.number_input(
            "Required source length (0 = off)",
            min_value=0,
            value=int(pv("expected_length", 0)) or 31,
            key=f"mrg_len_{key}_{ptag()}",
        )
        gs_prefix = st.text_input(
            "Require GS before (blank = off)",
            pv("gs_prefix", core.DEFAULT_GS_PREFIX),
            key=f"mrg_gs_{key}_{ptag()}",
        )
        dedupe = st.checkbox("Deduplicate source first", value=True, key=f"mrg_dd_{key}")
        if st.button("Run", type="primary", key=f"run_mrg_{key}"):
            show_result(
                lambda: ops.merge_validated(
                    df_a,
                    df_b,
                    column,
                    target_length=int(target) or None,
                    gs_prefix=gs_prefix or None,
                    dedupe_source=dedupe,
                )
            )
    else:
        column = pick_column(df_a, f"{key}col2")
        n = st.number_input("How many to sample from file 2", min_value=1, value=min(100, len(df_b)), key=f"cs_{key}")
        shuffle = st.checkbox("Shuffle final result", value=True, key=f"cs_sh_{key}")
        if st.button("Run", type="primary", key=f"run_cs_{key}"):
            show_result(lambda: ops.combine_sample(df_a, df_b, column, int(n), shuffle))


SINGLE_OPS = {
    "Full audit (all checks)": op_full_audit,
    "GS separator check": op_gs,
    "Prefix validation": op_prefix,
    "Length audit": op_length,
    "Duplicate check": op_duplicate,
    "Deduplicate (remove)": op_dedupe,
    "Randomize / sample": op_random,
    "Filter / clean": op_filter,
}


def main():
    email = require_auth()
    st.title("Barcode / GS Report Toolkit")
    st.caption("Upload a report, run a check, download the result. All processing happens on the server for this session only.")

    with st.sidebar:
        st.header("Choose a task")
        group = st.radio(
            "Category",
            ["Single file", "Two files"],
            label_visibility="collapsed",
        )
        st.divider()
        st.subheader("Product preset")
        if "presets" not in st.session_state:
            st.session_state["presets"] = presets.load_presets()
        preset_list = st.session_state["presets"]
        preset_name = st.selectbox(
            "Preset", presets.names(preset_list), key="preset_name"
        )
        st.session_state["preset"] = presets.get(preset_list, preset_name)
        st.session_state["ptag"] = re.sub(r"\W+", "_", preset_name)
        st.download_button(
            "Download presets (.json)",
            data=presets.to_json_bytes(preset_list),
            file_name="presets.json",
            mime="application/json",
            key="dl_presets",
        )
        uploaded_preset = st.file_uploader(
            "Upload presets (.json)", type=["json"], key="up_presets"
        )
        if uploaded_preset is not None:
            try:
                imported = presets.parse_presets(uploaded_preset.getvalue())
                merged = preset_list
                for preset in imported:
                    merged = presets.upsert(merged, preset)
                st.session_state["presets"] = merged
                st.success(f"Loaded {len(imported)} preset(s).")
            except Exception as exc:
                st.error(f"Could not read presets: {exc}")
        st.divider()
        st.subheader("Display")
        st.session_state.setdefault("show_invis", True)
        st.session_state.setdefault("gs_token", "[GS]")
        st.checkbox("Reveal invisible characters", key="show_invis")
        st.text_input("GS token", key="gs_token", disabled=not st.session_state.get("show_invis", True))

    if group == "Single file":
        uploaded = st.file_uploader("Report file", type=UPLOAD_TYPES, key="single_up")
        single_file_section(SINGLE_OPS, uploaded, "single")
    else:
        c1, c2 = st.columns(2)
        up_a = c1.file_uploader("File 1 (base)", type=UPLOAD_TYPES, key="two_a")
        up_b = c2.file_uploader("File 2", type=UPLOAD_TYPES, key="two_b")
        two_file_section(up_a, up_b, "two")

    sidebar_footer(email)


if __name__ == "__main__":
    main()
