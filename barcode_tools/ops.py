import random
import re
import string
from dataclasses import dataclass, field

import pandas as pd

from . import core
from .core import (
    DEFAULT_CORE_LENGTH,
    DEFAULT_GS_PREFIX,
    DEFAULT_MARKER,
)


@dataclass
class OpResult:
    title: str
    summary: dict = field(default_factory=dict)
    table: pd.DataFrame | None = None
    table_caption: str = ""
    result: pd.DataFrame | None = None
    output_name: str = "result.xlsx"
    output_header: bool = True
    sheets: dict | None = None
    verdict: dict | None = None

    @property
    def has_sheets(self):
        return bool(self.sheets)


def _values(df, column):
    return df[column]


def _offender_preview(df, column, limit=None):
    out = df[[column]].copy()
    out.insert(0, "Excel Row", out.index + 2)
    out = out.rename(columns={column: "Value"})
    out = out.reset_index(drop=True)
    if limit is not None:
        out = out.head(limit)
    return out


def check_gs(
    df,
    column,
    mode="presence",
    marker=DEFAULT_MARKER,
    gs_prefix=DEFAULT_GS_PREFIX,
):
    values = _values(df, column)
    if mode == "strict_surround":
        valid = values.map(lambda v: core.has_gs_surrounding(v, marker))
        desc = f"GS on both sides of '{marker}'"
    elif mode == "gs_before":
        valid = values.map(lambda v: core.has_gs_before(v, gs_prefix))
        desc = f"GS directly before '{gs_prefix}'"
    else:
        valid = values.map(core.has_gs)
        desc = "GS separator present (\\x1d or _x001D_)"

    offenders = df[~valid].copy()
    total = int(len(df))
    valid_count = int(valid.sum())
    summary = {
        "Total codes checked": total,
        "Valid": valid_count,
        "Invalid / missing GS": total - valid_count,
        "Check mode": desc,
    }
    return OpResult(
        title="GS separator check",
        summary=summary,
        table=_offender_preview(offenders, column, limit=500),
        table_caption="Offending rows (missing/invalid GS placement)",
        result=offenders,
        output_name="gs_offenders.xlsx",
    )


def duplicate_check(
    df,
    column,
    key_mode="core_id",
    length=DEFAULT_CORE_LENGTH,
    gs_prefix=None,
    action="report",
    view="all",
):
    keys = df[column].map(lambda v: core.match_key(v, key_mode, length))
    work = df.assign(_match_key=keys)
    valid = work[work["_match_key"] != ""]
    dup_mask = valid["_match_key"].duplicated(keep=False)
    dup_df = valid[dup_mask]

    groups = list(dup_df.groupby("_match_key"))
    exact_groups = 0
    variation_groups = 0
    table_rows = []
    for match_key, group in groups:
        raws = group[column].astype(str)
        has_variation = raws.nunique() > 1
        if has_variation:
            variation_groups += 1
        else:
            exact_groups += 1
        if len(table_rows) < 5000:
            for idx, row in group.iterrows():
                gs_ok = (
                    None
                    if gs_prefix is None
                    else core.has_gs_before(row[column], gs_prefix)
                )
                table_rows.append(
                    {
                        "Excel Row": idx + 2,
                        "Group": str(match_key)[:40],
                        "Type": "Variation" if has_variation else "Exact",
                        "GS ok": gs_ok,
                        "Value": row[column],
                    }
                )
    table = pd.DataFrame(table_rows)
    if not table.empty and view in ("exact", "variation"):
        table = table[table["Type"] == ("Variation" if view == "variation" else "Exact")]

    removed = 0
    if action == "remove_all":
        keep = ~dup_df["_match_key"].duplicated(keep="first")
        drop_index = dup_df[~keep].index
        result = work.drop(index=drop_index).drop(columns=["_match_key"])
        removed = len(drop_index)
        output_name = "deduplicated.xlsx"
    elif action == "remove_corrupted":
        rows_to_drop = []
        for _, group in groups:
            gs_mask = group[column].map(
                lambda v: core.has_gs_before(v, gs_prefix) if gs_prefix else False
            )
            if gs_prefix and gs_mask.any():
                rows_to_drop.extend(group[~gs_mask].index.tolist())
            else:
                rows_to_drop.extend(group.index[1:].tolist())
        result = work.drop(index=rows_to_drop).drop(columns=["_match_key"])
        removed = len(rows_to_drop)
        output_name = "deduplicated.xlsx"
    else:
        result = dup_df.drop(columns=["_match_key"])
        output_name = "duplicates.xlsx"

    summary = {
        "Total rows": int(len(df)),
        "Duplicate rows": int(len(dup_df)),
        "Duplicate groups": len(groups),
        "Exact groups": exact_groups,
        "Corrupted / variation groups": variation_groups,
        "Rows removed": removed,
    }
    return OpResult(
        title="Duplicate check",
        summary=summary,
        table=table.head(500) if not table.empty else table,
        table_caption="Duplicate groups (Excel Row, group key, exact vs variation, GS status)",
        result=result,
        output_name=output_name,
    )


def check_prefix(df, column, prefix, strip_aim=True, action="report"):
    values = _values(df, column)
    cleaned = values.map(
        lambda v: core.normalize(v, upper=False, strip_aim_prefix=strip_aim)
    )
    valid = cleaned.str.startswith(prefix)
    offenders = df[~valid].copy()
    total = int(len(df))
    if action == "remove":
        result = df[valid].copy()
        output_name = "prefix_cleaned.xlsx"
    else:
        result = offenders
        output_name = "prefix_offenders.xlsx"
    summary = {
        "Target prefix": prefix,
        "Total rows": total,
        "Valid": int(valid.sum()),
        "Non-matching prefix": total - int(valid.sum()),
        "Action": "Removed offenders" if action == "remove" else "Report only",
    }
    return OpResult(
        title="Prefix validation",
        summary=summary,
        table=_offender_preview(offenders, column, limit=500),
        table_caption=f"Rows not starting with {prefix}",
        result=result,
        output_name=output_name,
    )


def check_length(df, column, target_length):
    values = _values(df, column)
    lengths = values.map(lambda v: len(core.unxml(v)))
    offenders = df[lengths != target_length].copy()
    total = int(len(df))
    summary = {
        "Target length": int(target_length),
        "Total rows": total,
        "Matching length": int((lengths == target_length).sum()),
        "Mismatched length": int((lengths != target_length).sum()),
    }
    table = offenders[[column]].copy()
    table.insert(0, "Excel Row", table.index + 2)
    table["Length"] = lengths.loc[offenders.index].values
    table["Diff"] = table["Length"] - int(target_length)
    table = table.rename(columns={column: "Value"}).reset_index(drop=True).head(500)
    return OpResult(
        title="Length audit",
        summary=summary,
        table=table,
        table_caption=f"Rows whose cleaned length != {target_length}",
        result=offenders,
        output_name="length_offenders.xlsx",
    )


def deduplicate(df, column, key_mode="core_id", length=DEFAULT_CORE_LENGTH, keep_blanks=True):
    values = _values(df, column)
    keys = values.map(lambda v: core.match_key(v, key_mode=key_mode, length=length))
    work = df.assign(_match_key=keys)
    blanks = work[work["_match_key"] == ""]
    non_blanks = work[work["_match_key"] != ""]
    deduped = non_blanks.drop_duplicates(subset=["_match_key"], keep="first")
    if keep_blanks:
        final = pd.concat([deduped, blanks]).sort_index()
    else:
        final = deduped.sort_index()
    final = final.drop(columns=["_match_key"])
    removed = int(len(df) - len(final))
    summary = {
        "Total rows": int(len(df)),
        "Duplicates removed": removed,
        "Remaining rows": int(len(final)),
        "Match basis": "core ID (first 31 chars)" if key_mode == "core_id" else "full normalized string",
    }
    return OpResult(
        title="Deduplicate",
        summary=summary,
        table=None,
        result=final,
        output_name="deduplicated.xlsx",
    )


def compare_files(
    df_a,
    df_b,
    mode="in_a_not_b",
    key_mode="core_id",
    length=DEFAULT_CORE_LENGTH,
    dedupe=False,
    keep="first",
):
    col_a = df_a.columns[0]
    col_b = df_b.columns[0]
    keys_a = df_a[col_a].map(lambda v: core.match_key(v, key_mode, length))
    keys_b = df_b[col_b].map(lambda v: core.match_key(v, key_mode, length))
    set_a = set(keys_a)
    set_b = set(keys_b)

    if mode == "in_b_not_a":
        mask = ~keys_b.isin(set_a)
        result = df_b[mask].reset_index(drop=True)
        result_keys = keys_b[mask].reset_index(drop=True)
        caption = "Rows in the second file that are missing from the first"
        out_name = "missing_codes.xlsx"
    elif mode == "common":
        mask = keys_a.isin(set_b)
        result = df_a[mask].reset_index(drop=True)
        result_keys = keys_a[mask].reset_index(drop=True)
        caption = "Rows present in both files"
        out_name = "common_codes.xlsx"
    elif mode == "symmetric":
        mask_a = ~keys_a.isin(set_b)
        mask_b = ~keys_b.isin(set_a)
        left = df_a[mask_a].copy()
        right = df_b[mask_b].copy()
        common = list(left.columns.intersection(right.columns))
        result = pd.concat([left[common], right[common]], ignore_index=True)
        result_keys = pd.concat(
            [keys_a[mask_a], keys_b[mask_b]], ignore_index=True
        )
        caption = "Rows unique to either file (symmetric difference)"
        out_name = "unique_codes.xlsx"
    else:
        mask = ~keys_a.isin(set_b)
        result = df_a[mask].reset_index(drop=True)
        result_keys = keys_a[mask].reset_index(drop=True)
        caption = "Rows in the first file that are missing from the second"
        out_name = "missing_codes.xlsx"

    internal_removed = 0
    if dedupe and len(result):
        keep_mask = ~result_keys.duplicated(keep=keep)
        internal_removed = int((~keep_mask).sum())
        result = result[keep_mask.values].reset_index(drop=True)

    summary = {
        "Rows in file 1": int(len(df_a)),
        "Rows in file 2": int(len(df_b)),
        "Unique codes in file 1": len(set_a),
        "Unique codes in file 2": len(set_b),
        "Internal duplicates removed": internal_removed,
        "Result rows": int(len(result)),
    }
    return OpResult(
        title="Compare files",
        summary=summary,
        table=_offender_preview(result, result.columns[0], limit=200),
        table_caption=caption,
        result=result,
        output_name=out_name,
    )


def sample_rows(df, n, seed=None):
    n = min(int(n), len(df))
    result = df.sample(n=n, random_state=seed).reset_index(drop=True)
    summary = {"Total rows": int(len(df)), "Sampled rows": int(len(result))}
    return OpResult("Random sample", summary, None, "", result, "sample.xlsx")


def shuffle_rows(df, seed=None):
    result = df.sample(frac=1.0, random_state=seed).reset_index(drop=True)
    summary = {"Total rows": int(len(df)), "Action": "Fully shuffled"}
    return OpResult("Shuffle", summary, None, "", result, "shuffled.xlsx")


def shuffle_middle(df, first_n=1000, last_n=1000, seed=None):
    if len(df) <= first_n + last_n:
        first_n = last_n = 0
    head = df.iloc[:first_n]
    tail = df.iloc[len(df) - last_n:] if last_n else df.iloc[0:0]
    middle = df.iloc[first_n: len(df) - last_n] if last_n else df.iloc[first_n:]
    shuffled = middle.sample(frac=1.0, random_state=seed)
    result = pd.concat([head, shuffled, tail], ignore_index=True)
    summary = {
        "Total rows": int(len(df)),
        "Kept at top": int(first_n),
        "Kept at bottom": int(last_n),
        "Shuffled middle": int(len(middle)),
    }
    return OpResult("Shuffle middle", summary, None, "", result, "shuffled_middle.xlsx")


def scatter_last(df, n=50, seed=None):
    n = min(int(n), len(df))
    rng = random.Random(seed)
    main = df.iloc[: len(df) - n].copy()
    tail = df.iloc[len(df) - n:].copy()
    keys = list(range(len(main))) + [rng.uniform(0, max(len(main) - 1, 1)) for _ in range(len(tail))]
    combined = pd.concat([main, tail], ignore_index=True)
    combined["_sort"] = keys
    result = combined.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
    summary = {"Total rows": int(len(df)), "Scattered rows": int(n)}
    return OpResult("Scatter last rows", summary, None, "", result, "scattered.xlsx")


def _generate_dummy_codes(rng, prefix1, infix, count, symbols="!@#$%^&*()?,<>'=-;"):
    codes = set()
    while len(codes) < count:
        part2 = "".join(rng.choices(string.ascii_letters + string.digits + symbols, k=12))
        part4 = "".join(rng.choices(string.ascii_letters + string.digits + "+/", k=43)) + "="
        codes.add(f"{prefix1}{part2}{infix}{part4}")
    return list(codes)


def insert_dummies(
    df,
    column,
    n,
    prefix1="0108606018940011215",
    infix="91EE1292",
    buffer=100,
    seed=None,
):
    n = int(n)
    rng = random.Random(seed)
    data = df.reset_index(drop=True)
    n_orig = len(data)
    codes = _generate_dummy_codes(rng, prefix1, infix, n)
    lo = min(buffer, max(n_orig - 1, 0))
    hi = max(n_orig - buffer, lo)
    mu = n_orig / 2
    sigma = max(n_orig / 6, 1)
    positions = sorted(
        max(lo, min(hi, int(rng.gauss(mu, sigma)))) for _ in range(n)
    )
    records = []
    di = 0
    for i, row in enumerate(data.to_dict("records")):
        while di < n and positions[di] <= i:
            blank = {c: None for c in data.columns}
            blank[column] = codes[di]
            records.append(blank)
            di += 1
        records.append(row)
    while di < n:
        blank = {c: None for c in data.columns}
        blank[column] = codes[di]
        records.append(blank)
        di += 1
    result = pd.DataFrame(records, columns=data.columns)
    summary = {
        "Original rows": n_orig,
        "Dummy rows inserted": n,
        "New total rows": int(len(result)),
        "Insertion spread": "Gaussian around the middle",
    }
    return OpResult("Insert dummy codes", summary, None, "", result, "with_dummies.xlsx")


def combine_sample(base_df, source_df, column, n, shuffle=True, seed=None):
    base_col = base_df.columns[0]
    source_col = source_df.columns[0]
    n = min(int(n), len(source_df))
    picked = source_df.sample(n=n, random_state=seed)
    sample = pd.DataFrame({c: None for c in base_df.columns}, index=picked.index)
    sample[base_col] = picked[source_col].values
    combined = pd.concat([base_df, sample], ignore_index=True)
    if shuffle:
        combined = combined.sample(frac=1.0, random_state=seed).reset_index(drop=True)
    summary = {
        "Base rows": int(len(base_df)),
        "Added from source": n,
        "Final rows": int(len(combined)),
    }
    return OpResult("Combine + sample", summary, None, "", combined, "combined.xlsx")


def filter_clean(
    df,
    column,
    min_length=None,
    split_prefix=None,
    remove_prefix=None,
    dedupe=False,
    key_mode="core_id",
    length=DEFAULT_CORE_LENGTH,
):
    initial = len(df)
    work = df.copy()
    parts_created = 0
    removed_prefix = 0
    removed_length = 0

    if split_prefix:
        records = []
        for _, row in work.iterrows():
            parts = [
                p.strip()
                for p in re.split(f"(?={re.escape(split_prefix)})", str(row[column]))
            ]
            parts = [p for p in parts if p]
            parts_created += len(parts)
            for part in parts:
                new_row = row.copy()
                new_row[column] = part
                records.append(new_row)
        work = pd.DataFrame(records) if records else work.iloc[0:0]

    if remove_prefix:
        cleaned = work[column].map(lambda v: core.normalize(v, upper=False))
        before = len(work)
        work = work[~cleaned.str.startswith(remove_prefix)].copy()
        removed_prefix = before - len(work)

    if min_length is not None:
        lens = work[column].map(lambda v: len(core.unxml(v)))
        before = len(work)
        work = work[lens >= int(min_length)].copy()
        removed_length = before - len(work)

    dedupe_removed = 0
    if dedupe:
        keys = work[column].map(lambda v: core.match_key(v, key_mode, length))
        before = len(work)
        work = work[~keys.duplicated(keep="first")].copy()
        dedupe_removed = before - len(work)

    result = work.reset_index(drop=True)
    summary = {
        "Initial rows": int(initial),
        "Parts created by split": int(parts_created) if split_prefix else 0,
        "Removed (prefix filter)": int(removed_prefix),
        "Removed (length filter)": int(removed_length),
        "Duplicates removed": int(dedupe_removed),
        "Final rows": int(len(result)),
    }
    return OpResult(
        title="Filter / clean",
        summary=summary,
        table=None,
        result=result,
        output_name="filtered.xlsx",
    )


def _report_sheet(df):
    out = df.copy()
    out.insert(0, "Excel Row", out.index + 2)
    return out.reset_index(drop=True)


def full_audit(
    df,
    column,
    prefix=None,
    expected_length=None,
    gs_mode="presence",
    marker=DEFAULT_MARKER,
    gs_prefix=DEFAULT_GS_PREFIX,
    dedupe_mode="core_id",
    core_length=DEFAULT_CORE_LENGTH,
    include_passed=False,
    clean_duplicates="keep_first",
):
    values = df[column]
    total = len(df)

    gs_valid = values.map(core.has_gs)
    gs_missing = df[~gs_valid]

    if gs_mode == "strict_surround":
        placement_valid = values.map(lambda v: core.has_gs_surrounding(v, marker))
        placement_desc = f"both sides of {marker}"
    elif gs_mode == "gs_before":
        placement_valid = values.map(lambda v: core.has_gs_before(v, gs_prefix))
        placement_desc = f"directly before {gs_prefix}"
    else:
        placement_valid = gs_valid
        placement_desc = None
    gs_placement = df[~placement_valid]

    if prefix:
        prefix_valid = values.map(
            lambda v: core.normalize(v, upper=False).startswith(prefix)
        )
    else:
        prefix_valid = pd.Series(True, index=df.index)
    prefix_offenders = df[~prefix_valid]

    if expected_length:
        lengths = values.map(lambda v: len(core.unxml(v)))
        length_valid = lengths == int(expected_length)
    else:
        length_valid = pd.Series(True, index=df.index)
    length_offenders = df[~length_valid]

    dup_result = duplicate_check(
        df, column, dedupe_mode, core_length, gs_prefix=gs_prefix, action="report"
    )
    duplicates = dup_result.result if dup_result.result is not None else df.iloc[0:0]

    clean_mask = gs_valid & placement_valid & prefix_valid & length_valid
    cleaned = df[clean_mask].copy()
    clean_keys = cleaned[column].map(
        lambda v: core.match_key(v, dedupe_mode, core_length)
    )
    if clean_duplicates == "remove_all":
        cleaned = cleaned[~clean_keys.duplicated(keep=False)].copy()
    elif clean_duplicates == "keep_first":
        cleaned = cleaned[~clean_keys.duplicated(keep="first")].copy()

    summary = {
        "Total rows": total,
        "GS missing": int(len(gs_missing)),
    }
    if placement_desc:
        summary[f"GS placement issues ({placement_desc})"] = int(len(gs_placement))
    if prefix:
        summary["Prefix mismatches"] = int(len(prefix_offenders))
    else:
        summary["Prefix check"] = "skipped"
    if expected_length:
        summary["Length mismatches"] = int(len(length_offenders))
    else:
        summary["Length check"] = "skipped"
    summary["Duplicate rows"] = int(len(duplicates))
    summary["Rows after cleaning"] = int(len(cleaned))
    summary_df = pd.DataFrame(
        {"Metric": list(summary.keys()), "Value": list(summary.values())}
    )

    issues = []
    if len(gs_missing):
        issues.append(f"{len(gs_missing):,} rows missing a GS separator")
    if placement_desc and len(gs_placement):
        issues.append(
            f"{len(gs_placement):,} rows where GS is not {placement_desc}"
        )
    if prefix and len(prefix_offenders):
        issues.append(f"{len(prefix_offenders):,} rows with the wrong prefix")
    if expected_length and len(length_offenders):
        issues.append(f"{len(length_offenders):,} rows with the wrong length")
    if len(duplicates):
        issues.append(f"{len(duplicates):,} duplicate rows")
    verdict = {"passed": not issues, "issues": issues}

    sheets = {"Summary": summary_df, "GS Missing": _report_sheet(gs_missing)}
    if placement_desc:
        sheets[f"GS Placement ({placement_desc})"] = _report_sheet(gs_placement)
    sheets["Prefix Offenders"] = _report_sheet(prefix_offenders)
    sheets["Length Offenders"] = _report_sheet(length_offenders)
    sheets["Duplicates"] = _report_sheet(duplicates)

    if include_passed:
        sheets["Passed"] = _report_sheet(cleaned)

    return OpResult(
        title="Full audit",
        summary=summary,
        table=summary_df,
        table_caption="Audit summary",
        result=cleaned,
        output_name="full_audit_cleaned.xlsx",
        sheets=sheets,
        verdict=verdict,
    )


def merge_validated(
    base_df,
    source_df,
    column,
    target_length=None,
    gs_prefix=None,
    dedupe_source=True,
):
    base_col = base_df.columns[0]
    source_col = source_df.columns[0]
    source_codes = source_df[source_col].map(core.unxml)

    mask = pd.Series(True, index=source_df.index)
    if target_length is not None:
        mask = mask & (source_codes.str.len() == int(target_length))
    if gs_prefix:
        mask = mask & source_df[source_col].map(
            lambda v: core.has_gs_before(v, gs_prefix)
        )
    valid = source_df[mask].copy()
    valid["_key"] = source_codes[mask].values
    if dedupe_source:
        valid = valid.drop_duplicates(subset=["_key"], keep="first")
    existing = set(base_df[base_col].map(lambda v: core.normalize(v, upper=True)))
    new_rows = valid[~valid["_key"].isin(existing)].drop(columns=["_key"])
    new_rows = new_rows.rename(columns={source_col: base_col})
    combined = pd.concat([base_df, new_rows[base_df.columns]], ignore_index=True)
    summary = {
        "Base rows": int(len(base_df)),
        "Source rows": int(len(source_df)),
        "Valid in source": int(mask.sum()),
        "New codes appended": int(len(new_rows)),
        "Final rows": int(len(combined)),
    }
    return OpResult("Merge validated", summary, None, "", combined, "merged.xlsx")
