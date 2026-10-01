import html as _html

import pandas as pd

from . import core


def display_frame(df, gs_token="[GS]"):
    out = df.copy()
    for col in out.columns:
        out[col] = out[col].map(lambda v: core.visualize(v, gs_token))
    return out


def highlight_html(values, gs_token="[GS]", limit=20):
    token = _html.escape(gs_token)
    span = (
        '<span style="background:#ffe08a;color:#a00000;font-weight:600">'
        + token
        + "</span>"
    )
    rows = []
    for value in list(values)[:limit]:
        shown = _html.escape(str(core.visualize(value, gs_token)))
        shown = shown.replace(token, span)
        rows.append(
            '<div style="font-family:monospace;white-space:pre-wrap;margin:2px 0">'
            + shown
            + "</div>"
        )
    return "<div>" + "".join(rows) + "</div>"


def with_excel_rows(df):
    out = df.copy()
    out.insert(0, "Excel Row", out.index + 2)
    return out


def offender_table(df, column):
    table = df[[column]].copy()
    table.insert(0, "Excel Row", table.index + 2)
    return table.rename(columns={column: "Value"}).reset_index(drop=True)


def format_summary_markdown(title, summary):
    lines = [f"**{title}**", ""]
    for key, value in summary.items():
        if isinstance(value, float):
            value = f"{value:,.0f}"
        elif isinstance(value, int):
            value = f"{value:,}"
        lines.append(f"- {key}: {value}")
    return "\n".join(lines)


def summary_frame(summary):
    return pd.DataFrame(
        {"Metric": list(summary.keys()), "Value": list(summary.values())}
    )
