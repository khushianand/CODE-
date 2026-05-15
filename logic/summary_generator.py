"""Summary table generators used by the Summary Data sheet charts."""

from __future__ import annotations
import re
from typing import Dict, List
import pandas as pd

SEVERITY_ORDER={
    "Critical": "#9B0F06",
    "High": "#D53E0F",
    "Medium": "#F77F00",
    "Low": "#FCBF49",
    }
   
def _normalized_series(df: pd.DataFrame, column: str) -> pd.Series:
    """Handle the normalized series step for this module workflow."""
    return df.get(column, pd.Series(dtype=str)).fillna("").astype(str).str.strip()


#def severity_summary(df: pd.DataFrame) -> pd.DataFrame:


"""def _severity_counts(df: pd.DataFrame) -> List[Dict[str, object]]:
    counts = []
    risk = _normalized_series(df, "Risk").str.lower()
    for sev in SEVERITY_ORDER:
        count = int == ((sev.lower()).sum())
        count = int((risk == sev.lower()).sum())
        counts.append({"Severity": sev, "Count": count})
    #total = sum(item["Count"] for item in counts)
    return counts
"""

# logic/summary_generator.py

def _severity_counts(df: pd.DataFrame) -> List[Dict[str, object]]:
    """Handle the severity counts step for this module workflow."""
    counts = []
    risk = _normalized_series(df, "Risk").str.lower()

    for sev in SEVERITY_ORDER:
        # FIXED: Removed the malformed 'count = int == ...' line
        count = int((risk == sev.lower()).sum())

        counts.append({
            "Severity": sev,
            "Count": count
        })

    return counts

def severity_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Handle the severity summary step for this module workflow."""
    counts = _severity_counts(df)
    total = sum(int(item["Count"]) for item in counts)
    counts.append({"Severity": "Total", "Count": total})
    return pd.DataFrame(counts)


#def severity_chart_summary(df: pd.DataFrame) -> pd.DataFrame:
#    """Return severity counts without the total row for chart data sources."""
#    return pd.DataFrame(_severity_counts(df))

def severity_chart_summary(df: pd.DataFrame, include_total: bool = True) -> pd.DataFrame:
    """Return severity counts for dashboard chart data sources."""

    summary_df = pd.DataFrame(_severity_counts(df))

    if include_total:
        total_row = pd.DataFrame({
            summary_df.columns[0]: ["Total"],
            summary_df.columns[1]: [summary_df.iloc[:, 1].sum()]
        })

        summary_df = pd.concat([summary_df, total_row], ignore_index=True)

    return summary_df


def expert_severity_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Return reported vs expert severity counts for the VAMS dashboard bar chart."""
    reported = _normalized_series(df, "Risk").str.lower()
    expert = _normalized_series(df, "Expert Severity").str.lower()
    rows = []
    for sev in SEVERITY_ORDER:
        rows.append(
            {
                "Severity": sev,
                "Reported": int((reported == sev.lower()).sum()),
                "Expert": int((expert == sev.lower()).sum()),
            }
        )
    return pd.DataFrame(rows)


def _split_summary_values(value: object) -> List[str]:
    """Handle the split summary values step for this module workflow."""
    text = str(value or "").strip()
    if not text:
        return []
    return [part.strip() for part in re.split(r"\s*;\s*", text) if part.strip()]


def disposition_summary(df: pd.DataFrame, disposition_order: List[str]) -> pd.DataFrame:
    """Return disposition counts, including semicolon-merged VAMS values."""
    normalized_order = {value.lower(): value for value in disposition_order}
    counts = {value: 0 for value in disposition_order}
    for value in _normalized_series(df, "Disposition"):
        for token in _split_summary_values(value):
            display = normalized_order.get(token.lower(), token)
            counts[display] = counts.get(display, 0) + 1
    return pd.DataFrame({"Disposition": key, "Count": int(value)} for key, value in counts.items())


#def product_summary(unique_df: pd.DataFrame) -> pd.DataFrame:
    # Product dimension is not present in template rows; fallback uses Host distribution.
#    series = unique_df.get("Host / Image", pd.Series(dtype=str)).fillna("").astype(str).str.strip().replace("", "Unknown")
#    counts = series.value_counts()
#    rows = [{"Product": k, "Count": int(v)} for k, v in counts.items()]
#    rows.append({"Product": "Total", "Count": int(counts.sum())})
#    return pd.DataFrame(rows)


#def vulnerability_summary(unique_df: pd.DataFrame) -> pd.DataFrame:
#    series = unique_df.get("Name", pd.Series(dtype=str)).fillna("").astype(str).str.strip().replace("", "Unknown")
#    counts = series.value_counts()
#    rows = [{"Name": k, "Count": int(v)} for k, v in counts.items()]
#    rows.append({"Name": "Total", "Count": int(counts.sum())})
#    return pd.DataFrame(rows)
