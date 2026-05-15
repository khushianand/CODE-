"""Performance helpers for large vulnerability datasets."""

from __future__ import annotations

import pandas as pd

# Above this many rows, avoid expensive workbook-wide formatting passes.
LARGE_DATASET_ROW_THRESHOLD = 200_000


def optimize_dataframe_memory(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with lighter dtypes where safe for large in-memory processing."""
    if df is None or df.empty:
        return df

    out = df.copy()

    for col in out.columns:
        series = out[col]
        if pd.api.types.is_integer_dtype(series):
            out[col] = pd.to_numeric(series, downcast="integer")
        elif pd.api.types.is_float_dtype(series):
            out[col] = pd.to_numeric(series, downcast="float")
        elif pd.api.types.is_object_dtype(series):
            # Convert to category only when cardinality is relatively low.
            nunique = series.nunique(dropna=True)
            total = len(series)
            if total and nunique / total < 0.5:
                try:
                    out[col] = series.astype("category")
                except Exception:
                    pass

    return out


def should_skip_full_formatting(*dfs: pd.DataFrame) -> bool:
    """Heuristic gate for expensive full-sheet formatting in openpyxl."""
    for df in dfs:
        if df is not None and hasattr(df, "shape") and len(df.index) >= LARGE_DATASET_ROW_THRESHOLD:
            return True
    return False
