import pandas as pd
import os
import ast
from typing import Optional, List, Union

def _read_any(path):
    path = str(path)
    if path.lower().endswith('.csv'):
        return pd.read_csv(path)
    return pd.read_excel(path)

def _save_any(df, output_path):
    output_path = str(output_path)
    if output_path.lower().endswith('.csv'):
        df.to_csv(output_path, index=False)
    else:
        df.to_excel(output_path, index=False)

def _normalize_target_columns(target_columns: Union[None, str, List[str]]):
    """
    Επιστρέφει None ή list[str].
    Accepts:
      - None
      - "A,B,C"
      - "['A','B']"  (stringified list)
      - ['A','B']
      - 'A'  -> ['A']
    """
    if target_columns is None or (isinstance(target_columns, str) and target_columns.strip() == ""):
        return None

    # If already list/tuple/set -> flatten and coerce to str
    if isinstance(target_columns, (list, tuple, set)):
        out = []
        for v in target_columns:
            if isinstance(v, (list, tuple, set)):
                out.extend([str(x).strip() for x in v])
            else:
                out.append(str(v).strip())
        return [c for c in out if c != ""]

    # If string -> try literal_eval if looks like list, else split by comma
    if isinstance(target_columns, str):
        s = target_columns.strip()
        # try literal eval for strings like "['A','B']" or '("A","B")'
        if (s.startswith('[') and s.endswith(']')) or (s.startswith('(') and s.endswith(')')):
            try:
                vals = ast.literal_eval(s)
                if isinstance(vals, (list, tuple, set)):
                    return [str(x).strip() for x in vals if str(x).strip() != ""]
            except Exception:
                pass
        # comma separated
        if ',' in s:
            return [c.strip() for c in s.split(',') if c.strip() != ""]
        # single column name
        return [s]

    # fallback: coerce to string
    return [str(target_columns).strip()]

def _apply_case(series: pd.Series, case_option: Optional[str]) -> pd.Series:
    if case_option == 'upper':
        return series.astype(str).str.upper()
    elif case_option == 'lower':
        return series.astype(str).str.lower()
    elif case_option == 'capitalize':
        return series.astype(str).str.title()
    return series

def clean_file(file_path: str,
               output_path: str,
               case_option: Optional[str] = None,
               drop_duplicates: bool = True,
               trim_spaces: bool = True,
               target_columns: Union[None, str, List[str]] = None):
    """
    Clean an Excel/CSV input and save to output_path (.xlsx or .csv).

    - drop_duplicates: remove duplicate rows
    - trim_spaces: strip leading/trailing spaces on string columns (or target_columns)
    - case_option: 'upper' | 'lower' | 'capitalize' or None
    - target_columns: None => apply to all object columns; or list of column names / csv string / "['A','B']"
    """

    file_path = str(file_path)
    output_path = str(output_path)

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")

    df = _read_any(file_path)

    # drop rows/cols that are all NaN
    df = df.dropna(how='all').dropna(axis=1, how='all')

    # drop duplicates
    if drop_duplicates:
        df = df.drop_duplicates()

    # normalize target_columns -> list[str] or None
    cols = _normalize_target_columns(target_columns)

    # if target columns not provided -> operate over object (string) dtype columns
    if cols is None:
        cols = df.select_dtypes(include=['object']).columns.tolist()

    # Validate that requested columns exist
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise KeyError(f"The following target columns were not found: {missing}")

    # Trim spaces
    if trim_spaces and cols:
        for col in cols:
            # coerce to string then strip
            df[col] = df[col].astype(str).str.strip()

    # Case conversion
    if case_option:
        if case_option not in ('upper', 'lower', 'capitalize'):
            raise ValueError("case_option must be one of: 'upper','lower','capitalize' or None")
        for col in cols:
            df[col] = _apply_case(df[col], case_option)

    # Save
    _save_any(df, output_path)
    return output_path
