from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Workbook Viewer", layout="wide")

DATA_PATH = Path(__file__).parent / "data/GP 2025 with MAIN.xlsx"
REGULAR_SHEET_CANDIDATES: tuple[str, ...] = ("Regular", "REGULAR", "REGULAR-25")


def _sheet_name_candidates() -> Iterable[str]:
    """Yield distinct sheet names to try when loading the regular view."""

    # Preserve the first occurrence of each candidate to avoid duplicate attempts.
    seen: set[str] = set()
    for name in REGULAR_SHEET_CANDIDATES:
        if name not in seen:
            seen.add(name)
            yield name


@st.cache_data(show_spinner=False)
def load_regular_sheet() -> pd.DataFrame:
    """Load the workbook sheet that should be displayed in the Regular tab."""

    for sheet_name in _sheet_name_candidates():
        try:
            return pd.read_excel(DATA_PATH, sheet_name=sheet_name)
        except ValueError:
            continue

    candidate_list = ", ".join(REGULAR_SHEET_CANDIDATES)
    raise ValueError(
        f"None of the expected sheets ({candidate_list}) were found in '{DATA_PATH.name}'."
    )


main_tab, regular_tab = st.tabs(["Main", "Regular"])

with main_tab:
    st.write("")

with regular_tab:
    try:
        regular_data = load_regular_sheet()
    except FileNotFoundError:
        st.error(
            "The Excel workbook could not be found. Please ensure it is located at "
            f"'{DATA_PATH}'."
        )
    except ImportError:
        st.error(
            "Reading the Excel workbook requires optional dependencies such as "
            "'openpyxl'. Please install them and try again."
        )
    except ValueError as exc:
        st.error(str(exc))
    else:
        st.dataframe(regular_data, use_container_width=True)
