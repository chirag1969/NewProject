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
            # The sheet places headers in the second row (index 1), so we use
            # ``header=1`` to promote that row to column names while dropping the
            # extraneous first row from the data frame.
            return pd.read_excel(DATA_PATH, sheet_name=sheet_name, header=1)
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
        regular_data = regular_data.copy()
        st.markdown(
            """
            <style>
            [data-testid="stDataFrame"] div[role="gridcell"],
            [data-testid="stDataFrame"] div[role="columnheader"] {
                font-size: 0.8rem !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        integer_columns = [col for col in ("Qty", "Count") if col in regular_data.columns]
        for column in integer_columns:
            regular_data[column] = (
                pd.to_numeric(regular_data[column], errors="coerce")
                .round()
                .astype("Int64")
            )

        float_columns = [
            column
            for column in regular_data.select_dtypes(include="number").columns
            if column not in integer_columns
        ]

        format_dict: dict[str, str] = {}
        if integer_columns:
            format_dict.update({column: "{:.0f}" for column in integer_columns})
        if float_columns:
            format_dict.update({column: "{:,.2f}" for column in float_columns})

        styled_regular_data = regular_data.style.format(format_dict, na_rep="")

        st.dataframe(styled_regular_data, use_container_width=True, height=650)
