from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

import hashlib

import pandas as pd
import streamlit as st
from pandas.api.types import is_datetime64_any_dtype

st.set_page_config(page_title="Workbook Viewer", layout="wide")

ACCESS_CODE_HASH = "712dca40936b39ce670dc803736fe3735cf99311030a928de039a36f77926230"


def _access_code_matches(candidate: str) -> bool:
    """Return ``True`` when the provided candidate matches the stored code."""

    candidate_hash = hashlib.sha256(candidate.encode("utf-8")).hexdigest()
    return candidate_hash == ACCESS_CODE_HASH


def _require_access_code() -> None:
    """Prompt for the access code and halt execution until it is validated."""

    if "access_granted" not in st.session_state:
        st.session_state.access_granted = False

    if st.session_state.access_granted:
        return

    st.title("Access Restricted")

    with st.form("access_code"):
        code = st.text_input("Enter access code", type="password")
        submitted = st.form_submit_button("Submit")

    if submitted:
        if _access_code_matches(code):
            st.session_state.access_granted = True
            st.rerun()
        else:
            st.error("Incorrect access code. Please try again.")

    st.stop()


_require_access_code()

# Enhanced CSS to remove all unwanted elements and optimize layout
st.markdown(
    """
    <style>
    /* Hide all default Streamlit elements */
    header, footer, #MainMenu {visibility: hidden;}
    .stDeployButton {display:none;}
    
    /* Remove padding and margins from main container */
    .main .block-container {
        padding-top: 0rem !important;
        padding-bottom: 0rem !important;
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: 100%;
        margin-top: -2rem !important;
    }
    
    /* Remove padding from tabs container */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0px;
        padding: 0 0 0 0;
        margin-top: 0 !important;
    }
    
    /* Style tab content to remove extra space */
    .stTabs [data-baseweb="tab-panel"] {
        padding-top: 0.5rem;
    }
    
    /* Remove footer completely */
    .st-emotion-cache-1avcm9n {
        display: none;
    }
    
    /* Remove blank space above tabs */
    .stMarkdown {
        margin-top: 0;
        margin-bottom: 0;
    }
    
    /* Remove top margin from the app */
    .element-container {
        margin-top: 0 !important;
    }
    
    /* Remove top margin from the first element */
    div[data-testid="stVerticalBlock"] > div:first-child {
        margin-top: 0 !important;
    }
    
    /* Optimize data table styling */
    [data-testid="stDataFrame"] {
        padding: 0;
        margin: 0;
    }
    
    [data-testid="stDataFrame"] div[role="gridcell"],
    [data-testid="stDataFrame"] div[role="columnheader"] {
        font-size: 0.8rem !important;
        padding: 4px 8px !important;
    }
    
    /* Remove scrollbar from main content */
    .main {
        overflow: hidden !important;
    }
    
    /* Make the entire app fit viewport height */
    .streamlit-container {
        height: 100vh;
        overflow: hidden;
    }
    
    /* Remove padding from the main content area */
    .stApp {
        padding-top: 0 !important;
    }
    
    /* Remove any top margin from the app */
    .app-view-container {
        margin-top: 0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DATA_PATH = Path(__file__).parent / "data/GP 2025 with MAIN.xlsx"
REGULAR_SHEET_CANDIDATES: tuple[str, ...] = ("Regular", "REGULAR", "REGULAR-25")
MAIN_SHEET_CANDIDATES: tuple[str, ...] = ("Main", "MAIN")


def _sheet_name_candidates() -> Iterable[str]:
    """Yield distinct sheet names to try when loading the regular view."""

    # Preserve the first occurrence of each candidate to avoid duplicate attempts.
    seen: set[str] = set()
    for name in REGULAR_SHEET_CANDIDATES:
        if name not in seen:
            seen.add(name)
            yield name


def _main_sheet_name_candidates() -> Iterable[str]:
    """Yield distinct sheet names to try when loading the main view."""

    # Preserve the first occurrence of each candidate to avoid duplicate attempts.
    seen: set[str] = set()
    for name in MAIN_SHEET_CANDIDATES:
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


@st.cache_data(show_spinner=False)
def load_main_sheet() -> pd.DataFrame:
    """Load the workbook sheet that should be displayed in the Main tab."""

    for sheet_name in _main_sheet_name_candidates():
        try:
            # The sheet places headers in the fourth row (index 3), so we use
            # ``header=3`` to promote that row to column names while dropping the
            # extraneous first three rows from the data frame.
            return pd.read_excel(DATA_PATH, sheet_name=sheet_name, header=3)
        except ValueError:
            continue

    candidate_list = ", ".join(MAIN_SHEET_CANDIDATES)
    raise ValueError(
        f"None of the expected sheets ({candidate_list}) were found in '{DATA_PATH.name}'."
    )


def display_dataframe(data: pd.DataFrame) -> None:
    """Display a DataFrame with optimized styling and responsive height."""
    
    # Create a copy to avoid modifying the original
    df = data.copy()
    
    # Remove completely empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # Handle integer columns
    integer_columns = [col for col in ("Qty", "Count") if col in df.columns]
    for column in integer_columns:
        df[column] = (
            pd.to_numeric(df[column], errors="coerce")
            .round()
            .astype("Int64")
        )

    # Handle float columns
    float_columns = [
        column
        for column in df.select_dtypes(include="number").columns
        if column not in integer_columns
    ]

    # Handle date columns
    date_columns = [
        column
        for column in df.columns
        if is_datetime64_any_dtype(df[column])
    ]

    for column in date_columns:
        df[column] = pd.to_datetime(df[column], errors="coerce")

    # Configure column display
    column_config: dict[str, Any] = {}

    for column in integer_columns:
        column_config[column] = st.column_config.NumberColumn(
            column, format="%d", help="Integer values"
        )

    for column in float_columns:
        column_config[column] = st.column_config.NumberColumn(
            column, format="%.2f", help="Numeric values"
        )

    for column in date_columns:
        column_config[column] = st.column_config.DatetimeColumn(
            column, format="DD-MM-YYYY", help="Date values"
        )

    # Calculate dynamic height based on viewport
    viewport_height = 600  # Approximate viewport height minus tabs and padding
    row_height = 35  # Height per row including padding
    header_height = 40  # Header height
    padding = 20  # Additional padding
    
    max_rows = max(1, min(len(df.index), (viewport_height - header_height - padding) // row_height))
    table_height = header_height + (max_rows * row_height) + padding

    st.dataframe(
        df,
        use_container_width=True,
        height=table_height,
        column_config=column_config,
        hide_index=True,
    )


# Create tabs without any spacing above
regular_tab, main_tab = st.tabs(["Regular", "Main"])

with main_tab:
    try:
        main_data = load_main_sheet()
        display_dataframe(main_data)
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

with regular_tab:
    try:
        regular_data = load_regular_sheet()
        display_dataframe(regular_data)
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