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

# 减少主容器的内边距
st.markdown(
    """
    <style>
    header, footer, #MainMenu {visibility: hidden;}
    
    /* 减少主容器的内边距 */
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: 100%;
    }
    
    /* 减少标签页的内边距 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 4px 4px 0 0;
        gap: 1px;
        padding-left: 10px;
        padding-right: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DATA_PATH = Path(__file__).parent / "data/GP Current Month.xlsx"
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
            # 使用header=1来将第2行作为列标题，跳过第1行
            # 然后手动跳过第2行（索引为1），从第3行（索引为2）开始加载数据
            df = pd.read_excel(DATA_PATH, sheet_name=sheet_name, header=1)
            # 跳过索引为1的行（即第2行），从第3行开始保留数据
            return df.drop(1).reset_index(drop=True)
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


# 修改标签页顺序，使Regular成为默认选中的标签页
regular_tab, main_tab = st.tabs(["Regular", "Main"])

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

        date_columns = [
            column
            for column in regular_data.columns
            if is_datetime64_any_dtype(regular_data[column])
        ]

        for column in date_columns:
            regular_data[column] = pd.to_datetime(regular_data[column], errors="coerce")

        column_config: dict[str, Any] = {}

        for column in integer_columns:
            column_config[column] = st.column_config.NumberColumn(
                column, format="%d", help="Integer values"
            )

        for column in float_columns:
            # 使用会计格式显示数值，包括千位分隔符和两位小数
            column_config[column] = st.column_config.NumberColumn(
                column, format="$%,.2f", help="Currency values in accounting format"
            )

        for column in date_columns:
            column_config[column] = st.column_config.DatetimeColumn(
                column, format="DD-MM-YYYY", help="Date values"
            )

        visible_rows = min(len(regular_data.index), 15)
        row_height = 33
        base_height = 90
        table_height = base_height + row_height * visible_rows

        st.dataframe(
            regular_data,
            use_container_width=True,
            height=table_height,
            column_config=column_config,
        )

with main_tab:
    try:
        main_data = load_main_sheet()
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
        main_data = main_data.copy()
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

        integer_columns = [col for col in ("Qty", "Count") if col in main_data.columns]
        for column in integer_columns:
            main_data[column] = (
                pd.to_numeric(main_data[column], errors="coerce")
                .round()
                .astype("Int64")
            )

        float_columns = [
            column
            for column in main_data.select_dtypes(include="number").columns
            if column not in integer_columns
        ]

        date_columns = [
            column
            for column in main_data.columns
            if is_datetime64_any_dtype(main_data[column])
        ]

        for column in date_columns:
            main_data[column] = pd.to_datetime(main_data[column], errors="coerce")

        column_config: dict[str, Any] = {}

        for column in integer_columns:
            column_config[column] = st.column_config.NumberColumn(
                column, format="%d", help="Integer values"
            )

        for column in float_columns:
            # 使用会计格式显示数值，包括千位分隔符和两位小数
            column_config[column] = st.column_config.NumberColumn(
                column, format="$%,.2f", help="Currency values in accounting format"
            )

        for column in date_columns:
            column_config[column] = st.column_config.DatetimeColumn(
                column, format="DD-MM-YYYY", help="Date values"
            )

        visible_rows = min(len(main_data.index), 15)
        row_height = 33
        base_height = 90
        table_height = base_height + row_height * visible_rows

        st.dataframe(
            main_data,
            use_container_width=True,
            height=table_height,
            column_config=column_config,
        )