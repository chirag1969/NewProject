from pathlib import Path
import typing as t

import altair as alt
import math
import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="GP 2025 Sales Dashboard",
    page_icon=":bar_chart:",
    layout="wide",
)


st.markdown(
    """
    <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        [data-testid="stToolbar"] {display: none !important;}
    </style>
    """,
    unsafe_allow_html=True,
)


DATA_PATH = Path(__file__).parent / "data/GP 2025 with MAIN.xlsx"


class WorkbookDependencyError(RuntimeError):
    """Raised when reading the Excel workbook requires a missing dependency."""


class WorkbookNotFoundError(RuntimeError):
    """Raised when the expected Excel workbook cannot be located."""

CHANNEL_LABELS = {
    "AMZN": "Amazon",
    "EBAY/Walmart": "eBay / Walmart",
    "WooCommerce": "WooCommerce",
}

GENERAL_RENAME = {
    "DC LIST": "dc_list",
    "DC": "dc",
    "CAT": "category",
    "SR. NO.": "serial_number",
    "SKU": "sku",
    "SKU(G)": "sku_g",
    "PLAIN SKU": "plain_sku",
    "AMZN": "amazon_manager",
    "EBAY": "ebay_manager",
    "WEBSITE": "website_manager",
    "FOCUS": "focus",
    "NEW/OLD": "new_old",
    "LISTING OWNER": "listing_owner",
    "PLATFORM": "platform",
    "CHECKOUT": "checkout",
    "FILTER STORE": "filter_store",
}

CHANNEL_RENAME = {
    "TOTAL TARGET SALES": "total_target_sales",
    "Sales Qty": "sales_quantity",
    "ACHIVED REV": "achieved_revenue",
    "Diff in rev": "revenue_gap",
    "Desired IP till Date": "desired_ip_to_date",
    "Gross IP": "gross_ip",
    "ADVT SPEND": "ad_spend",
    "Achieved IP": "achieved_ip",
    "Diff in IP": "ip_gap",
    "IP MARGIN": "ip_margin",
    "Storage Fees": "storage_fees",
}

NUMERIC_COLUMNS = list(CHANNEL_RENAME.values())
DEFAULT_TOP_N = 10


def _nice_upper_bound(max_value: float) -> float:
    """Return a pleasant upper bound for chart axes based on the maximum value."""

    if pd.isna(max_value) or max_value <= 0:
        return 1.0

    padded = max_value * 1.05
    magnitude = 10 ** max(int(math.log10(padded)), 0)
    for factor in (1, 2, 5, 10):
        candidate = factor * magnitude
        if candidate >= padded:
            return candidate

    return padded


@st.cache_data
def load_raw_workbook() -> pd.DataFrame:
    """Read the Excel workbook and normalise the header structure."""

    try:
        raw = pd.read_excel(DATA_PATH, sheet_name="MAIN", header=[2, 3])
    except FileNotFoundError as exc:
        raise WorkbookNotFoundError(
            f"Workbook not found at '{DATA_PATH}'."
        ) from exc
    except ImportError as exc:  # Missing optional dependency such as openpyxl
        raise WorkbookDependencyError(
            "Reading the GP 2025 workbook requires the optional 'openpyxl' package."
        ) from exc

    original_second_level = raw.columns.get_level_values(1)
    valid_columns = ~original_second_level.isna()
    raw = raw.loc[:, valid_columns]

    general_columns = set(GENERAL_RENAME.keys()) | {"NEW/OLD"}

    first_level_labels = []
    current_section = "General"
    level_zero = raw.columns.get_level_values(0)
    level_one = raw.columns.get_level_values(1)

    for top_label, bottom_label in zip(level_zero, level_one):
        if isinstance(bottom_label, str):
            bottom_trimmed = bottom_label.strip()
        else:
            bottom_trimmed = bottom_label

        if isinstance(top_label, str):
            top_trimmed = top_label.strip()
        else:
            top_trimmed = top_label

        if (
            isinstance(bottom_trimmed, str)
            and bottom_trimmed in general_columns
        ):
            first_level_labels.append("General")
            continue

        if (
            pd.isna(top_trimmed)
            or top_trimmed == ""
            or (isinstance(top_trimmed, str) and top_trimmed.lower().startswith("unnamed"))
        ):
            first_level_labels.append(current_section)
        elif top_trimmed == "#N/A":
            first_level_labels.append("General")
        else:
            current_section = top_trimmed
            first_level_labels.append(current_section)

    raw.columns = pd.MultiIndex.from_arrays(
        [first_level_labels, raw.columns.get_level_values(1)]
    )

    raw = raw.dropna(how="all")
    return raw


@st.cache_data
def load_sales_data() -> pd.DataFrame:
    """Transform the workbook into a tidy table ready for analysis."""

    workbook = load_raw_workbook()

    cleaned_columns = {}
    for original in workbook["General"].columns:
        if isinstance(original, str):
            key = original.strip()
        else:
            key = original

        lookup_key = key.upper() if isinstance(key, str) else key
        cleaned_columns[original] = GENERAL_RENAME.get(lookup_key, key)

    general = workbook["General"].rename(columns=cleaned_columns)
    general = general.dropna(subset=["sku"]).copy()
    general["sku"] = general["sku"].astype(str).str.strip()
    general = general[general["sku"].str.lower() != "nan"]

    def _ensure_text_column(frame: pd.DataFrame, column: str, default: str) -> None:
        if column not in frame.columns:
            frame[column] = default
        frame[column] = frame[column].astype("string").str.strip()
        frame[column] = frame[column].fillna(default)

    _ensure_text_column(general, "focus", "Unspecified")
    _ensure_text_column(general, "category", "Uncategorised")
    _ensure_text_column(general, "dc_list", "Unknown")
    _ensure_text_column(general, "new_old", "Unspecified")
    _ensure_text_column(general, "listing_owner", "Unassigned")
    _ensure_text_column(general, "checkout", "Not set")
    _ensure_text_column(general, "filter_store", "All Stores")

    if "dc" not in general.columns:
        general["dc"] = general["dc_list"]
    _ensure_text_column(general, "dc", "Unknown")

    if "plain_sku" not in general.columns:
        general["plain_sku"] = general["sku"]
    _ensure_text_column(general, "plain_sku", "Unspecified")

    if "sku_g" not in general.columns:
        general["sku_g"] = general["sku"]
    _ensure_text_column(general, "sku_g", "Unspecified")

    if "platform" in general.columns:
        general["platform"] = general["platform"].astype("string").str.strip()

    channel_frames = []
    first_level = workbook.columns.get_level_values(0)

    for raw_key, display_name in CHANNEL_LABELS.items():
        if raw_key not in first_level:
            continue

        channel = workbook[raw_key].rename(columns=CHANNEL_RENAME)
        channel = channel.apply(pd.to_numeric, errors="coerce")

        combined = pd.concat([general, channel], axis=1)
        combined = combined.dropna(subset=NUMERIC_COLUMNS, how="all")
        combined = combined.assign(channel=display_name)

        if "platform" not in combined.columns:
            combined["platform"] = display_name
        else:
            combined["platform"] = combined["platform"].fillna(display_name)

        channel_frames.append(combined)

    if not channel_frames:
        return pd.DataFrame()

    tidy = pd.concat(channel_frames, ignore_index=True)

    achievement_ratio = tidy["achieved_revenue"] / tidy["total_target_sales"].replace(
        {0: pd.NA}
    )
    tidy["achievement_ratio"] = achievement_ratio

    return tidy


@st.cache_data
def load_order_history() -> pd.DataFrame:
    """Load the order level history used for time-based visualisations."""

    try:
        orders = pd.read_excel(
            DATA_PATH,
            sheet_name="REGULAR-25",
            header=1,
            dtype={"Filter Store": "string", "LISTING OWNER": "string"},
        )
    except FileNotFoundError as exc:
        raise WorkbookNotFoundError(
            f"Workbook not found at '{DATA_PATH}'."
        ) from exc
    except ImportError as exc:
        raise WorkbookDependencyError(
            "Reading the GP 2025 workbook requires the optional 'openpyxl' package."
        ) from exc
    except ValueError:
        return pd.DataFrame()

    rename_map = {
        "Checkout": "checkout",
        "Total Revenue": "total_revenue",
        "NET": "net",
        "Filter Store": "filter_store",
        "LISTING OWNER": "listing_owner",
        "Platform": "platform",
        "Plain SKU": "plain_sku",
        "SKU": "sku",
    }

    orders = orders.rename(columns=rename_map)

    numeric_columns = [
        column for column in ("total_revenue", "net") if column in orders.columns
    ]

    for column in numeric_columns:
        orders[column] = pd.to_numeric(orders[column], errors="coerce")

    if "checkout" in orders.columns:
        orders["checkout"] = pd.to_datetime(orders["checkout"], errors="coerce")

    text_columns = [
        column
        for column in (
            "filter_store",
            "listing_owner",
            "platform",
            "plain_sku",
            "sku",
        )
        if column in orders.columns
    ]

    for column in text_columns:
        orders[column] = orders[column].astype("string").str.strip()
        orders[column] = orders[column].fillna("")

    required_columns = ["checkout"] + numeric_columns + text_columns
    orders = orders[[column for column in required_columns if column in orders.columns]]

    if "checkout" not in orders.columns:
        return pd.DataFrame()

    if "filter_store" in orders.columns:
        orders.loc[orders["filter_store"] == "", "filter_store"] = "All Stores"

    orders = orders.dropna(subset=["checkout"])

    if numeric_columns:
        orders = orders.dropna(subset=numeric_columns, how="all")

    return orders


def _render_dashboard(sales_data: pd.DataFrame, order_history: pd.DataFrame) -> None:
    filter_definitions = [
        ("Listing owner", "listing_owner"),
        ("Platform", "platform"),
        ("Category", "category"),
        ("Checkout", "checkout"),
        ("SKU (G)", "sku_g"),
        ("Plain SKU", "plain_sku"),
        ("DC", "dc"),
        ("Filter store", "filter_store"),
    ]

    available_filters = [
        (label, column)
        for label, column in filter_definitions
        if column in sales_data.columns
    ]

    if "filters" not in st.session_state:
        st.session_state.filters = {column: [] for _, column in available_filters}
    else:
        for _, column in available_filters:
            st.session_state.filters.setdefault(column, [])

    if "top_n" not in st.session_state:
        st.session_state.top_n = DEFAULT_TOP_N

    if hasattr(st, "popover"):
        filter_container = st.popover("Filters", use_container_width=True)
    else:
        filter_container = st.container(border=True)

    with filter_container:
        st.subheader("Filter data")

        new_selections: dict[str, list[str]] = {}
        for label, column in available_filters:
            options = sorted(sales_data[column].dropna().unique())
            new_selections[column] = st.multiselect(
                label,
                options,
                default=st.session_state.filters.get(column, []),
            )

        new_top_n = st.slider(
            "Top SKUs to show",
            min_value=5,
            max_value=25,
            value=st.session_state.top_n,
            step=1,
        )

        action_cols = st.columns(3)
        if action_cols[0].button("Apply", use_container_width=True, key="apply_filters"):
            st.session_state.filters.update(new_selections)
            st.session_state.top_n = new_top_n
            st.rerun()

        if action_cols[1].button("Reset", use_container_width=True, key="reset_filters"):
            st.session_state.filters = {column: [] for _, column in available_filters}
            st.session_state.top_n = DEFAULT_TOP_N
            st.rerun()

        if action_cols[2].button("Close", use_container_width=True, key="close_filters"):
            st.rerun()

    filter_selections = {
        column: st.session_state.filters.get(column, [])
        for _, column in available_filters
    }

    top_n = st.session_state.top_n

    def _apply_filter(series: pd.Series, choices: list[str]) -> pd.Series:
        if not choices:
            return pd.Series(True, index=series.index)
        return series.isin(choices)

    filter_mask = pd.Series(True, index=sales_data.index)
    for column, selected in filter_selections.items():
        filter_mask &= _apply_filter(sales_data[column], selected)

    filtered = sales_data[filter_mask].copy()

    if filtered.empty:
        st.warning("No records match the current filter selection.")
        st.stop()
        return

    total_achieved = filtered["achieved_revenue"].sum(min_count=1)
    total_target = filtered["total_target_sales"].sum(min_count=1)
    total_quantity = filtered["sales_quantity"].sum(min_count=1)
    total_ad_spend = filtered["ad_spend"].sum(min_count=1)

    if pd.isna(total_achieved):
        total_achieved = 0.0
    if pd.isna(total_target):
        total_target = 0.0
    if pd.isna(total_quantity):
        total_quantity = 0.0
    if pd.isna(total_ad_spend):
        total_ad_spend = 0.0

    revenue_weights = filtered["achieved_revenue"].fillna(0)
    margin_values = filtered["ip_margin"].fillna(0)
    weighted_margin = (
        (revenue_weights * margin_values).sum() / revenue_weights.sum()
        if revenue_weights.sum() > 0
        else pd.NA
    )

    delta_pct = (total_achieved / total_target - 1) * 100 if total_target else None

    metric_cols = st.columns(4)
    metrics = [
        (
            "Achieved revenue",
            f"{total_achieved:,.2f}",
            f"{delta_pct:.2f}%" if delta_pct is not None else "n/a",
        ),
        ("Target sales", f"{total_target:,.2f}", None),
        ("Units sold", f"{total_quantity:,.2f}", None),
        (
            "Weighted IP margin",
            f"{weighted_margin:.2%}" if pd.notna(weighted_margin) else "n/a",
            None,
        ),
    ]

    for column, (label, value, delta) in zip(metric_cols, metrics):
        with column:
            with st.container(border=True):
                if delta is None:
                    st.metric(label, value)
                else:
                    st.metric(label, value, delta=delta)

    order_filtered = order_history.copy()

    for column, selected in filter_selections.items():
        if not selected or column not in order_filtered.columns:
            continue

        series = order_filtered[column]
        if pd.api.types.is_datetime64_any_dtype(series):
            formatted_series = series.dt.strftime("%Y-%m-%d")
            order_filtered = order_filtered[formatted_series.isin(selected)]
        else:
            order_filtered = order_filtered[series.isin(selected)]

    with st.container(border=True):
        st.subheader("Revenue vs net by listing owner")

        if order_filtered.empty:
            st.info("No order history data available for the current selection.")
        elif "listing_owner" not in order_filtered.columns:
            st.info("Listing owner information is not available in the order history.")
        else:
            performance = (
                order_filtered.copy()
                .assign(
                    listing_owner=lambda frame: frame["listing_owner"].fillna("Unassigned")
                )
                .groupby("listing_owner")[["total_revenue", "net"]]
                .sum(min_count=1)
                .dropna(how="all")
                .reset_index()
                .rename(columns={"listing_owner": "Listing owner"})
            )

            if performance.empty:
                st.info("No order history data available for the current selection.")
            else:
                performance = performance.sort_values("total_revenue", ascending=False)
                performance = performance.head(max(st.session_state.top_n, 1))

                y_axis = alt.Y(
                    "total_revenue:Q",
                    title="Amount",
                    axis=alt.Axis(labelAngle=0),
                    scale=alt.Scale(domain=[0, _nice_upper_bound(performance["total_revenue"].max())]),
                )

                base = alt.Chart(performance).encode(
                    x=alt.X(
                        "Listing owner:N",
                        sort=list(performance["Listing owner"]),
                        title="Listing owner",
                        axis=alt.Axis(labelAngle=0),
                    ),
                    tooltip=[
                        alt.Tooltip("Listing owner:N", title="Listing owner"),
                        alt.Tooltip("total_revenue:Q", title="Revenue", format=",.2f"),
                        alt.Tooltip("net:Q", title="Net", format=",.2f"),
                    ],
                )

                revenue_bars = base.mark_bar(color="#4C78A8").encode(y=y_axis)

                revenue_labels = revenue_bars.mark_text(
                    align="center", baseline="bottom", dy=-4, color="#4C78A8"
                ).encode(text=alt.Text("total_revenue:Q", format=",.2f"))

                net_line = base.mark_line(point=True, color="#F58518").encode(
                    y=alt.Y("net:Q", axis=alt.Axis(title=None))
                )

                net_labels = net_line.mark_text(
                    align="center", baseline="bottom", dy=-12, color="#F58518"
                ).encode(text=alt.Text("net:Q", format=",.2f"))

                combined_chart = (
                    alt.layer(revenue_bars, net_line, revenue_labels, net_labels)
                    .resolve_scale(y="shared")
                    .properties(height=360)
                )

                st.altair_chart(combined_chart, use_container_width=True)

    metric_options = {
        "Achieved revenue": "achieved_revenue",
        "Units sold": "sales_quantity",
        "Advertising spend": "ad_spend",
        "Target sales": "total_target_sales",
    }

    selected_metric_label = st.selectbox(
        "Metric for category breakdown",
        list(metric_options.keys()),
        index=0,
        key="category_metric",
    )
    selected_metric = metric_options[selected_metric_label]

    category_summary = (
        filtered.groupby("category")[selected_metric]
        .sum(min_count=1)
        .sort_values(ascending=False)
        .head(12)
    )

    focus_summary = (
        filtered.groupby("focus")["achieved_revenue"]
        .sum(min_count=1)
        .sort_values(ascending=False)
        .head(12)
    )

    category_chart_data = (
        category_summary.reset_index().rename(
            columns={"category": "Category", selected_metric: "value"}
        )
    )
    focus_chart_data = (
        focus_summary.reset_index().rename(
            columns={"focus": "Focus", "achieved_revenue": "value"}
        )
    )

    col_a, col_b = st.columns(2)

    with col_a:
        with st.container(border=True):
            st.subheader("Top categories")
            if category_chart_data.empty:
                st.info("No category data available for the current selection.")
            else:
                category_upper = _nice_upper_bound(category_chart_data["value"].max())
                category_chart = (
                    alt.Chart(category_chart_data)
                    .mark_bar()
                    .encode(
                        x=alt.X(
                            "value:Q",
                            title=selected_metric_label,
                            scale=alt.Scale(domain=[0, category_upper]),
                            axis=alt.Axis(labelAngle=0),
                        ),
                        y=alt.Y("Category:N", sort="-x"),
                        tooltip=[
                            alt.Tooltip("Category:N", title="Category"),
                            alt.Tooltip(
                                "value:Q", title=selected_metric_label, format=",.2f"
                            ),
                        ],
                    )
                    .properties(height=360)
                )
                st.altair_chart(category_chart, use_container_width=True)

    with col_b:
        with st.container(border=True):
            st.subheader("Focus mix (by achieved revenue)")
            if focus_chart_data.empty:
                st.info("No focus data available for the current selection.")
            else:
                focus_upper = _nice_upper_bound(focus_chart_data["value"].max())
                focus_chart = (
                    alt.Chart(focus_chart_data)
                    .mark_bar()
                    .encode(
                        x=alt.X(
                            "value:Q",
                            title="Achieved revenue",
                            scale=alt.Scale(domain=[0, focus_upper]),
                            axis=alt.Axis(labelAngle=0),
                        ),
                        y=alt.Y("Focus:N", sort="-x"),
                        tooltip=[
                            alt.Tooltip("Focus:N", title="Focus"),
                            alt.Tooltip("value:Q", title="Achieved revenue", format=",.2f"),
                        ],
                    )
                    .properties(height=360)
                )
                st.altair_chart(focus_chart, use_container_width=True)

    with st.container(border=True):
        st.subheader("Top SKU performance")
        sku_channel_options = sorted(filtered["channel"].dropna().unique())

        if not sku_channel_options:
            st.info("No channels available after applying the current filters.")
        else:
            sku_channel = st.selectbox(
                "Channel for SKU ranking",
                sku_channel_options,
                index=0,
                key="sku_channel",
            )

            sku_subset = filtered[filtered["channel"] == sku_channel]

            if sku_subset.empty:
                st.info("No SKU data available for the selected channel.")
            else:
                top_skus = (
                    sku_subset.sort_values("achieved_revenue", ascending=False)
                    .loc[
                        :,
                        [
                            "sku",
                            "plain_sku",
                            "sku_g",
                            "category",
                            "focus",
                            "total_target_sales",
                            "achieved_revenue",
                            "sales_quantity",
                            "ad_spend",
                            "achievement_ratio",
                        ],
                    ]
                    .head(top_n)
                )

                if top_skus.empty:
                    st.info("No SKU data available for the selected channel.")
                else:
                    top_sku_chart_data = top_skus[["sku", "achieved_revenue"]].rename(
                        columns={"sku": "SKU", "achieved_revenue": "Achieved"}
                    )
                    sku_upper = _nice_upper_bound(top_sku_chart_data["Achieved"].max())
                    sku_chart = (
                        alt.Chart(top_sku_chart_data)
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "Achieved:Q",
                                title="Achieved revenue",
                                scale=alt.Scale(domain=[0, sku_upper]),
                                axis=alt.Axis(labelAngle=0),
                            ),
                            y=alt.Y("SKU:N", sort="-x"),
                            tooltip=[
                                alt.Tooltip("SKU:N", title="SKU"),
                                alt.Tooltip("Achieved:Q", title="Achieved revenue", format=",.2f"),
                            ],
                        )
                        .properties(height=400)
                    )

                    st.subheader(f"Top {len(top_skus)} SKUs by achieved revenue")
                    st.altair_chart(sku_chart, use_container_width=True)
                    display_table = top_skus.rename(
                        columns={
                            "plain_sku": "Plain SKU",
                            "sku_g": "SKU (G)",
                            "total_target_sales": "Target",
                            "achieved_revenue": "Achieved",
                            "sales_quantity": "Units",
                            "ad_spend": "Ad spend",
                            "achievement_ratio": "% to target",
                        }
                    )

                    table_formatters: dict[str, t.Callable[[float], str]] = {
                        "Target": lambda value: f"{value:,.2f}" if pd.notna(value) else "n/a",
                        "Achieved": lambda value: f"{value:,.2f}" if pd.notna(value) else "n/a",
                        "Units": lambda value: f"{value:,.2f}" if pd.notna(value) else "n/a",
                        "Ad spend": lambda value: f"{value:,.2f}" if pd.notna(value) else "n/a",
                        "% to target": lambda value: f"{value:.2%}" if pd.notna(value) else "n/a",
                    }

                    st.dataframe(
                        display_table.style.format(table_formatters),
                        use_container_width=True,
                    )

    with st.expander("View filtered records"):
        display_columns = [
            "channel",
            "platform",
            "listing_owner",
            "filter_store",
            "dc",
            "dc_list",
            "category",
            "sku",
            "plain_sku",
            "sku_g",
            "focus",
            "new_old",
            "checkout",
            "total_target_sales",
            "achieved_revenue",
            "sales_quantity",
            "ad_spend",
            "ip_margin",
        ]
        sortable = filtered[display_columns].sort_values(
            "achieved_revenue", ascending=False
        )
        value_columns = [
            column
            for column in (
                "total_target_sales",
                "achieved_revenue",
                "sales_quantity",
                "ad_spend",
                "ip_margin",
            )
            if column in sortable.columns
        ]
        value_formatters = {
            column: (lambda value: f"{value:,.2f}" if pd.notna(value) else "n/a")
            for column in value_columns
        }
        st.dataframe(
            sortable.style.format(value_formatters),
            use_container_width=True,
        )


def main() -> None:
    st.title("GP 2025 Sales Performance Dashboard")

    try:
        sales_data = load_sales_data()
        order_history = load_order_history()
    except WorkbookDependencyError:
        st.error(
            "Unable to read the workbook because the optional dependency `openpyxl` is "
            "missing. Install it with `pip install openpyxl` or `pip install -r "
            "requirements.txt` and rerun the app."
        )
        st.stop()
        return
    except WorkbookNotFoundError as exc:
        st.error(str(exc))
        st.stop()
        return

    if sales_data.empty:
        st.error("No sales data found in the workbook.")
        st.stop()
        return

    try:
        _render_dashboard(sales_data, order_history)
    except Exception as exc:  # pragma: no cover - defensive safety net
        if exc.__class__.__name__ in {"RerunException", "StopException"}:
            raise

        st.exception(exc)
        st.stop()


if __name__ == "__main__":
    main()
