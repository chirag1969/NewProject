import pandas as pd
import streamlit as st
from pathlib import Path


st.set_page_config(
    page_title="GP 2025 Sales Dashboard",
    page_icon=":bar_chart:",
    layout="wide",
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
    "CAT": "category",
    "Sr. No.": "serial_number",
    "SKU": "sku",
    "AMZN": "amazon_manager",
    "EBAY": "ebay_manager",
    "Website": "website_manager",
    "FOCUS": "focus",
    "NEW/OLD": "new_old",
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

    general = workbook["General"].rename(columns=GENERAL_RENAME)
    general = general.dropna(subset=["sku"]).copy()
    general["sku"] = general["sku"].astype(str).str.strip()
    general = general[general["sku"].str.lower() != "nan"]

    general["focus"] = general["focus"].fillna("Unspecified")
    general["category"] = general["category"].fillna("Uncategorised")
    general["dc_list"] = general["dc_list"].fillna("Unknown")
    general["new_old"] = general["new_old"].fillna("Unspecified")

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

        channel_frames.append(combined)

    if not channel_frames:
        return pd.DataFrame()

    tidy = pd.concat(channel_frames, ignore_index=True)

    achievement_ratio = tidy["achieved_revenue"] / tidy["total_target_sales"].replace(
        {0: pd.NA}
    )
    tidy["achievement_ratio"] = achievement_ratio

    return tidy


def main() -> None:
    st.title("GP 2025 Sales Performance Dashboard")
    st.caption(
        "Explore the GP 2025 workbook and compare channel performance against targets. "
        "Monetary values are shown using the original workbook units."
    )

    try:
        sales_data = load_sales_data()
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

    st.sidebar.header("Filters")

    channel_options = list(CHANNEL_LABELS.values())
    selected_channels = st.sidebar.multiselect(
        "Channels",
        channel_options,
        default=channel_options,
    )

    category_options = sorted(sales_data["category"].dropna().unique())
    selected_categories = st.sidebar.multiselect(
        "Categories",
        category_options,
        default=category_options,
    )

    focus_options = sorted(sales_data["focus"].dropna().unique())
    selected_focus = st.sidebar.multiselect(
        "Focus buckets",
        focus_options,
        default=focus_options,
    )

    status_options = sorted(sales_data["new_old"].dropna().unique())
    selected_status = st.sidebar.multiselect(
        "Assortment status",
        status_options,
        default=status_options,
    )

    top_n = st.sidebar.slider(
        "Top SKUs to show", min_value=5, max_value=25, value=10, step=1
    )

    def _apply_filter(series: pd.Series, choices: list[str]) -> pd.Series:
        if not choices:
            return pd.Series(True, index=series.index)
        return series.isin(choices)

    filtered = sales_data[
        _apply_filter(sales_data["channel"], selected_channels)
        & _apply_filter(sales_data["category"], selected_categories)
        & _apply_filter(sales_data["focus"], selected_focus)
        & _apply_filter(sales_data["new_old"], selected_status)
    ].copy()

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

    st.markdown(f"**{len(filtered):,}** channel records after filtering")

    metric_cols = st.columns(4)

    metric_cols[0].metric(
        "Achieved revenue",
        f"{total_achieved:,.0f}",
        delta=f"{delta_pct:.1f}%" if delta_pct is not None else "n/a",
    )
    metric_cols[1].metric("Target sales", f"{total_target:,.0f}")
    metric_cols[2].metric("Units sold", f"{total_quantity:,.0f}")
    metric_cols[3].metric(
        "Weighted IP margin",
        f"{weighted_margin:.1%}" if pd.notna(weighted_margin) else "n/a",
    )

    st.divider()

    channel_summary = (
        filtered.groupby("channel")[
            ["total_target_sales", "achieved_revenue", "sales_quantity", "ad_spend"]
        ]
        .sum()
        .sort_values("achieved_revenue", ascending=False)
    )

    st.subheader("Channel performance")
    st.bar_chart(channel_summary[["total_target_sales", "achieved_revenue"]])
    st.dataframe(
        channel_summary.rename(
            columns={
                "total_target_sales": "Total target",
                "achieved_revenue": "Achieved revenue",
                "sales_quantity": "Units sold",
                "ad_spend": "Advertising spend",
            }
        ),
        use_container_width=True,
    )

    st.divider()

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

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Top categories")
        st.bar_chart(category_summary)

    with col_b:
        st.subheader("Focus mix (by achieved revenue)")
        st.bar_chart(focus_summary)

    st.divider()

    sku_channel_options = selected_channels or channel_options
    sku_channel = st.selectbox(
        "Channel for SKU ranking",
        sku_channel_options,
        index=0,
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

        st.subheader(f"Top {len(top_skus)} SKUs by achieved revenue")
        st.bar_chart(
            top_skus.set_index("sku")["achieved_revenue"],
        )
        st.dataframe(
            top_skus.rename(
                columns={
                    "total_target_sales": "Target",
                    "achieved_revenue": "Achieved",
                    "sales_quantity": "Units",
                    "ad_spend": "Ad spend",
                    "achievement_ratio": "% to target",
                }
            ),
            use_container_width=True,
        )

    with st.expander("View filtered records"):
        display_columns = [
            "channel",
            "dc_list",
            "category",
            "sku",
            "focus",
            "new_old",
            "total_target_sales",
            "achieved_revenue",
            "sales_quantity",
            "ad_spend",
            "ip_margin",
        ]
        st.dataframe(
            filtered[display_columns].sort_values("achieved_revenue", ascending=False),
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
