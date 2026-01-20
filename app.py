import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================
# Configuration
# =========================
DEFAULT_FILENAME_HINT = "CR Pipeline-Master_02Dec2025_PO_EY.xlsx"
SHEET_FACT = "CR Master"
SHEET_LOOKUP = "Look Up"  # loaded for potential future validation

TOP_N_DEFAULT = 10  # per your requirement


# =========================
# Cleaning & standardisation
# =========================

def _to_blank(x) -> str:
    """Match Power BI '(Blank)' behaviour: null/empty -> '(Blank)'."""
    if pd.isna(x):
        return "(Blank)"
    if isinstance(x, str) and x.strip() == "":
        return "(Blank)"
    return x


def safe_clean_text(x):
    """
    Safe, non-semantic normalisation to reduce label drift:
    - trim
    - collapse multiple spaces
    - normalise dash types
    - normalise spacing around slashes
    - keep '(Blank)' explicit
    """
    x = _to_blank(x)
    if x == "(Blank)":
        return x
    if not isinstance(x, str):
        return x

    s = x.strip()
    # normalise unicode dashes to hyphen
    s = s.replace("–", "-").replace("—", "-")
    # collapse multiple spaces
    s = re.sub(r"\s+", " ", s)
    # normalise slash spacing: "A / B" -> "A/B"
    s = re.sub(r"\s*/\s*", "/", s)
    return s


def derive_cleaned_timeline(val):
    """
    Best-effort 'Cleaned Timeline' derivation:
    - blank / TBD-like -> 'TBD'
    - parseable date -> 'MMM-YYYY'
    - otherwise keep trimmed text (e.g., '1H 2026', 'Q1 2025')
    """
    if pd.isna(val):
        return "TBD"
    if isinstance(val, str):
        s = val.strip()
        if s == "":
            return "TBD"
        if re.search(r"\bTBD\b|\bTBC\b|to be confirmed|to be determined|pending", s, flags=re.I):
            return "TBD"
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%b-%Y")
        return safe_clean_text(s)

    dt = pd.to_datetime(val, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%b-%Y")
    return "TBD"


def year_from_cleaned(cleaned_timeline: str) -> str:
    m = re.search(r"(\d{4})$", str(cleaned_timeline))
    return m.group(1) if m else "TBD"


def safe_sum(series: pd.Series) -> float:
    s = pd.to_numeric(series, errors="coerce")
    return float(s.fillna(0).sum())


def force_category_order(series: pd.Series, preferred_order: list[str]) -> pd.Categorical:
    """
    Fixes legend order and slice order like Power BI:
    - use preferred_order when values exist
    - include any unexpected values at the end (sorted) so nothing disappears
    """
    uniq = [x for x in preferred_order if x in set(series.dropna().unique())]
    extras = [x for x in series.dropna().unique() if x not in uniq]
    final = uniq + sorted(extras)
    return pd.Categorical(series, categories=final, ordered=True)


def pick_first_existing_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def apply_common_normalisation(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise core text columns + derive timeline fields consistently."""
    df = df.copy()

    cat_cols = [
        "Division",
        "PO",
        "Functional Module",
        "CR Prep Status",
        "MoSCoW Priorities",
        "Agency",
        "Delivery Status",
        "Delivery Timeline",
        "Cleaned Timeline",
        "CR Number",
        "Title",
    ]
    for c in cat_cols:
        if c in df.columns:
            df[c] = df[c].apply(safe_clean_text)

    if "Cleaned Timeline" not in df.columns:
        if "Delivery Timeline" in df.columns:
            df["Cleaned Timeline"] = df["Delivery Timeline"].apply(derive_cleaned_timeline)
        else:
            df["Cleaned Timeline"] = "TBD"
    else:
        df["Cleaned Timeline"] = df["Cleaned Timeline"].apply(safe_clean_text)

    if "Delivery Timeline (Year)" not in df.columns:
        df["Delivery Timeline (Year)"] = df["Cleaned Timeline"].apply(year_from_cleaned)
    else:
        df["Delivery Timeline (Year)"] = df["Delivery Timeline (Year)"].apply(safe_clean_text)

    return df


# =========================
# Top N + Others
# =========================

def top_n_with_others(value_counts: pd.Series, top_n: int = TOP_N_DEFAULT) -> pd.DataFrame:
    """Returns Top N categories + 'Others (n=XX)' where XX is total count of remaining rows."""
    vc = value_counts.copy()
    if len(vc) <= top_n:
        out = vc.reset_index()
        out.columns = ["Category", "Count"]
        return out

    top = vc.head(top_n)
    others_count = int(vc.iloc[top_n:].sum())
    out = top.reset_index()
    out.columns = ["Category", "Count"]
    out.loc[len(out)] = [f"Others (n={others_count})", others_count]
    return out


def counts_for_chart(df: pd.DataFrame, group_col: str, show_all: bool, top_n: int = TOP_N_DEFAULT) -> pd.DataFrame:
    vc = df[group_col].fillna("(Blank)").value_counts(dropna=False)
    if show_all:
        out = vc.reset_index()
        out.columns = ["Category", "Count"]
        return out
    return top_n_with_others(vc, top_n=top_n)


# =========================
# Data loading
# =========================
@st.cache_data(show_spinner=False)
def load_excel(uploaded_file_or_path):
    """Load both sheets from Excel."""
    df_fact = pd.read_excel(uploaded_file_or_path, sheet_name=SHEET_FACT)
    try:
        df_lookup = pd.read_excel(uploaded_file_or_path, sheet_name=SHEET_LOOKUP)
    except Exception:
        df_lookup = pd.DataFrame()
    return df_fact, df_lookup


# =========================
# Page 1: CR Overview
# =========================

def build_overview_page(df: pd.DataFrame):
    df = apply_common_normalisation(df)

    # Column mapping (expected from your source)
    COL_CR_ID = "CR Number" if "CR Number" in df.columns else "#"
    COL_EFFORT = "Appx Effort (Only for pipeline estimation)"
    COL_DIV = "Division"
    COL_PO = "PO"
    COL_MOSCOW = "MoSCoW Priorities"
    COL_PREP = "CR Prep Status"
    COL_DELIVERY_STATUS = "Delivery Status"

    # Sidebar filters (match the Power BI overview experience)
    with st.sidebar:
        st.header("Filters")

        year_opts = sorted(df["Delivery Timeline (Year)"].dropna().unique().tolist())
        year_sel = st.selectbox("Delivery Timeline (Year)", ["All"] + year_opts, index=0)

        cleaned_opts = sorted(df["Cleaned Timeline"].dropna().unique().tolist())
        cleaned_sel = st.selectbox("Cleaned Timeline", ["All"] + cleaned_opts, index=0)

        div_opts = sorted(df[COL_DIV].dropna().unique().tolist()) if COL_DIV in df.columns else []
        div_sel = st.selectbox("Division", ["All"] + div_opts, index=0)

        po_opts = sorted(df[COL_PO].dropna().unique().tolist()) if COL_PO in df.columns else []
        po_sel = st.selectbox("PO", ["All"] + po_opts, index=0)

        st.divider()
        st.subheader("Display options")
        show_all_div = st.checkbox("Show all divisions", value=False)
        show_all_po = st.checkbox("Show all POs", value=False)

    # Apply filters
    df_f = df.copy()
    if year_sel != "All":
        df_f = df_f[df_f["Delivery Timeline (Year)"] == year_sel]
    if cleaned_sel != "All":
        df_f = df_f[df_f["Cleaned Timeline"] == cleaned_sel]
    if div_sel != "All" and COL_DIV in df_f.columns:
        df_f = df_f[df_f[COL_DIV] == div_sel]
    if po_sel != "All" and COL_PO in df_f.columns:
        df_f = df_f[df_f[COL_PO] == po_sel]

    # KPIs
    if COL_CR_ID in df_f.columns:
        # Option A: Exclude blanks from distinct CR count (aligns with earlier 205)
        total_cr = int(df_f.loc[df_f[COL_CR_ID] != "(Blank)", COL_CR_ID].nunique())
    else:
        total_cr = int(len(df_f))
    total_effort = int(safe_sum(df_f[COL_EFFORT])) if COL_EFFORT in df_f.columns else 0

    st.subheader("CR Overview")

    kpi1, kpi2 = st.columns([1, 1])
    with kpi1:
        st.markdown("<div class='kpi-title'>No of Unique Change Requests<br>(CR)</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='kpi-value'>{total_cr:,}</div>", unsafe_allow_html=True)

    with kpi2:
        st.markdown("<div class='kpi-title'>Estimated Effort<br>(Man-days)</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='kpi-value'>{total_effort:,}</div>", unsafe_allow_html=True)

    # Preferred orders (to match PBIX-stage numbering)
    ORDER_MOSCOW = ["Must Have", "Should Have", "(Blank)", "Could Have", "Won't Have"]
    ORDER_PREP = [
        "(Blank)",
        "0. CR Not Initiated",
        "1. CR Initiated",
        "4. CR Presented to PO/DTI",
        "5. CR Presented to Business",
        "6. CR Approved",
        "7. CR Cancelled",
    ]
    ORDER_DELIVERY = [
        "(Blank)",
        "0. Not Started",
        "1. CR Approved",
        "2. CR Retracted/On-hold",
        "3. Dev in Progress",
        "4. Dev Completed",
        "6. UAT completed",
        "8. Deployed",
    ]

    def make_pie(col: str, title: str, order: list[str]):
        if col not in df_f.columns:
            st.info(f"Column not found: {col}")
            return
        s = df_f[col].fillna("(Blank)")
        s_cat = force_category_order(s, order)
        counts = pd.DataFrame({col: s_cat}).value_counts().reset_index(name="Count")

        fig = px.pie(
            counts,
            names=col,
            values="Count",
            title=title,
            hole=0.0,
            category_orders={col: list(s_cat.categories)},
        )
        fig.update_traces(
            textinfo="percent",
            textposition="inside",
            insidetextorientation="auto",
            hovertemplate=f"{col}: %{{label}}<br>Count: %{{value}}<br>%{{percent}}<extra></extra>",
        )
        fig.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=40, b=10),
            legend=dict(font=dict(size=10), itemsizing="constant"),
            title=dict(x=0.0, xanchor="left"),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    pie1, pie2, pie3 = st.columns([1, 1, 1])
    with pie1:
        make_pie(COL_MOSCOW, "MoSCoW", ORDER_MOSCOW)
    with pie2:
        make_pie(COL_PREP, "CR Prep Status", ORDER_PREP)
    with pie3:
        make_pie(COL_DELIVERY_STATUS, "Delivery Status", ORDER_DELIVERY)

    st.divider()
    b1, b2 = st.columns([1, 1])

    def make_hbar(group_col: str, title: str, show_all: bool):
        if group_col not in df_f.columns:
            st.info(f"Column not found: {group_col}")
            return

        counts_df = counts_for_chart(df_f, group_col, show_all=show_all, top_n=TOP_N_DEFAULT)
        counts_df = counts_df.sort_values("Count", ascending=True)

        fig = px.bar(
            counts_df,
            x="Count",
            y="Category",
            orientation="h",
            text="Count",
            title=title,
        )
        fig.update_traces(textposition="outside", hovertemplate=f"{group_col}: %{{y}}<br>Count: %{{x}}<extra></extra>")
        fig.update_layout(height=360, margin=dict(l=0, r=0, t=40, b=0), title=dict(x=0.0, xanchor="left"))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with b1:
        make_hbar(COL_DIV, f"Change Requests by Divisions (Top {TOP_N_DEFAULT})", show_all=show_all_div)
    with b2:
        make_hbar(COL_PO, f"Change Requests by PO (Top {TOP_N_DEFAULT})", show_all=show_all_po)

    st.caption("Prototype: Excel-driven Streamlit dashboard (internal use).")


# =========================
# Page 2: CR Details (row-grain)
# =========================

def build_details_page(df: pd.DataFrame):
    df = apply_common_normalisation(df)

    # Column mapping (robust to slight schema drift)
    COL_CR = pick_first_existing_col(df, ["CR Number", "CR No", "CR"])
    COL_MODULE = pick_first_existing_col(df, ["Functional Module", "Module"])
    COL_AGENCY = pick_first_existing_col(df, ["Agency"])
    COL_DIV = pick_first_existing_col(df, ["Division"])
    COL_TITLE = pick_first_existing_col(df, ["Title", "CR Title"])
    COL_PO = pick_first_existing_col(df, ["PO"])
    COL_MOSCOW = pick_first_existing_col(df, ["MoSCoW Priorities", "MoSCoW", "MoSCoW Priority"])
    COL_PREP = pick_first_existing_col(df, ["CR Prep Status", "Prep Status"])

    # Effort column: support both the pipeline column and the screenshot label
    COL_EFFORT = pick_first_existing_col(
        df,
        [
            "Estimated Effort",
            "Appx Effort (Only for pipeline estimation)",
            "Effort",
            "Appx Effort",
        ],
    )

    # --- Sidebar filters (CR selector must remain) ---
    with st.sidebar:
        st.header("Filters")

        year_opts = sorted(df["Delivery Timeline (Year)"].dropna().unique().tolist())
        year_sel = st.selectbox("Delivery Timeline (Year)", ["All"] + year_opts, index=0)

        cleaned_opts = sorted(df["Cleaned Timeline"].dropna().unique().tolist())
        cleaned_sel = st.selectbox("Cleaned Timeline", ["All"] + cleaned_opts, index=0)

        div_opts = sorted(df[COL_DIV].dropna().unique().tolist()) if COL_DIV else []
        div_sel = st.selectbox("Division", ["All"] + div_opts, index=0)

        po_opts = sorted(df[COL_PO].dropna().unique().tolist()) if COL_PO else []
        po_sel = st.selectbox("PO", ["All"] + po_opts, index=0)

        # Retain CR selector (not optional / not removed)
        cr_opts = sorted(df[COL_CR].dropna().unique().tolist()) if COL_CR else []
        cr_sel = st.selectbox("CR Number", ["All"] + cr_opts, index=0)

        mod_opts = sorted(df[COL_MODULE].dropna().unique().tolist()) if COL_MODULE else []
        mod_sel = st.selectbox("Functional Module", ["All"] + mod_opts, index=0)

        prep_opts = sorted(df[COL_PREP].dropna().unique().tolist()) if COL_PREP else []
        prep_sel = st.selectbox("CR Prep Status", ["All"] + prep_opts, index=0)
    # --- Main header + search box (right under subheader) ---
    st.subheader("CR Details")

    search_col, clear_col = st.columns([5, 1])

    with search_col:
        search_text = st.text_input(
            "Search CR Number (e.g. 422)",
            key="cr_search_input",
            placeholder="Type part of a CR Number (matches anywhere)",
        )

    with clear_col:
        st.write("")
        if st.button("Clear", use_container_width=True):
            # Clear BOTH: the table filter and the visible text in the input box
            st.session_state.pop("cr_search_input", None)
            st.rerun()

    search_text = (search_text or "").strip()

    # --- Apply filters (row-grain filtering) ---
    df_f = df.copy()
    if year_sel != "All":
        df_f = df_f[df_f["Delivery Timeline (Year)"] == year_sel]
    if cleaned_sel != "All":
        df_f = df_f[df_f["Cleaned Timeline"] == cleaned_sel]
    if div_sel != "All" and COL_DIV:
        df_f = df_f[df_f[COL_DIV] == div_sel]
    if po_sel != "All" and COL_PO:
        df_f = df_f[df_f[COL_PO] == po_sel]
    if cr_sel != "All" and COL_CR:
        df_f = df_f[df_f[COL_CR] == cr_sel]
    if mod_sel != "All" and COL_MODULE:
        df_f = df_f[df_f[COL_MODULE] == mod_sel]
    if prep_sel != "All" and COL_PREP:
        df_f = df_f[df_f[COL_PREP] == prep_sel]

    # CR search: substring contains (case-insensitive)
    if search_text and COL_CR:
        mask = df_f[COL_CR].astype(str).str.contains(re.escape(search_text), case=False, na=False)
        df_f = df_f[mask]

    # --- KPIs (IMPORTANT: row count, not distinct CR) ---
    total_rows = int(len(df_f))
    total_effort = int(safe_sum(df_f[COL_EFFORT])) if COL_EFFORT else 0

    k1, k2 = st.columns([1, 1])
    with k1:
        st.markdown("<div class='kpi-title'>No of Change Request</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='kpi-value'>{total_rows:,}</div>", unsafe_allow_html=True)
    with k2:
        st.markdown("<div class='kpi-title'>Total Estimated Effort</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='kpi-value'>{total_effort:,}</div>", unsafe_allow_html=True)

    # --- Clear filtered state indicator (KPIs + table) ---
    active_filters = []
    if year_sel != "All":
        active_filters.append(f"Year={year_sel}")
    if cleaned_sel != "All":
        active_filters.append(f"Timeline={cleaned_sel}")
    if div_sel != "All":
        active_filters.append(f"Division={div_sel}")
    if po_sel != "All":
        active_filters.append(f"PO={po_sel}")
    if cr_sel != "All":
        active_filters.append(f"CR={cr_sel}")
    if mod_sel != "All":
        active_filters.append(f"Module={mod_sel}")
    if prep_sel != "All":
        active_filters.append(f"Prep={prep_sel}")
    if search_text:
        active_filters.append(f"CR contains '{search_text}'")

    if active_filters:
        st.info("Filtered by: " + " | ".join(active_filters))
    else:
        st.caption("Showing all rows (no filters applied).")

    # --- No results message ---
    if df_f.empty:
        if search_text:
            st.warning(f"No results found for CR search: '{search_text}'.")
        else:
            st.warning("No results found for the selected filters.")

    st.divider()

    # --- Display table ---
    display_cols = []
    for c in [
        COL_CR,
        COL_MODULE,
        COL_AGENCY,
        COL_DIV,
        COL_TITLE,
        COL_PO,
        "Cleaned Timeline",
        "Delivery Timeline (Year)",
        COL_EFFORT,
        COL_MOSCOW,
        COL_PREP,
    ]:
        if c and c in df_f.columns and c not in display_cols:
            display_cols.append(c)

    # Default sort
    sort_cols = [c for c in ["Delivery Timeline (Year)", "Cleaned Timeline", COL_CR] if c and c in df_f.columns]
    if sort_cols:
        df_f = df_f.sort_values(sort_cols, ascending=True, kind="mergesort")

    st.dataframe(
        df_f[display_cols] if display_cols else df_f,
        use_container_width=True,
        hide_index=True,
        height=620,
    )

    st.caption("Row-grain table: duplicates and blank CR Numbers are intentionally kept to reflect workload items.")


def main():
    st.set_page_config(page_title="CR Dashboard Prototype", layout="wide")

    st.markdown(
        """
        <style>
        /* Prevent top title clipping (esp. after reruns / on Streamlit Cloud) */
        .block-container { padding-top: 2rem; padding-bottom: 1rem; }
        h1, h2 { margin-top: 0.25rem; padding-top: 0.25rem; line-height: 1.15; }

        /* Simple KPI styling */
        .kpi-title { font-size: 0.95rem; color: #666; }
        .kpi-value { font-size: 2.1rem; font-weight: 700; margin-top: -0.25rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("## CR Dashboard Prototype (Streamlit)")
    st.write("Upload the Excel source file to render the dashboards.")

    # --- Seamless navigation: switcher at top of sidebar + session persistence ---
    if "selected_dashboard" not in st.session_state:
        st.session_state.selected_dashboard = "CR Overview"

    with st.sidebar:
        st.header("Dashboard")
        st.radio(
            "",
            ["CR Overview", "CR Details"],
            key="selected_dashboard",
        )
        st.divider()

    uploaded = st.file_uploader(
        "Upload Excel (.xlsx)",
        type=["xlsx"],
        help=f"Expected sheets: '{SHEET_FACT}' (required), '{SHEET_LOOKUP}' (optional). File hint: {DEFAULT_FILENAME_HINT}",
    )

    # Optional: allow a local file path for internal testing (dev only)
    with st.expander("Developer option: load from local path (optional)"):
        path_str = st.text_input("Local file path", value="")
        use_path = st.checkbox("Use local file path instead of upload", value=False)

    if uploaded is None and not (use_path and path_str.strip()):
        st.info("Please upload the Excel file to proceed.")
        return

    source = Path(path_str.strip()) if (use_path and path_str.strip()) else uploaded

    try:
        df_fact, df_lookup = load_excel(source)
    except Exception as e:
        st.error("Failed to read the Excel file. Please verify the file format and sheet names.")
        st.exception(e)
        return

    if st.session_state.selected_dashboard == "CR Overview":
        build_overview_page(df_fact)
    else:
        build_details_page(df_fact)


if __name__ == "__main__":
    main()
