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


def sort_delivery_periods(periods: list[str]) -> list[str]:
    """Sort delivery timeline labels in a human-friendly chronological order.

    Handles common label formats found in the CR file / PBIX:
      - 'MMM-YYYY' (e.g., 'Jan-2026')
      - 'Q1 2025' / 'Q4 2026'
      - '1H 2026' / '2H 2026'
      - 'TBD' and '(Blank)' (sent to the end)
    Unrecognised values are placed after recognised ones, sorted lexicographically.
    """
    if not periods:
        return []

    def key(p: str):
        s = str(p).strip()
        if s in {"", "(Blank)"}:
            return (9, 9999, 99)
        if s.upper() == "TBD":
            return (8, 9999, 99)

        # MMM-YYYY
        dt = pd.to_datetime(s, errors="coerce", format="%b-%Y")
        if pd.notna(dt):
            return (0, int(dt.year), int(dt.month))

        # Qx YYYY
        m = re.match(r"^Q([1-4])\s*(\d{4})$", s, flags=re.I)
        if m:
            q = int(m.group(1))
            y = int(m.group(2))
            # map quarter to month anchor for sorting
            return (1, y, q * 3)

        # 1H/2H YYYY
        m = re.match(r"^([12])H\s*(\d{4})$", s, flags=re.I)
        if m:
            h = int(m.group(1))
            y = int(m.group(2))
            return (2, y, 6 if h == 1 else 12)

        # Plain year
        m = re.match(r"^(\d{4})$", s)
        if m:
            return (3, int(m.group(1)), 0)

        return (7, 9999, 99, s.lower())

    return sorted([safe_clean_text(p) for p in periods], key=key)


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
# Session-state helpers
# =========================

def clear_cr_search():
    """Clear the CR search text input safely via Streamlit callback."""
    st.session_state['cr_search'] = ''



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

    # Dynamic titles to reflect whether Top-N is applied
    div_title = (
        'Change Requests by Divisions'
        if show_all_div
        else f'Change Requests by Divisions (Top {TOP_N_DEFAULT})'
    )
    po_title = (
        'Change Requests by PO'
        if show_all_po
        else f'Change Requests by PO (Top {TOP_N_DEFAULT})'
    )

    with b1:
        make_hbar(COL_DIV, div_title, show_all=show_all_div)
    with b2:
        make_hbar(COL_PO, po_title, show_all=show_all_po)

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
            key="cr_search",
            placeholder="Type part of a CR Number (matches anywhere)",
        )

    with clear_col:
        st.write("")
        st.button("Clear", use_container_width=True, on_click=clear_cr_search)

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


    # --- No results messaging (esp. for CR search) ---
    if len(df_f) == 0:
        if search_text:
            st.warning(f"No results found for CR search: '{search_text}'")
        else:
            st.warning("No results found for the selected filters.")

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


# =========================
# Page 3: CR Division (SSG)
# =========================

def build_ssg_division_page(df: pd.DataFrame):
    """Division-focused dashboard (stacked bars by Division / Year, per screenshot)."""
    df = apply_common_normalisation(df)

    COL_DIV = pick_first_existing_col(df, ["Division"])
    COL_CR = pick_first_existing_col(df, ["CR Number", "CR No", "CR"])
    COL_EFFORT = pick_first_existing_col(
        df,
        [
            "Estimated Effort",
            "Appx Effort (Only for pipeline estimation)",
            "Effort",
            "Appx Effort",
        ],
    )

    if COL_DIV is None or COL_CR is None:
        st.error("Missing required columns for this dashboard (need Division and CR Number).")
        st.write("Columns found:", list(df.columns))
        return

    # ---- Sidebar filters (match screenshot) ----
    with st.sidebar:
        st.header("Filters")

        div_opts = sorted(df[COL_DIV].dropna().unique().tolist())
        div_sel = st.selectbox("Division", ["All"] + div_opts, index=0)

        cleaned_opts = sort_delivery_periods(df["Cleaned Timeline"].dropna().unique().tolist())
        cleaned_sel = st.selectbox("Cleaned Timeline", ["All"] + cleaned_opts, index=0)

        year_opts = sorted(df["Delivery Timeline (Year)"].dropna().unique().tolist())
        year_sel = st.multiselect("Delivery Timeline (Year)", year_opts, default=year_opts)

    # ---- Apply filters ----
    df_f = df.copy()
    if div_sel != "All":
        df_f = df_f[df_f[COL_DIV] == div_sel]
    if cleaned_sel != "All":
        df_f = df_f[df_f["Cleaned Timeline"] == cleaned_sel]
    if year_sel:
        df_f = df_f[df_f["Delivery Timeline (Year)"].isin(year_sel)]

    # For distinct CR counts: exclude blank CR numbers (align with Overview KPI definition)
    df_nonblank_cr = df_f[df_f[COL_CR] != "(Blank)"].copy()

    st.subheader("CR Division (SSG)")

    # =========================
    # Chart 1: No of unique CRs by delivery period, stacked by Division
    # =========================
    st.markdown("### Number of Change Requests By Period (SSG)")

    periods = sort_delivery_periods(df_f["Cleaned Timeline"].dropna().unique().tolist())
    if not periods:
        periods = sort_delivery_periods(df["Cleaned Timeline"].dropna().unique().tolist())

    g1 = (
        df_nonblank_cr
        .groupby(["Cleaned Timeline", COL_DIV])[COL_CR]
        .nunique()
        .reset_index(name="No of CRs")
    )
    if len(g1) == 0:
        st.info("No data available for the selected filters.")
    else:
        g1["Cleaned Timeline"] = pd.Categorical(g1["Cleaned Timeline"], categories=periods, ordered=True)
        g1 = g1.sort_values(["Cleaned Timeline", COL_DIV], kind="mergesort")

        fig1 = px.bar(
            g1,
            x="Cleaned Timeline",
            y="No of CRs",
            color=COL_DIV,
            barmode="stack",
            title="",
        )
        fig1.update_layout(
            height=380,
            margin=dict(l=0, r=0, t=10, b=0),
            xaxis_title="Delivery Timeline",
            yaxis_title="Number of CRs",
            legend_title_text="Division",
        )
        st.plotly_chart(fig1, use_container_width=True, config={"displayModeBar": False})

    st.divider()

    # =========================
    # Chart 2: Effort by Division, stacked by Delivery Timeline (Year)
    # =========================
    st.markdown("### Estimated Effort (Man-days) By Divisions")
    if COL_EFFORT is None:
        st.info("Effort column not found in the dataset (expected something like 'Appx Effort (Only for pipeline estimation)').")
    else:
        g2 = (
            df_f
            .groupby([COL_DIV, "Delivery Timeline (Year)"])[COL_EFFORT]
            .apply(safe_sum)
            .reset_index(name="Estimated Effort (Man-days)")
        )
        if len(g2) == 0:
            st.info("No effort data available for the selected filters.")
        else:
            # Keep Division order stable for readability
            div_order = sorted(df_f[COL_DIV].dropna().unique().tolist())
            g2[COL_DIV] = pd.Categorical(g2[COL_DIV], categories=div_order, ordered=True)
            g2 = g2.sort_values([COL_DIV, "Delivery Timeline (Year)"], kind="mergesort")

            fig2 = px.bar(
                g2,
                x=COL_DIV,
                y="Estimated Effort (Man-days)",
                color="Delivery Timeline (Year)",
                barmode="stack",
                title="",
            )
            fig2.update_layout(
                height=380,
                margin=dict(l=0, r=0, t=10, b=0),
                xaxis_title="Division",
                yaxis_title="Estimated Effort (Man-days)",
                legend_title_text="Delivery Timeline (Year)",
            )
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

    st.divider()

    # =========================
    # Chart 3: No of unique CRs by Division, stacked by Delivery Timeline (Year)
    # =========================
    st.markdown("### No of Change Requests By Divisions")
    g3 = (
        df_nonblank_cr
        .groupby([COL_DIV, "Delivery Timeline (Year)"])[COL_CR]
        .nunique()
        .reset_index(name="No of CRs")
    )
    if len(g3) == 0:
        st.info("No data available for the selected filters.")
    else:
        div_order = sorted(df_f[COL_DIV].dropna().unique().tolist())
        g3[COL_DIV] = pd.Categorical(g3[COL_DIV], categories=div_order, ordered=True)
        g3 = g3.sort_values([COL_DIV, "Delivery Timeline (Year)"], kind="mergesort")

        fig3 = px.bar(
            g3,
            x=COL_DIV,
            y="No of CRs",
            color="Delivery Timeline (Year)",
            barmode="stack",
            title="",
        )
        fig3.update_layout(
            height=380,
            margin=dict(l=0, r=0, t=10, b=0),
            xaxis_title="Division",
            yaxis_title="No of CRs",
            legend_title_text="Delivery Timeline (Year)",
        )
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

    st.caption("Distinct CR counts exclude '(Blank)' CR Numbers. Effort follows row-grain summation and is not deduplicated.")


# =========================
# Page 4: CR Delivery Period
# =========================

def build_delivery_period_page(df: pd.DataFrame):
    """Delivery-period dashboard (3 stacked charts per screenshot)."""
    df = apply_common_normalisation(df)

    COL_CR = pick_first_existing_col(df, ["CR Number", "CR No", "CR"])
    COL_PREP = pick_first_existing_col(df, ["CR Prep Status", "Prep Status"])
    COL_DELIVERY = pick_first_existing_col(df, ["Delivery Status", "Delivery Status (Clean)"])
    COL_MODULE = pick_first_existing_col(df, ["Functional Module", "Module"])
    COL_EFFORT = pick_first_existing_col(
        df,
        [
            "Estimated Effort",
            "Appx Effort (Only for pipeline estimation)",
            "Effort",
            "Appx Effort",
        ],
    )

    required = [COL_CR, COL_PREP, COL_DELIVERY, COL_MODULE]
    if any(c is None for c in required):
        st.error(
            "Missing required columns for this dashboard. Need CR Number, CR Prep Status, Delivery Status, and Functional Module."
        )
        st.write("Columns found:", list(df.columns))
        return

    # ---- Sidebar filters ----
    with st.sidebar:
        st.header("Filters")

        year_opts = sorted(df["Delivery Timeline (Year)"].dropna().unique().tolist())
        year_sel = st.multiselect("Delivery Timeline (Year)", year_opts, default=year_opts)

        cleaned_opts = sort_delivery_periods(df["Cleaned Timeline"].dropna().unique().tolist())
        cleaned_sel = st.selectbox("Cleaned Timeline", ["All"] + cleaned_opts, index=0)

    # ---- Apply filters ----
    df_f = df.copy()
    if year_sel:
        df_f = df_f[df_f["Delivery Timeline (Year)"].isin(year_sel)]
    if cleaned_sel != "All":
        df_f = df_f[df_f["Cleaned Timeline"] == cleaned_sel]

    # Timeline order for charts
    period_order = sort_delivery_periods(df_f["Cleaned Timeline"].dropna().unique().tolist())
    if not period_order:
        period_order = sort_delivery_periods(df["Cleaned Timeline"].dropna().unique().tolist())

    # Distinct CR counts exclude blank CR Number (align with Overview KPI definition)
    df_nonblank_cr = df_f[df_f[COL_CR] != "(Blank)"].copy()

    st.subheader("CR Delivery Period")

    # Preferred legend orders (match PBIX numbering where possible)
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

    # =========================
    # Chart 1: Number of Change Requests By Delivery Period (stacked by CR Prep Status)
    # =========================
    st.markdown("### Number of Change Requests By Delivery Period")

    g1 = (
        df_nonblank_cr
        .groupby(["Cleaned Timeline", COL_PREP])[COL_CR]
        .nunique()
        .reset_index(name="Number of CRs")
    )
    if len(g1) == 0:
        st.info("No data available for the selected filters.")
    else:
        g1["Cleaned Timeline"] = pd.Categorical(g1["Cleaned Timeline"], categories=period_order, ordered=True)
        g1[COL_PREP] = force_category_order(g1[COL_PREP], ORDER_PREP)
        g1 = g1.sort_values(["Cleaned Timeline", COL_PREP], kind="mergesort")

        fig1 = px.bar(
            g1,
            x="Cleaned Timeline",
            y="Number of CRs",
            color=COL_PREP,
            barmode="stack",
        )
        fig1.update_layout(
            height=360,
            margin=dict(l=0, r=0, t=10, b=0),
            xaxis_title="Delivery Timeline",
            yaxis_title="Number of CRs",
            legend_title_text="CR Prep Status",
        )
        st.plotly_chart(fig1, use_container_width=True, config={"displayModeBar": False})

    st.divider()

    # =========================
    # Chart 2: Delivery Status By Delivery Period (Effort = row count)
    # =========================
    st.markdown("### Change Requests Delivery Status By Delivery Period")

    # Row-grain effort: each row is one unit of effort/workload
    g2 = (
        df_f
        .groupby(["Cleaned Timeline", COL_DELIVERY])
        .size()
        .reset_index(name="Effort")
    )
    if len(g2) == 0:
        st.info("No data available for the selected filters.")
    else:
        g2["Cleaned Timeline"] = pd.Categorical(g2["Cleaned Timeline"], categories=period_order, ordered=True)
        g2[COL_DELIVERY] = force_category_order(g2[COL_DELIVERY], ORDER_DELIVERY)
        g2 = g2.sort_values(["Cleaned Timeline", COL_DELIVERY], kind="mergesort")

        fig2 = px.bar(
            g2,
            x="Cleaned Timeline",
            y="Effort",
            color=COL_DELIVERY,
            barmode="stack",
        )
        fig2.update_layout(
            height=360,
            margin=dict(l=0, r=0, t=10, b=0),
            xaxis_title="Delivery Timeline",
            yaxis_title="Effort",
            legend_title_text="Delivery Status",
        )
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

    st.divider()

    # =========================
    # Chart 3: Estimated Effort Across Delivery Period (stacked by Functional Module)
    # =========================
    st.markdown("### Estimated Effort Across Delivery Period")
    if COL_EFFORT is None:
        st.info("Effort column not found in the dataset (expected something like 'Appx Effort (Only for pipeline estimation)').")
    else:
        g3 = (
            df_f
            .groupby(["Cleaned Timeline", COL_MODULE])[COL_EFFORT]
            .apply(safe_sum)
            .reset_index(name="Estimated Effort")
        )
        if len(g3) == 0:
            st.info("No effort data available for the selected filters.")
        else:
            g3["Cleaned Timeline"] = pd.Categorical(g3["Cleaned Timeline"], categories=period_order, ordered=True)
            g3 = g3.sort_values(["Cleaned Timeline", COL_MODULE], kind="mergesort")

            fig3 = px.bar(
                g3,
                x="Cleaned Timeline",
                y="Estimated Effort",
                color=COL_MODULE,
                barmode="stack",
            )
            fig3.update_layout(
                height=360,
                margin=dict(l=0, r=0, t=10, b=0),
                xaxis_title="Delivery Timeline",
                yaxis_title="Effort",
                legend_title_text="Functional Module",
            )
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

    st.caption(
        "Distinct CR counts exclude '(Blank)' CR Numbers. Delivery Status chart uses row count as Effort (row-grain workload)."
    )


# =========================
# Page 5: CR By Period (WOG)
# =========================

def build_wog_period_page(df: pd.DataFrame):
    """Whole-of-Government style view: unique CRs by delivery period, stacked by Division (per screenshot)."""
    df = apply_common_normalisation(df)

    COL_DIV = pick_first_existing_col(df, ["Division"])
    COL_CR = pick_first_existing_col(df, ["CR Number", "CR No", "CR"])

    if COL_DIV is None or COL_CR is None:
        st.error("Missing required columns for this dashboard (need Division and CR Number).")
        st.write("Columns found:", list(df.columns))
        return

    # ---- Sidebar filter (match screenshot: only Division) ----
    with st.sidebar:
        st.header("Filters")
        div_opts = sorted(df[COL_DIV].dropna().unique().tolist())
        div_sel = st.selectbox("Division", ["All"] + div_opts, index=0)

    df_f = df.copy()
    if div_sel != "All":
        df_f = df_f[df_f[COL_DIV] == div_sel]

    # Distinct CR counts exclude blank CR Number (align with Overview KPI definition)
    df_nonblank_cr = df_f[df_f[COL_CR] != "(Blank)"].copy()

    st.subheader("CR By Period (WOG)")
    st.markdown("### Number of Change Requests By Period (SSG)")

    period_order = sort_delivery_periods(df_f["Cleaned Timeline"].dropna().unique().tolist())
    if not period_order:
        period_order = sort_delivery_periods(df["Cleaned Timeline"].dropna().unique().tolist())

    g = (
        df_nonblank_cr
        .groupby(["Cleaned Timeline", COL_DIV])[COL_CR]
        .nunique()
        .reset_index(name="Number of CRs")
    )

    if len(g) == 0:
        st.info("No data available for the selected filters.")
        return

    g["Cleaned Timeline"] = pd.Categorical(g["Cleaned Timeline"], categories=period_order, ordered=True)
    g = g.sort_values(["Cleaned Timeline", COL_DIV], kind="mergesort")

    fig = px.bar(
        g,
        x="Cleaned Timeline",
        y="Number of CRs",
        color=COL_DIV,
        barmode="stack",
        title="",
    )
    fig.update_layout(
        height=520,
        margin=dict(l=0, r=0, t=10, b=0),
        xaxis_title="Delivery Timeline",
        yaxis_title="Number of CRs",
        legend_title_text="Division",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.caption("Distinct CR counts exclude '(Blank)' CR Numbers.")




# =========================
# Page 6: 2026 Workplan
# =========================

def build_workplan_page(df: pd.DataFrame):
    """Card/grid view of CRs by delivery period (workplan style, per screenshot)."""
    df = apply_common_normalisation(df)

    COL_CR = pick_first_existing_col(df, ["CR Number", "CR No", "CR"])
    COL_TITLE = pick_first_existing_col(df, ["Title", "CR Title"])
    COL_EFFORT = pick_first_existing_col(
        df,
        [
            "Estimated Effort",
            "Appx Effort (Only for pipeline estimation)",
            "Effort",
            "Appx Effort",
        ],
    )

    if COL_CR is None or COL_TITLE is None:
        st.error("Missing required columns for this dashboard (need CR Number and Title).")
        st.write("Columns found:", list(df.columns))
        return

    # ---- Sidebar filters ----
    with st.sidebar:
        st.header("Filters")

        year_opts = sorted(df["Delivery Timeline (Year)"].dropna().unique().tolist())
        # Default to 2026 if present; else all
        default_year = "2026" if "2026" in year_opts else (year_opts[-1] if year_opts else "All")
        year_sel = st.selectbox("Delivery Timeline (Year)", ["All"] + year_opts, index=(1 + year_opts.index(default_year)) if default_year in year_opts else 0)

        # Optional: filter to a single period (keeps layout manageable when needed)
        period_opts_all = sort_delivery_periods(df["Cleaned Timeline"].dropna().unique().tolist())
        period_sel = st.selectbox("Cleaned Timeline", ["All"] + period_opts_all, index=0)

        # Layout option
        cols_per_row = st.selectbox("Cards per row", [2, 3, 4], index=2)

    # ---- Apply filters ----
    df_f = df.copy()
    if year_sel != "All":
        df_f = df_f[df_f["Delivery Timeline (Year)"] == year_sel]

    if period_sel != "All":
        df_f = df_f[df_f["Cleaned Timeline"] == period_sel]

    st.subheader("CR Workplan")

    # Period order (include TBD if present)
    period_order = sort_delivery_periods(df_f["Cleaned Timeline"].dropna().unique().tolist())
    if not period_order:
        st.info("No data available for the selected filters.")
        return

    # CSS to approximate PBIX card styling
    st.markdown(
        """
        <style>
        .wp-card-title {
            background: #1f4e79;
            color: white;
            padding: 8px 10px;
            font-weight: 700;
            border-radius: 10px 10px 0 0;
            font-size: 0.95rem;
        }
        .wp-card-body {
            border: 1px solid #d0d0d0;
            border-top: 0;
            border-radius: 0 0 10px 10px;
            padding: 8px 10px 10px 10px;
            background: white;
        }
        .wp-metric-left {
            background: #f3e6b3;
            padding: 6px 10px;
            border-radius: 0 0 0 10px;
            text-align: center;
            font-weight: 700;
            font-size: 0.9rem;
            line-height: 1.15;
        }
        .wp-metric-right {
            background: #f7c8b5;
            padding: 6px 10px;
            border-radius: 0 0 10px 0;
            text-align: center;
            font-weight: 700;
            font-size: 0.9rem;
            line-height: 1.15;
            white-space: normal;
            word-break: break-word;
        }
        .wp-period-label {
            background: #bcd7f4;
            color: #0b2d4d;
            padding: 6px 10px;
            border-radius: 8px;
            text-align: center;
            font-weight: 700;
            margin-bottom: 6px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Build cards
    n_cols = int(cols_per_row)
    rows = [period_order[i:i+n_cols] for i in range(0, len(period_order), n_cols)]

    for row_periods in rows:
        cols = st.columns(n_cols)
        for i in range(n_cols):
            if i >= len(row_periods):
                cols[i].empty()
                continue
            period = row_periods[i]
            with cols[i]:
                # Period label should appear ABOVE the table (more intuitive than placing it below)
                st.markdown(f'<div class="wp-period-label">{period}</div>', unsafe_allow_html=True)

                # Filter per card period
                df_p = df_f[df_f["Cleaned Timeline"] == period].copy()

                # Table (CR Number + Title), keep duplicates/blanks as-is (row grain)
                table_df = df_p[[COL_CR, COL_TITLE]].copy()
                # Sort stable for readability
                table_df = table_df.sort_values([COL_CR, COL_TITLE], kind="mergesort")

                # Metrics
                # Unique CRs: distinct non-blank CR Number (align with your KPI definition)
                uniq_cr = int(table_df.loc[table_df[COL_CR] != "(Blank)", COL_CR].nunique())

                # Effort: sum of effort column if present; show '(Blank)' when all blank
                effort_display = "(Blank)"
                if COL_EFFORT is not None and COL_EFFORT in df_p.columns:
                    eff_num = pd.to_numeric(df_p[COL_EFFORT], errors="coerce")
                    if eff_num.notna().any():
                        effort_display = f"{int(eff_num.fillna(0).sum()):,}"

                # Render header
                st.markdown('<div class="wp-card-title">CR Number&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Title</div>', unsafe_allow_html=True)

                with st.container(border=True):
                    st.dataframe(
                        table_df,
                        use_container_width=True,
                        hide_index=True,
                        height=200,
                    )

                    # Bottom metrics band (2 cells)
                    m1, m2 = st.columns(2)
                    with m1:
                        st.markdown(f'<div class="wp-metric-left">{uniq_cr}</div>', unsafe_allow_html=True)
                    with m2:
                        st.markdown(f'<div class="wp-metric-right">{effort_display}</div>', unsafe_allow_html=True)


    st.caption(
        "Card view: table preserves row-grain records (duplicates + '(Blank)' CR Numbers). "
        "The left metric is distinct non-blank CR count; the right metric is summed estimated effort (if available)."
    )


# =========================
# Page 7: CR Category Dashboard (Pie + Division stacked bar)
# =========================

def build_cr_category_page(df: pd.DataFrame):
    """Combined Broad Category dashboard (per your screenshots).

    One page with shared filters + three visuals:
      1) Pie: Efforts Across Broad Categories (effort = sum of effort values by rows)
      2) Pie: Change Requests by Broad Categories (distinct, non-blank CR count)
         + KPI card: total distinct, non-blank CRs
      3) Stacked bar: CR Category by Divisions (distinct, non-blank CR count)

    Grain rules preserved:
      - Row = workload / effort grain (no dedup for effort)
      - Multi-row CRs and blank CR Numbers remain legitimate
      - Distinct CR counts exclude '(Blank)' CR Numbers (align with Overview KPI)
    """

    df = apply_common_normalisation(df)

    COL_CR = pick_first_existing_col(df, ["CR Number", "CR No", "CR"])
    COL_DIV = pick_first_existing_col(df, ["Division"])
    COL_PREP = pick_first_existing_col(df, ["CR Prep Status", "Prep Status"])
    COL_CAT = pick_first_existing_col(df, ["Broad Category", "Category", "CR Category"])
    COL_EFFORT = pick_first_existing_col(
        df,
        [
            "Estimated Effort",
            "Appx Effort (Only for pipeline estimation)",
            "Effort",
            "Appx Effort",
        ],
    )

    if any(c is None for c in [COL_CR, COL_DIV, COL_PREP, COL_CAT]):
        st.error(
            "Missing required columns for this dashboard. Need CR Number, Division, CR Prep Status, and Broad Category."
        )
        st.write("Columns found:", list(df.columns))
        return

    # ---- Sidebar filters (shared; match screenshots) ----
    with st.sidebar:
        st.header("Filters")

        year_opts = sorted(df["Delivery Timeline (Year)"].dropna().unique().tolist())
        year_sel = st.selectbox("Delivery Timeline (Year)", ["All"] + year_opts, index=0)

        cleaned_opts = sort_delivery_periods(df["Cleaned Timeline"].dropna().unique().tolist())
        cleaned_sel = st.selectbox("Cleaned Timeline", ["All"] + cleaned_opts, index=0)

        prep_opts = sorted(df[COL_PREP].dropna().unique().tolist())
        prep_sel = st.multiselect("CR Prep Status", prep_opts, default=prep_opts)

        div_opts = sorted(df[COL_DIV].dropna().unique().tolist())
        div_sel = st.multiselect("Division", div_opts, default=div_opts)

    # ---- Apply filters ----
    df_f = df.copy()
    if year_sel != "All":
        df_f = df_f[df_f["Delivery Timeline (Year)"] == year_sel]
    if cleaned_sel != "All":
        df_f = df_f[df_f["Cleaned Timeline"] == cleaned_sel]
    if prep_sel:
        df_f = df_f[df_f[COL_PREP].isin(prep_sel)]
    if div_sel:
        df_f = df_f[df_f[COL_DIV].isin(div_sel)]

    # Distinct CR counts exclude blank CR Number (align with Overview KPI definition)
    df_nonblank_cr = df_f[df_f[COL_CR] != "(Blank)"].copy()
    total_unique_cr = int(df_nonblank_cr[COL_CR].nunique())

    st.subheader("CR Broad Categories")

    # Preferred legend order (match screenshot)
    ORDER_CAT = [
        "Operations",
        "(Blank)",
        "Policy",
        "Improve Stakeholder Experience",
        "IM8 / Governance",
        "Increase Productivity",
        "New Feature(s)",
        "Technical",
    ]

    # =========================
    # Visual 1: Efforts Across Broad Categories (pie)
    # =========================
    st.markdown("### Efforts Across Broad Categories")
    if COL_EFFORT is None:
        st.info("Effort column not found (expected something like 'Appx Effort (Only for pipeline estimation)').")
    else:
        g_eff = (
            df_f
            .groupby(COL_CAT)[COL_EFFORT]
            .apply(safe_sum)
            .reset_index(name="Effort")
        )
        if len(g_eff) == 0 or g_eff["Effort"].sum() == 0:
            st.info("No effort data available for the selected filters.")
        else:
            g_eff[COL_CAT] = force_category_order(g_eff[COL_CAT], ORDER_CAT)
            g_eff = g_eff.sort_values(COL_CAT, kind="mergesort")

            fig_eff = px.pie(
                g_eff,
                names=COL_CAT,
                values="Effort",
                title="",
                category_orders={COL_CAT: list(g_eff[COL_CAT].cat.categories)},
            )
            fig_eff.update_traces(
                textinfo="value+percent",
                textposition="outside",
                textfont_size=12,
                hovertemplate=f"{COL_CAT}: %{{label}}<br>Effort: %{{value}}<br>%{{percent}}<extra></extra>",
            )
            fig_eff.update_layout(
                # Larger margins so outside labels are not clipped (esp. left/top)
                height=520,
                margin=dict(l=60, r=220, t=30, b=60),
                uniformtext_minsize=10,
                uniformtext_mode="hide",
                legend_title_text="Broad Category",
            )
            st.plotly_chart(fig_eff, use_container_width=True, config={"displayModeBar": False})

    st.divider()

    # =========================
    # Visual 2: Change Requests by Broad Categories (pie) + KPI
    # =========================
    st.markdown("### Change Requests by Broad Categories")
    top_left, top_right = st.columns([4, 1])

    with top_left:
        g_pie = (
            df_nonblank_cr
            .groupby(COL_CAT)[COL_CR]
            .nunique()
            .reset_index(name="Count")
        )

        if len(g_pie) == 0:
            st.info("No data available for the selected filters.")
        else:
            g_pie[COL_CAT] = force_category_order(g_pie[COL_CAT], ORDER_CAT)
            g_pie = g_pie.sort_values(COL_CAT, kind="mergesort")

            fig_pie = px.pie(
                g_pie,
                names=COL_CAT,
                values="Count",
                title="",
                category_orders={COL_CAT: list(g_pie[COL_CAT].cat.categories)},
            )
            fig_pie.update_traces(
                textinfo="value+percent",
                textposition="outside",
                textfont_size=12,
                hovertemplate=f"{COL_CAT}: %{{label}}<br>Count: %{{value}}<br>%{{percent}}<extra></extra>",
            )
            fig_pie.update_layout(
                # Larger margins so outside labels are not clipped (esp. left/top)
                height=520,
                margin=dict(l=60, r=220, t=30, b=60),
                uniformtext_minsize=10,
                uniformtext_mode="hide",
                legend_title_text="Broad Category",
            )
            st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})

    with top_right:
        st.markdown(
            """
            <div style="text-align:right; padding-top: 40px;">
              <div style="font-size: 18px; color: #333;">No of Change Requests<br>(CR)</div>
              <div style="font-size: 64px; font-weight: 800; color: #1f4e79; line-height: 1;">{}</div>
            </div>
            """.format(total_unique_cr),
            unsafe_allow_html=True,
        )

    st.divider()

    # =========================
    # Visual 3: CR Category by Divisions (stacked)
    # =========================
    st.markdown("### CR Category by Divisions")

    g_bar = (
        df_nonblank_cr
        .groupby([COL_DIV, COL_CAT])[COL_CR]
        .nunique()
        .reset_index(name="Count")
    )

    if len(g_bar) == 0:
        st.info("No data available for the selected filters.")
        return

    div_order = sorted(df_nonblank_cr[COL_DIV].dropna().unique().tolist())
    g_bar[COL_DIV] = pd.Categorical(g_bar[COL_DIV], categories=div_order, ordered=True)
    g_bar[COL_CAT] = force_category_order(g_bar[COL_CAT], ORDER_CAT)
    g_bar = g_bar.sort_values([COL_DIV, COL_CAT], kind="mergesort")

    fig_bar = px.bar(
        g_bar,
        x=COL_DIV,
        y="Count",
        color=COL_CAT,
        barmode="stack",
        title="",
        category_orders={COL_DIV: div_order, COL_CAT: list(g_bar[COL_CAT].cat.categories)},
    )
    fig_bar.update_layout(
        height=430,
        margin=dict(l=0, r=0, t=10, b=0),
        xaxis_title="Division",
        yaxis_title="Count of #",
        legend_title_text="Broad Category",
    )
    st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})

    st.caption(
        "Counts are distinct non-blank CR Numbers (same definition as the Overview 'No of Unique CR'). "
        "Effort is summed by rows (no CR deduplication)."
    )


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
        st.header("Dashboards")
        st.radio(
            "",
            [
                "CR Overview",
                "CR Details",
                "CR Division (SSG)",
                "CR Delivery Period",
                "CR By Period (WOG)",
                "CR Broad Categories",
                "CR Workplan",
            ],
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
    elif st.session_state.selected_dashboard == "CR Details":
        build_details_page(df_fact)
    elif st.session_state.selected_dashboard == "CR Division (SSG)":
        build_ssg_division_page(df_fact)
    elif st.session_state.selected_dashboard == "CR Delivery Period":
        build_delivery_period_page(df_fact)
    elif st.session_state.selected_dashboard == "CR By Period (WOG)":
        build_wog_period_page(df_fact)
    elif st.session_state.selected_dashboard == "CR Broad Categories":
        build_cr_category_page(df_fact)
    elif st.session_state.selected_dashboard == "CR Workplan":
        build_workplan_page(df_fact)
    else:
        # Safe fallback
        build_overview_page(df_fact)


if __name__ == "__main__":
    main()