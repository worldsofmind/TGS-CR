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
    Best-effort 'Cleaned Timeline' derivation (since it may exist only in PBIX):
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


# =========================
# Top N + Others
# =========================
def top_n_with_others(value_counts: pd.Series, top_n: int = TOP_N_DEFAULT) -> pd.DataFrame:
    """
    Returns Top N categories + 'Others (n=XX)' where XX is the total count of remaining rows.
    """
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
    """
    Load both sheets from Excel.
    Accepts either:
    - an UploadedFile (from st.file_uploader)
    - a filesystem path (string/Path)
    """
    df_fact = pd.read_excel(uploaded_file_or_path, sheet_name=SHEET_FACT)
    try:
        df_lookup = pd.read_excel(uploaded_file_or_path, sheet_name=SHEET_LOOKUP)
    except Exception:
        df_lookup = pd.DataFrame()
    return df_fact, df_lookup


def build_overview_page(df: pd.DataFrame):
    # Column mapping (expected from your source)
    COL_CR_ID = "CR Number" if "CR Number" in df.columns else "#"
    COL_EFFORT = "Appx Effort (Only for pipeline estimation)"
    COL_DIV = "Division"
    COL_PO = "PO"
    COL_MOSCOW = "MoSCoW Priorities"
    COL_PREP = "CR Prep Status"
    COL_DELIVERY_STATUS = "Delivery Status"
    COL_DELIVERY_TIMELINE = "Delivery Timeline"

    # Normalise key categorical columns
    for c in [COL_DIV, COL_PO, COL_MOSCOW, COL_PREP, COL_DELIVERY_STATUS]:
        if c in df.columns:
            df[c] = df[c].apply(safe_clean_text)

    # Derive timeline fields if needed
    if "Cleaned Timeline" not in df.columns:
        df["Cleaned Timeline"] = df[COL_DELIVERY_TIMELINE].apply(derive_cleaned_timeline) if COL_DELIVERY_TIMELINE in df.columns else "TBD"
    else:
        df["Cleaned Timeline"] = df["Cleaned Timeline"].apply(safe_clean_text)

    if "Delivery Timeline (Year)" not in df.columns:
        df["Delivery Timeline (Year)"] = df["Cleaned Timeline"].apply(year_from_cleaned)
    else:
        df["Delivery Timeline (Year)"] = df["Delivery Timeline (Year)"].apply(safe_clean_text)

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
    total_cr = int(df_f[COL_CR_ID].nunique()) if COL_CR_ID in df_f.columns else int(len(df_f))
    total_effort = int(safe_sum(df_f[COL_EFFORT])) if COL_EFFORT in df_f.columns else 0

    # Layout
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.0rem; padding-bottom: 1.0rem; }
        .kpi-title { font-size: 18px; font-weight: 600; margin-bottom: 0.25rem; }
        .kpi-value { font-size: 64px; font-weight: 800; color: #1f4e79; line-height: 1; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("CR Overview")

    kpi1, kpi2, pie1, pie2, pie3 = st.columns([1.25, 1.25, 1.1, 1.1, 1.1])

    with kpi1:
        st.markdown("<div class='kpi-title'>No of Change Requests<br>(CR)</div>", unsafe_allow_html=True)
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
            hole=0.0,  # Pie (not donut) to resemble your screenshot
            category_orders={col: list(s_cat.categories)},
        )
        fig.update_traces(
            textinfo="percent",
            textposition="outside",
            hovertemplate=f"{col}: %{{label}}<br>Count: %{{value}}<br>%{{percent}}<extra></extra>",
        )
        fig.update_layout(height=260, margin=dict(l=0, r=0, t=40, b=0), title=dict(x=0.0, xanchor="left"))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

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
        # Sort so the largest appears at the top (Power BI-like for horizontal bars)
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


def main():
    st.set_page_config(page_title="CR Dashboard Prototype", layout="wide")

    st.title("CR Dashboard Prototype (Streamlit)")
    st.write("Upload the Excel source file to render the prototype overview page.")

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

    # For now, we only build the first visualisation page (CR Overview).
    build_overview_page(df_fact)


if __name__ == "__main__":
    main()
