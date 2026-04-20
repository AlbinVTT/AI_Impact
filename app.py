import io
import re
from pathlib import Path
from typing import List, Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="GenAI Revenue Impact Dashboard",
    page_icon="📈",
    layout="wide",
)

DEFAULT_WORKBOOK_NAME = "Ai_Impact_Analysis.xlsx"
APP_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK_PATH = APP_DIR / DEFAULT_WORKBOOK_NAME
SCENARIO_ORDER = ["Low", "Base", "High"]


# -----------------------------
# Styling
# -----------------------------
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.4rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    .hero-card {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 1.4rem 1.6rem;
        border-radius: 22px;
        color: white;
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.20);
        margin-bottom: 1rem;
    }
    .subtle-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 1rem 1.1rem;
        border-radius: 18px;
        box-shadow: 0 6px 20px rgba(15, 23, 42, 0.05);
        min-height: 128px;
        color: #0f172a;
    }
    .highlight-good {
        background: #ecfdf5;
        border: 1px solid #a7f3d0;
        padding: 1rem 1.1rem;
        border-radius: 18px;
        min-height: 130px;
        color: #0f172a;
    }
    .highlight-risk {
        background: #fff7ed;
        border: 1px solid #fed7aa;
        padding: 1rem 1.1rem;
        border-radius: 18px;
        min-height: 130px;
        color: #0f172a;
    }
    .rank-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 0.9rem 1rem;
        box-shadow: 0 6px 20px rgba(15, 23, 42, 0.05);
        min-height: 122px;
        color: #0f172a;
    }
    .small-muted {
        color: #475569;
        font-size: 0.92rem;
    }
    .section-title {
        font-size: 1.15rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        color: inherit;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Helper functions
# -----------------------------
def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_key(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalize_text(value).lower())


def parse_number(value):
    if value is None or pd.isna(value):
        return pd.NA
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    if s == "":
        return pd.NA

    for token in ["$", ",", "USD", "usd"]:
        s = s.replace(token, "")
    s = s.replace("B", "").replace("b", "")
    s = s.replace("bn", "").replace("Bn", "").replace("%", "").strip()

    try:
        return float(s)
    except ValueError:
        return pd.NA


def to_num_series(series: pd.Series) -> pd.Series:
    return series.apply(parse_number)


def to_share_value(value):
    v = parse_number(value)
    if pd.isna(v):
        return pd.NA
    v = float(v)
    return v / 100 if v > 1 else v


def fmt_money(value) -> str:
    if value is None or pd.isna(value):
        return "N/A"
    value = float(value)
    if abs(value) < 1:
        return f"${value * 1000:,.0f}M"
    return f"${value:,.2f}B"


def fmt_pct(value) -> str:
    if value is None or pd.isna(value):
        return "N/A"
    value = float(value)
    if abs(value) <= 1:
        return f"{value * 100:.1f}%"
    return f"{value:.1f}%"


def format_portfolio_table(df: pd.DataFrame) -> pd.DataFrame:
    display_df = df.copy()
    if "Growth 2025-2027" in display_df.columns:
        display_df["Growth 2025-2027"] = display_df["Growth 2025-2027"].apply(fmt_pct)
    if "AI Uplift 2027" in display_df.columns:
        display_df["AI Uplift 2027"] = display_df["AI Uplift 2027"].apply(fmt_money)
    if "Cannibalization 2027" in display_df.columns:
        display_df["Cannibalization 2027"] = display_df["Cannibalization 2027"].apply(fmt_money)
    if "Risk Score" in display_df.columns:
        display_df["Risk Score"] = display_df["Risk Score"].apply(
            lambda x: "N/A" if pd.isna(x) else round(float(x), 3)
        )
    return display_df


@st.cache_data
def read_sheet_with_detected_header(
    workbook_bytes: bytes,
    sheet_name: str,
    search_terms: Optional[List[str]] = None,
    search_rows: int = 20,
) -> pd.DataFrame:
    if search_terms is None:
        search_terms = ["Company"]

    preview = pd.read_excel(io.BytesIO(workbook_bytes), sheet_name=sheet_name, header=None, nrows=search_rows)
    header_row = 0
    normalized_terms = [normalize_key(term) for term in search_terms]

    for idx in range(len(preview)):
        row_values = [normalize_key(v) for v in preview.iloc[idx].tolist() if pd.notna(v)]
        if any(term in row_values for term in normalized_terms):
            header_row = idx
            break

    df = pd.read_excel(io.BytesIO(workbook_bytes), sheet_name=sheet_name, header=header_row)
    return df.dropna(how="all")


@st.cache_data
def load_workbook_data(workbook_bytes: bytes):
    raw_data = read_sheet_with_detected_header(workbook_bytes, "Raw_Data", ["Company"])
    assumptions = read_sheet_with_detected_header(workbook_bytes, "Assumptions", ["Company", "Scenario"])
    revenue_mix_raw = read_sheet_with_detected_header(workbook_bytes, "Revenue_Mix", ["Company"])
    sources = read_sheet_with_detected_header(workbook_bytes, "Sources", ["Company"])
    return raw_data, assumptions, revenue_mix_raw, sources


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned.columns = [normalize_text(c) for c in cleaned.columns]
    cleaned = cleaned.loc[:, ~cleaned.columns.str.contains(r"^Unnamed", case=False, regex=True)]
    return cleaned.dropna(how="all")


def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    if df.empty:
        return None

    col_map = {normalize_key(c): c for c in df.columns}

    for candidate in candidates:
        key = normalize_key(candidate)
        if key in col_map:
            return col_map[key]

    for col in df.columns:
        col_key = normalize_key(col)
        for candidate in candidates:
            candidate_key = normalize_key(candidate)
            if candidate_key in col_key or col_key in candidate_key:
                return col

    return None


def find_secondary_company_col(df: pd.DataFrame, primary_company_col: Optional[str]) -> Optional[str]:
    for col in df.columns:
        if primary_company_col and col == primary_company_col:
            continue
        if normalize_key(col).startswith("company"):
            return col
    return None


def normalize_company_series(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.lower()


def keep_valid_companies(df: pd.DataFrame, company_col: Optional[str], valid_companies: set[str]) -> pd.DataFrame:
    if df.empty or not company_col or company_col not in df.columns:
        return df.copy()

    cleaned = df.copy()
    normalized = normalize_company_series(cleaned[company_col])

    invalid_tokens = {"total", "grand total", "review point", "reviewpoint", "review point:"}
    mask = normalized.isin(valid_companies)
    mask &= ~normalized.isin(invalid_tokens)
    mask &= ~normalized.str.contains(r"review", na=False)
    mask &= ~normalized.str.contains(r"^total$", na=False)

    return cleaned.loc[mask].copy()


def build_interpretation_map(
    revenue_mix_raw: pd.DataFrame,
    primary_company_col: Optional[str],
) -> dict[str, str]:
    interpretation_col = find_col(revenue_mix_raw, ["Interpretation"])
    secondary_company_col = find_secondary_company_col(revenue_mix_raw, primary_company_col)

    if not interpretation_col or not secondary_company_col:
        return {}

    interp_df = revenue_mix_raw[[secondary_company_col, interpretation_col]].copy()
    interp_df = interp_df.dropna(subset=[secondary_company_col, interpretation_col])
    interp_df[secondary_company_col] = normalize_company_series(interp_df[secondary_company_col])

    return dict(zip(interp_df[secondary_company_col], interp_df[interpretation_col].astype(str)))


def calculate_scenario_model(raw_data: pd.DataFrame, assumptions: pd.DataFrame, selected_scenario: str) -> pd.DataFrame:
    raw_company_col = find_col(raw_data, ["Company"])
    ass_company_col = find_col(assumptions, ["Company"])
    ass_scenario_col = find_col(assumptions, ["Scenario"])

    revenue_col = find_col(raw_data, ["Revenue_2025_USD_Bn", "Revenue_2025", "Revenue 2025"])
    existing_ai_share_col = find_col(
        raw_data,
        ["Existing_AI_Revenue_Share_2025_%", "Existing AI Revenue Share 2025 %", "AI Share 2025"],
    )

    if not raw_company_col or not ass_company_col or not ass_scenario_col or not revenue_col:
        return pd.DataFrame()

    ass_filtered = assumptions[
        normalize_company_series(assumptions[ass_scenario_col]) == selected_scenario.lower()
    ].copy()

    col_map = {
        "Baseline_Growth_2026_%": find_col(ass_filtered, ["Baseline_Growth_2026_%", "Base_Growth_2026_%"]),
        "Baseline_Growth_2027_%": find_col(ass_filtered, ["Baseline_Growth_2027_%", "Base_Growth_2027_%"]),
        "AI_Exposed_Revenue_%": find_col(ass_filtered, ["AI_Exposed_Revenue_%", "AI_Exposed_%"]),
        "AI_Adoption_2026_%": find_col(ass_filtered, ["AI_Adoption_2026_%", "Adoption_2026_%"]),
        "AI_Adoption_2027_%": find_col(ass_filtered, ["AI_Adoption_2027_%", "Adoption_2027_%"]),
        "Monetization_Yield_2026_%": find_col(
            ass_filtered, ["Monetization_Yield_2026_%", "Monetization_2026_%"]
        ),
        "Monetization_Yield_2027_%": find_col(
            ass_filtered, ["Monetization_Yield_2027_%", "Monetization_2027_%"]
        ),
        "Cannibalization_2026_%": find_col(ass_filtered, ["Cannibalization_2026_%"]),
        "Cannibalization_2027_%": find_col(ass_filtered, ["Cannibalization_2027_%"]),
    }

    base = raw_data.copy()
    base[raw_company_col] = base[raw_company_col].astype(str).str.strip()

    ass_subset_cols = [ass_company_col] + [v for v in col_map.values() if v]
    ass_subset = ass_filtered[ass_subset_cols].copy()
    ass_subset[ass_company_col] = ass_subset[ass_company_col].astype(str).str.strip()

    merged = base.merge(ass_subset, left_on=raw_company_col, right_on=ass_company_col, how="left")
    if ass_company_col in merged.columns and ass_company_col != raw_company_col:
        merged = merged.drop(columns=[ass_company_col])

    merged["Revenue_2025"] = to_num_series(merged[revenue_col])
    merged["Existing_AI_Revenue_Share_2025_%"] = (
        merged[existing_ai_share_col].fillna(0).apply(to_share_value) if existing_ai_share_col else 0
    )

    for target, source in col_map.items():
        if source:
            merged[target] = to_num_series(merged[source])
        else:
            merged[target] = 0

    merged["Revenue_2026_Baseline"] = merged["Revenue_2025"] * (1 + merged["Baseline_Growth_2026_%"])
    merged["AI_Base_2026"] = merged["Revenue_2026_Baseline"] * merged["AI_Exposed_Revenue_%"]
    merged["AI_Uplift_2026"] = (
        merged["AI_Base_2026"] * merged["AI_Adoption_2026_%"] * merged["Monetization_Yield_2026_%"]
    )
    merged["Cannibalization_2026"] = merged["AI_Base_2026"] * merged["Cannibalization_2026_%"]
    merged["Net_Revenue_2026"] = (
        merged["Revenue_2026_Baseline"] + merged["AI_Uplift_2026"] - merged["Cannibalization_2026"]
    )

    merged["Revenue_2027_Baseline"] = merged["Net_Revenue_2026"] * (1 + merged["Baseline_Growth_2027_%"])
    merged["AI_Base_2027"] = merged["Revenue_2027_Baseline"] * merged["AI_Exposed_Revenue_%"]
    merged["AI_Uplift_2027"] = (
        merged["AI_Base_2027"] * merged["AI_Adoption_2027_%"] * merged["Monetization_Yield_2027_%"]
    )
    merged["Cannibalization_2027"] = merged["AI_Base_2027"] * merged["Cannibalization_2027_%"]
    merged["Net_Revenue_2027"] = (
        merged["Revenue_2027_Baseline"] + merged["AI_Uplift_2027"] - merged["Cannibalization_2027"]
    )

    merged["Growth_2025_2027_Pct"] = pd.NA
    valid_growth = merged["Revenue_2025"].notna() & merged["Net_Revenue_2027"].notna() & (merged["Revenue_2025"] != 0)
    merged.loc[valid_growth, "Growth_2025_2027_Pct"] = (
        merged.loc[valid_growth, "Net_Revenue_2027"] / merged.loc[valid_growth, "Revenue_2025"] - 1
    )

    merged["Risk_Score"] = pd.NA
    valid_risk = merged["AI_Uplift_2027"].notna() & merged["Cannibalization_2027"].notna() & (merged["AI_Uplift_2027"] > 0)
    merged.loc[valid_risk, "Risk_Score"] = (
        merged.loc[valid_risk, "Cannibalization_2027"] / merged.loc[valid_risk, "AI_Uplift_2027"]
    )

    return merged


def calculate_revenue_mix(
    scenario_df: pd.DataFrame,
    company_col: str,
    interpretation_map: dict[str, str],
) -> pd.DataFrame:
    mix = scenario_df[[company_col, "Existing_AI_Revenue_Share_2025_%", "AI_Uplift_2026", "Net_Revenue_2026", "AI_Uplift_2027", "Net_Revenue_2027"]].copy()
    mix = mix.rename(columns={"Existing_AI_Revenue_Share_2025_%": "AI_Share_2025_%"}).copy()

    uplift_share_2026 = pd.Series(0.0, index=mix.index)
    uplift_share_2027 = pd.Series(0.0, index=mix.index)

    valid_2026 = mix["Net_Revenue_2026"].notna() & (mix["Net_Revenue_2026"] != 0)
    valid_2027 = mix["Net_Revenue_2027"].notna() & (mix["Net_Revenue_2027"] != 0)

    uplift_share_2026.loc[valid_2026] = mix.loc[valid_2026, "AI_Uplift_2026"] / mix.loc[valid_2026, "Net_Revenue_2026"]
    uplift_share_2027.loc[valid_2027] = mix.loc[valid_2027, "AI_Uplift_2027"] / mix.loc[valid_2027, "Net_Revenue_2027"]

    mix["Legacy_Share_2025_%"] = 1 - mix["AI_Share_2025_%"]
    mix["AI_Share_2026_%"] = (mix["AI_Share_2025_%"] + uplift_share_2026).clip(upper=1)
    mix["Legacy_Share_2026_%"] = 1 - mix["AI_Share_2026_%"]
    mix["AI_Share_2027_%"] = (mix["AI_Share_2025_%"] + uplift_share_2026 + uplift_share_2027).clip(upper=1)
    mix["Legacy_Share_2027_%"] = 1 - mix["AI_Share_2027_%"]
    mix["Shift_in_AI_Share_2027_vs_2025_%"] = mix["AI_Share_2027_%"] - mix["AI_Share_2025_%"]
    mix["Interpretation"] = normalize_company_series(mix[company_col]).map(interpretation_map)

    return mix


def render_rank_card(
    rank_label: str,
    company: str,
    primary_label: str,
    primary_value: str,
    secondary_label: str,
    secondary_value: str,
):
    st.markdown(
        f"""
        <div class='rank-card'>
            <div class='small-muted'>{rank_label}</div>
            <div style='font-size:1.12rem;font-weight:800;margin-top:0.15rem'>{company}</div>
            <div style='margin-top:0.65rem;font-size:1.05rem;font-weight:700'>{primary_value}</div>
            <div class='small-muted'>{primary_label}</div>
            <div style='margin-top:0.45rem;font-size:0.98rem;font-weight:600'>{secondary_value}</div>
            <div class='small-muted'>{secondary_label}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------
# App shell
# -----------------------------
st.markdown(
    """
    <div class='hero-card'>
        <div style='font-size:0.95rem;opacity:0.85;'>Interactive decision view</div>
        <div style='font-size:3rem;font-weight:900;line-height:1.08;margin-top:0.35rem;'>Agentic AI / GenAI Revenue Impact Dashboard</div>
        <div style='font-size:1.04rem;opacity:0.88;margin-top:0.7rem;'>Assess 2026–2027 AI-driven revenue growth, revenue mix shift, monetization upside, and downside risk across selected US-listed software companies.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.header("Controls")
st.sidebar.caption(f"Default workbook: {DEFAULT_WORKBOOK_NAME}")

uploaded_workbook = st.sidebar.file_uploader(
    "Optional: upload another workbook",
    type=["xlsx"],
    help="If nothing is uploaded, the app will use the workbook stored in the same folder as app.py.",
)

if uploaded_workbook is not None:
    workbook_bytes = uploaded_workbook.getvalue()
    workbook_label = uploaded_workbook.name
else:
    if not DEFAULT_WORKBOOK_PATH.exists():
        st.error(
            f"Workbook not found. Keep `{DEFAULT_WORKBOOK_NAME}` in the same folder as app.py, or upload a workbook from the sidebar."
        )
        st.stop()
    workbook_bytes = DEFAULT_WORKBOOK_PATH.read_bytes()
    workbook_label = DEFAULT_WORKBOOK_PATH.name

try:
    raw_data, assumptions, revenue_mix_raw, sources = load_workbook_data(workbook_bytes)
except Exception as exc:
    st.error(f"Could not read the workbook. Please check sheet names and structure. Error: {exc}")
    st.stop()

raw_data = normalize_columns(raw_data)
assumptions = normalize_columns(assumptions)
revenue_mix_raw = normalize_columns(revenue_mix_raw)
sources = normalize_columns(sources)

raw_company_col = find_col(raw_data, ["Company"])
source_company_col = find_col(sources, ["Company"])

if raw_company_col is None:
    st.error("Could not find a Company column in Raw_Data.")
    st.stop()

company_list = sorted(
    raw_data[raw_company_col].dropna().astype(str).str.strip().unique().tolist()
)
valid_company_keys = {company.strip().lower() for company in company_list}

raw_data = keep_valid_companies(raw_data, raw_company_col, valid_company_keys)
if source_company_col:
    sources = keep_valid_companies(sources, source_company_col, valid_company_keys)

company_list = sorted(
    raw_data[raw_company_col].dropna().astype(str).str.strip().unique().tolist()
)

selected_company = st.sidebar.selectbox("Company", company_list)
selected_scenario = st.sidebar.selectbox("Scenario", SCENARIO_ORDER, index=1)

interpretation_map = build_interpretation_map(revenue_mix_raw, find_col(revenue_mix_raw, ["Company"]))

scenario_model = calculate_scenario_model(raw_data, assumptions, selected_scenario)
if scenario_model.empty:
    st.error("Could not calculate scenario outputs from Raw_Data and Assumptions.")
    st.stop()

scenario_company_col = find_col(scenario_model, ["Company"])
revenue_mix = calculate_revenue_mix(scenario_model, scenario_company_col, interpretation_map)

raw_company = raw_data[
    normalize_company_series(raw_data[raw_company_col]) == selected_company.strip().lower()
].copy()

company_scenario = scenario_model[
    normalize_company_series(scenario_model[scenario_company_col]) == selected_company.strip().lower()
].copy()

mix_company_col = find_col(revenue_mix, ["Company"])
mix_company = revenue_mix[
    normalize_company_series(revenue_mix[mix_company_col]) == selected_company.strip().lower()
].copy()

portfolio_scored = scenario_model[
    [
        scenario_company_col,
        "Growth_2025_2027_Pct",
        "AI_Uplift_2027",
        "Cannibalization_2027",
        "Risk_Score",
    ]
].copy()
portfolio_scored = portfolio_scored.rename(
    columns={
        "AI_Uplift_2027": "AI_Uplift_2027_Calc",
        "Cannibalization_2027": "Cannibalization_2027_Calc",
    }
)

winner_company = "N/A"
winner_growth = None
if not portfolio_scored.empty and portfolio_scored["Growth_2025_2027_Pct"].notna().any():
    winner_row = portfolio_scored.sort_values("Growth_2025_2027_Pct", ascending=False).iloc[0]
    winner_company = str(winner_row[scenario_company_col])
    winner_growth = winner_row["Growth_2025_2027_Pct"]

risk_company = "N/A"
risk_score = None
if not portfolio_scored.empty and portfolio_scored["Risk_Score"].notna().any():
    risk_row = portfolio_scored.sort_values("Risk_Score", ascending=False).iloc[0]
    risk_company = str(risk_row[scenario_company_col])
    risk_score = risk_row["Risk_Score"]

portfolio_revenue_2025 = scenario_model["Revenue_2025"].sum(min_count=1)
portfolio_revenue_2027 = scenario_model["Net_Revenue_2027"].sum(min_count=1)
portfolio_ai_uplift_2027 = scenario_model["AI_Uplift_2027"].sum(min_count=1)

# -----------------------------
# Portfolio header metrics
# -----------------------------
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Portfolio Revenue 2025", fmt_money(portfolio_revenue_2025))
with m2:
    delta_display = None
    if (
        pd.notna(portfolio_revenue_2025)
        and pd.notna(portfolio_revenue_2027)
        and float(portfolio_revenue_2025) != 0
    ):
        delta_display = f"{((float(portfolio_revenue_2027) / float(portfolio_revenue_2025)) - 1) * 100:.1f}% vs 2025"
    st.metric(f"Portfolio Revenue 2027 ({selected_scenario})", fmt_money(portfolio_revenue_2027), delta=delta_display)
with m3:
    st.metric("Total AI Uplift 2027", fmt_money(portfolio_ai_uplift_2027))
with m4:
    st.metric("Companies Evaluated", str(len(company_list)))

# -----------------------------
# Winner / risk highlights
# -----------------------------
h1, h2 = st.columns(2)
with h1:
    st.markdown(
        f"""
        <div class='highlight-good'>
            <div class='small-muted'>Winner highlight</div>
            <div style='font-size:1.35rem;font-weight:800;margin-top:0.2rem'>{winner_company}</div>
            <div style='margin-top:0.65rem;font-size:1.02rem;'>Highest modeled 2025–2027 revenue expansion under the selected scenario.</div>
            <div style='margin-top:0.6rem;font-weight:700'>{fmt_pct(winner_growth)}</div>
            <div class='small-muted'>Growth from 2025 to 2027</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with h2:
    risk_note = "Highest relative downside pressure from cannibalization versus AI uplift."
    if risk_score is None or pd.isna(risk_score):
        risk_note = "Risk score unavailable in the current workbook structure."
    risk_value = "N/A" if risk_score is None or pd.isna(risk_score) else round(float(risk_score), 2)
    st.markdown(
        f"""
        <div class='highlight-risk'>
            <div class='small-muted'>High-risk highlight</div>
            <div style='font-size:1.35rem;font-weight:800;margin-top:0.2rem'>{risk_company}</div>
            <div style='margin-top:0.65rem;font-size:1.02rem;'>{risk_note}</div>
            <div style='margin-top:0.6rem;font-weight:700'>{risk_value}</div>
            <div class='small-muted'>Risk score = cannibalization / uplift</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# -----------------------------
# Ranking cards
# -----------------------------
st.markdown("<div class='section-title'>Company ranking cards</div>", unsafe_allow_html=True)
rank_cols = st.columns(3)
if not portfolio_scored.empty:
    growth_rank = portfolio_scored.sort_values("Growth_2025_2027_Pct", ascending=False).head(3)
    for idx, (_, rank_row) in enumerate(growth_rank.iterrows()):
        with rank_cols[idx]:
            render_rank_card(
                rank_label=f"Top {idx + 1} by growth",
                company=str(rank_row[scenario_company_col]),
                primary_label="2025–2027 growth",
                primary_value=fmt_pct(rank_row["Growth_2025_2027_Pct"]),
                secondary_label="AI uplift 2027",
                secondary_value=fmt_money(rank_row["AI_Uplift_2027_Calc"]),
            )

st.markdown("---")
tab1, tab2, tab3 = st.tabs(["Portfolio Overview", "Company Deep Dive", "Sources"])

# -----------------------------
# Tab 1: Portfolio Overview
# -----------------------------
with tab1:
    row1_left, row1_right = st.columns([1.3, 1])

    with row1_left:
        st.markdown("<div class='section-title'>Revenue outlook</div>", unsafe_allow_html=True)
        chart_df = scenario_model[[scenario_company_col, "Revenue_2025", "Net_Revenue_2027"]].copy()
        chart_df.columns = ["Company", "Revenue 2025", "Net Revenue 2027"]

        fig_compare = go.Figure()
        fig_compare.add_bar(name="2025", x=chart_df["Company"], y=chart_df["Revenue 2025"])
        fig_compare.add_bar(name="2027", x=chart_df["Company"], y=chart_df["Net Revenue 2027"])
        fig_compare.update_layout(
            barmode="group",
            height=420,
            title=f"2025 vs 2027 revenue ({selected_scenario} scenario)",
            legend_title_text="",
            margin=dict(l=10, r=10, t=50, b=10),
        )
        st.plotly_chart(fig_compare, use_container_width=True)

    with row1_right:
        st.markdown("<div class='section-title'>AI uplift vs downside</div>", unsafe_allow_html=True)
        impact_df = scenario_model[[scenario_company_col, "AI_Uplift_2027", "Cannibalization_2027"]].copy()
        impact_df.columns = ["Company", "AI Uplift 2027", "Cannibalization 2027"]
        impact_melt = impact_df.melt(id_vars="Company", var_name="Metric", value_name="Value")

        fig_impact = px.bar(
            impact_melt,
            x="Company",
            y="Value",
            color="Metric",
            barmode="group",
            height=420,
            title="2027 AI uplift compared with cannibalization",
        )
        fig_impact.update_layout(margin=dict(l=10, r=10, t=50, b=10), legend_title_text="")
        st.plotly_chart(fig_impact, use_container_width=True)

    row2_left, row2_right = st.columns([1.1, 1.2])

    with row2_left:
        st.markdown("<div class='section-title'>Scenario ranking table</div>", unsafe_allow_html=True)
        table_df = portfolio_scored.copy().rename(
            columns={
                scenario_company_col: "Company",
                "Growth_2025_2027_Pct": "Growth 2025-2027",
                "AI_Uplift_2027_Calc": "AI Uplift 2027",
                "Cannibalization_2027_Calc": "Cannibalization 2027",
                "Risk_Score": "Risk Score",
            }
        )
        display_table_df = format_portfolio_table(table_df.sort_values("Growth 2025-2027", ascending=False))
        st.dataframe(display_table_df, use_container_width=True, hide_index=True)

    with row2_right:
        st.markdown("<div class='section-title'>Revenue mix shift</div>", unsafe_allow_html=True)
        mix_chart_df = revenue_mix[[mix_company_col, "AI_Share_2025_%", "AI_Share_2027_%"]].copy()
        mix_chart_df.columns = ["Company", "AI Share 2025", "AI Share 2027"]
        mix_long = mix_chart_df.melt(id_vars="Company", var_name="Period", value_name="AI Share")

        fig_mix = px.bar(
            mix_long,
            x="Company",
            y="AI Share",
            color="Period",
            barmode="group",
            height=420,
            title="AI share of revenue: 2025 vs 2027",
        )
        fig_mix.update_yaxes(tickformat=".0%")
        fig_mix.update_layout(margin=dict(l=10, r=10, t=50, b=10), legend_title_text="")
        st.plotly_chart(fig_mix, use_container_width=True)

# -----------------------------
# Tab 2: Company Deep Dive
# -----------------------------
with tab2:
    st.markdown(f"<div class='section-title'>{selected_company} deep dive</div>", unsafe_allow_html=True)

    top_left, top_mid, top_right = st.columns(3)
    row = company_scenario.iloc[0] if not company_scenario.empty else pd.Series(dtype=object)

    rev2025_value = parse_number(row.get("Revenue_2025"))
    net2026_value = parse_number(row.get("Net_Revenue_2026"))
    net2027_value = parse_number(row.get("Net_Revenue_2027"))

    with top_left:
        st.metric("2025 Revenue", fmt_money(rev2025_value))

    with top_mid:
        delta_2026 = None
        if pd.notna(rev2025_value) and pd.notna(net2026_value) and float(rev2025_value) != 0:
            delta_2026 = f"{((float(net2026_value) / float(rev2025_value)) - 1) * 100:.1f}%"
        st.metric(f"2026 Net Revenue ({selected_scenario})", fmt_money(net2026_value), delta=delta_2026)

    with top_right:
        delta_2027 = None
        if pd.notna(rev2025_value) and pd.notna(net2027_value) and float(rev2025_value) != 0:
            delta_2027 = f"{((float(net2027_value) / float(rev2025_value)) - 1) * 100:.1f}%"
        st.metric(f"2027 Net Revenue ({selected_scenario})", fmt_money(net2027_value), delta=delta_2027)

    d1, d2 = st.columns([1.1, 1])

    segment_name_cols = [c for c in raw_data.columns if "segment" in c.lower() and "name" in c.lower()]
    segment_value_cols = [c for c in raw_data.columns if "segment" in c.lower() and "revenue" in c.lower()]
    ai_product_col = find_col(raw_data, ["Main_AI_Products", "AI_Products", "Main AI Products"])
    monetization_col = find_col(raw_data, ["AI_Monetization_Model", "Monetization", "AI Monetization Model"])
    notes_col = find_col(raw_data, ["Notes"])

    with d1:
        st.markdown("<div class='section-title'>Business snapshot</div>", unsafe_allow_html=True)
        if not raw_company.empty:
            display_cols = [raw_company_col]
            if ai_product_col:
                display_cols.append(ai_product_col)
            if monetization_col:
                display_cols.append(monetization_col)

            st.dataframe(raw_company[display_cols], use_container_width=True, hide_index=True)

            company_row = raw_company.iloc[0]
            segment_rows = []
            for n_col, v_col in zip(segment_name_cols, segment_value_cols):
                seg_name = company_row.get(n_col)
                seg_value = parse_number(company_row.get(v_col))
                if pd.notna(seg_name) and pd.notna(seg_value):
                    segment_rows.append({"Segment": seg_name, "Revenue ($B)": seg_value})

            if segment_rows:
                seg_df = pd.DataFrame(segment_rows)
                fig_seg = px.bar(seg_df, x="Segment", y="Revenue ($B)", title="Latest segment revenue")
                fig_seg.update_layout(height=350, margin=dict(l=10, r=10, t=45, b=10))
                st.plotly_chart(fig_seg, use_container_width=True)

            if notes_col and pd.notna(company_row.get(notes_col)):
                st.markdown(
                    f"""
                    <div class='subtle-card'>
                        <div class='small-muted'>Notes</div>
                        <div style='margin-top:0.45rem;font-size:1rem;line-height:1.55'>
                            {company_row.get(notes_col)}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    with d2:
        st.markdown("<div class='section-title'>AI impact profile</div>", unsafe_allow_html=True)
        impact_rows = [
            {"Metric": "AI Uplift 2026", "Value": parse_number(row.get("AI_Uplift_2026"))},
            {"Metric": "Cannibalization 2026", "Value": parse_number(row.get("Cannibalization_2026"))},
            {"Metric": "AI Uplift 2027", "Value": parse_number(row.get("AI_Uplift_2027"))},
            {"Metric": "Cannibalization 2027", "Value": parse_number(row.get("Cannibalization_2027"))},
        ]
        impact_rows = [item for item in impact_rows if pd.notna(item["Value"])]

        if impact_rows:
            company_impact_df = pd.DataFrame(impact_rows)
            fig_company_impact = px.bar(company_impact_df, x="Metric", y="Value", title="AI upside vs downside")
            fig_company_impact.update_layout(height=350, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig_company_impact, use_container_width=True)

    b1, b2 = st.columns([1, 1])

    with b1:
        st.markdown("<div class='section-title'>Scenario comparison</div>", unsafe_allow_html=True)
        scenario_compare_rows = []
        for scen in SCENARIO_ORDER:
            scen_df = calculate_scenario_model(raw_data, assumptions, scen)
            scen_company = scen_df[
                normalize_company_series(scen_df[scenario_company_col]) == selected_company.strip().lower()
            ]
            if not scen_company.empty:
                scenario_compare_rows.append(
                    {"Scenario": scen, "Net Revenue 2027": scen_company.iloc[0]["Net_Revenue_2027"]}
                )

        if scenario_compare_rows:
            scen_df = pd.DataFrame(scenario_compare_rows)
            scen_df["Scenario"] = pd.Categorical(scen_df["Scenario"], categories=SCENARIO_ORDER, ordered=True)
            scen_df = scen_df.sort_values("Scenario")

            fig_scen = px.bar(
                scen_df,
                x="Scenario",
                y="Net Revenue 2027",
                title=f"{selected_company}: 2027 scenario comparison",
            )
            fig_scen.update_layout(height=340, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig_scen, use_container_width=True)

    with b2:
        st.markdown("<div class='section-title'>Revenue mix visuals</div>", unsafe_allow_html=True)

        if not mix_company.empty:
            mix_row = mix_company.iloc[0]

            ai_2025 = to_share_value(mix_row.get("AI_Share_2025_%"))
            ai_2027 = to_share_value(mix_row.get("AI_Share_2027_%"))
            core_2025 = to_share_value(mix_row.get("Legacy_Share_2025_%"))
            core_2027 = to_share_value(mix_row.get("Legacy_Share_2027_%"))

            if pd.isna(core_2025) and pd.notna(ai_2025):
                core_2025 = 1 - ai_2025
            if pd.isna(core_2027) and pd.notna(ai_2027):
                core_2027 = 1 - ai_2027

            mix_clean = pd.DataFrame(
                {
                    "Period": ["2025", "2025", "2027", "2027"],
                    "Revenue Type": ["Core/Legacy", "AI", "Core/Legacy", "AI"],
                    "Share": [core_2025, ai_2025, core_2027, ai_2027],
                }
            ).dropna()

            if not mix_clean.empty:
                fig_company_mix = px.bar(
                    mix_clean,
                    x="Period",
                    y="Share",
                    color="Revenue Type",
                    barmode="stack",
                    title="2025 to 2027 revenue mix shift",
                )
                fig_company_mix.update_yaxes(tickformat=".0%")
                fig_company_mix.update_layout(
                    height=340,
                    margin=dict(l=10, r=10, t=45, b=10),
                    legend_title_text="",
                )
                st.plotly_chart(fig_company_mix, use_container_width=True)

            interpretation = mix_row.get("Interpretation")
            if pd.notna(interpretation):
                st.markdown(
                    f"""
                    <div class='subtle-card'>
                        <div class='small-muted'>Interpretation</div>
                        <div style='margin-top:0.45rem;font-size:1rem;line-height:1.55'>
                            {interpretation}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

# -----------------------------
# Tab 3: Sources
# -----------------------------
with tab3:
    st.markdown("<div class='section-title'>Sources</div>", unsafe_allow_html=True)
    if source_company_col:
        source_view = st.radio("Source view", ["Selected company", "All companies"], horizontal=True)
        if source_view == "Selected company":
            src_df = sources[
                normalize_company_series(sources[source_company_col]) == selected_company.strip().lower()
            ]
        else:
            src_df = sources.copy()
        st.dataframe(src_df, use_container_width=True, hide_index=True)
    else:
        st.dataframe(sources, use_container_width=True, hide_index=True)

with st.expander("Diagnostics"):
    st.write("Workbook loaded:", workbook_label)
    st.write("Selected scenario:", selected_scenario)
    st.write("Scenario model preview:", scenario_model[[scenario_company_col, "Revenue_2025", "Net_Revenue_2027", "AI_Uplift_2027"]])
    st.write("Revenue mix preview:", revenue_mix[[mix_company_col, "AI_Share_2025_%", "AI_Share_2027_%", "Shift_in_AI_Share_2027_vs_2025_%"]])

st.caption("Run locally with: streamlit run app.py")
