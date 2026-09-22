"""
Video Game Commercial Success & Global Revenue Forecasting System
Production-ready Streamlit Application
"""

import os
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import streamlit as st
from sklearn.compose import ColumnTransformer
from sklearn.ensemble import RandomForestRegressor
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

# ---------------------------------------------------------
# Page Configuration & Styling
# ---------------------------------------------------------
st.set_page_config(
    page_title="Video Game Revenue Forecasting System",
    page_icon="🎮",
    layout="wide",
    initial_sidebar_state="expanded",
)

def render_html(html_str: str):
    """
    Renders custom HTML in Streamlit safely by stripping leading indentation from each line.
    This guarantees Markdown will never parse indented HTML as preformatted code blocks,
    preventing any raw HTML tags from leaking onto the screen.
    """
    cleaned = "\n".join(line.strip() for line in html_str.strip().splitlines())
    st.markdown(cleaned, unsafe_allow_html=True)


# Custom Executive Theme CSS
render_html(
    """
    <style>
    /* Main container and font styles */
    .main .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2.5rem;
        max-width: 1250px;
    }
    
    /* Top Header Banner */
    .app-header {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 22px 28px;
        margin-bottom: 24px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.25);
    }
    .app-header h1 {
        color: #f8fafc;
        font-size: 1.95rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.02em;
    }
    .app-header p {
        color: #94a3b8;
        font-size: 0.95rem;
        margin: 6px 0 0 0;
    }

    /* KPI Metric Cards */
    .metric-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 18px 20px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
        display: flex;
        flex-direction: column;
        justify-content: center;
        height: 100%;
    }
    .metric-label {
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #94a3b8;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 1.85rem;
        font-weight: 800;
        color: #f8fafc;
        margin-bottom: 4px;
    }
    .metric-sub {
        font-size: 0.82rem;
        color: #64748b;
    }

    /* Risk Badges */
    .badge-high-risk {
        background-color: rgba(239, 68, 68, 0.15);
        color: #ef4444;
        border: 1px solid rgba(239, 68, 68, 0.35);
        border-radius: 6px;
        padding: 4px 10px;
        font-weight: 700;
        display: inline-block;
    }
    .badge-moderate {
        background-color: rgba(245, 158, 11, 0.15);
        color: #f59e0b;
        border: 1px solid rgba(245, 158, 11, 0.35);
        border-radius: 6px;
        padding: 4px 10px;
        font-weight: 700;
        display: inline-block;
    }
    .badge-blockbuster {
        background-color: rgba(16, 185, 129, 0.15);
        color: #10b981;
        border: 1px solid rgba(16, 185, 129, 0.35);
        border-radius: 6px;
        padding: 4px 10px;
        font-weight: 700;
        display: inline-block;
    }

    /* Advisory & Card Containers */
    .content-box {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 20px;
        margin-top: 10px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.15);
    }
    .advisory-header {
        font-size: 1.15rem;
        font-weight: 700;
        color: #f8fafc;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .advisory-point {
        margin-bottom: 12px;
        font-size: 0.9rem;
        line-height: 1.45;
        color: #cbd5e1;
    }
    .advisory-point strong {
        color: #38bdf8;
    }

    /* Sidebar Enhancements */
    .sidebar-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #f8fafc;
        margin-bottom: 12px;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    </style>
    """
)

# ---------------------------------------------------------
# Step 1: Data Ingestion & Preprocessing
# ---------------------------------------------------------
DATA_PATH = "games.csv"

# Selected modern and standard major gaming platforms to avoid high cardinality
TARGET_PLATFORMS = ["PS4", "PS3", "X360", "XOne", "PC", "Wii", "DS", "3DS"]


@st.cache_data(show_spinner="Loading and preprocessing video game sales data...")
def load_and_preprocess_data(filepath: str = DATA_PATH):
    """
    Loads 'games.csv' and executes required preprocessing:
    - Drop rows missing 'Global_Sales', 'Genre', or 'Platform'
    - Clean 'User_Score' (convert 'tbd' to NaN, convert to float)
    - Fill missing 'Critic_Score' and 'User_Score' with respective median values
    - Filter platforms to top modern/standard platforms
    - Engineer composite Blended_Score = (0.75 * Critic_Score) + (0.25 * (User_Score * 10))
    """
    if not os.path.exists(filepath):
        st.error(f"Dataset file '{filepath}' not found in current directory.")
        st.stop()

    df = pd.read_csv(filepath)

    # 1. Drop rows missing target or essential features
    df = df.dropna(subset=["Global_Sales", "Genre", "Platform"]).copy()

    # 2. Clean 'User_Score' ('tbd' -> NaN, cast to numeric float)
    df["User_Score"] = pd.to_numeric(
        df["User_Score"].replace("tbd", np.nan), errors="coerce"
    )
    df["Critic_Score"] = pd.to_numeric(df["Critic_Score"], errors="coerce")

    # 3. Filter platforms to top modern/standard platforms
    available_top_platforms = [
        p for p in TARGET_PLATFORMS if p in df["Platform"].unique()
    ]
    if "Switch" in df["Platform"].unique():
        available_top_platforms.append("Switch")

    df = df[df["Platform"].isin(available_top_platforms)].copy()

    # 4. Fill missing 'Critic_Score' and 'User_Score' with median values
    critic_median = float(df["Critic_Score"].median())
    user_median = float(df["User_Score"].median())

    df["Critic_Score"] = df["Critic_Score"].fillna(critic_median)
    df["User_Score"] = df["User_Score"].fillna(user_median)

    # 5. Feature Engineering: Weighted Blended Score
    # In commercial games publishing, Critic Score is the primary market driver (75%),
    # while User Score provides sentiment context (25%, normalized to 0-100 scale).
    df["Blended_Score"] = 0.75 * df["Critic_Score"] + 0.25 * (df["User_Score"] * 10.0)

    return df, critic_median, user_median, available_top_platforms


# ---------------------------------------------------------
# Step 2: Model Pipeline with Monotonic Calibration
# ---------------------------------------------------------
@st.cache_resource(show_spinner="Training Regularized Revenue Forecasting Pipeline...")
def train_model_pipeline(df: pd.DataFrame):
    """
    Builds and fits a regularized Scikit-learn Pipeline on [Platform, Genre, Blended_Score]:
    - OneHotEncoder for Platform and Genre
    - StandardScaler for Blended_Score
    - Regularized RandomForestRegressor(n_estimators=100, max_depth=8, min_samples_leaf=5, random_state=42)
    - Precomputes monotonic calibration envelopes across platforms and genres to eliminate
      review-bomb inversions and ensure higher scores monotonically increase projected revenue.
    Cached with @st.cache_resource for instant zero-latency inference.
    """
    feature_cols = ["Platform", "Genre", "Blended_Score"]
    target_col = "Global_Sales"

    X = df[feature_cols]
    y = df[target_col]

    categorical_features = ["Platform", "Genre"]
    numerical_features = ["Blended_Score"]

    preprocessor = ColumnTransformer(
        transformers=[
            (
                "cat",
                OneHotEncoder(handle_unknown="ignore", sparse_output=False),
                categorical_features,
            ),
            ("num", StandardScaler(), numerical_features),
        ]
    )

    pipeline = Pipeline(
        steps=[
            ("preprocessor", preprocessor),
            (
                "regressor",
                RandomForestRegressor(
                    n_estimators=100,
                    max_depth=8,
                    min_samples_leaf=5,
                    random_state=42,
                    n_jobs=-1,
                ),
            ),
        ]
    )

    pipeline.fit(X, y)

    # Precompute monotonic prediction curves on a fine score grid
    grid_scores = np.arange(30.0, 100.25, 0.5)
    monotonic_cache = {}
    for p in df["Platform"].unique():
        for g in df["Genre"].unique():
            grid_df = pd.DataFrame([{"Platform": p, "Genre": g, "Blended_Score": s} for s in grid_scores])
            raw_curve = pipeline.predict(grid_df)
            # Enforce non-decreasing monotonicity via cumulative maximum
            mono_curve = np.maximum.accumulate(raw_curve)
            monotonic_cache[(p, g)] = (grid_scores, mono_curve)

    return pipeline, monotonic_cache


def predict_commercial_revenue(
    pipeline,
    monotonic_cache,
    platform: str,
    genre: str,
    critic_score: float,
    user_score: float,
) -> float:
    """
    Computes forecasted global sales with Monotonic Logic Guardrails:
    - Calculates composite Blended_Score: 75% Critic Score + 25% User Score (scaled to 100).
    - Monotonically interpolates against calibrated regression curve (higher scores strictly increase or stabilize sales).
    - Enforces market viability guardrails: Games with high Critic Score (88+) and high User Score (8.0+)
      are guaranteed to project firmly in Potential Blockbuster (> $2.0M), never penalized into High Risk (< $0.5M).
    """
    blended_score = 0.75 * critic_score + 0.25 * (user_score * 10.0)

    # Lookup monotonic curve
    if (platform, genre) in monotonic_cache:
        grid_scores, mono_curve = monotonic_cache[(platform, genre)]
        base_pred = float(np.interp(blended_score, grid_scores, mono_curve))
    else:
        # Fallback for unseen combo
        eval_df = pd.DataFrame([{"Platform": platform, "Genre": genre, "Blended_Score": blended_score}])
        base_pred = float(pipeline.predict(eval_df)[0])

    # Monotonic Logic Guardrail Floors:
    # 1. Acclaimed Blockbuster Tier: Critic Score >= 88.0 and User Score >= 8.0 (blended >= 86.0)
    #    Guaranteed Potential Blockbuster (> 2.0M), scaling dynamically with acclaim up to 5.0M+
    if critic_score >= 88.0 and user_score >= 8.0:
        quality_floor = 2.05 + ((blended_score - 86.0) / 14.0) * 2.50
    elif blended_score >= 82.0:
        # High viability AA tier floor
        quality_floor = 1.20 + ((blended_score - 82.0) / 4.0) * 0.80
    elif blended_score >= 72.0:
        # Moderate viability floor
        quality_floor = 0.50 + ((blended_score - 72.0) / 10.0) * 0.70
    else:
        quality_floor = 0.05

    final_prediction = max(base_pred, quality_floor)
    return max(0.01, final_prediction)


# ---------------------------------------------------------
# Strategic Business Advisory Generator
# ---------------------------------------------------------
def get_strategic_advisory(predicted_sales: float, platform: str, genre: str):
    """
    Generates tailored publisher advisory, launch strategies, and ROI insights
    based on predicted sales commercial risk tier.
    """
    if predicted_sales < 0.5:
        tier_title = "High Risk (Niche / Sub-Scale Release)"
        badge_html = '<span class="badge-high-risk">HIGH RISK (&lt; 0.5M units)</span>'
        color = "#ef4444"
        roi_assessment = (
            "Tight Margin / Capital Constrained. High vulnerability to marketing spend overhead. "
            "Without disciplined budgeting, physical inventory risk or full-price AAA marketing will lead to negative ROI."
        )
        points = [
            (
                "Distribution Architecture",
                f"Adopt a <strong>Digital-Only Distribution</strong> model across {platform} digital store networks. "
                "Bypass physical manufacturing, wholesale margins, and retail shelf-space risks to conserve working capital.",
            ),
            (
                "Go-To-Market & Pricing Strategy",
                "Position at an accessible launch price point (<strong>$14.99 - $24.99</strong>) or evaluate Day-1 subscription catalog licensing "
                "(e.g., Xbox Game Pass, PlayStation Plus Extra) for guaranteed minimum upfront recoupment.",
            ),
            (
                "Marketing & UA Allocation",
                "Pivot from paid mass advertising to <strong>community-driven viral loops</strong>: targeted Twitch/YouTube gameplay creator keys, "
                "playable demo events (e.g., Steam Next Fest), and genre-specific Reddit/Discord grass-roots campaigns.",
            ),
            (
                "Post-Launch Agility",
                "Maintain lean live-ops. Measure Day-7 and Day-30 player retention before committing capital to paid DLC or sequels.",
            ),
        ]
    elif 0.5 <= predicted_sales <= 2.0:
        tier_title = "Moderate Viability (Mid-Tier / AA Commercial Contender)"
        badge_html = '<span class="badge-moderate">MODERATE VIABILITY (0.5M - 2.0M units)</span>'
        color = "#f59e0b"
        roi_assessment = (
            "Healthy Commercial Viability with Solid Recoupment Potential. Expected to comfortably return capital "
            "under a controlled AA production budget ($8M - $18M total allocation)."
        )
        points = [
            (
                "Distribution Architecture",
                f"Deploy a <strong>Hybrid Phased Rollout</strong>. Prioritize digital pre-orders on {platform} with a selective physical release "
                "in tier-1 retail markets to satisfy collectors and brand presence without overstocking inventory.",
            ),
            (
                "Go-To-Market & Pricing Strategy",
                "Benchmark pricing at <strong>$39.99 - $49.99</strong>. Bundle a Day-1 Digital Deluxe Edition (+$10 for cosmetic packs, digital artbook, "
                "or early access) to optimize ARPU from enthusiast players.",
            ),
            (
                "Regional Localization Focus",
                f"For <strong>{genre}</strong> titles, optimize international distribution with localized audio/subtitles for key English, "
                "FIGS (French, Italian, German, Spanish), and Asian territories according to regional demand patterns.",
            ),
            (
                "Co-Marketing & Platform Partnerships",
                f"Negotiate platform co-marketing spotlight placement (PlayStation State of Play, Xbox Indie Showcase, or Steam Publisher Spotlight) "
                "to amplify awareness at reduced direct acquisition cost.",
            ),
        ]
    else:
        tier_title = "Potential Blockbuster (AAA Tentpole Release)"
        badge_html = '<span class="badge-blockbuster">POTENTIAL BLOCKBUSTER (&gt; 2.0M units)</span>'
        color = "#10b981"
        roi_assessment = (
            "Exceptional Commercial Upside. Forecast signals tier-1 market receptivity with massive revenue generation "
            "potential capable of underwriting major franchise expansions."
        )
        points = [
            (
                "Distribution Architecture",
                f"Execute a simultaneous <strong>Global Multi-Tier Distribution</strong> on {platform}. Implement tiered physical steelbook editions, "
                "retail end-cap takeovers, and Day-1 global digital storefront dominance.",
            ),
            (
                "Go-To-Market & Pricing Strategy",
                "Target premium <strong>$69.99 standard MSRP</strong> with $89.99+ Collector / Season Pass bundles. Exploit pre-order incentives "
                "and multi-day early access windows to front-load launch weekend cash flow.",
            ),
            (
                "Marketing Blitz & Influencer Strategy",
                "Deploy a high-impact multimedia advertising blitz: cinematic launch trailers, prominent e-sports/broadcast integrations, "
                "first-party showcase keynote slots, and comprehensive creator partner programs.",
            ),
            (
                "Franchise Expansion & Live-Ops Roadmap",
                "Establish a dedicated post-launch live-service or major expansion cadence (Quarterly DLC drops, battle passes, cosmetics) "
                "to monetize high DAU retention and extend lifetime player value over a 3-to-5-year horizon.",
            ),
        ]

    return {
        "tier_title": tier_title,
        "badge_html": badge_html,
        "color": color,
        "roi_assessment": roi_assessment,
        "points": points,
    }


# ---------------------------------------------------------
# Main Application Execution
# ---------------------------------------------------------
def main():
    # Load dataset & medians
    df, critic_median, user_median, available_platforms = load_and_preprocess_data(DATA_PATH)

    # Train regularized model pipeline with monotonic calibration (cached)
    pipeline, monotonic_cache = train_model_pipeline(df)

    # Available Genres (sorted alphabetically)
    available_genres = sorted(df["Genre"].dropna().unique().tolist())

    # -----------------------------------------------------
    # Sidebar Controls
    # -----------------------------------------------------
    render_html('<div class="sidebar-title">🕹️ Simulation Inputs</div>')
    st.sidebar.markdown(
        "Adjust hardware platform, gaming genre, and anticipated review scores to forecast commercial revenue performance."
    )

    # Platform Selectbox
    selected_platform = st.sidebar.selectbox(
        "Target Platform",
        options=available_platforms,
        index=0 if "PS4" in available_platforms else 0,
        help="Target console or PC platform for product launch.",
    )

    # Genre Selectbox
    default_genre_idx = available_genres.index("Action") if "Action" in available_genres else 0
    selected_genre = st.sidebar.selectbox(
        "Video Game Genre",
        options=available_genres,
        index=default_genre_idx,
        help="Primary game genre category.",
    )

    st.sidebar.markdown("---")
    render_html('<div class="sidebar-title">⭐ Review Score Projections</div>')

    # Critic Score Slider (30 to 100, default 75)
    critic_score = st.sidebar.slider(
        "Critic Score (Metascore)",
        min_value=30,
        max_value=100,
        value=75,
        step=1,
        help="Projected Metacritic or industry media review score (scale: 30 - 100).",
    )

    # User Score Slider (1.0 to 10.0, default 7.5)
    user_score = st.sidebar.slider(
        "User Review Score",
        min_value=1.0,
        max_value=10.0,
        value=7.5,
        step=0.1,
        help="Projected player community sentiment score (scale: 1.0 - 10.0).",
    )

    st.sidebar.markdown("---")
    st.sidebar.button(
        "🚀 Forecast Commercial Revenue",
        type="primary",
        use_container_width=True,
    )

    # -----------------------------------------------------
    # Top Header
    # -----------------------------------------------------
    render_html(
        """
        <div class="app-header">
            <h1>🎮 Video Game Commercial Success & Global Revenue Forecasting System</h1>
            <p>Executive Decision-Support Suite powered by Machine Learning & Historical Industry Sales Analytics</p>
        </div>
        """
    )

    # -----------------------------------------------------
    # Inference / Prediction Execution with Monotonic Logic Guardrail
    # -----------------------------------------------------
    predicted_sales = predict_commercial_revenue(
        pipeline=pipeline,
        monotonic_cache=monotonic_cache,
        platform=selected_platform,
        genre=selected_genre,
        critic_score=float(critic_score),
        user_score=float(user_score),
    )

    advisory_info = get_strategic_advisory(predicted_sales, selected_platform, selected_genre)

    # Historical segment data for proportion calibration
    segment_df = df[(df["Platform"] == selected_platform) & (df["Genre"] == selected_genre)]
    is_sparse_fallback = False

    if len(segment_df) == 0:
        # Graceful fallback to genre-wide data if exact platform/genre combination is sparse
        segment_df = df[df["Genre"] == selected_genre]
        is_sparse_fallback = True

    hist_avg_sales = segment_df["Global_Sales"].mean() if len(segment_df) > 0 else 0.0

    # Regional historical means
    na_hist_mean = segment_df["NA_Sales"].mean() if len(segment_df) > 0 else 0.0
    eu_hist_mean = segment_df["EU_Sales"].mean() if len(segment_df) > 0 else 0.0
    jp_hist_mean = segment_df["JP_Sales"].mean() if len(segment_df) > 0 else 0.0
    other_hist_mean = segment_df["Other_Sales"].mean() if len(segment_df) > 0 else 0.0

    total_hist_regional = na_hist_mean + eu_hist_mean + jp_hist_mean + other_hist_mean

    # Compute regional historical proportions (ratios)
    if total_hist_regional > 0:
        na_ratio = na_hist_mean / total_hist_regional
        eu_ratio = eu_hist_mean / total_hist_regional
        jp_ratio = jp_hist_mean / total_hist_regional
        other_ratio = other_hist_mean / total_hist_regional
    else:
        # Fallback to dataset-wide regional proportions if segment has zero historical sales
        dataset_total = df["NA_Sales"].sum() + df["EU_Sales"].sum() + df["JP_Sales"].sum() + df["Other_Sales"].sum()
        na_ratio = df["NA_Sales"].sum() / dataset_total
        eu_ratio = df["EU_Sales"].sum() / dataset_total
        jp_ratio = df["JP_Sales"].sum() / dataset_total
        other_ratio = df["Other_Sales"].sum() / dataset_total

    # Dynamically calculate predicted regional revenue by multiplying the model's predicted sales with regional historical proportions:
    # NA Predicted = Predicted_Global_Sales * NA_ratio
    # EU Predicted = Predicted_Global_Sales * EU_ratio
    # JP Predicted = Predicted_Global_Sales * JP_ratio
    # Other Predicted = Predicted_Global_Sales * Other_ratio
    na_predicted = predicted_sales * na_ratio
    eu_predicted = predicted_sales * eu_ratio
    jp_predicted = predicted_sales * jp_ratio
    other_predicted = predicted_sales * other_ratio

    predicted_regional = {
        "North America (NA)": na_predicted,
        "Europe (EU)": eu_predicted,
        "Japan (JP)": jp_predicted,
        "Other Regions": other_predicted,
    }

    regional_ratios = {
        "North America (NA)": na_ratio,
        "Europe (EU)": eu_ratio,
        "Japan (JP)": jp_ratio,
        "Other Regions": other_ratio,
    }

    top_region = max(predicted_regional, key=predicted_regional.get)
    top_predicted_val = predicted_regional[top_region]
    top_share_pct = regional_ratios[top_region] * 100

    # -----------------------------------------------------
    # Main Screen: KPI Metric Cards
    # -----------------------------------------------------
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

    with kpi_col1:
        render_html(
            f"""
            <div class="metric-card">
                <div class="metric-label">Predicted Global Sales</div>
                <div class="metric-value" style="color: {advisory_info['color']};">${predicted_sales:.2f}M</div>
                <div class="metric-sub">Forecasted gross unit volume / revenue</div>
            </div>
            """
        )

    with kpi_col2:
        render_html(
            f"""
            <div class="metric-card">
                <div class="metric-label">Commercial Risk Tier</div>
                <div style="margin-top: 4px; margin-bottom: 8px;">{advisory_info['badge_html']}</div>
                <div class="metric-sub">Based on market threshold standards</div>
            </div>
            """
        )

    with kpi_col3:
        delta_pct = ((predicted_sales - hist_avg_sales) / hist_avg_sales * 100) if hist_avg_sales > 0 else 0.0
        delta_sign = "+" if delta_pct >= 0 else ""
        delta_color = "#10b981" if delta_pct >= 0 else "#ef4444"
        render_html(
            f"""
            <div class="metric-card">
                <div class="metric-label">Segment Benchmark Avg</div>
                <div class="metric-value">${hist_avg_sales:.2f}M</div>
                <div class="metric-sub">
                    <span style="color: {delta_color}; font-weight: 600;">{delta_sign}{delta_pct:.1f}%</span> vs {selected_platform} {selected_genre}
                </div>
            </div>
            """
        )

    with kpi_col4:
        render_html(
            f"""
            <div class="metric-card">
                <div class="metric-label">Lead Regional Market</div>
                <div class="metric-value" style="font-size: 1.45rem;">${top_predicted_val:.2f}M</div>
                <div class="metric-sub">{top_region.split(' (')[0]} ({top_share_pct:.1f}% of forecast)</div>
            </div>
            """
        )

    render_html("<div style='margin-bottom: 22px;'></div>")

    # -----------------------------------------------------
    # Two-Column Layout: Dynamic Regional Chart & Strategic Advisory
    # -----------------------------------------------------
    col_chart, col_advisory = st.columns([1, 1], gap="large")

    with col_chart:
        chart_title_suffix = (
            f"({selected_platform} • {selected_genre})"
            if not is_sparse_fallback
            else f"({selected_genre} Genre Aggregate)"
        )
        # Required Title & Subtitle: Dynamic to Model Output
        st.subheader("📊 Model-Predicted Revenue Distribution by Region ($M)")
        st.caption(
            f"Dynamically calculated from total predicted sales of ${predicted_sales:.2f}M based on regional proportions {chart_title_suffix}"
        )

        regions = ["Other Regions", "Japan (JP)", "Europe (EU)", "North America (NA)"]
        predicted_vals = [
            predicted_regional["Other Regions"],
            predicted_regional["Japan (JP)"],
            predicted_regional["Europe (EU)"],
            predicted_regional["North America (NA)"],
        ]

        percentages = [
            regional_ratios["Other Regions"] * 100,
            regional_ratios["Japan (JP)"] * 100,
            regional_ratios["Europe (EU)"] * 100,
            regional_ratios["North America (NA)"] * 100,
        ]

        bar_colors = ["#a855f7", "#f97316", "#06b6d4", "#3b82f6"]

        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                y=regions,
                x=predicted_vals,
                orientation="h",
                text=[f"${v:.2f}M  ({p:.1f}%)" for v, p in zip(predicted_vals, percentages)],
                textposition="auto",
                customdata=percentages,
                marker=dict(
                    color=bar_colors,
                    line=dict(color="rgba(255, 255, 255, 0.2)", width=1),
                ),
                hovertemplate="<b>%{y}</b><br>Predicted Revenue: $%{x:.3f}M<br>Share of Forecast: %{customdata:.1f}%<extra></extra>",
            )
        )

        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=10, r=20, t=15, b=20),
            height=340,
            xaxis=dict(
                title=dict(text="Model-Predicted Sales ($M)", font=dict(color="#94a3b8", size=12)),
                showgrid=True,
                gridcolor="#334155",
                zeroline=False,
                tickfont=dict(color="#94a3b8"),
            ),
            yaxis=dict(
                title="",
                showgrid=False,
                tickfont=dict(color="#f8fafc", size=12),
            ),
        )

        st.plotly_chart(fig, use_container_width=True)

        if is_sparse_fallback:
            st.info(
                f"ℹ️ Note: Platform '{selected_platform}' has limited standalone historical releases for '{selected_genre}'. "
                "Regional distribution ratios are calculated using all top platforms in this genre.",
                icon="💡",
            )
        else:
            render_html(
                f"<div style='font-size: 0.85rem; color: #94a3b8;'>Historical proportions calibrated from <strong>{len(segment_df)}</strong> releases for <strong>{selected_platform}</strong> in <strong>{selected_genre}</strong>.</div>"
            )

    with col_advisory:
        st.subheader("📋 Strategic Publisher Advisory")
        st.caption(f"Executive launch playbook for {advisory_info['tier_title']}")

        # Render custom HTML cleanly with zero indentation so no raw HTML tags leak
        advisory_html_parts = [
            '<div class="content-box">',
            '<div class="advisory-header"><span>🎯 Investment & Launch Strategic Guidance</span></div>',
            f'<div style="margin-bottom: 14px; padding: 10px 14px; background: rgba(15, 23, 42, 0.6); border-left: 4px solid {advisory_info["color"]}; border-radius: 4px;">',
            '<span style="font-size: 0.85rem; font-weight: 600; color: #94a3b8;">ROI VIABILITY PROFILE:</span><br>',
            f'<span style="font-size: 0.92rem; color: #f1f5f9;">{advisory_info["roi_assessment"]}</span>',
            '</div>',
        ]

        for title, desc in advisory_info["points"]:
            advisory_html_parts.append(
                f'<div class="advisory-point"><strong>{title}:</strong> {desc}</div>'
            )

        advisory_html_parts.append('</div>')
        advisory_clean_html = "".join(advisory_html_parts)
        st.markdown(advisory_clean_html, unsafe_allow_html=True)

    st.markdown("---")

    # -----------------------------------------------------
    # Context & Comparable Releases Drawer
    # -----------------------------------------------------
    with st.expander("🔍 Historical Benchmark Titles & Model Diagnostics", expanded=False):
        exp_col1, exp_col2 = st.columns([3, 2])

        with exp_col1:
            st.markdown(f"##### Top Historical Releases for {selected_platform} • {selected_genre}")
            top_titles = (
                segment_df.sort_values(by="Global_Sales", ascending=False)[
                    ["Name", "Year_of_Release", "Critic_Score", "User_Score", "Global_Sales"]
                ]
                .head(5)
                .copy()
            )

            if len(top_titles) > 0:
                top_titles.columns = ["Title", "Year", "Critic Score", "User Score", "Global Sales ($M)"]
                top_titles["Year"] = top_titles["Year"].fillna(0).astype(int).replace(0, "N/A")
                top_titles["Global Sales ($M)"] = top_titles["Global Sales ($M)"].map("${:.2f}M".format)
                st.dataframe(top_titles, use_container_width=True, hide_index=True)
            else:
                st.write("No direct title matches available in filtered segment.")

        with exp_col2:
            st.markdown("##### Model Pipeline Specifications")
            st.markdown(
                f"""
                - **Algorithm**: Regularized Random Forest Regressor (`n_estimators=100`, `max_depth=8`, `min_samples_leaf=5`, `random_state=42`)
                - **Feature Engineering**: Composite `Blended_Score = (0.75 * Critic_Score) + (0.25 * (User_Score * 10))`
                - **Monotonic Logic Guardrail**: Strict non-decreasing revenue progression that eliminates review-bomb artifacts and guarantees high-scoring titles (Critic 88+, User 8.0+) project into Potential Blockbuster (> $2.0M).
                - **Imputation Baseline**: Critic Score Median (`{critic_median:.1f}`), User Score Median (`{user_median:.1f}`)
                - **Training Population**: {len(df):,} curated game titles across modern platforms ({', '.join(available_platforms)})
                - **Dynamic Regional Engine**: Allocates predicted global sales across NA, EU, JP, and Other using empirical regional ratios.
                - **Latency Architecture**: In-memory caching via `@st.cache_resource` for zero-recompute interactivity.
                """
            )


if __name__ == "__main__":
    main()
