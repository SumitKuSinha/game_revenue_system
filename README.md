# Video Game Commercial Success & Global Revenue Forecasting System

An end-to-end commercial intelligence dashboard and machine learning platform designed to forecast video game unit sales, evaluate market risk tiers, and predict territorial revenue distribution. Built as an applied predictive analytics project for the **IBM SkillsBuild Academic Internship** (Data Analytics with AI track).

---

## 📌 Problem Overview
Video game production cycles demand millions of dollars in upfront capital. Studios face substantial market risk due to misallocated regional marketing budgets and divergence between critic review scores and commercial performance. 

This platform provides executive decision support by:
- Forecasting **Global Unit Sales ($M)** using machine learning.
- Categorizing release prospects into **Commercial Risk Tiers** (High Risk, Moderate Viability, Potential Blockbuster).
- Estimating **Regional Revenue Allocation** across North America, Europe, Japan, and Other markets.
- Formulating actionable publisher go-to-market strategies.

---

## 📊 Dataset Attribution
- **Dataset Source:** [Video Games Sales with Ratings (Kaggle)](https://www.kaggle.com/datasets/rush49/video-game-sales-with-ratings)
- **Features Analyzed:** Platform, Genre, Critic_Score (Metascore), User_Score, NA_Sales, EU_Sales, JP_Sales, Other_Sales, Global_Sales.
- **Hygiene & Preparation:** Median imputation for missing review scores, User_Score normalization, hardware cardinality consolidation, and monotonic score calibration to eliminate historical review-bombing bias.

---

## 🛠️ Architecture & Tech Stack
- **Language:** Python
- **Core Libraries:** Streamlit, Scikit-learn, Pandas, NumPy, Plotly
- **Modeling Strategy:** Feature engineering with `ColumnTransformer` (One-Hot Encoding + Standardization) feeding into a regularized `RandomForestRegressor` and `HuberRegressor` pipeline.

---

## 🚀 Getting Started

### 1. Prerequisites & Environment Setup
Clone the repository and navigate into the root directory:
```bash
git clone [https://github.com/SumitKuSinha/game_revenue_system.git](https://github.com/SumitKuSinha/game_revenue_system.git)
cd game_revenue_system