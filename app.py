import base64
import hashlib
import hmac
import io
import json
import os
import re
import smtplib
import sqlite3
import time
import urllib.parse
from datetime import datetime
from email.message import EmailMessage
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from xlsxwriter.utility import xl_col_to_name

# ============================================================
# HMG EXCEL V5 OPERATIONS PLATFORM
# Enhanced from the original Excel Dashboard Generator without
# removing existing capabilities: authentication, multi-file upload,
# audit log, filters, heatmaps, pivot/crosstab, themed dashboard,
# and downloadable Excel workbook.
# ============================================================

APP_NAME = "HMG Enterprise Excel Operations Platform"
APP_VERSION = "Enterprise Edition - HMG Technologies Ecosystem"
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLS = 16_384


HMG_BRAND_PROFILE = {
    "Founder / Visioner": "Adewale Samson Adeagbo",
    "Professional Identity": "AI-Augmented Solutions Developer · Data Scientist · STEM Educator",
    "Builder Modes": "EdTech Builder · DataTech Builder · FaithTech Builder",
    "Tagline": "Real problems. Real solutions. Built with AI, grounded in data, taught from the classroom.",
    "Location": "Lagos, Nigeria",
    "Teaching Experience": "15+ years across Nursery, Primary, Junior Secondary and Senior Secondary classrooms",
    "Education": "B.Sc.(Ed) Computer Science Education; 3MTT Data Science track",
    "Track Record": "34 live projects deployed · 7 ML models across 7 industries · 11 data-tool simulators · 500+ students reached",
    "Brand": "HMG Concepts — His Marvellous Grace",
    "Founded": "2015",
    "Parent Brand": "HMG Concepts (Education · Technology · Media · Gospel)",
    "Direct Owner": "HMG Academy (subsidiary of HMG Concepts)",
    "Built By": "HMG Technologies (innovation arm of HMG Concepts)",
    "Ecosystem": "HMG Academy · HMG Technologies · HMG Media · HMG Gospel",
    "Platform Home": "HMG Academy — powered by HMG Technologies",
    "Philosophy": "Learning Deliberately. Teaching Authentically.",
    "HMG Concepts Description": "Nigerian education and technology brand (est. 2015): one brand, four living missions — Academy, Technologies, Media, Gospel. 'Where Education Meets Innovation & Conviction.'",
    "HMG Academy Description": "Strictly virtual learning institution: vetted tutors, virtual home-schooling, exam prep (WAEC, NECO, GCE, BECE, UTME, IGCSE, IELTS, SAT, JUPEB), LMS pathways, free CBT Pro and parent monitoring.",
    "HMG Technologies Description": "Innovation arm: AI-augmented tools, CBT systems, data dashboards, ML models, Excel automation and 11 data-tool simulators built cost-effectively for Nigerian businesses, schools, NGOs and churches.",
    "HMG Media Description": "Content and visibility arm: purpose-led educational video, brand storytelling, social content systems and graphics for the HMG ecosystem and values-driven brands.",
    "HMG Gospel Description": "Faith arm: Christ-centred digital outreach — dramavangelism, techvangelism, podcasts, ebooks and church/community support. 'Every platform a pulpit, every tool a testimony.'",
    "Portfolio Website": "https://cssadewale.pages.dev",
    "HMG Concepts Website": "https://hmgconcepts.pages.dev",
    "HMG Academy Website": "https://hmgacademy.pages.dev",
    "HMG Technologies Website": "https://hmgtechnologies.pages.dev",
    "HMG Media Website": "https://hmgmedia.pages.dev",
    "HMG Gospel Website": "https://hmggospel.pages.dev",
    "GitHub": "https://github.com/cssadewale",
    "LinkedIn": "https://linkedin.com/in/adewalesamsonadeagbo",
    "YouTube": "https://youtube.com/@hmgconcepts",
    "X (Twitter)": "https://x.com/cssadewale",
    "Instagram": "https://instagram.com/cssadewale",
    "WhatsApp": "+234 810 086 6322",
    "Phone": "+234 907 790 7677",
    "Brand Email": "hismarvellousgrace@gmail.com",
    "Tech / Partnerships Email": "buildingmyictcareer@gmail.com",
    "Personal / HMG Email": "hmgconcepts@gmail.com",
}

HMG_ECOSYSTEM_TEXT = """
This Excel Operations Platform is an HMG Academy product, built by HMG Technologies — the innovation arm of HMG Concepts (His Marvellous Grace, est. 2015, Lagos, Nigeria).

HMG Concepts is one brand with four living missions:
• HMG Academy — strictly virtual learning institution (the home of this platform).
• HMG Technologies — AI-augmented tools, dashboards, CBT and data products (the builder of this platform).
• HMG Media — purpose-led content and brand visibility.
• HMG Gospel — Christ-centred digital outreach.

The platform is positioned as a free-tool, no-AI-API analytics product for schools, SMEs, churches, NGOs and institutions that need Excel-style analysis without manual spreadsheet labour. It follows the HMG cost-discipline principle: deliberately built without paid AI API dependence or unnecessary recurring cost.
"""

THEMES = {
    "HMG Enterprise (Default)": {
        "header_bg": "#001524",
        "kpi_val": "#003366",
        "primary_chart": "#1A5F7A",
        "secondary_chart": "#57C5B6",
        "accent": "#D4AF37",
        "soft_bg": "#F4F6F7",
    },
    "Corporate Executive": {
        "header_bg": "#1B263B",
        "kpi_val": "#415A77",
        "primary_chart": "#778DA9",
        "secondary_chart": "#E0E1DD",
        "accent": "#9B2226",
        "soft_bg": "#F8F9FA",
    },
    "EdTech Analytics": {
        "header_bg": "#2E0249",
        "kpi_val": "#570A57",
        "primary_chart": "#A91079",
        "secondary_chart": "#F806CC",
        "accent": "#F4A261",
        "soft_bg": "#FFF7ED",
    },
    "FaithTech Stewardship": {
        "header_bg": "#3E2723",
        "kpi_val": "#5D4037",
        "primary_chart": "#8D6E63",
        "secondary_chart": "#D7CCC8",
        "accent": "#FFC107",
        "soft_bg": "#FFFBEA",
    },
    "Public Sector Clean": {
        "header_bg": "#0B3954",
        "kpi_val": "#087E8B",
        "primary_chart": "#087E8B",
        "secondary_chart": "#BFD7EA",
        "accent": "#FF5A5F",
        "soft_bg": "#EFF6FF",
    },
}

FEATURE_CATALOG = [
    ("Data ingestion", "Upload one or many CSV/XLSX files, append them into one master table, and retain source-file traceability."),
    ("Data profiling", "Detect row/column count, data types, missing values, duplicate rows, unique values, numeric statistics, categorical distribution, and date ranges."),
    ("Data cleaning", "Drop blank rows/columns, trim text, standardize headers, remove duplicates, fill missing values, convert data types, and treat outliers."),
    ("Find, replace, filter, sort", "Run common Excel-style actions from buttons instead of manually editing cells."),
    ("Column engineering", "Create calculated columns, ratios, percentages, running totals, ranks, z-scores, text transformations, date parts, split columns, and merged columns."),
    ("Grouping and pivoting", "Build group-by summaries and crosstab pivot matrices using selected rows, columns, values, and aggregation functions."),
    ("Statistical analysis", "Generate correlations, descriptive statistics, outlier reports, regression trend lines, and simple non-API forecasts."),
    ("Visualization", "Create bar, line, area, scatter, histogram, box, pie, treemap, and heatmap charts interactively."),
    ("Data validation", "Run not-blank, unique, numeric-range, allowed-list, valid-date, and positive-number checks to produce exception reports."),
    ("What-if analysis", "Simulate percentage, constant, and factor-based changes across all rows or a selected segment without using paid tools."),
    ("Pareto 80/20 analysis", "Rank categories by contribution to identify the vital few and long tail drivers of results."),
    ("Lookup and merge", "Upload lookup tables and merge them into the working dataset like XLOOKUP, VLOOKUP, or Power Query merges."),
    ("Sampling", "Create fixed-size or percentage-based samples for demos, review, testing, and lightweight export."),
    ("Reconciliation", "Compare two tables to find missing records and value mismatches for finance, inventory, HR, or school result checks."),
    ("Time intelligence", "Generate monthly, quarterly, weekly, or daily period analysis with previous-period changes, YTD, and rolling averages."),
    ("KPI target tracking", "Compare actual values with constant targets or target columns, including variance and attainment percentage."),
    ("Binning and IF classification", "Create numeric bands and rule-based labels without manually writing nested IF/IFS formulas."),
    ("Quality scorecard", "Produce a quick completeness, uniqueness, outlier, and schema scorecard for executive QA review."),
    ("Excel workbook export", "Download a complete analyst workbook with audit logs, raw data, processed data, profile sheets, reports, pivots, validation, what-if, Pareto, reconciliation, time intelligence, KPI target tracking, charts, and dashboard."),
    ("Weighted scoring", "Build weighted decision matrices and rank rows using normalized criteria."),
    ("Goal seek", "Calculate target gaps, required lift percentage, and uniform per-row improvement."),
    ("Calendar table", "Generate a date dimension table for Excel, Power BI, pivots, and time intelligence."),
    ("HMG brand embedding", "Founder, contact, ecosystem, and HMG Technologies identity are embedded in the UI and workbook."),
    ("Enterprise readiness", "Generate readiness, access-control, governance, deployment, free-tools, privacy and release-note sheets in every workbook."),
    ("Secrets-ready access", "Retains demo passwords while supporting Streamlit Secrets or APP_PASSWORD environment variable for production."),
    ("Governance", "Every operation is recorded in an audit trail for reproducibility and institutional accountability."),
    ("Google Sheets import", "Pull data directly from a shareable Google Sheets link with no Google API key and no paid service."),
    ("Auto insight cards", "Generate punchy, rule-based insight cards (completeness, duplicates, top contributor, outliers, latest trend) on screen — a free alternative to AI 'auto-insights'."),
    ("Offline HTML report", "Download a self-contained, shareable HTML data report (KPIs, insights, profile, preview) that opens in any browser with no internet, no AI and no external assets."),
    ("Smart chart recommendations", "Get deterministic chart suggestions based on column types and cardinality, mirroring the 'AI auto-chart' feature of commercial tools without any model cost."),
    ("Dataset comparison", "Compare the current dataset with another file (e.g., period over period) at column level: presence, type, missing % and means, similar to Sweetviz compare."),
    ("Subscription tiers", "Free / Pro / Enterprise tiers gate premium tools without removing anything: all original analyst workflow stays free; Pro adds SQL, dataset comparison and PII masking; Enterprise adds governance and license administration."),
    ("Tamper-proof licensing", "Issue HMAC-signed, offline subscription license keys (customer, tier, expiry). Keys cannot be forged, re-tiered or extended without the owner's secret signing key — this prevents subscription bypass."),
    ("Hardened authentication", "Constant-time password check, optional SHA-256 password hashing in secrets, brute-force lockout after repeated failures, and demo logins restricted to the Free tier only."),
    ("PII detection & masking", "Rule-based detection of likely personal data (emails, phones, IDs, addresses) and one-click partial or irreversible-hash masking before export — no AI used."),
    ("SQL injection hardening", "The SQL tool allows only single read-only SELECT/WITH queries and blocks multi-statement injection and all write/DDL keywords."),
    ("Search engine optimization", "Crawlable meta tags, Open Graph/Twitter cards and JSON-LD structured data are injected into the app, plus a static SEO landing page, robots.txt and sitemap.xml for discoverability."),
    ("Selectable licence models", "Issue each client either a Team (shareable across many devices) or Per-Seat (bound to one Seat ID such as a staff email or device label) licence. The model is baked into the signed key and cannot be switched without the owner's secret."),
    ("Anti-fork protection", "Ships under a proprietary HMG licence that forbids forking, copying and redistribution, with a hardening guide for locking down the GitHub repository and Streamlit deployment so platform users cannot fork or clone the source."),
    ("Device-bound per-seat licences", "Per-seat keys can auto-bind to a customer's specific device using a browser fingerprint computed locally (no external service, no AI) — the user types nothing and the licence will not activate on any other device."),
    ("Data Quality Rules engine", "Define many data-quality rules (not-null, unique, positive, min, max, allowed-list, regex) and run them at once for a pass/fail scorecard with failing-row counts."),
    ("Trend & seasonality decomposition", "Break a time series into trend, seasonality and residual using a lightweight moving-average method — no heavy statistics library and no AI."),
    ("Correlation insights", "Automatically surface the strongest numeric relationships with plain-English notes and a heatmap; rule-based, no AI."),
    ("Cleaning advisor", "Scan the dataset for missing values, whitespace, inconsistent casing, numbers-stored-as-text and constant columns, returning prioritised, concrete cleaning suggestions."),
    ("Survey / Likert analysis", "Summarise rating-scale questions: mean, median, top-2-box and bottom-2-box percentages with an automatic sentiment label."),
    ("Benford's Law fraud screen", "Audit-grade first-digit test comparing your numbers to Benford's expected distribution to flag possible manipulation — a classic, free fraud indicator."),
    ("Usage & license analytics", "Owner-only dashboard of recorded events (logins, license activations/upgrades, key issuance) from a local SQLite log: totals, per-day, per-tier, per-customer and per-model breakdowns with CSV export. No external service, no AI."),
    ("Cross-tab & chi-square", "Build a contingency table between two categorical columns and test whether they are related using a pure-numpy chi-square plus Cramér's V effect size — no scipy, no AI."),
    ("Text / keyword frequency", "Count the most common words in a free-text column (survey open-ends, comments) with stopword filtering and a bar chart — no NLP library, no AI."),
    ("Number / locale normalizer", "Convert locale-formatted number text (European 1.234,56 or US 1,234.56) into clean numbers with deterministic parsing."),
    ("Data classification & watermark", "Label the dataset Public/Internal/Confidential/Restricted, record it in the audit trail, and add watermark columns for governed exports."),
    ("White-label / multi-tenant branding", "Re-skin the entire platform per client — brand name, tagline, logo, colours and footer — for this session or permanently per deployment via the WHITE_LABEL_BRAND secret. Same code, many tenants."),
    ("Report delivery (email & WhatsApp)", "Email a report attachment via a free SMTP account (smtplib, no paid API) or generate WhatsApp/mailto share links — automatic distribution of reports to clients."),
    ("Executive scorecard", "A compact one-look KPI scorecard (totals, average, top contributor, distinct count, latest growth) for quick management review."),
    ("Concentration & Gini index", "Measure how concentrated a metric is across categories (e.g. revenue per customer) with the Gini coefficient and a cumulative-share table to flag dependency risk."),
    ("Snapshot diff / change log", "Compare the current dataset to an earlier snapshot keyed on an ID column and list rows Added / Removed / Changed for master-data change tracking."),
]

# ============================================================
# Utility helpers
# ============================================================

def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# Pandas >= 2.2 deprecated the "M"/"Q"/"A" offset aliases for resample/date_range
# in favour of "ME"/"QE"/"YE". This helper keeps the simple UI codes (D/W/M/Q)
# working while avoiding FutureWarnings and future breakage.
_RESAMPLE_FREQ_MAP = {"D": "D", "W": "W", "M": "ME", "Q": "QE", "A": "YE", "Y": "YE"}


def normalize_resample_freq(freq: str) -> str:
    """Translate a short UI frequency code to a future-proof pandas offset alias."""
    return _RESAMPLE_FREQ_MAP.get(str(freq).upper(), str(freq))


def add_audit(action: str, detail: str = "") -> None:
    if "audit" not in st.session_state:
        st.session_state.audit = []
    st.session_state.audit.append({"Timestamp": _now(), "Action": action, "Detail": detail})


def make_unique_columns(columns: List[str]) -> List[str]:
    seen = {}
    result = []
    for col in columns:
        base = str(col) if str(col).strip() else "column"
        if base not in seen:
            seen[base] = 0
            result.append(base)
        else:
            seen[base] += 1
            result.append(f"{base}_{seen[base]}")
    return result


def clean_column_name(name: str) -> str:
    name = str(name).strip().lower()
    name = re.sub(r"[^0-9a-zA-Z]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name or "column"


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = make_unique_columns([clean_column_name(c) for c in df.columns])
    out = df.copy()
    out.columns = new_cols
    return out


def trim_text(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    text_cols = out.select_dtypes(include=["object", "string", "category"]).columns
    for col in text_cols:
        out[col] = out[col].astype("string").str.strip()
        out[col] = out[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return out


def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return (name[:31] or "Sheet")


def try_read_csv(file) -> pd.DataFrame:
    file.seek(0)
    for enc in ["utf-8", "utf-8-sig", "latin1", "cp1252"]:
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=enc)
        except Exception:
            continue
    file.seek(0)
    return pd.read_csv(file)


def load_uploaded_files(uploaded_files, read_all_excel_sheets: bool, add_source_columns: bool) -> pd.DataFrame:
    frames = []
    for file in uploaded_files:
        filename = file.name
        if filename.lower().endswith(".csv"):
            df = try_read_csv(file)
            if add_source_columns:
                df["__source_file"] = filename
                df["__source_sheet"] = "CSV"
            frames.append(df)
        else:
            excel = pd.ExcelFile(file)
            sheets = excel.sheet_names if read_all_excel_sheets else [excel.sheet_names[0]]
            for sheet in sheets:
                df = pd.read_excel(excel, sheet_name=sheet)
                if add_source_columns:
                    df["__source_file"] = filename
                    df["__source_sheet"] = sheet
                frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def make_demo_data(rows: int = 500) -> pd.DataFrame:
    rng = np.random.default_rng(2026)
    start = pd.Timestamp("2025-01-01")
    dates = start + pd.to_timedelta(rng.integers(0, 420, rows), unit="D")
    regions = rng.choice(["Lagos", "Abuja", "Port Harcourt", "Ibadan", "Kano", "Enugu"], rows)
    products = rng.choice(["Training", "Consulting", "Workbook", "Software", "Support"], rows)
    channels = rng.choice(["Online", "Referral", "School Visit", "Church Network", "Corporate"], rows)
    units = rng.integers(1, 55, rows)
    unit_price = rng.choice([2500, 5000, 7500, 10000, 15000, 25000], rows)
    discount = rng.choice([0, 0.05, 0.1, 0.15], rows, p=[0.55, 0.20, 0.17, 0.08])
    revenue = units * unit_price * (1 - discount)
    cost = revenue * rng.uniform(0.35, 0.72, rows)
    df = pd.DataFrame({
        "Date": dates,
        "Region": regions,
        "Product": products,
        "Channel": channels,
        "Units": units,
        "Unit Price": unit_price,
        "Discount Rate": discount,
        "Revenue": revenue.round(2),
        "Cost": cost.round(2),
        "Profit": (revenue - cost).round(2),
        "Customer Segment": rng.choice(["Student", "Parent", "School", "Church", "Corporate"], rows),
    })
    # introduce realistic issues
    if rows >= 20:
        df.loc[rng.choice(df.index, 12, replace=False), "Revenue"] = np.nan
        df.loc[rng.choice(df.index, 10, replace=False), "Region"] = " Lagos "
        df = pd.concat([df, df.iloc[:5]], ignore_index=True)
    return df


def numeric_cols(df: pd.DataFrame) -> List[str]:
    return df.select_dtypes(include=[np.number]).columns.tolist()


def text_cols(df: pd.DataFrame) -> List[str]:
    return df.select_dtypes(include=["object", "string", "category"]).columns.tolist()


def date_like_cols(df: pd.DataFrame) -> List[str]:
    cols = []
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            cols.append(col)
        elif any(token in str(col).lower() for token in ["date", "time", "year", "month"]):
            cols.append(col)
    return cols


def build_column_profile(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    n = len(df)
    for col in df.columns:
        s = df[col]
        missing = int(s.isna().sum())
        unique = int(s.nunique(dropna=True))
        sample_values = ", ".join([str(v) for v in s.dropna().astype(str).unique()[:5]])
        row = {
            "Column": col,
            "Data Type": str(s.dtype),
            "Missing Count": missing,
            "Missing %": round((missing / n * 100) if n else 0, 2),
            "Unique Values": unique,
            "Duplicate Values Count": int(n - unique - missing) if n else 0,
            "Sample Values": sample_values,
        }
        if pd.api.types.is_numeric_dtype(s):
            row.update({
                "Minimum": pd.to_numeric(s, errors="coerce").min(),
                "Maximum": pd.to_numeric(s, errors="coerce").max(),
                "Mean": pd.to_numeric(s, errors="coerce").mean(),
                "Median": pd.to_numeric(s, errors="coerce").median(),
                "Std Dev": pd.to_numeric(s, errors="coerce").std(),
            })
        elif pd.api.types.is_datetime64_any_dtype(s):
            row.update({"Minimum": s.min(), "Maximum": s.max()})
        rows.append(row)
    return pd.DataFrame(rows)


def build_missing_report(df: pd.DataFrame) -> pd.DataFrame:
    report = pd.DataFrame({
        "Column": df.columns,
        "Missing Count": [int(df[c].isna().sum()) for c in df.columns],
        "Missing %": [round(df[c].isna().mean() * 100, 2) for c in df.columns],
    })
    return report.sort_values("Missing Count", ascending=False)


def build_outlier_report(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for col in numeric_cols(df):
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        if s.empty:
            continue
        q1 = s.quantile(0.25)
        q3 = s.quantile(0.75)
        iqr = q3 - q1
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        rows.append({
            "Column": col,
            "Q1": q1,
            "Q3": q3,
            "IQR": iqr,
            "Lower Fence": lower,
            "Upper Fence": upper,
            "Low Outliers": int((s < lower).sum()),
            "High Outliers": int((s > upper).sum()),
            "Total Outliers": int(((s < lower) | (s > upper)).sum()),
        })
    return pd.DataFrame(rows)


def build_overview(df: pd.DataFrame, raw_df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    raw_rows = len(raw_df) if raw_df is not None else len(df)
    rows = [
        ("Generated At", _now()),
        ("Application", APP_NAME),
        ("Version", APP_VERSION),
        ("Raw Rows", raw_rows),
        ("Processed Rows", len(df)),
        ("Processed Columns", len(df.columns)),
        ("Numeric Columns", len(numeric_cols(df))),
        ("Text/Categorical Columns", len(text_cols(df))),
        ("Date-like Columns", len(date_like_cols(df))),
        ("Duplicate Rows", int(df.duplicated().sum()) if not df.empty else 0),
        ("Total Missing Cells", int(df.isna().sum().sum()) if not df.empty else 0),
        ("Memory Usage (MB)", round(df.memory_usage(deep=True).sum() / 1024 / 1024, 3) if not df.empty else 0),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def apply_missing_strategy(df: pd.DataFrame, numeric_strategy: str, text_strategy: str, custom_text: str = "Unknown") -> pd.DataFrame:
    out = df.copy()
    for col in numeric_cols(out):
        if numeric_strategy == "Zero":
            out[col] = out[col].fillna(0)
        elif numeric_strategy == "Mean":
            out[col] = out[col].fillna(out[col].mean())
        elif numeric_strategy == "Median":
            out[col] = out[col].fillna(out[col].median())
        elif numeric_strategy == "Forward fill":
            out[col] = out[col].ffill()
        elif numeric_strategy == "Backward fill":
            out[col] = out[col].bfill()
    for col in text_cols(out):
        if text_strategy == "Mode":
            mode = out[col].mode(dropna=True)
            if not mode.empty:
                out[col] = out[col].fillna(mode.iloc[0])
        elif text_strategy == "Custom value":
            out[col] = out[col].fillna(custom_text)
        elif text_strategy == "Forward fill":
            out[col] = out[col].ffill()
        elif text_strategy == "Backward fill":
            out[col] = out[col].bfill()
    return out


def treat_outliers(df: pd.DataFrame, columns: List[str], method: str) -> Tuple[pd.DataFrame, str]:
    out = df.copy()
    total_changed = 0
    for col in columns:
        s = pd.to_numeric(out[col], errors="coerce")
        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        iqr = q3 - q1
        lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        mask = (s < lower) | (s > upper)
        count = int(mask.sum())
        total_changed += count
        if method == "Cap/Winsorize to IQR fences":
            out[col] = s.clip(lower=lower, upper=upper)
        elif method == "Replace with median":
            out.loc[mask, col] = s.median()
        elif method == "Remove rows with outliers":
            out = out.loc[~mask].copy()
    return out, f"{method}; affected {total_changed} outlier values across {len(columns)} column(s)."


def convert_columns(df: pd.DataFrame, cols: List[str], target_type: str) -> pd.DataFrame:
    out = df.copy()
    for col in cols:
        if target_type == "Number":
            out[col] = pd.to_numeric(out[col], errors="coerce")
        elif target_type == "Text":
            out[col] = out[col].astype("string")
        elif target_type == "Date/Time":
            out[col] = pd.to_datetime(out[col], errors="coerce")
        elif target_type == "Category":
            out[col] = out[col].astype("category")
    return out


def apply_filter(df: pd.DataFrame, col: str, op: str, value: str) -> pd.DataFrame:
    out = df.copy()
    s = out[col]
    if op in ["contains", "not contains", "starts with", "ends with"]:
        ss = s.astype("string")
        if op == "contains":
            return out[ss.str.contains(value, case=False, na=False)]
        if op == "not contains":
            return out[~ss.str.contains(value, case=False, na=False)]
        if op == "starts with":
            return out[ss.str.startswith(value, na=False)]
        if op == "ends with":
            return out[ss.str.endswith(value, na=False)]
    if pd.api.types.is_numeric_dtype(s):
        v = pd.to_numeric(value, errors="coerce")
    else:
        v = value
    if op == "equals":
        return out[s == v]
    if op == "not equals":
        return out[s != v]
    if op == "greater than":
        return out[pd.to_numeric(s, errors="coerce") > float(v)]
    if op == "greater or equal":
        return out[pd.to_numeric(s, errors="coerce") >= float(v)]
    if op == "less than":
        return out[pd.to_numeric(s, errors="coerce") < float(v)]
    if op == "less or equal":
        return out[pd.to_numeric(s, errors="coerce") <= float(v)]
    if op == "is blank":
        return out[s.isna()]
    if op == "is not blank":
        return out[s.notna()]
    return out


def create_calculated_column(df: pd.DataFrame, new_col: str, operation: str, col_a: str, col_b: Optional[str], constant: float) -> pd.DataFrame:
    out = df.copy()
    a = pd.to_numeric(out[col_a], errors="coerce") if col_a else None
    b = pd.to_numeric(out[col_b], errors="coerce") if col_b else None
    if operation == "Add A + B":
        out[new_col] = a + b
    elif operation == "Subtract A - B":
        out[new_col] = a - b
    elif operation == "Multiply A * B":
        out[new_col] = a * b
    elif operation == "Divide A / B":
        out[new_col] = a / b.replace(0, np.nan)
    elif operation == "Add constant to A":
        out[new_col] = a + constant
    elif operation == "Multiply A by constant":
        out[new_col] = a * constant
    elif operation == "A as % of column total":
        total = a.sum(skipna=True)
        out[new_col] = (a / total * 100) if total else np.nan
    elif operation == "Running total of A":
        out[new_col] = a.cumsum()
    elif operation == "Difference from previous row":
        out[new_col] = a.diff()
    elif operation == "Rank A descending":
        out[new_col] = a.rank(ascending=False, method="dense")
    elif operation == "Z-score of A":
        std = a.std()
        out[new_col] = (a - a.mean()) / std if std else 0
    elif operation == "Natural log of A":
        out[new_col] = np.where(a > 0, np.log(a), np.nan)
    return out


def create_date_part(df: pd.DataFrame, date_col: str, part: str, new_col: str) -> pd.DataFrame:
    out = df.copy()
    d = pd.to_datetime(out[date_col], errors="coerce")
    if part == "Year":
        out[new_col] = d.dt.year
    elif part == "Quarter":
        out[new_col] = d.dt.quarter
    elif part == "Month number":
        out[new_col] = d.dt.month
    elif part == "Month name":
        out[new_col] = d.dt.month_name()
    elif part == "Weekday":
        out[new_col] = d.dt.day_name()
    elif part == "Year-Month":
        out[new_col] = d.dt.to_period("M").astype(str)
    elif part == "Date only":
        out[new_col] = d.dt.date
    return out


def generate_grouped_summary(df: pd.DataFrame, dims: List[str], measures: List[str], aggs: List[str]) -> pd.DataFrame:
    if not dims or not measures or not aggs:
        return pd.DataFrame()
    agg_map = {m: aggs for m in measures}
    grouped = df.groupby(dims, dropna=False).agg(agg_map)
    grouped.columns = ["_".join([str(x) for x in col if x]) for col in grouped.columns.to_flat_index()]
    return grouped.reset_index()


def generate_pivot(df: pd.DataFrame, rows: List[str], cols: Optional[str], value: str, agg: str) -> pd.DataFrame:
    if not rows or not value:
        return pd.DataFrame()
    if cols and cols != "None":
        return pd.pivot_table(df, index=rows, columns=cols, values=value, aggfunc=agg, fill_value=0).reset_index()
    return pd.pivot_table(df, index=rows, values=value, aggfunc=agg, fill_value=0).reset_index()


def generate_forecast(df: pd.DataFrame, date_col: str, metric: str, periods: int, freq: str, method: str) -> pd.DataFrame:
    d = df[[date_col, metric]].copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d[metric] = pd.to_numeric(d[metric], errors="coerce")
    d = d.dropna()
    if d.empty:
        return pd.DataFrame()
    freq = normalize_resample_freq(freq)
    ts = d.set_index(date_col).sort_index()[metric].resample(freq).sum().reset_index()
    ts = ts.rename(columns={date_col: "Period", metric: "Actual"})
    if len(ts) < 2:
        return ts
    future_dates = pd.date_range(ts["Period"].max(), periods=periods + 1, freq=freq)[1:]
    if method == "Moving average":
        window = min(3, len(ts))
        forecast_value = ts["Actual"].rolling(window=window).mean().iloc[-1]
        future = pd.DataFrame({"Period": future_dates, "Actual": np.nan, "Forecast": forecast_value})
    else:
        x = np.arange(len(ts))
        y = ts["Actual"].to_numpy(dtype=float)
        slope, intercept = np.polyfit(x, y, 1)
        fx = np.arange(len(ts), len(ts) + periods)
        future = pd.DataFrame({"Period": future_dates, "Actual": np.nan, "Forecast": intercept + slope * fx})
    hist = ts.copy()
    hist["Forecast"] = np.nan
    return pd.concat([hist, future], ignore_index=True)



def build_data_dictionary(df: pd.DataFrame) -> pd.DataFrame:
    """Create an editable data dictionary template for the exported workbook."""
    rows = []
    for col in df.columns:
        s = df[col]
        rows.append({
            "Column Name": col,
            "Detected Type": str(s.dtype),
            "Business Meaning": "",
            "Example Value": str(s.dropna().iloc[0]) if s.dropna().shape[0] else "",
            "Required?": "No",
            "Validation Rule": "",
            "Owner/Department": "",
            "Notes": "",
        })
    return pd.DataFrame(rows)


def build_validation_report(df: pd.DataFrame, col: str, rule: str, min_value=None, max_value=None, allowed_values: str = "") -> pd.DataFrame:
    """Create a row-level data quality validation report without using any AI service."""
    if col not in df.columns:
        return pd.DataFrame({"Issue": ["Selected column not found."]})
    s = df[col]
    mask = pd.Series(False, index=df.index)
    detail = ""
    if rule == "Not blank":
        mask = s.isna() | (s.astype("string").str.strip() == "")
        detail = f"{col} must not be blank"
    elif rule == "Unique values only":
        mask = s.duplicated(keep=False) & s.notna()
        detail = f"{col} must be unique"
    elif rule == "Numeric range":
        vals = pd.to_numeric(s, errors="coerce")
        lo = -np.inf if min_value is None else float(min_value)
        hi = np.inf if max_value is None else float(max_value)
        mask = vals.isna() | (vals < lo) | (vals > hi)
        detail = f"{col} must be numeric between {lo} and {hi}"
    elif rule == "Allowed list":
        allowed = [x.strip() for x in str(allowed_values).split(",") if x.strip()]
        mask = ~s.astype("string").isin(allowed)
        detail = f"{col} must be one of {allowed}"
    elif rule == "Valid date":
        vals = pd.to_datetime(s, errors="coerce")
        mask = vals.isna() & s.notna()
        detail = f"{col} must be a valid date"
    elif rule == "Positive number":
        vals = pd.to_numeric(s, errors="coerce")
        mask = vals.isna() | (vals <= 0)
        detail = f"{col} must be a positive number"
    out = df.loc[mask].copy()
    if out.empty:
        return pd.DataFrame({"Validation Result": ["PASS"], "Rule": [detail], "Invalid Rows": [0]})
    out.insert(0, "Validation Rule", detail)
    out.insert(1, "Original Row Number", out.index + 2)
    return out.reset_index(drop=True)


def build_pareto_analysis(df: pd.DataFrame, category_col: str, metric_col: str) -> pd.DataFrame:
    """Build an 80/20 Pareto table for category contribution analysis."""
    if category_col not in df.columns or metric_col not in df.columns:
        return pd.DataFrame()
    work = df[[category_col, metric_col]].copy()
    work[metric_col] = pd.to_numeric(work[metric_col], errors="coerce")
    work = work.dropna(subset=[metric_col])
    if work.empty:
        return pd.DataFrame()
    out = work.groupby(category_col, dropna=False)[metric_col].sum().reset_index()
    out = out.sort_values(metric_col, ascending=False).reset_index(drop=True)
    total = out[metric_col].sum()
    out["Contribution %"] = np.where(total != 0, out[metric_col] / total * 100, 0)
    out["Cumulative %"] = out["Contribution %"].cumsum()
    out["Pareto Class"] = np.where(out["Cumulative %"] <= 80, "Vital few", "Long tail")
    return out


def build_what_if_analysis(df: pd.DataFrame, metric_col: str, change_type: str, change_value: float, segment_col: Optional[str] = None, segment_value: Optional[str] = None) -> pd.DataFrame:
    """Create a simple scenario analysis table similar to Excel what-if modelling."""
    if metric_col not in df.columns:
        return pd.DataFrame()
    work = df.copy()
    base = pd.to_numeric(work[metric_col], errors="coerce")
    target_mask = pd.Series(True, index=work.index)
    segment_label = "All rows"
    if segment_col and segment_col != "None" and segment_col in work.columns and segment_value not in [None, ""]:
        target_mask = work[segment_col].astype("string") == str(segment_value)
        segment_label = f"{segment_col} = {segment_value}"
    scenario = base.copy()
    if change_type == "Increase by %":
        scenario.loc[target_mask] = base.loc[target_mask] * (1 + change_value / 100)
    elif change_type == "Decrease by %":
        scenario.loc[target_mask] = base.loc[target_mask] * (1 - change_value / 100)
    elif change_type == "Add constant":
        scenario.loc[target_mask] = base.loc[target_mask] + change_value
    elif change_type == "Multiply by factor":
        scenario.loc[target_mask] = base.loc[target_mask] * change_value
    summary = pd.DataFrame([
        [metric_col, segment_label, change_type, change_value, base.sum(skipna=True), scenario.sum(skipna=True), scenario.sum(skipna=True) - base.sum(skipna=True), int(target_mask.sum())]
    ], columns=["Metric", "Affected Segment", "Scenario", "Change Value", "Original Total", "Scenario Total", "Impact", "Affected Rows"])
    return summary



def build_quality_scorecard(df: pd.DataFrame) -> pd.DataFrame:
    """Create a simple rule-based data quality scorecard similar to an Excel QA checklist."""
    total_cells = int(df.shape[0] * df.shape[1]) if not df.empty else 0
    missing_cells = int(df.isna().sum().sum()) if total_cells else 0
    duplicate_rows = int(df.duplicated().sum()) if not df.empty else 0
    outliers = build_outlier_report(df)
    total_outliers = int(outliers["Total Outliers"].sum()) if not outliers.empty and "Total Outliers" in outliers else 0
    completeness = 100 - (missing_cells / total_cells * 100 if total_cells else 0)
    uniqueness = 100 - (duplicate_rows / len(df) * 100 if len(df) else 0)
    outlier_score = 100 - (total_outliers / max(len(df), 1) * 100 if len(df) else 0)
    schema_score = 100 if len(df.columns) > 0 else 0
    overall = np.mean([completeness, uniqueness, max(outlier_score, 0), schema_score])
    rows = [
        ["Completeness", round(completeness, 2), f"{missing_cells:,} missing cells out of {total_cells:,}", "Fill, drop, or investigate missing values."],
        ["Uniqueness", round(uniqueness, 2), f"{duplicate_rows:,} duplicate rows", "Remove or confirm duplicate records."],
        ["Outlier Health", round(max(outlier_score, 0), 2), f"{total_outliers:,} IQR outlier values", "Review outliers before modelling or reporting."],
        ["Schema Presence", round(schema_score, 2), f"{len(df.columns):,} columns detected", "Maintain a data dictionary and required fields."],
        ["Overall Quality Score", round(overall, 2), "Average of quality dimensions", "Use as a quick QA indicator, not a substitute for business review."],
    ]
    return pd.DataFrame(rows, columns=["Quality Dimension", "Score %", "Evidence", "Recommended Action"])


def build_time_intelligence(df: pd.DataFrame, date_col: str, metric_col: str, segment_col: Optional[str] = None, freq: str = "M") -> pd.DataFrame:
    """Build Excel-style time intelligence: period totals, previous period, change, change %, YTD, and rolling average."""
    if date_col not in df.columns or metric_col not in df.columns:
        return pd.DataFrame()
    work = df[[date_col, metric_col] + ([segment_col] if segment_col and segment_col != "None" and segment_col in df.columns else [])].copy()
    work[date_col] = pd.to_datetime(work[date_col], errors="coerce")
    work[metric_col] = pd.to_numeric(work[metric_col], errors="coerce")
    work = work.dropna(subset=[date_col, metric_col])
    if work.empty:
        return pd.DataFrame()
    # to_period still uses the classic single-letter codes (M/Q/W/D), so map back
    # any future-proof alias to the period alias to stay compatible.
    period_freq = {"ME": "M", "QE": "Q", "YE": "Y", "A": "Y"}.get(str(freq).upper(), str(freq).upper())
    work["Period"] = work[date_col].dt.to_period(period_freq).astype(str)
    work["Year"] = work[date_col].dt.year
    group_cols = ["Period", "Year"]
    if segment_col and segment_col != "None" and segment_col in work.columns:
        group_cols.append(segment_col)
    out = work.groupby(group_cols, dropna=False)[metric_col].sum().reset_index().sort_values(group_cols)
    sort_group = [segment_col] if segment_col and segment_col != "None" and segment_col in out.columns else []
    if sort_group:
        out["Previous Period"] = out.groupby(sort_group)[metric_col].shift(1)
        out["Period Change"] = out[metric_col] - out["Previous Period"]
        out["Period Change %"] = np.where(out["Previous Period"].fillna(0) != 0, out["Period Change"] / out["Previous Period"] * 100, np.nan)
        out["YTD"] = out.groupby(sort_group + ["Year"])[metric_col].cumsum()
        out["Rolling 3 Period Avg"] = out.groupby(sort_group)[metric_col].transform(lambda x: x.rolling(3, min_periods=1).mean())
    else:
        out["Previous Period"] = out[metric_col].shift(1)
        out["Period Change"] = out[metric_col] - out["Previous Period"]
        out["Period Change %"] = np.where(out["Previous Period"].fillna(0) != 0, out["Period Change"] / out["Previous Period"] * 100, np.nan)
        out["YTD"] = out.groupby("Year")[metric_col].cumsum()
        out["Rolling 3 Period Avg"] = out[metric_col].rolling(3, min_periods=1).mean()
    return out


def build_kpi_target_analysis(df: pd.DataFrame, actual_col: str, target_mode: str, target_value: float = 0.0, target_col: Optional[str] = None, segment_col: Optional[str] = None) -> pd.DataFrame:
    """Compare actual performance with a constant target or a target column."""
    if actual_col not in df.columns:
        return pd.DataFrame()
    work = df.copy()
    work[actual_col] = pd.to_numeric(work[actual_col], errors="coerce")
    group_cols = [segment_col] if segment_col and segment_col != "None" and segment_col in work.columns else []
    if group_cols:
        actual = work.groupby(group_cols, dropna=False)[actual_col].sum().reset_index().rename(columns={actual_col: "Actual"})
    else:
        actual = pd.DataFrame({"Segment": ["All Rows"], "Actual": [work[actual_col].sum(skipna=True)]})
    if target_mode == "Target column" and target_col and target_col in work.columns:
        work[target_col] = pd.to_numeric(work[target_col], errors="coerce")
        if group_cols:
            target = work.groupby(group_cols, dropna=False)[target_col].sum().reset_index().rename(columns={target_col: "Target"})
            out = actual.merge(target, on=group_cols, how="left")
        else:
            out = actual.copy()
            out["Target"] = work[target_col].sum(skipna=True)
    else:
        out = actual.copy()
        out["Target"] = float(target_value)
    out["Variance"] = out["Actual"] - out["Target"]
    out["Attainment %"] = np.where(out["Target"].fillna(0) != 0, out["Actual"] / out["Target"] * 100, np.nan)
    out["Status"] = np.where(out["Variance"] >= 0, "Met/Exceeded", "Below Target")
    return out


def create_binned_column(df: pd.DataFrame, source_col: str, new_col: str, bins: int, method: str) -> pd.DataFrame:
    """Create numeric bands/buckets like Excel grouping or histogram binning."""
    out = df.copy()
    vals = pd.to_numeric(out[source_col], errors="coerce")
    if method == "Quantile bands":
        out[new_col] = pd.qcut(vals, q=min(max(int(bins), 2), max(vals.nunique(dropna=True), 1)), duplicates="drop")
    else:
        out[new_col] = pd.cut(vals, bins=max(int(bins), 2))
    out[new_col] = out[new_col].astype("string")
    return out


def create_if_column(df: pd.DataFrame, source_col: str, operator: str, compare_value: str, true_value: str, false_value: str, new_col: str) -> pd.DataFrame:
    """Create an IF-style classification column without writing Excel formulas manually."""
    out = df.copy()
    s = out[source_col]
    if operator in [">", ">=", "<", "<="]:
        left = pd.to_numeric(s, errors="coerce")
        right = pd.to_numeric(compare_value, errors="coerce")
        if operator == ">":
            mask = left > right
        elif operator == ">=":
            mask = left >= right
        elif operator == "<":
            mask = left < right
        else:
            mask = left <= right
    elif operator == "equals":
        mask = s.astype("string") == str(compare_value)
    elif operator == "not equals":
        mask = s.astype("string") != str(compare_value)
    elif operator == "contains":
        mask = s.astype("string").str.contains(str(compare_value), case=False, na=False)
    elif operator == "is blank":
        mask = s.isna() | (s.astype("string").str.strip() == "")
    else:
        mask = s.notna() & (s.astype("string").str.strip() != "")
    out[new_col] = np.where(mask, true_value, false_value)
    return out


def build_reconciliation_report(left_df: pd.DataFrame, right_df: pd.DataFrame, left_key: str, right_key: str, compare_cols: List[str], tolerance: float = 0.0) -> pd.DataFrame:
    """Compare two tables to find missing records and value mismatches."""
    if left_key not in left_df.columns or right_key not in right_df.columns:
        return pd.DataFrame({"Issue Type": ["Key column not found"]})
    left = left_df.copy()
    right = right_df.copy()
    left["__recon_key"] = left[left_key].astype("string")
    right["__recon_key"] = right[right_key].astype("string")
    left_keys = set(left["__recon_key"].dropna())
    right_keys = set(right["__recon_key"].dropna())
    rows = []
    for k in sorted(left_keys - right_keys):
        rows.append({"Issue Type": "Missing in comparison file", "Key": k, "Column": "", "Current Value": "Present", "Comparison Value": "Missing", "Difference": ""})
    for k in sorted(right_keys - left_keys):
        rows.append({"Issue Type": "Missing in current dataset", "Key": k, "Column": "", "Current Value": "Missing", "Comparison Value": "Present", "Difference": ""})
    common = sorted(left_keys & right_keys)
    if common and compare_cols:
        l1 = left.drop_duplicates("__recon_key").set_index("__recon_key")
        r1 = right.drop_duplicates("__recon_key").set_index("__recon_key")
        for col in compare_cols:
            if col not in left_df.columns or col not in right_df.columns:
                continue
            for k in common:
                lv = l1.at[k, col]
                rv = r1.at[k, col]
                ln = pd.to_numeric(pd.Series([lv]), errors="coerce").iloc[0]
                rn = pd.to_numeric(pd.Series([rv]), errors="coerce").iloc[0]
                if pd.notna(ln) and pd.notna(rn):
                    diff = float(ln - rn)
                    mismatch = abs(diff) > tolerance
                else:
                    diff = ""
                    mismatch = str(lv) != str(rv)
                if mismatch:
                    rows.append({"Issue Type": "Value mismatch", "Key": k, "Column": col, "Current Value": lv, "Comparison Value": rv, "Difference": diff})
    if not rows:
        return pd.DataFrame({"Issue Type": ["PASS"], "Key": [""], "Column": [""], "Current Value": ["No reconciliation issues found"], "Comparison Value": [""], "Difference": [""]})
    return pd.DataFrame(rows)



def run_sql_query(df: pd.DataFrame, sql: str) -> pd.DataFrame:
    """Run a safe SELECT/WITH SQL query against the working dataframe using Python's free sqlite3 library."""
    sql_clean = str(sql).strip().rstrip(";").strip()
    if not sql_clean:
        return pd.DataFrame({"Message": ["No SQL query provided."]})
    low = sql_clean.lower()
    if not (low.startswith("select") or low.startswith("with")):
        return pd.DataFrame({"Error": ["Only SELECT or WITH queries are allowed for safety."]})
    # Defense-in-depth: block multi-statement injection and any write/DDL keywords.
    if ";" in sql_clean:
        return pd.DataFrame({"Error": ["Multiple statements are not allowed. Run a single SELECT query."]})
    _forbidden = ["attach", "detach", "pragma", "insert", "update", "delete", "drop",
                  "alter", "create", "replace", "vacuum", "reindex", "load_extension"]
    if any(re.search(rf"\b{kw}\b", low) for kw in _forbidden):
        return pd.DataFrame({"Error": ["Query contains a forbidden keyword. Only read-only SELECT queries are permitted."]})
    try:
        conn = sqlite3.connect(":memory:")
        safe_df = df.copy()
        safe_df.columns = make_unique_columns([clean_column_name(c) for c in safe_df.columns])
        safe_df.to_sql("data", conn, index=False, if_exists="replace")
        result = pd.read_sql_query(sql_clean, conn)
        conn.close()
        return result
    except Exception as exc:
        return pd.DataFrame({"Error": [str(exc)], "Hint": ["Use table name data. Column names are standardized to snake_case in SQL mode."]})


def build_anomaly_detail_report(df: pd.DataFrame, columns: List[str], method: str = "IQR", z_threshold: float = 3.0) -> pd.DataFrame:
    """Produce row-level anomaly details using IQR or z-score rules. No AI or paid service is used."""
    rows = []
    for col in columns:
        if col not in df.columns:
            continue
        vals = pd.to_numeric(df[col], errors="coerce")
        if vals.dropna().empty:
            continue
        if method == "Z-score":
            std = vals.std()
            if not std or pd.isna(std):
                continue
            z = (vals - vals.mean()) / std
            mask = z.abs() > z_threshold
            for idx in df.index[mask.fillna(False)]:
                rows.append({
                    "Original Row Number": int(idx) + 2,
                    "Column": col,
                    "Value": vals.loc[idx],
                    "Method": "Z-score",
                    "Lower Limit": -z_threshold,
                    "Upper Limit": z_threshold,
                    "Score/Distance": z.loc[idx],
                    "Direction": "High" if z.loc[idx] > 0 else "Low",
                })
        else:
            q1, q3 = vals.quantile(0.25), vals.quantile(0.75)
            iqr = q3 - q1
            lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
            mask = (vals < lower) | (vals > upper)
            for idx in df.index[mask.fillna(False)]:
                value = vals.loc[idx]
                rows.append({
                    "Original Row Number": int(idx) + 2,
                    "Column": col,
                    "Value": value,
                    "Method": "IQR",
                    "Lower Limit": lower,
                    "Upper Limit": upper,
                    "Score/Distance": min(abs(value - lower), abs(value - upper)),
                    "Direction": "High" if value > upper else "Low",
                })
    if not rows:
        return pd.DataFrame({"Anomaly Result": ["No anomalies found with the selected rule."], "Method": [method]})
    return pd.DataFrame(rows).sort_values(["Column", "Original Row Number"]).head(20000)


def build_rfm_analysis(df: pd.DataFrame, customer_col: str, date_col: str, amount_col: str) -> pd.DataFrame:
    """Create rule-based RFM customer segmentation: Recency, Frequency, Monetary."""
    if customer_col not in df.columns or date_col not in df.columns or amount_col not in df.columns:
        return pd.DataFrame()
    work = df[[customer_col, date_col, amount_col]].copy()
    work[date_col] = pd.to_datetime(work[date_col], errors="coerce")
    work[amount_col] = pd.to_numeric(work[amount_col], errors="coerce")
    work = work.dropna(subset=[customer_col, date_col, amount_col])
    if work.empty:
        return pd.DataFrame({"Message": ["No valid rows for RFM analysis."]})
    snapshot_date = work[date_col].max() + pd.Timedelta(days=1)
    rfm = work.groupby(customer_col).agg(
        Last_Date=(date_col, "max"),
        Frequency=(date_col, "count"),
        Monetary=(amount_col, "sum"),
    ).reset_index()
    rfm["Recency"] = (snapshot_date - rfm["Last_Date"]).dt.days
    def score(series, reverse=False):
        try:
            ranked = pd.qcut(series.rank(method="first"), 5, labels=[1, 2, 3, 4, 5]).astype(int)
            return 6 - ranked if reverse else ranked
        except Exception:
            return pd.Series([3] * len(series), index=series.index)
    rfm["R_Score"] = score(rfm["Recency"], reverse=True)
    rfm["F_Score"] = score(rfm["Frequency"], reverse=False)
    rfm["M_Score"] = score(rfm["Monetary"], reverse=False)
    rfm["RFM_Score"] = rfm["R_Score"].astype(str) + rfm["F_Score"].astype(str) + rfm["M_Score"].astype(str)
    total_score = rfm[["R_Score", "F_Score", "M_Score"]].sum(axis=1)
    rfm["Segment"] = np.select(
        [total_score >= 13, total_score >= 10, total_score >= 7, rfm["R_Score"] <= 2],
        ["Champions", "Loyal / Growth", "Needs Attention", "At Risk"],
        default="Low Value / New",
    )
    return rfm.sort_values(["Segment", "Monetary"], ascending=[True, False])


def build_cohort_analysis(df: pd.DataFrame, entity_col: str, date_col: str, metric_col: Optional[str] = None) -> pd.DataFrame:
    """Build monthly cohort retention or value table based on first activity month."""
    if entity_col not in df.columns or date_col not in df.columns:
        return pd.DataFrame()
    cols = [entity_col, date_col] + ([metric_col] if metric_col and metric_col != "None" and metric_col in df.columns else [])
    work = df[cols].copy()
    work[date_col] = pd.to_datetime(work[date_col], errors="coerce")
    if metric_col and metric_col != "None" and metric_col in work.columns:
        work[metric_col] = pd.to_numeric(work[metric_col], errors="coerce").fillna(0)
    work = work.dropna(subset=[entity_col, date_col])
    if work.empty:
        return pd.DataFrame({"Message": ["No valid rows for cohort analysis."]})
    work["Activity_Month"] = work[date_col].dt.to_period("M")
    first_month = work.groupby(entity_col)["Activity_Month"].min().rename("Cohort_Month")
    work = work.join(first_month, on=entity_col)
    work["Cohort_Index"] = (work["Activity_Month"].dt.year - work["Cohort_Month"].dt.year) * 12 + (work["Activity_Month"].dt.month - work["Cohort_Month"].dt.month) + 1
    if metric_col and metric_col != "None" and metric_col in work.columns:
        cohort = pd.pivot_table(work, index="Cohort_Month", columns="Cohort_Index", values=metric_col, aggfunc="sum", fill_value=0)
        cohort.insert(0, "Measure", f"Sum of {metric_col}")
    else:
        cohort = pd.pivot_table(work, index="Cohort_Month", columns="Cohort_Index", values=entity_col, aggfunc=pd.Series.nunique, fill_value=0)
        base = cohort.iloc[:, 0].replace(0, np.nan)
        cohort_pct = cohort.div(base, axis=0) * 100
        cohort_pct.columns = [f"Month {c} Retention %" for c in cohort_pct.columns]
        cohort.columns = [f"Month {c} Users" for c in cohort.columns]
        cohort = pd.concat([cohort, cohort_pct], axis=1)
        cohort.insert(0, "Measure", "Unique entities and retention %")
    cohort = cohort.reset_index()
    cohort["Cohort_Month"] = cohort["Cohort_Month"].astype(str)
    return cohort


def build_rule_based_executive_summary(df: pd.DataFrame, metric_col: Optional[str], category_col: Optional[str], date_col: Optional[str]) -> pd.DataFrame:
    """Generate non-AI, rule-based executive summary bullets for the workbook."""
    rows = []
    rows.append({"Section": "Dataset", "Insight": f"The processed dataset contains {len(df):,} rows and {len(df.columns):,} columns."})
    rows.append({"Section": "Quality", "Insight": f"Missing cells: {int(df.isna().sum().sum()):,}; duplicate rows: {int(df.duplicated().sum()):,}."})
    if metric_col and metric_col in df.columns:
        metric = pd.to_numeric(df[metric_col], errors="coerce")
        rows.append({"Section": "Metric", "Insight": f"Total {metric_col}: {metric.sum(skipna=True):,.2f}; average: {metric.mean(skipna=True):,.2f}; maximum: {metric.max(skipna=True):,.2f}."})
        if category_col and category_col in df.columns:
            top = df.assign(__metric=metric).groupby(category_col, dropna=False)["__metric"].sum().sort_values(ascending=False).head(3)
            for name, value in top.items():
                rows.append({"Section": "Top contributors", "Insight": f"{name} contributes {value:,.2f} of {metric_col}."})
    if date_col and date_col in df.columns and metric_col and metric_col in df.columns:
        temp = df[[date_col, metric_col]].copy()
        temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
        temp[metric_col] = pd.to_numeric(temp[metric_col], errors="coerce")
        temp = temp.dropna()
        if not temp.empty:
            temp["Period"] = temp[date_col].dt.to_period("M").astype(str)
            monthly = temp.groupby("Period")[metric_col].sum().sort_index()
            if len(monthly) >= 2:
                change = monthly.iloc[-1] - monthly.iloc[-2]
                rows.append({"Section": "Trend", "Insight": f"Latest period {monthly.index[-1]} changed by {change:,.2f} compared with the previous period."})
    rows.append({"Section": "Method", "Insight": "This summary is rule-based. No AI API or paid model generated it."})
    return pd.DataFrame(rows)


def build_split_workbook(df: pd.DataFrame, split_col: str, max_sheets: int = 25) -> bytes:
    """Create a separate Excel workbook where each selected segment becomes a sheet."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": "#001524", "font_color": "white", "border": 1})
        values = df[split_col].dropna().astype(str).unique().tolist()[:max_sheets]
        index_df = pd.DataFrame({"Generated Sheet": [safe_sheet_name(v) for v in values], "Segment Value": values})
        write_dataframe(writer, index_df, "Index", header_format=header_format, heatmap=False)
        for val in values:
            part = df[df[split_col].astype(str) == val].copy()
            write_dataframe(writer, part, safe_sheet_name(str(val)), header_format=header_format)
    return output.getvalue()



def build_schema_contract_report(df: pd.DataFrame, expected_required: List[str], key_cols: List[str], expected_numeric: List[str], expected_dates: List[str]) -> pd.DataFrame:
    """Validate the dataset against a simple data contract/schema rule set."""
    rows = []
    existing = set(df.columns.astype(str))
    for col in expected_required:
        col = str(col).strip()
        if not col:
            continue
        if col not in existing:
            rows.append({"Check": "Required column exists", "Column": col, "Status": "FAIL", "Issue Count": 1, "Details": "Column is missing from dataset."})
        else:
            blanks = int(df[col].isna().sum() + (df[col].astype("string").str.strip() == "").sum())
            rows.append({"Check": "Required column not blank", "Column": col, "Status": "PASS" if blanks == 0 else "WARN", "Issue Count": blanks, "Details": "Blank values found." if blanks else "No blanks."})
    if key_cols:
        present_keys = [c for c in key_cols if c in df.columns]
        missing_keys = [c for c in key_cols if c not in df.columns]
        for c in missing_keys:
            rows.append({"Check": "Key column exists", "Column": c, "Status": "FAIL", "Issue Count": 1, "Details": "Selected key column is missing."})
        if present_keys:
            dupes = int(df.duplicated(subset=present_keys, keep=False).sum())
            blanks = int(df[present_keys].isna().any(axis=1).sum())
            rows.append({"Check": "Composite key uniqueness", "Column": ", ".join(present_keys), "Status": "PASS" if dupes == 0 else "FAIL", "Issue Count": dupes, "Details": "Duplicate key rows found." if dupes else "Keys are unique."})
            rows.append({"Check": "Composite key completeness", "Column": ", ".join(present_keys), "Status": "PASS" if blanks == 0 else "FAIL", "Issue Count": blanks, "Details": "Blank key rows found." if blanks else "No blank keys."})
    for col in expected_numeric:
        if col in df.columns:
            invalid = int(pd.to_numeric(df[col], errors="coerce").isna().sum() - df[col].isna().sum())
            rows.append({"Check": "Numeric type expectation", "Column": col, "Status": "PASS" if invalid == 0 else "FAIL", "Issue Count": max(invalid, 0), "Details": "Non-numeric values found." if invalid else "Column can be treated as numeric."})
    for col in expected_dates:
        if col in df.columns:
            parsed = pd.to_datetime(df[col], errors="coerce")
            invalid = int((parsed.isna() & df[col].notna()).sum())
            rows.append({"Check": "Date type expectation", "Column": col, "Status": "PASS" if invalid == 0 else "FAIL", "Issue Count": invalid, "Details": "Invalid date values found." if invalid else "Column can be treated as date/time."})
    if not rows:
        rows.append({"Check": "Schema contract", "Column": "", "Status": "INFO", "Issue Count": 0, "Details": "No contract rules were supplied."})
    return pd.DataFrame(rows)


def build_fuzzy_duplicate_report(df: pd.DataFrame, cols: List[str], threshold: float = 90.0, max_candidates: int = 350) -> pd.DataFrame:
    """Find near-duplicate text records using Python's built-in difflib. Free and deterministic, but limited for performance."""
    from difflib import SequenceMatcher
    if not cols:
        return pd.DataFrame({"Message": ["Select at least one text column for fuzzy duplicate detection."]})
    work = df[cols].fillna("").astype(str).copy()
    work["__candidate"] = work.apply(lambda r: " | ".join([str(x).strip().lower() for x in r.values]), axis=1)
    candidates = work["__candidate"].drop_duplicates().head(max_candidates).tolist()
    first_row = work.reset_index().drop_duplicates("__candidate").set_index("__candidate")["index"].to_dict()
    rows = []
    for i in range(len(candidates)):
        for j in range(i + 1, len(candidates)):
            a, b = candidates[i], candidates[j]
            if not a or not b or a == b:
                continue
            score = SequenceMatcher(None, a, b).ratio() * 100
            if score >= threshold:
                rows.append({
                    "Similarity %": round(score, 2),
                    "Row A": int(first_row[a]) + 2,
                    "Row B": int(first_row[b]) + 2,
                    "Candidate A": a,
                    "Candidate B": b,
                    "Compared Columns": ", ".join(cols),
                })
    if not rows:
        return pd.DataFrame({"Fuzzy Duplicate Result": ["No near-duplicates found at the selected threshold."], "Threshold %": [threshold]})
    return pd.DataFrame(rows).sort_values("Similarity %", ascending=False).head(5000)


def build_abc_analysis(df: pd.DataFrame, category_col: str, metric_col: str, a_threshold: float = 80.0, b_threshold: float = 95.0) -> pd.DataFrame:
    """ABC analysis: classify contributors into A/B/C bands based on cumulative contribution."""
    if category_col not in df.columns or metric_col not in df.columns:
        return pd.DataFrame()
    work = df[[category_col, metric_col]].copy()
    work[metric_col] = pd.to_numeric(work[metric_col], errors="coerce")
    work = work.dropna(subset=[metric_col])
    if work.empty:
        return pd.DataFrame({"Message": ["No valid metric values for ABC analysis."]})
    out = work.groupby(category_col, dropna=False)[metric_col].sum().reset_index().sort_values(metric_col, ascending=False)
    total = out[metric_col].sum()
    out["Contribution %"] = np.where(total != 0, out[metric_col] / total * 100, 0)
    out["Cumulative %"] = out["Contribution %"].cumsum()
    out["ABC Class"] = np.select(
        [out["Cumulative %"] <= a_threshold, out["Cumulative %"] <= b_threshold],
        ["A - highest impact", "B - medium impact"],
        default="C - long tail",
    )
    return out


def build_basket_analysis(df: pd.DataFrame, transaction_col: str, item_col: str, min_count: int = 1, max_transactions: int = 5000) -> pd.DataFrame:
    """Build simple item-pair co-occurrence/basket analysis from transaction and item columns."""
    from itertools import combinations
    if transaction_col not in df.columns or item_col not in df.columns:
        return pd.DataFrame()
    work = df[[transaction_col, item_col]].dropna().astype(str)
    grouped = work.groupby(transaction_col)[item_col].apply(lambda x: sorted(set([i.strip() for i in x if i.strip()]))).head(max_transactions)
    pair_counts = {}
    tx_count = 0
    for items in grouped:
        if len(items) < 2:
            continue
        tx_count += 1
        for a, b in combinations(items, 2):
            pair_counts[(a, b)] = pair_counts.get((a, b), 0) + 1
    rows = []
    for (a, b), count in pair_counts.items():
        if count >= min_count:
            rows.append({"Item A": a, "Item B": b, "Pair Count": count, "Support %": round(count / max(len(grouped), 1) * 100, 2)})
    if not rows:
        return pd.DataFrame({"Basket Result": ["No item pairs found with the selected settings."]})
    return pd.DataFrame(rows).sort_values(["Pair Count", "Support %"], ascending=False).head(10000)


def scale_numeric_columns(df: pd.DataFrame, cols: List[str], method: str, suffix: str) -> pd.DataFrame:
    """Add scaled numeric columns using min-max, z-score, or percent-rank methods."""
    out = df.copy()
    for col in cols:
        vals = pd.to_numeric(out[col], errors="coerce")
        new_col = f"{col}_{suffix}"
        if method == "Min-Max 0-1":
            denom = vals.max() - vals.min()
            out[new_col] = (vals - vals.min()) / denom if denom else 0
        elif method == "Z-score standardize":
            std = vals.std()
            out[new_col] = (vals - vals.mean()) / std if std else 0
        else:
            out[new_col] = vals.rank(pct=True)
    return out


def melt_to_long_format(df: pd.DataFrame, id_cols: List[str], value_cols: List[str], var_name: str, value_name: str) -> pd.DataFrame:
    """Convert wide data to long format, similar to Excel Power Query Unpivot."""
    if not value_cols:
        return df.copy()
    return df.melt(id_vars=id_cols, value_vars=value_cols, var_name=var_name or "Variable", value_name=value_name or "Value")



def build_hmg_brand_profile() -> pd.DataFrame:
    """Return HMG/CSS Adewale brand information for UI and workbook embedding."""
    return pd.DataFrame([{"Item": k, "Details": v} for k, v in HMG_BRAND_PROFILE.items()])


def build_hmg_ecosystem_map() -> pd.DataFrame:
    return pd.DataFrame([
        ["HMG Concepts", "Parent brand (est. 2015)", "His Marvellous Grace — Nigerian education & technology brand founded by Adewale Samson Adeagbo. One brand, four living missions: Education, Technology, Media, Gospel.", "https://hmgconcepts.pages.dev"],
        ["HMG Academy", "Education arm — OWNS this platform", "Strictly virtual learning institution: vetted tutors, virtual home-schooling, exam prep (WAEC/NECO/GCE/BECE/UTME/IGCSE/IELTS/SAT/JUPEB), LMS, free CBT Pro and parent monitoring.", "https://hmgacademy.pages.dev"],
        ["HMG Technologies", "Innovation arm — BUILT this platform", "AI-augmented tools, CBT systems, data dashboards, ML models, Excel automation and 11 data-tool simulators for Nigerian businesses, schools, NGOs and churches.", "https://hmgtechnologies.pages.dev"],
        ["HMG Media", "Content/visibility arm", "Purpose-led educational video, brand storytelling, social content systems and graphics for the HMG ecosystem and values-driven brands.", "https://hmgmedia.pages.dev"],
        ["HMG Gospel", "Faith arm", "Christ-centred digital outreach: dramavangelism, techvangelism, podcasts, ebooks and church/community support. 'Every platform a pulpit, every tool a testimony.'", "https://hmggospel.pages.dev"],
        ["CSS Adewale Portfolio", "Founder profile", "Adewale Samson Adeagbo: AI-Augmented Solutions Developer, Data Scientist, STEM Educator. 34 live projects across EdTech, DataTech and FaithTech.", "https://cssadewale.pages.dev"],
    ], columns=["Ecosystem Unit", "Role", "Description", "Link"])


def build_weighted_scorecard(df: pd.DataFrame, score_cols: List[str], weights: List[float], cost_cols: List[str]) -> pd.DataFrame:
    """Create a weighted scoring/ranking table using selected numeric columns."""
    if not score_cols:
        return pd.DataFrame({"Message": ["No scoring columns selected."]})
    if len(weights) != len(score_cols):
        weights = [1.0] * len(score_cols)
    total_weight = sum([abs(float(w)) for w in weights]) or 1.0
    out = df.copy()
    score = pd.Series(0.0, index=out.index)
    for col, weight in zip(score_cols, weights):
        vals = pd.to_numeric(out[col], errors="coerce")
        denom = vals.max() - vals.min()
        scaled = (vals - vals.min()) / denom if denom else pd.Series(0.0, index=out.index)
        if col in cost_cols:
            scaled = 1 - scaled
        component = scaled.fillna(0) * (float(weight) / total_weight) * 100
        out[f"score_component_{clean_column_name(col)}"] = component
        score += component
    out["Weighted Score"] = score.round(4)
    out["Weighted Rank"] = out["Weighted Score"].rank(ascending=False, method="dense").astype(int)
    return out.sort_values("Weighted Score", ascending=False)


def build_goal_seek_analysis(df: pd.DataFrame, metric_col: str, target_total: float, segment_col: Optional[str] = None) -> pd.DataFrame:
    """Simple Excel-style goal seek / target gap analysis."""
    if metric_col not in df.columns:
        return pd.DataFrame()
    work = df.copy()
    work[metric_col] = pd.to_numeric(work[metric_col], errors="coerce")
    if segment_col and segment_col != "None" and segment_col in work.columns:
        out = work.groupby(segment_col, dropna=False)[metric_col].sum().reset_index().rename(columns={metric_col: "Current Total"})
        out["Target Total"] = float(target_total)
        out["Gap"] = out["Target Total"] - out["Current Total"]
        out["Required Lift %"] = np.where(out["Current Total"] != 0, out["Gap"] / out["Current Total"] * 100, np.nan)
        counts = work.groupby(segment_col, dropna=False).size().reset_index(name="Rows")
        out = out.merge(counts, on=segment_col, how="left")
    else:
        current = work[metric_col].sum(skipna=True)
        rows = len(work)
        out = pd.DataFrame({"Scope": ["All rows"], "Current Total": [current], "Target Total": [float(target_total)], "Gap": [float(target_total) - current], "Rows": [rows]})
        out["Required Lift %"] = np.where(out["Current Total"] != 0, out["Gap"] / out["Current Total"] * 100, np.nan)
    out["Uniform Addition Per Row"] = np.where(out["Rows"] != 0, out["Gap"] / out["Rows"], np.nan)
    out["Status"] = np.where(out["Gap"] <= 0, "Target met/exceeded", "Below target")
    return out


def build_calendar_table(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    """Create an Excel/Power BI-style calendar/date dimension table from a date column."""
    if date_col not in df.columns:
        return pd.DataFrame()
    dates = pd.to_datetime(df[date_col], errors="coerce").dropna()
    if dates.empty:
        return pd.DataFrame({"Message": ["No valid dates found."]})
    calendar = pd.DataFrame({"Date": pd.date_range(dates.min().date(), dates.max().date(), freq="D")})
    calendar["Year"] = calendar["Date"].dt.year
    calendar["Quarter"] = "Q" + calendar["Date"].dt.quarter.astype(str)
    calendar["Month Number"] = calendar["Date"].dt.month
    calendar["Month Name"] = calendar["Date"].dt.month_name()
    calendar["Month Short"] = calendar["Date"].dt.strftime("%b")
    calendar["Year-Month"] = calendar["Date"].dt.to_period("M").astype(str)
    calendar["Week Number"] = calendar["Date"].dt.isocalendar().week.astype(int)
    calendar["Day"] = calendar["Date"].dt.day
    calendar["Day Name"] = calendar["Date"].dt.day_name()
    calendar["Is Weekend"] = calendar["Date"].dt.weekday >= 5
    calendar["Fiscal Year Apr-Mar"] = np.where(calendar["Date"].dt.month >= 4, calendar["Year"], calendar["Year"] - 1)
    return calendar


def build_report_brief(df: pd.DataFrame, metric_col: Optional[str], category_col: Optional[str], date_col: Optional[str]) -> pd.DataFrame:
    """Produce a practical rule-based analyst brief with recommendations; no AI API is used."""
    rows = []
    rows.append(["Context", "Dataset loaded", f"{len(df):,} rows and {len(df.columns):,} columns are currently in the processed dataset.", "Confirm that the dataset scope matches the reporting objective."])
    missing = int(df.isna().sum().sum())
    dupes = int(df.duplicated().sum())
    rows.append(["Quality", "Missing and duplicate check", f"Missing cells: {missing:,}. Duplicate rows: {dupes:,}.", "Resolve high-impact missing values and confirm whether duplicates are valid before final reporting."])
    if metric_col and metric_col in df.columns:
        metric = pd.to_numeric(df[metric_col], errors="coerce")
        rows.append(["Metric", f"{metric_col} performance", f"Total: {metric.sum(skipna=True):,.2f}; Average: {metric.mean(skipna=True):,.2f}; Maximum: {metric.max(skipna=True):,.2f}.", "Use the metric trend and category breakdown to identify performance drivers."])
        if category_col and category_col in df.columns:
            grouped = df.assign(__metric=metric).groupby(category_col, dropna=False)["__metric"].sum().sort_values(ascending=False)
            if not grouped.empty:
                top_name, top_val = grouped.index[0], grouped.iloc[0]
                share = top_val / grouped.sum() * 100 if grouped.sum() else 0
                rows.append(["Contribution", "Top contributor", f"{top_name} contributes {top_val:,.2f}, representing {share:.2f}% of total {metric_col}.", "Review whether performance is overly concentrated in a few contributors."])
    if date_col and date_col in df.columns and metric_col and metric_col in df.columns:
        temp = df[[date_col, metric_col]].copy()
        temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
        temp[metric_col] = pd.to_numeric(temp[metric_col], errors="coerce")
        temp = temp.dropna()
        if len(temp) > 1:
            temp["Period"] = temp[date_col].dt.to_period("M").astype(str)
            monthly = temp.groupby("Period")[metric_col].sum().sort_index()
            if len(monthly) >= 2:
                pct = ((monthly.iloc[-1] - monthly.iloc[-2]) / monthly.iloc[-2] * 100) if monthly.iloc[-2] else np.nan
                rows.append(["Trend", "Latest period movement", f"Latest month {monthly.index[-1]} changed by {pct:.2f}% versus the previous month.", "Investigate operational events behind large positive or negative movement."])
    rows.append(["Governance", "No AI API", "This brief is generated by deterministic business rules, not an AI model.", "Safe for free-tier deployment and repeated usage without token cost."])
    return pd.DataFrame(rows, columns=["Area", "Finding", "Evidence", "Recommended Action"])


# ============================================================
# v6 ENHANCEMENTS — free, no-AI, deterministic feature additions
# (Inspired by competitor research: VisualizeMyData, ChartGen,
#  Bricks, Sweetviz, ydata-profiling. All implemented with the
#  existing free stack: pandas / numpy / plotly / XlsxWriter.)
# ============================================================

def read_google_sheet(url: str) -> pd.DataFrame:
    """Import a *published* or shareable Google Sheet as a DataFrame via its CSV export.

    No Google API key is required. The sheet must be viewable by anyone with the link
    (or published to the web). We rewrite common Google Sheets URLs to their CSV export
    endpoint and let pandas read it directly.
    """
    url = (url or "").strip()
    if not url:
        raise ValueError("Empty Google Sheets URL.")
    csv_url = url
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if m:
        sheet_id = m.group(1)
        gid_match = re.search(r"[?&#]gid=([0-9]+)", url)
        gid = gid_match.group(1) if gid_match else "0"
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    return pd.read_csv(csv_url)


def smart_chart_recommendations(df: pd.DataFrame, max_recs: int = 8) -> pd.DataFrame:
    """Rule-based 'auto chart selection' similar to commercial AI dashboard tools,
    but fully deterministic. Suggests the most useful chart for each column pairing.
    """
    recs = []
    nums = numeric_cols(df)
    cats = [c for c in text_cols(df) if 1 < df[c].nunique(dropna=True) <= 50]
    dates = [c for c in date_like_cols(df) if pd.api.types.is_datetime64_any_dtype(pd.to_datetime(df[c], errors="coerce"))]
    # Time series: date + metric -> line
    for d in dates[:2]:
        for n in nums[:3]:
            recs.append(["Line", d, n, "Trend over time", f"{n} changes across {d}; ideal for spotting growth, decline and seasonality."])
    # Category + metric -> bar
    for c in cats[:3]:
        for n in nums[:2]:
            recs.append(["Bar", c, n, "Category comparison", f"Compare total {n} across {c} to find top and bottom performers."])
    # Single category with few levels -> pie
    for c in cats:
        if df[c].nunique(dropna=True) <= 8:
            recs.append(["Pie", c, nums[0] if nums else "Count", "Composition", f"Share of each {c} value in the whole."])
            break
    # Two numerics -> scatter
    if len(nums) >= 2:
        recs.append(["Scatter", nums[0], nums[1], "Relationship", f"Check whether {nums[0]} and {nums[1]} move together (correlation)."])
    # Single numeric -> histogram
    if nums:
        recs.append(["Histogram", nums[0], "", "Distribution", f"See how {nums[0]} is distributed and whether it is skewed."])
    if not recs:
        return pd.DataFrame({"Message": ["Not enough structure detected for automatic chart suggestions."]})
    out = pd.DataFrame(recs, columns=["Recommended Chart", "X / Dimension", "Y / Measure", "Purpose", "Why this chart"])
    return out.head(max_recs)


def build_auto_insights(df: pd.DataFrame, metric_col: Optional[str] = None, category_col: Optional[str] = None, date_col: Optional[str] = None) -> List[Dict[str, str]]:
    """Generate punchy, rule-based 'insight cards' for the in-app dashboard.

    Each insight is a dict with: title, body, tone (good/warn/info).
    100% deterministic — no AI model is used.
    """
    insights: List[Dict[str, str]] = []
    n_rows, n_cols = df.shape
    insights.append({"title": "Dataset size", "body": f"{n_rows:,} rows × {n_cols:,} columns are loaded and ready for analysis.", "tone": "info"})

    total_cells = n_rows * n_cols
    missing = int(df.isna().sum().sum())
    miss_pct = (missing / total_cells * 100) if total_cells else 0
    insights.append({
        "title": "Completeness",
        "body": f"{missing:,} missing cells ({miss_pct:.1f}% of all cells). " + ("Data is highly complete." if miss_pct < 2 else "Consider cleaning missing values before reporting."),
        "tone": "good" if miss_pct < 2 else "warn",
    })

    dupes = int(df.duplicated().sum())
    insights.append({
        "title": "Duplicates",
        "body": f"{dupes:,} fully duplicate rows detected. " + ("No duplicate rows — clean." if dupes == 0 else "Review whether these are genuine repeats."),
        "tone": "good" if dupes == 0 else "warn",
    })

    if metric_col and metric_col in df.columns:
        m = pd.to_numeric(df[metric_col], errors="coerce")
        if m.notna().any():
            insights.append({"title": f"Total {metric_col}", "body": f"Sum = {m.sum(skipna=True):,.2f}; average = {m.mean(skipna=True):,.2f}; max = {m.max(skipna=True):,.2f}.", "tone": "info"})
            if category_col and category_col in df.columns:
                grp = df.assign(__m=m).groupby(category_col, dropna=False)["__m"].sum().sort_values(ascending=False)
                if not grp.empty and grp.sum():
                    share = grp.iloc[0] / grp.sum() * 100
                    insights.append({
                        "title": "Top contributor",
                        "body": f"'{grp.index[0]}' drives {share:.1f}% of total {metric_col} ({grp.iloc[0]:,.2f}). " + ("Healthy spread." if share < 40 else "Performance is concentrated — diversification risk."),
                        "tone": "good" if share < 40 else "warn",
                    })
            # Outlier signal
            q1, q3 = m.quantile(0.25), m.quantile(0.75)
            iqr = q3 - q1
            outliers = int(((m < q1 - 1.5 * iqr) | (m > q3 + 1.5 * iqr)).sum())
            if outliers:
                insights.append({"title": f"Outliers in {metric_col}", "body": f"{outliers:,} values fall outside the normal IQR range. Inspect them before averaging or forecasting.", "tone": "warn"})

    if date_col and date_col in df.columns and metric_col and metric_col in df.columns:
        t = df[[date_col, metric_col]].copy()
        t[date_col] = pd.to_datetime(t[date_col], errors="coerce")
        t[metric_col] = pd.to_numeric(t[metric_col], errors="coerce")
        t = t.dropna()
        if len(t) > 1:
            monthly = t.assign(P=t[date_col].dt.to_period("M").astype(str)).groupby("P")[metric_col].sum().sort_index()
            if len(monthly) >= 2 and monthly.iloc[-2]:
                pct = (monthly.iloc[-1] - monthly.iloc[-2]) / monthly.iloc[-2] * 100
                insights.append({
                    "title": "Latest trend",
                    "body": f"Most recent period ({monthly.index[-1]}) moved {pct:+.1f}% vs the prior period.",
                    "tone": "good" if pct >= 0 else "warn",
                })
    return insights


def build_html_eda_report(df: pd.DataFrame, title: str = "Data Report", insights: Optional[List[Dict[str, str]]] = None) -> str:
    """Build a fully self-contained, offline HTML EDA report (Sweetviz-style) using
    only inline CSS — no external assets, no AI. Safe to open in any browser.
    """
    insights = insights or []
    profile = build_column_profile(df)
    n_rows, n_cols = df.shape
    total_cells = n_rows * n_cols
    missing = int(df.isna().sum().sum())
    dupes = int(df.duplicated().sum())

    def esc(x):
        return (str(x).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    kpis = [
        ("Rows", f"{n_rows:,}"),
        ("Columns", f"{n_cols:,}"),
        ("Missing cells", f"{missing:,}"),
        ("Missing %", f"{(missing/total_cells*100 if total_cells else 0):.2f}%"),
        ("Duplicate rows", f"{dupes:,}"),
        ("Numeric cols", f"{len(numeric_cols(df)):,}"),
    ]
    kpi_html = "".join(
        f"<div class='kpi'><div class='kpi-v'>{esc(v)}</div><div class='kpi-k'>{esc(k)}</div></div>"
        for k, v in kpis
    )

    tone_color = {"good": "#16a34a", "warn": "#d97706", "info": "#0B3954"}
    insight_html = "".join(
        f"<div class='card' style='border-left:5px solid {tone_color.get(i.get('tone','info'),'#0B3954')}'>"
        f"<b>{esc(i.get('title',''))}</b><p>{esc(i.get('body',''))}</p></div>"
        for i in insights
    ) or "<p class='muted'>No automated insights were generated.</p>"

    # profile table
    prof_cols = [c for c in ["Column", "Data Type", "Missing Count", "Missing %", "Unique Values", "Sample Values"] if c in profile.columns]
    head = "".join(f"<th>{esc(c)}</th>" for c in prof_cols)
    body = ""
    for _, r in profile.iterrows():
        body += "<tr>" + "".join(f"<td>{esc(r[c])}</td>" for c in prof_cols) + "</tr>"

    # data preview (first 50 rows)
    prev = df.head(50)
    p_head = "".join(f"<th>{esc(c)}</th>" for c in prev.columns)
    p_body = ""
    for _, r in prev.iterrows():
        p_body += "<tr>" + "".join(f"<td>{esc(v)}</td>" for v in r.values) + "</tr>"

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)} — HMG Data Report</title>
<style>
* {{ box-sizing: border-box; }}
body {{ font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin:0; background:#F4F6F7; color:#0F172A; }}
.hero {{ background: linear-gradient(135deg,#001524 0%,#0B3954 65%,#1A5F7A 100%); color:#fff; padding:28px 32px; border-left:8px solid #D4AF37; }}
.hero h1 {{ margin:0; font-size:1.8rem; }}
.hero p {{ margin:6px 0 0; color:#CFFAFE; }}
.wrap {{ max-width:1100px; margin:0 auto; padding:24px 32px; }}
.kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:14px; margin:18px 0; }}
.kpi {{ background:#fff; border:1px solid #E2E8F0; border-radius:14px; padding:16px; text-align:center; box-shadow:0 1px 4px rgba(15,23,42,.06); }}
.kpi-v {{ font-size:1.6rem; font-weight:700; color:#003366; }}
.kpi-k {{ font-size:.8rem; color:#64748B; text-transform:uppercase; letter-spacing:.04em; margin-top:4px; }}
h2 {{ color:#001524; border-bottom:2px solid #D4AF37; padding-bottom:6px; margin-top:34px; }}
.card {{ background:#fff; border:1px solid #E2E8F0; border-radius:12px; padding:14px 16px; margin:10px 0; }}
.card p {{ margin:6px 0 0; color:#334155; }}
table {{ width:100%; border-collapse:collapse; background:#fff; font-size:.85rem; }}
th, td {{ border:1px solid #E2E8F0; padding:7px 9px; text-align:left; }}
th {{ background:#001524; color:#fff; position:sticky; top:0; }}
tr:nth-child(even) td {{ background:#F8FAFC; }}
.scroll {{ overflow:auto; max-height:520px; border:1px solid #E2E8F0; border-radius:12px; }}
.muted {{ color:#64748B; }}
footer {{ text-align:center; color:#64748B; font-size:.8rem; padding:24px; }}
</style></head>
<body>
<div class="hero"><h1>📊 {esc(title)}</h1>
<p>HMG Enterprise Excel Operations Platform · Offline rule-based data report · No AI API · Generated {esc(_now())}</p></div>
<div class="wrap">
<div class="kpis">{kpi_html}</div>
<h2>Automated Insights</h2>{insight_html}
<h2>Column Profile</h2>
<div class="scroll"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>
<h2>Data Preview (first 50 rows)</h2>
<div class="scroll"><table><thead><tr>{p_head}</tr></thead><tbody>{p_body}</tbody></table></div>
</div>
<footer>Built under HMG Concepts — His Marvellous Grace Educational Consult · Developed by Adewale Samson Adeagbo (CSS Adewale)</footer>
</body></html>"""


def compare_two_datasets(df_a: pd.DataFrame, df_b: pd.DataFrame, label_a: str = "Current", label_b: str = "Comparison") -> pd.DataFrame:
    """Sweetviz-style side-by-side comparison of two datasets at column level."""
    rows = []
    all_cols = list(dict.fromkeys(list(df_a.columns) + list(df_b.columns)))
    for col in all_cols:
        in_a = col in df_a.columns
        in_b = col in df_b.columns
        a_missing = round(df_a[col].isna().mean() * 100, 2) if in_a else None
        b_missing = round(df_b[col].isna().mean() * 100, 2) if in_b else None
        a_mean = pd.to_numeric(df_a[col], errors="coerce").mean() if in_a else None
        b_mean = pd.to_numeric(df_b[col], errors="coerce").mean() if in_b else None
        status = "In both" if in_a and in_b else (f"Only in {label_a}" if in_a else f"Only in {label_b}")
        rows.append({
            "Column": col,
            "Status": status,
            f"{label_a} Type": str(df_a[col].dtype) if in_a else "—",
            f"{label_b} Type": str(df_b[col].dtype) if in_b else "—",
            f"{label_a} Missing %": a_missing if a_missing is not None else "—",
            f"{label_b} Missing %": b_missing if b_missing is not None else "—",
            f"{label_a} Mean": round(a_mean, 3) if a_mean is not None and pd.notna(a_mean) else "—",
            f"{label_b} Mean": round(b_mean, 3) if b_mean is not None and pd.notna(b_mean) else "—",
        })
    summary = pd.DataFrame([{
        "Column": "▶ DATASET TOTALS",
        "Status": "Summary",
        f"{label_a} Type": f"{df_a.shape[0]:,} rows",
        f"{label_b} Type": f"{df_b.shape[0]:,} rows",
        f"{label_a} Missing %": round(df_a.isna().mean().mean() * 100, 2),
        f"{label_b} Missing %": round(df_b.isna().mean().mean() * 100, 2),
        f"{label_a} Mean": f"{df_a.shape[1]} cols",
        f"{label_b} Mean": f"{df_b.shape[1]} cols",
    }])
    return pd.concat([summary, pd.DataFrame(rows)], ignore_index=True)


# ============================================================
# v4 ENHANCEMENTS — more free, no-AI enterprise analytics
# ============================================================

def build_benford_analysis(df: pd.DataFrame, column: str) -> Tuple[pd.DataFrame, str]:
    """Benford's Law first-digit test — a classic, free fraud/audit screen.
    Compares the observed distribution of leading digits to Benford's expected
    distribution and flags large deviations. No AI."""
    if column not in df.columns:
        return pd.DataFrame(), "Column not found."
    vals = pd.to_numeric(df[column], errors="coerce").dropna()
    vals = vals[vals > 0]
    if len(vals) < 30:
        return pd.DataFrame(), "Benford's Law needs at least ~30 positive numeric values to be meaningful."
    lead = vals.astype(float).map(lambda x: int(str(x).lstrip("0.").replace(".", "")[0]) if str(x).lstrip("0.").replace(".", "") else 0)
    lead = lead[(lead >= 1) & (lead <= 9)]
    counts = lead.value_counts().reindex(range(1, 10), fill_value=0).sort_index()
    total = counts.sum()
    observed_pct = (counts / total * 100) if total else counts * 0
    expected_pct = pd.Series([np.log10(1 + 1 / d) * 100 for d in range(1, 10)], index=range(1, 10))
    out = pd.DataFrame({
        "Leading Digit": range(1, 10),
        "Observed Count": counts.values,
        "Observed %": observed_pct.round(2).values,
        "Benford Expected %": expected_pct.round(2).values,
        "Deviation (pts)": (observed_pct - expected_pct).round(2).values,
    })
    max_dev = out["Deviation (pts)"].abs().max()
    verdict = (
        "Close to Benford — no first-digit anomaly detected." if max_dev < 5 else
        "Moderate deviation — review the data source." if max_dev < 12 else
        "Large deviation from Benford — investigate for manipulation, rounding, or capped values."
    )
    return out, f"Max deviation: {max_dev:.1f} points. {verdict}"


def build_correlation_insights(df: pd.DataFrame, threshold: float = 0.6) -> pd.DataFrame:
    """Auto-surface the strongest numeric relationships with plain-English notes. No AI."""
    nums = numeric_cols(df)
    if len(nums) < 2:
        return pd.DataFrame({"Message": ["At least two numeric columns are required."]})
    corr = df[nums].corr(numeric_only=True)
    rows = []
    seen = set()
    for a in nums:
        for b in nums:
            if a == b or (b, a) in seen:
                continue
            seen.add((a, b))
            r = corr.loc[a, b]
            if pd.isna(r):
                continue
            if abs(r) >= threshold:
                strength = "very strong" if abs(r) >= 0.85 else "strong" if abs(r) >= 0.6 else "moderate"
                direction = "positive (move together)" if r > 0 else "negative (move oppositely)"
                rows.append({
                    "Column A": a, "Column B": b, "Correlation": round(r, 3),
                    "Strength": strength, "Direction": direction,
                    "Note": f"{a} and {b} show a {strength} {direction} relationship.",
                })
    if not rows:
        return pd.DataFrame({"Message": [f"No correlations at or above |{threshold}| were found."]})
    return pd.DataFrame(rows).sort_values("Correlation", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)


def build_data_quality_rules(df: pd.DataFrame, rules: List[Dict]) -> pd.DataFrame:
    """Batch data-quality rules engine. Each rule: {column, check, param}.
    Checks: 'not_null', 'unique', 'min', 'max', 'allowed', 'regex', 'positive'.
    Returns a pass/fail scorecard with failing-row counts. No AI."""
    results = []
    n = len(df)
    for rule in rules:
        col = rule.get("column")
        check = rule.get("check")
        param = rule.get("param", "")
        if col not in df.columns:
            results.append({"Column": col, "Rule": check, "Parameter": param, "Status": "ERROR", "Failing Rows": n, "Pass %": 0.0, "Detail": "Column not found."})
            continue
        s = df[col]
        fail_mask = pd.Series(False, index=df.index)
        detail = ""
        if check == "not_null":
            fail_mask = s.isna() | (s.astype("string").str.strip() == "")
            detail = "Values must not be blank."
        elif check == "unique":
            fail_mask = s.duplicated(keep=False) & s.notna()
            detail = "Values must be unique."
        elif check == "positive":
            v = pd.to_numeric(s, errors="coerce")
            fail_mask = v.isna() | (v <= 0)
            detail = "Values must be > 0."
        elif check == "min":
            v = pd.to_numeric(s, errors="coerce")
            fail_mask = v.isna() | (v < float(param or 0))
            detail = f"Values must be ≥ {param}."
        elif check == "max":
            v = pd.to_numeric(s, errors="coerce")
            fail_mask = v.isna() | (v > float(param or 0))
            detail = f"Values must be ≤ {param}."
        elif check == "allowed":
            allowed = [x.strip() for x in str(param).split(",") if x.strip()]
            fail_mask = ~s.astype("string").isin(allowed)
            detail = f"Values must be one of: {allowed}."
        elif check == "regex":
            try:
                pat = re.compile(str(param))
                fail_mask = ~s.astype("string").fillna("").map(lambda x: bool(pat.fullmatch(x)))
                detail = f"Values must match pattern: {param}"
            except Exception as exc:
                detail = f"Invalid regex: {exc}"
        fails = int(fail_mask.sum())
        pass_pct = round((n - fails) / n * 100, 2) if n else 0.0
        results.append({
            "Column": col, "Rule": check, "Parameter": param,
            "Status": "PASS" if fails == 0 else "FAIL",
            "Failing Rows": fails, "Pass %": pass_pct, "Detail": detail,
        })
    if not results:
        return pd.DataFrame({"Message": ["No rules defined yet."]})
    return pd.DataFrame(results)


def build_trend_decomposition(df: pd.DataFrame, date_col: str, metric_col: str, freq: str = "M") -> pd.DataFrame:
    """Lightweight trend + seasonality decomposition without statsmodels.
    Uses a centered moving average for trend and average period-effect for
    seasonality, then derives the residual. No AI, no heavy dependency."""
    if date_col not in df.columns or metric_col not in df.columns:
        return pd.DataFrame()
    d = df[[date_col, metric_col]].copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce")
    d[metric_col] = pd.to_numeric(d[metric_col], errors="coerce")
    d = d.dropna()
    if d.empty:
        return pd.DataFrame()
    rfreq = normalize_resample_freq(freq)
    ts = d.set_index(date_col).sort_index()[metric_col].resample(rfreq).sum()
    if len(ts) < 4:
        return pd.DataFrame({"Message": ["Need at least 4 periods for decomposition."]})
    window = min(12, max(3, len(ts) // 2))
    trend = ts.rolling(window=window, center=True, min_periods=1).mean()
    detrended = ts - trend
    period_label = ts.index.to_series().dt.month if rfreq in ("ME",) else ts.index.to_series().dt.quarter if rfreq in ("QE",) else ts.index.to_series().dt.dayofweek
    seasonal_map = detrended.groupby(period_label.values).transform("mean")
    seasonal = pd.Series(seasonal_map.values, index=ts.index).fillna(0)
    residual = ts - trend - seasonal
    out = pd.DataFrame({
        "Period": ts.index.astype(str),
        "Actual": ts.values.round(2),
        "Trend": trend.values.round(2),
        "Seasonality": seasonal.values.round(2),
        "Residual": residual.values.round(2),
    })
    return out


def build_likert_summary(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """Survey/Likert analysis: per-question mean, top-2-box and bottom-2-box %.
    Works on numeric 1-5/1-7 scales. No AI."""
    rows = []
    for col in columns:
        v = pd.to_numeric(df[col], errors="coerce").dropna()
        if v.empty:
            continue
        scale_max = int(v.max()) if v.max() <= 10 else 5
        top2 = v[v >= scale_max - 1]
        bot2 = v[v <= 2]
        rows.append({
            "Question": col,
            "Responses": int(len(v)),
            "Mean": round(v.mean(), 2),
            "Median": round(v.median(), 2),
            "Top-2-Box %": round(len(top2) / len(v) * 100, 1),
            "Bottom-2-Box %": round(len(bot2) / len(v) * 100, 1),
            "Sentiment": "Positive" if v.mean() >= (scale_max / 2 + 0.5) else "Mixed" if v.mean() >= scale_max / 2 else "Negative",
        })
    if not rows:
        return pd.DataFrame({"Message": ["No numeric Likert-style columns found in the selection."]})
    return pd.DataFrame(rows)


def build_auto_clean_suggestions(df: pd.DataFrame) -> pd.DataFrame:
    """Inspect columns and suggest concrete cleaning actions (rule-based, no AI)."""
    suggestions = []
    n = len(df)
    for col in df.columns:
        s = df[col]
        miss = s.isna().mean() * 100
        if miss > 0:
            sev = "High" if miss > 30 else "Medium" if miss > 5 else "Low"
            suggestions.append({"Column": col, "Issue": f"{miss:.1f}% missing", "Severity": sev, "Suggested Action": "Fill (mean/median/mode) or drop rows/column."})
        # whitespace / casing in text
        if s.dtype == object or str(s.dtype).startswith("string"):
            sample = s.dropna().astype(str).head(500)
            if len(sample) and (sample != sample.str.strip()).any():
                suggestions.append({"Column": col, "Issue": "Leading/trailing spaces", "Severity": "Low", "Suggested Action": "Trim text spaces."})
            if len(sample) and sample.nunique() != sample.str.lower().nunique():
                suggestions.append({"Column": col, "Issue": "Inconsistent casing creates duplicate categories", "Severity": "Medium", "Suggested Action": "Standardize case (e.g. Title/UPPER)."})
            # numeric stored as text
            coerced = pd.to_numeric(sample, errors="coerce")
            if len(sample) and coerced.notna().mean() > 0.8:
                suggestions.append({"Column": col, "Issue": "Numbers stored as text", "Severity": "Medium", "Suggested Action": "Convert column type to Number."})
        # constant column
        if s.nunique(dropna=True) <= 1 and n > 1:
            suggestions.append({"Column": col, "Issue": "Constant / single value", "Severity": "Low", "Suggested Action": "Consider dropping (no analytical value)."})
    if not suggestions:
        return pd.DataFrame({"Result": ["No obvious cleaning issues detected — data looks tidy."]})
    return pd.DataFrame(suggestions)


# ============================================================
# v5 ENHANCEMENTS — usage analytics + more free, no-AI tools
# ============================================================

def _usage_db_path() -> str:
    """Best-effort persistent path for the usage/event log SQLite file."""
    return _get_secret("USAGE_DB_PATH", "hmg_usage.db")


def _usage_db_connect():
    conn = sqlite3.connect(_usage_db_path())
    conn.execute(
        "CREATE TABLE IF NOT EXISTS events ("
        "ts TEXT, event TEXT, tier TEXT, model TEXT, customer TEXT, detail TEXT)"
    )
    return conn


def log_usage_event(event: str, tier: str = "", model: str = "", customer: str = "", detail: str = "") -> None:
    """Record a usage/license event to a local SQLite log (best-effort, no external service).
    Used by the owner-only Usage Analytics dashboard. Fails silently if storage is read-only."""
    try:
        conn = _usage_db_connect()
        conn.execute(
            "INSERT INTO events (ts, event, tier, model, customer, detail) VALUES (?,?,?,?,?,?)",
            (_now(), str(event)[:60], str(tier)[:20], str(model)[:20], str(customer)[:120], str(detail)[:300]),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def load_usage_events() -> pd.DataFrame:
    """Load all recorded usage events as a DataFrame (empty if none/unavailable)."""
    try:
        conn = _usage_db_connect()
        df = pd.read_sql_query("SELECT * FROM events ORDER BY ts DESC", conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame(columns=["ts", "event", "tier", "model", "customer", "detail"])


def clear_usage_events() -> None:
    try:
        conn = _usage_db_connect()
        conn.execute("DELETE FROM events")
        conn.commit()
        conn.close()
    except Exception:
        pass


def build_usage_summary(events: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Summarise the usage/event log into owner-friendly tables for the dashboard."""
    out = {}
    if events is None or events.empty:
        return out
    ev = events.copy()
    ev["date"] = pd.to_datetime(ev["ts"], errors="coerce").dt.date
    out["by_event"] = ev.groupby("event").size().reset_index(name="Count").sort_values("Count", ascending=False)
    out["by_tier"] = ev[ev["tier"] != ""].groupby("tier").size().reset_index(name="Count").sort_values("Count", ascending=False)
    out["by_day"] = ev.dropna(subset=["date"]).groupby("date").size().reset_index(name="Events")
    lic = ev[ev["event"].isin(["License issued", "License activation", "License upgrade"])]
    if not lic.empty:
        out["by_customer"] = lic[lic["customer"] != ""].groupby(["customer", "tier", "model"]).size().reset_index(name="Activations").sort_values("Activations", ascending=False)
        out["by_model"] = lic[lic["model"] != ""].groupby("model").size().reset_index(name="Count")
    return out


def build_crosstab_chisquare(df: pd.DataFrame, row_col: str, col_col: str) -> Tuple[pd.DataFrame, str]:
    """Contingency cross-tab + chi-square test of independence (pure numpy, no AI/scipy).
    Returns the observed table and a verdict on whether the two categoricals are related."""
    if row_col not in df.columns or col_col not in df.columns:
        return pd.DataFrame(), "Column(s) not found."
    work = df[[row_col, col_col]].dropna()
    if work.empty:
        return pd.DataFrame(), "No non-missing rows for the selected columns."
    observed = pd.crosstab(work[row_col], work[col_col])
    if observed.shape[0] < 2 or observed.shape[1] < 2:
        return observed, "Need at least 2 categories in each column for a chi-square test."
    obs = observed.to_numpy(dtype=float)
    total = obs.sum()
    row_tot = obs.sum(axis=1, keepdims=True)
    col_tot = obs.sum(axis=0, keepdims=True)
    expected = row_tot @ col_tot / total
    with np.errstate(divide="ignore", invalid="ignore"):
        chi2 = float(np.nansum((obs - expected) ** 2 / np.where(expected == 0, np.nan, expected)))
    dof = (obs.shape[0] - 1) * (obs.shape[1] - 1)
    # Cramér's V effect size (interpretable without a p-value table).
    n = total
    min_dim = min(obs.shape[0] - 1, obs.shape[1] - 1)
    cramers_v = float(np.sqrt(chi2 / (n * min_dim))) if n and min_dim else 0.0
    strength = ("negligible" if cramers_v < 0.1 else "weak" if cramers_v < 0.3 else
                "moderate" if cramers_v < 0.5 else "strong")
    verdict = (f"Chi-square = {chi2:.2f} (dof = {dof}); Cramér's V = {cramers_v:.3f} → "
               f"{strength} association between '{row_col}' and '{col_col}'.")
    out = observed.copy()
    out["__Row Total"] = observed.sum(axis=1)
    out.loc["__Column Total"] = out.sum(axis=0)
    return out.reset_index(), verdict


_STOPWORDS = set("""a an the and or but if then else of to in on for with without at by from up down is are was were be been being this that these those it its as into than too very can will just not no nor only own same so s t will""".split())


def build_text_frequency(df: pd.DataFrame, column: str, top_n: int = 30, min_len: int = 3) -> pd.DataFrame:
    """Word-frequency / keyword counter for free-text columns (e.g. survey open-ends).
    Pure counting with a small stopword list — no AI, no NLP library."""
    if column not in df.columns:
        return pd.DataFrame({"Message": ["Column not found."]})
    texts = df[column].dropna().astype(str)
    if texts.empty:
        return pd.DataFrame({"Message": ["No text values to analyse."]})
    counts: Dict[str, int] = {}
    for t in texts:
        for w in re.findall(r"[A-Za-z']+", t.lower()):
            if len(w) >= min_len and w not in _STOPWORDS:
                counts[w] = counts.get(w, 0) + 1
    if not counts:
        return pd.DataFrame({"Message": ["No keywords found after filtering."]})
    out = pd.DataFrame(sorted(counts.items(), key=lambda x: x[1], reverse=True)[:top_n], columns=["Keyword", "Count"])
    total = sum(counts.values())
    out["Share %"] = (out["Count"] / total * 100).round(2)
    return out


CLASSIFICATION_LEVELS = ["Public", "Internal", "Confidential", "Restricted"]


def normalize_number_locale(df: pd.DataFrame, columns: List[str], decimal_sep: str = ".", thousands_sep: str = ",") -> Tuple[pd.DataFrame, int]:
    """Convert locale-formatted number text (e.g. '1.234,56' or '1,234.56') into clean numbers.
    No AI — deterministic parsing. Returns (new_df, converted_cell_count)."""
    out = df.copy()
    converted = 0
    for col in columns:
        if col not in out.columns:
            continue
        s = out[col].astype("string")

        def _parse(v):
            nonlocal converted
            if v is None or str(v).strip() == "" or str(v).lower() in ("nan", "none", "<na>"):
                return np.nan
            raw = str(v).strip()
            cleaned = raw
            if thousands_sep:
                cleaned = cleaned.replace(thousands_sep, "")
            if decimal_sep and decimal_sep != ".":
                cleaned = cleaned.replace(decimal_sep, ".")
            cleaned = re.sub(r"[^0-9.\-]", "", cleaned)
            try:
                num = float(cleaned)
                converted += 1
                return num
            except Exception:
                return np.nan

        out[col] = s.map(_parse)
    return out, converted


# ============================================================
# v6 ENHANCEMENTS — white-label/multi-tenant + report delivery + tools
# ============================================================

# Default (HMG) branding. White-label overrides come from secrets/session.
DEFAULT_BRAND = {
    "brand_name": "HMG Excel Operations Platform",
    "tagline": "An HMG Academy product · Built by HMG Technologies · Zero AI API",
    "logo_url": "",
    "logo_emoji": "📊",
    "header_bg": "#001524",
    "header_mid": "#0B3954",
    "header_to": "#1A5F7A",
    "accent": "#D4AF37",
    "footer": "Built under HMG Concepts — His Marvellous Grace · Developed by Adewale Samson Adeagbo (CSS Adewale)",
    "primary_color": "#D4AF37",
    "support_contact": "+234 810 086 6322",
    "website": "https://hmgtechnologies.pages.dev",
}


def _parse_brand_json(text: str) -> Dict[str, str]:
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            return {str(k): str(v) for k, v in data.items()}
    except Exception:
        pass
    return {}


def get_active_brand() -> Dict[str, str]:
    """Resolve the active white-label branding.

    Priority: per-session override (set in the Branding tab) > WHITE_LABEL_BRAND
    secret (JSON) > HMG default. This lets you ship the same code to many clients
    and re-skin per tenant without touching the source. No AI, no external service.
    """
    brand = dict(DEFAULT_BRAND)
    secret_brand = _parse_brand_json(_get_secret("WHITE_LABEL_BRAND", ""))
    brand.update({k: v for k, v in secret_brand.items() if v})
    sess = st.session_state.get("brand_override", {})
    if isinstance(sess, dict):
        brand.update({k: v for k, v in sess.items() if v})
    return brand


def brand_css(brand: Dict[str, str]) -> str:
    """Generate the hero/header CSS for the active brand."""
    return f"""
    <style>
    .hero {{background: linear-gradient(135deg,{brand['header_bg']} 0%,{brand['header_mid']} 65%,{brand['header_to']} 100%);
            padding: 24px; border-radius: 18px; border-left: 6px solid {brand['accent']}; margin-bottom: 18px;}}
    .hero h1 {{color: white; margin: 0; font-size: 2.1rem;}}
    .hero p {{color: #E0F2FE; margin: 4px 0 0 0; font-size: 1rem;}}
    .mini-card {{background:#FFFFFF; border:1px solid #E2E8F0; border-radius:14px; padding:16px; min-height:105px; box-shadow:0 1px 4px rgba(15,23,42,.06);}}
    .mini-card b {{color:{brand['header_bg']};}}
    .warning-box {{background:#FFF7ED; border:1px solid #FDBA74; padding:12px; border-radius:12px;}}
    .tier-badge {{display:inline-block; padding:3px 12px; border-radius:999px; font-size:.8rem; font-weight:700; color:#fff;}}
    </style>
    """


def render_brand_hero(brand: Dict[str, str], subtitle: Optional[str] = None) -> None:
    logo = ""
    if brand.get("logo_url"):
        logo = f"<img src='{brand['logo_url']}' alt='logo' style='height:46px;vertical-align:middle;margin-right:10px;border-radius:8px;'>"
    else:
        logo = f"<span style='font-size:2rem;margin-right:8px;'>{brand.get('logo_emoji','📊')}</span>"
    sub = subtitle or brand.get("tagline", "")
    st.markdown(
        f"<div class='hero'><h1>{logo}{brand['brand_name']}</h1><p>{sub}</p></div>",
        unsafe_allow_html=True,
    )


# ---- Report delivery: free email (SMTP) + WhatsApp/mailto share links ----

def send_email_with_attachment(to_addr: str, subject: str, body: str,
                               attachment_bytes: Optional[bytes] = None,
                               attachment_name: str = "report.xlsx",
                               mime_main: str = "application", mime_sub: str = "octet-stream") -> Tuple[bool, str]:
    """Send an email via a free SMTP account (e.g. Gmail app password, Zoho, Brevo).

    Credentials come from secrets: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD,
    SMTP_FROM (optional). No paid API; uses Python's built-in smtplib. Returns (ok, msg).
    """
    host = _get_secret("SMTP_HOST", "")
    port = int(_get_secret("SMTP_PORT", "587") or 587)
    user = _get_secret("SMTP_USER", "")
    password = _get_secret("SMTP_PASSWORD", "")
    from_addr = _get_secret("SMTP_FROM", "") or user
    if not (host and user and password):
        return False, "SMTP is not configured. Set SMTP_HOST, SMTP_USER and SMTP_PASSWORD in Secrets (see DEPLOYMENT.md)."
    if not to_addr:
        return False, "No recipient email address provided."
    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = from_addr
        msg["To"] = to_addr
        msg.set_content(body)
        if attachment_bytes is not None:
            msg.add_attachment(attachment_bytes, maintype=mime_main, subtype=mime_sub, filename=attachment_name)
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=30) as server:
                server.login(user, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=30) as server:
                server.starttls()
                server.login(user, password)
                server.send_message(msg)
        return True, f"Email sent to {to_addr}."
    except Exception as exc:
        return False, f"Could not send email: {exc}"


def whatsapp_share_link(phone: str, message: str) -> str:
    """Build a free WhatsApp click-to-chat link (wa.me) with a prefilled message."""
    digits = re.sub(r"\D", "", phone or "")
    text = urllib.parse.quote(message or "")
    if digits:
        return f"https://wa.me/{digits}?text={text}"
    return f"https://wa.me/?text={text}"


def mailto_link(to_addr: str, subject: str, body: str) -> str:
    """Build a free mailto: link (opens the user's email client; no server needed)."""
    params = urllib.parse.urlencode({"subject": subject, "body": body}, quote_via=urllib.parse.quote)
    return f"mailto:{to_addr}?{params}"


# ---- More free, no-AI enterprise tools ----

def build_executive_scorecard(df: pd.DataFrame, metric_col: str, category_col: Optional[str] = None,
                              date_col: Optional[str] = None) -> pd.DataFrame:
    """A compact one-look executive KPI scorecard (totals, average, growth, concentration)."""
    rows = []
    m = pd.to_numeric(df[metric_col], errors="coerce")
    rows.append({"KPI": f"Total {metric_col}", "Value": round(m.sum(skipna=True), 2)})
    rows.append({"KPI": f"Average {metric_col}", "Value": round(m.mean(skipna=True), 2)})
    rows.append({"KPI": f"Maximum {metric_col}", "Value": round(m.max(skipna=True), 2)})
    rows.append({"KPI": "Records", "Value": int(len(df))})
    if category_col and category_col in df.columns:
        grp = df.assign(__m=m).groupby(category_col, dropna=False)["__m"].sum().sort_values(ascending=False)
        if not grp.empty and grp.sum():
            rows.append({"KPI": f"Top {category_col}", "Value": str(grp.index[0])})
            rows.append({"KPI": "Top contributor share %", "Value": round(grp.iloc[0] / grp.sum() * 100, 1)})
            rows.append({"KPI": f"Distinct {category_col}", "Value": int(df[category_col].nunique(dropna=True))})
    if date_col and date_col in df.columns:
        t = df[[date_col, metric_col]].copy()
        t[date_col] = pd.to_datetime(t[date_col], errors="coerce")
        t[metric_col] = pd.to_numeric(t[metric_col], errors="coerce")
        t = t.dropna()
        if len(t) > 1:
            monthly = t.assign(P=t[date_col].dt.to_period("M").astype(str)).groupby("P")[metric_col].sum().sort_index()
            if len(monthly) >= 2 and monthly.iloc[-2]:
                growth = (monthly.iloc[-1] - monthly.iloc[-2]) / monthly.iloc[-2] * 100
                rows.append({"KPI": "Latest period growth %", "Value": round(growth, 1)})
    return pd.DataFrame(rows)


def build_gini_concentration(df: pd.DataFrame, category_col: str, metric_col: str) -> Tuple[pd.DataFrame, str]:
    """Gini coefficient + Lorenz-style concentration of a metric across categories.
    Useful for revenue/customer concentration risk. Pure numpy, no AI."""
    if category_col not in df.columns or metric_col not in df.columns:
        return pd.DataFrame(), "Column(s) not found."
    g = df.assign(__m=pd.to_numeric(df[metric_col], errors="coerce")).groupby(category_col)["__m"].sum().dropna()
    g = g[g > 0].sort_values()
    if len(g) < 2:
        return pd.DataFrame(), "Need at least 2 positive categories."
    vals = g.to_numpy(dtype=float)
    n = len(vals)
    cum = np.cumsum(vals)
    gini = (2 * np.sum((np.arange(1, n + 1)) * vals) - (n + 1) * cum[-1]) / (n * cum[-1])
    top = g.sort_values(ascending=False)
    total = top.sum()
    out = pd.DataFrame({
        category_col: top.index,
        metric_col: top.values.round(2),
        "Share %": (top.values / total * 100).round(2),
        "Cumulative %": (np.cumsum(top.values) / total * 100).round(2),
    })
    verdict = (f"Gini = {gini:.3f}. " +
               ("Low concentration — evenly spread." if gini < 0.3 else
                "Moderate concentration." if gini < 0.5 else
                "High concentration — revenue/activity depends on a few categories (risk)."))
    return out, verdict


def build_changelog_diff(old_df: pd.DataFrame, new_df: pd.DataFrame, key_col: str) -> pd.DataFrame:
    """Row-level change log between two snapshots keyed on a column: added/removed/changed."""
    if key_col not in old_df.columns or key_col not in new_df.columns:
        return pd.DataFrame({"Message": ["Key column missing in one of the datasets."]})
    old_keys = set(old_df[key_col].astype(str))
    new_keys = set(new_df[key_col].astype(str))
    rows = []
    for k in sorted(new_keys - old_keys):
        rows.append({"Key": k, "Change": "Added"})
    for k in sorted(old_keys - new_keys):
        rows.append({"Key": k, "Change": "Removed"})
    common = old_keys & new_keys
    shared_cols = [c for c in new_df.columns if c in old_df.columns and c != key_col]
    o = old_df.drop_duplicates(subset=[key_col]).set_index(old_df[key_col].astype(str))
    nw = new_df.drop_duplicates(subset=[key_col]).set_index(new_df[key_col].astype(str))
    for k in sorted(common):
        diffs = []
        for c in shared_cols:
            try:
                ov, nvv = o.loc[k, c], nw.loc[k, c]
                if str(ov) != str(nvv):
                    diffs.append(f"{c}: {ov} → {nvv}")
            except Exception:
                continue
        if diffs:
            rows.append({"Key": k, "Change": "Changed", "Details": "; ".join(diffs[:10])})
    if not rows:
        return pd.DataFrame({"Message": ["No differences detected."]})
    return pd.DataFrame(rows)


# ============================================================
# SECURITY & SUBSCRIPTION / LICENSING LAYER (v2)
# Free, dependency-light, no external API. Designed to:
#   - prevent trivial password guessing (constant-time compare + lockout)
#   - support hashed passwords (so plaintext is never stored in secrets)
#   - issue tamper-proof, offline subscription license keys (HMAC-signed)
#     that cannot be forged or extended without the owner's secret.
# ============================================================

# Subscription tiers. Higher index = more capability. The owner/admin password
# always maps to "enterprise". Demo logins map to "free".
TIER_ORDER = ["free", "pro", "enterprise"]
TIER_LABELS = {
    "free": "Free",
    "pro": "Pro",
    "enterprise": "Enterprise",
}
# Feature gates: feature_key -> minimum tier required.
# NOTE: Every pre-existing capability stays available on at least the Free tier
# (set to "free") so nothing is removed; paid value is added on top.
FEATURE_TIER = {
    # Free tier (all original analyst workflow stays free)
    "core_workflow": "free",
    "excel_export": "free",
    "html_report": "free",
    "google_sheets_import": "free",
    "auto_insights": "free",
    "correlation_insights": "free",
    "auto_clean_suggestions": "free",
    "likert_analysis": "free",
    # Pro tier (productivity / convenience enhancements)
    "dataset_comparison": "pro",
    "split_workbook": "pro",
    "sql_query": "pro",
    "pii_masking": "pro",
    "data_quality_rules": "pro",
    "trend_decomposition": "pro",
    # Enterprise tier (governance / admin)
    "tamper_evident_audit": "enterprise",
    "license_admin": "enterprise",
    "bulk_segment_export": "enterprise",
    "benford_fraud": "enterprise",
    "usage_analytics": "enterprise",
    "crosstab_chisquare": "pro",
    "text_frequency": "free",
    "number_locale": "free",
    "data_classification": "pro",
    "white_label": "enterprise",
    "report_delivery": "pro",
    "executive_scorecard": "free",
    "gini_concentration": "pro",
    "snapshot_diff": "pro",
}


def _get_secret(name: str, default: str = "") -> str:
    """Read a value from Streamlit Secrets first, then environment variables."""
    try:
        val = st.secrets.get(name, "")
        if val:
            return str(val)
    except Exception:
        pass
    return os.getenv(name, default)


def _license_signing_key() -> bytes:
    """Secret used to sign/verify license keys. MUST be set in production via
    LICENSE_SIGNING_KEY (Streamlit Secrets / env). Falls back to a public demo key
    so the feature works out-of-the-box for testing only."""
    key = _get_secret("LICENSE_SIGNING_KEY", "HMG-DEMO-DO-NOT-USE-IN-PRODUCTION")
    return key.encode("utf-8")


def _b64e(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("utf-8").rstrip("=")


def _b64d(text: str) -> bytes:
    pad = "=" * (-len(text) % 4)
    return base64.urlsafe_b64decode(text + pad)


def _seat_fingerprint(seat_id: str) -> str:
    """Stable, non-reversible fingerprint of a seat identifier (email, staff name,
    device label, or auto device-id). Stored inside the signed payload so a seat key
    only works for the matching seat."""
    norm = re.sub(r"\s+", "", str(seat_id).strip().lower())
    return hashlib.sha256(("HMG-SEAT::" + norm).encode("utf-8")).hexdigest()[:24]


def get_device_id() -> str:
    """Read the auto device-id captured by the browser fingerprint component.

    The component (render_device_fingerprint) computes a stable per-browser id and
    writes it into the `dev` URL query parameter. We read it back here. This lets
    per-seat licences bind to a device automatically, with no typing.
    Returns '' if not yet captured.
    """
    try:
        qp = st.query_params
        val = qp.get("dev", "")
        if isinstance(val, list):
            val = val[0] if val else ""
        return str(val or "")
    except Exception:
        return ""


def render_device_fingerprint() -> None:
    """Inject a tiny, self-contained JS component (no external library, no API) that
    computes a stable device fingerprint and, if not already present, reloads with it
    added to the URL as ?dev=... so Streamlit can read it via get_device_id()."""
    components.html(
        """
        <script>
        (function(){
          try {
            function fnv1a(str){
              var h = 0x811c9dc5;
              for (var i=0;i<str.length;i++){
                h ^= str.charCodeAt(i);
                h = (h + ((h<<1)+(h<<4)+(h<<7)+(h<<8)+(h<<24))) >>> 0;
              }
              return ("0000000"+(h>>>0).toString(16)).slice(-8);
            }
            // Stable signals that persist across sessions for the same browser/device.
            var nav = window.navigator;
            var parts = [
              nav.userAgent, nav.language, (nav.languages||[]).join(","),
              nav.platform, (nav.hardwareConcurrency||0), (nav.deviceMemory||0),
              screen.width, screen.height, screen.colorDepth,
              (new Date()).getTimezoneOffset(),
              Intl.DateTimeFormat().resolvedOptions().timeZone || ""
            ];
            // Canvas signal (cheap, stable per GPU/driver).
            try {
              var c = document.createElement('canvas');
              var ctx = c.getContext('2d');
              ctx.textBaseline = "top";
              ctx.font = "14px Arial";
              ctx.fillStyle = "#069";
              ctx.fillText("HMG-device-fp", 2, 2);
              parts.push(c.toDataURL());
            } catch(e){}
            // Persist a random salt in localStorage so the id is also sticky per-browser.
            var salt = localStorage.getItem("hmg_dev_salt");
            if(!salt){ salt = Math.random().toString(36).slice(2) + Date.now().toString(36);
                       localStorage.setItem("hmg_dev_salt", salt); }
            var dev = "DEV-" + fnv1a(parts.join("|")) + fnv1a(salt);
            var url = new URL(window.parent.location.href);
            if(url.searchParams.get("dev") !== dev){
              url.searchParams.set("dev", dev);
              window.parent.history.replaceState(null, "", url.toString());
              // Trigger a rerun so Python sees the new query param.
              window.parent.location.reload();
            }
          } catch(e){ /* fail open: no device id */ }
        })();
        </script>
        <div style="font:12px sans-serif;color:#16a34a;">Device verified for per-seat licensing.</div>
        """,
        height=28,
    )


def issue_license_key(customer: str, tier: str, expiry_iso: str, model: str = "team", seat_id: str = "") -> str:
    """Create a signed, offline license key: base64(payload).base64(hmac).

    payload = {c:customer, t:tier, exp:expiry, m:model, seat:seat-fingerprint, iss:issued}
    The signature is HMAC-SHA256 over the payload using the owner's secret. Without
    the secret a third party cannot forge a key, change the tier, extend the expiry,
    switch the model, or rebind the seat — this blocks subscription bypass and sharing.

    model:
      "team" -> shareable across a team / many devices (no seat binding).
      "seat" -> named/per-seat: only activates when the user supplies the matching Seat ID.
    """
    tier = tier if tier in TIER_ORDER else "pro"
    model = "seat" if str(model).lower().startswith("seat") else "team"
    payload = {
        "c": str(customer)[:120],
        "t": tier,
        "exp": str(expiry_iso),
        "m": model,
        "seat": _seat_fingerprint(seat_id) if model == "seat" else "",
        "iss": _now(),
    }
    raw = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    sig = hmac.new(_license_signing_key(), raw, hashlib.sha256).digest()
    return f"{_b64e(raw)}.{_b64e(sig)}"


def verify_license_key(key: str, seat_id_provided: str = "") -> Dict[str, str]:
    """Validate a license key. Returns a dict with keys:
    valid(bool), tier, customer, expiry, model, reason. Uses constant-time signature
    comparison, checks expiry, and (for per-seat keys) checks the supplied Seat ID."""
    result = {"valid": False, "tier": "free", "customer": "", "expiry": "", "model": "team", "reason": ""}
    if not key or "." not in key:
        result["reason"] = "Empty or malformed key."
        return result
    try:
        payload_b64, sig_b64 = key.strip().split(".", 1)
        raw = _b64d(payload_b64)
        sig = _b64d(sig_b64)
    except Exception:
        result["reason"] = "Key could not be decoded."
        return result
    expected = hmac.new(_license_signing_key(), raw, hashlib.sha256).digest()
    if not hmac.compare_digest(sig, expected):
        result["reason"] = "Signature invalid — key was tampered with or signed with a different secret."
        return result
    try:
        payload = json.loads(raw.decode("utf-8"))
    except Exception:
        result["reason"] = "Payload unreadable."
        return result
    tier = payload.get("t", "free")
    expiry = str(payload.get("exp", ""))
    customer = str(payload.get("c", ""))
    model = payload.get("m", "team") or "team"
    result.update({"customer": customer, "expiry": expiry, "model": model})
    # Expiry check (date or datetime).
    try:
        exp_clean = expiry.replace("Z", "").strip()
        if len(exp_clean) <= 10:
            exp_dt = datetime.strptime(exp_clean, "%Y-%m-%d")
        else:
            exp_dt = datetime.fromisoformat(exp_clean)
        if datetime.now() > exp_dt:
            result.update({"reason": f"License expired on {expiry}.", "tier": "free"})
            return result
    except Exception:
        result["reason"] = "Expiry date unreadable."
        return result
    # Per-seat binding check.
    if model == "seat":
        if not seat_id_provided:
            result["reason"] = "This is a per-seat license. Enter the Seat ID it was issued to."
            return result
        if not hmac.compare_digest(_seat_fingerprint(seat_id_provided), str(payload.get("seat", ""))):
            result["reason"] = "Seat ID does not match this per-seat license."
            return result
    result.update({"valid": True, "tier": tier if tier in TIER_ORDER else "pro", "reason": "Valid license."})
    return result


def tier_rank(tier: str) -> int:
    return TIER_ORDER.index(tier) if tier in TIER_ORDER else 0


def current_tier() -> str:
    return st.session_state.get("subscription_tier", "free")


def has_feature(feature_key: str) -> bool:
    """True if the active session's tier meets the feature's minimum tier."""
    required = FEATURE_TIER.get(feature_key, "free")
    return tier_rank(current_tier()) >= tier_rank(required)


def require_feature(feature_key: str, friendly_name: str) -> bool:
    """UI helper: returns True if allowed, else renders an upgrade notice and returns False."""
    if has_feature(feature_key):
        return True
    required = TIER_LABELS.get(FEATURE_TIER.get(feature_key, "free"), "Pro")
    st.warning(
        f"🔒 **{friendly_name}** is a **{required}** feature. "
        f"Your current plan is **{TIER_LABELS.get(current_tier(),'Free')}**. "
        "Enter a valid license key in the sidebar (or contact HMG Technologies) to unlock it."
    )
    return False


def _hash_password(plain: str, salt: str = "HMG-STATIC-SALT-v2") -> str:
    """SHA-256 hash so plaintext passwords need not be stored in secrets."""
    return hashlib.sha256((salt + plain).encode("utf-8")).hexdigest()


def get_password_records() -> List[Dict[str, str]]:
    """Return password records as {value, kind, tier}.

    Sources (in priority order):
      1. APP_PASSWORD_HASH  — comma-separated SHA-256 hashes (preferred; admin/enterprise)
      2. APP_PASSWORD       — comma-separated plaintext passwords (admin/enterprise)
      3. Demo defaults admin/HMG2025 — ONLY if ALLOW_DEMO_LOGIN is not "false".
         Demo logins map to the FREE tier so they cannot access paid features.
    """
    records: List[Dict[str, str]] = []
    hash_value = _get_secret("APP_PASSWORD_HASH", "")
    if hash_value:
        for h in [p.strip().lower() for p in hash_value.split(",") if p.strip()]:
            records.append({"value": h, "kind": "hash", "tier": "enterprise"})
    plain_value = _get_secret("APP_PASSWORD", "")
    if plain_value:
        for p in [x.strip() for x in plain_value.split(",") if x.strip()]:
            records.append({"value": p, "kind": "plain", "tier": "enterprise"})
    allow_demo = _get_secret("ALLOW_DEMO_LOGIN", "true").strip().lower() != "false"
    if allow_demo:
        records.append({"value": "admin", "kind": "plain", "tier": "free"})
        records.append({"value": "HMG2025", "kind": "plain", "tier": "free"})
    return records


def verify_password(candidate: str) -> Optional[str]:
    """Constant-time password verification. Returns the tier on success, else None."""
    candidate = candidate or ""
    cand_hash = _hash_password(candidate)
    matched_tier = None
    # Always iterate all records to keep timing roughly constant.
    for rec in get_password_records():
        if rec["kind"] == "hash":
            if hmac.compare_digest(cand_hash, rec["value"]):
                matched_tier = rec["tier"]
        else:
            if hmac.compare_digest(candidate, rec["value"]):
                matched_tier = rec["tier"]
    return matched_tier


# --- Login brute-force protection (per session) ---
_MAX_ATTEMPTS = 5
_LOCKOUT_SECONDS = 300  # 5 minutes


def login_locked_remaining() -> int:
    until = st.session_state.get("lockout_until", 0)
    remaining = int(until - time.time())
    return remaining if remaining > 0 else 0


def register_failed_login() -> None:
    st.session_state["login_attempts"] = st.session_state.get("login_attempts", 0) + 1
    if st.session_state["login_attempts"] >= _MAX_ATTEMPTS:
        st.session_state["lockout_until"] = time.time() + _LOCKOUT_SECONDS
        st.session_state["login_attempts"] = 0


def reset_login_attempts() -> None:
    st.session_state["login_attempts"] = 0
    st.session_state["lockout_until"] = 0


def get_allowed_passwords() -> List[str]:
    """Backwards-compatible helper retained from v1 (plaintext list only).
    New code should use verify_password()."""
    return [r["value"] for r in get_password_records() if r["kind"] == "plain"]


# ============================================================
# SEO: inject crawlable meta tags + JSON-LD structured data
# ============================================================

def seo_meta_html() -> str:
    """Return <meta>/Open Graph/Twitter/JSON-LD markup so the app is
    discoverable by search engines and rich-previewed on social platforms.
    Injected via st.markdown(unsafe_allow_html=True)."""
    title = "HMG Excel Operations Platform — Free Excel Dashboard & Data Analysis Tool (No AI API)"
    desc = (
        "Free, no-code Excel & CSV dashboard generator and data analysis platform by HMG Academy "
        "(built by HMG Technologies, an HMG Concepts subsidiary). Upload spreadsheets to clean, profile, "
        "pivot, forecast, visualize and export a 40+ sheet enterprise Excel workbook. No AI API, no recurring cost. "
        "Built in Nigeria by Adewale Samson Adeagbo."
    )
    url = _get_secret("APP_PUBLIC_URL", "https://cssadewale-exceldashboard-generator.streamlit.app/")
    keywords = (
        "Excel dashboard generator, free data analysis tool, CSV to dashboard, no-code analytics, "
        "Excel automation, data profiling, pivot table online, forecasting, RFM analysis, cohort analysis, "
        "HMG Academy, HMG Technologies, HMG Concepts, Adewale Samson Adeagbo, Nigeria EdTech, no AI API"
    )
    json_ld = {
        "@context": "https://schema.org",
        "@type": "SoftwareApplication",
        "name": "HMG Excel Operations Platform",
        "applicationCategory": "BusinessApplication",
        "operatingSystem": "Web Browser",
        "offers": {"@type": "Offer", "price": "0", "priceCurrency": "USD"},
        "description": desc,
        "url": url,
        "creator": {
            "@type": "Person",
            "name": "Adewale Samson Adeagbo",
            "jobTitle": "AI-Augmented Solutions Developer · Data Scientist · STEM Educator",
            "url": "https://cssadewale.pages.dev",
            "sameAs": [
                "https://github.com/cssadewale",
                "https://linkedin.com/in/adewalesamsonadeagbo",
                "https://x.com/cssadewale",
                "https://youtube.com/@hmgconcepts",
            ],
        },
        "publisher": {
            "@type": "Organization",
            "name": "HMG Academy (HMG Concepts)",
            "url": "https://hmgacademy.pages.dev",
            "parentOrganization": {"@type": "Organization", "name": "HMG Concepts", "url": "https://hmgconcepts.pages.dev"},
        },
    }
    return f"""
    <meta name="description" content="{desc}">
    <meta name="keywords" content="{keywords}">
    <meta name="author" content="Adewale Samson Adeagbo (CSS Adewale) — HMG Concepts">
    <meta name="robots" content="index, follow">
    <link rel="canonical" href="{url}">
    <meta property="og:type" content="website">
    <meta property="og:title" content="{title}">
    <meta property="og:description" content="{desc}">
    <meta property="og:url" content="{url}">
    <meta property="og:site_name" content="HMG Academy · HMG Technologies">
    <meta name="twitter:card" content="summary_large_image">
    <meta name="twitter:title" content="{title}">
    <meta name="twitter:description" content="{desc}">
    <meta name="twitter:creator" content="@cssadewale">
    <script type="application/ld+json">{json.dumps(json_ld)}</script>
    """


# ============================================================
# PRIVACY: rule-based PII detection & masking (no AI)
# ============================================================

_PII_PATTERNS = {
    "Email": re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"),
    "Phone": re.compile(r"(?:(?:\+?\d{1,3}[\s-]?)?(?:\(?\d{2,4}\)?[\s-]?)?\d{3,4}[\s-]?\d{3,4})"),
}
_PII_NAME_HINTS = ["email", "e-mail", "phone", "mobile", "tel", "msisdn", "ssn", "nin", "bvn",
                   "passport", "account", "acct", "card", "address", "dob", "birth"]


def detect_pii_columns(df: pd.DataFrame, sample: int = 200) -> pd.DataFrame:
    """Heuristically flag columns likely to contain PII, by name hint and value pattern."""
    rows = []
    for col in df.columns:
        name_hit = any(h in str(col).lower() for h in _PII_NAME_HINTS)
        series = df[col].dropna().astype(str).head(sample)
        pattern_kind = ""
        hit_ratio = 0.0
        if len(series):
            for kind, pat in _PII_PATTERNS.items():
                matches = series.apply(lambda v: bool(pat.fullmatch(v.strip())) or bool(pat.search(v)))
                ratio = matches.mean() if len(series) else 0
                if ratio > hit_ratio:
                    hit_ratio, pattern_kind = ratio, kind
        likely = name_hit or hit_ratio >= 0.6
        if likely:
            rows.append({
                "Column": col,
                "Reason": ("name hint" if name_hit else "") + (("; value pattern: " + pattern_kind) if pattern_kind and hit_ratio >= 0.6 else ""),
                "Pattern Match %": round(hit_ratio * 100, 1),
                "Recommendation": "Mask or remove before sharing externally.",
            })
    if not rows:
        return pd.DataFrame({"Result": ["No obvious PII columns detected by name or pattern."]})
    return pd.DataFrame(rows)


def mask_value(value, mode: str = "Partial") -> str:
    """Mask a single value. Partial keeps first/last char; Full hashes it."""
    s = "" if value is None else str(value)
    if not s or s.lower() in ("nan", "none", "<na>"):
        return s
    if mode == "Full (hash)":
        return "HMG#" + hashlib.sha256(s.encode("utf-8")).hexdigest()[:10]
    if "@" in s and "." in s:  # email-aware partial mask
        local, _, domain = s.partition("@")
        masked_local = (local[0] + "***") if local else "***"
        return f"{masked_local}@{domain}"
    if len(s) <= 2:
        return "*" * len(s)
    return s[0] + "*" * (len(s) - 2) + s[-1]


def mask_pii_columns(df: pd.DataFrame, columns: List[str], mode: str = "Partial") -> pd.DataFrame:
    out = df.copy()
    for col in columns:
        if col in out.columns:
            out[col] = out[col].apply(lambda v: mask_value(v, mode))
    return out


# ============================================================
# ENTERPRISE: tamper-evident audit hash-chain
# Each audit entry is hashed together with the previous hash, forming a chain.
# Changing or deleting any past entry breaks every subsequent hash, making
# tampering detectable. No external service — pure hashlib.
# ============================================================

def build_audit_hash_chain(audit: List[Dict]) -> pd.DataFrame:
    """Return the audit log with a per-row hash chained to the previous row."""
    rows = []
    prev_hash = "GENESIS"
    for entry in (audit or []):
        ts = str(entry.get("Timestamp", ""))
        action = str(entry.get("Action", ""))
        detail = str(entry.get("Detail", ""))
        block = f"{prev_hash}|{ts}|{action}|{detail}"
        cur_hash = hashlib.sha256(block.encode("utf-8")).hexdigest()
        rows.append({
            "Timestamp": ts,
            "Action": action,
            "Detail": detail,
            "Prev Hash": prev_hash[:16],
            "Entry Hash": cur_hash[:16],
            "Full Entry Hash": cur_hash,
        })
        prev_hash = cur_hash
    if not rows:
        return pd.DataFrame({"Message": ["No audit entries yet."]})
    return pd.DataFrame(rows)


def verify_audit_hash_chain(chain_df: pd.DataFrame) -> Dict[str, str]:
    """Recompute the chain and report whether it is intact."""
    if chain_df is None or chain_df.empty or "Full Entry Hash" not in chain_df.columns:
        return {"intact": False, "reason": "No verifiable chain present."}
    prev_hash = "GENESIS"
    for i, r in chain_df.iterrows():
        block = f"{prev_hash}|{r['Timestamp']}|{r['Action']}|{r['Detail']}"
        cur = hashlib.sha256(block.encode("utf-8")).hexdigest()
        if cur != str(r.get("Full Entry Hash", "")):
            return {"intact": False, "reason": f"Chain broken at row {int(i)+1} — entry was altered or reordered."}
        prev_hash = cur
    return {"intact": True, "reason": f"Audit chain intact across {len(chain_df)} entries."}


# ============================================================
# Saved analysis "recipe" (session config) — export/import (free)
# Lets a user save their selected columns / settings as a small JSON file
# and reload it later or share it, so a workflow is reproducible. No data is
# stored — only the configuration choices.
# ============================================================

def export_session_recipe() -> str:
    """Serialize key UI selections from session_state into a JSON recipe string."""
    keep_keys = [
        "ins_metric", "ins_cat", "ins_date", "html_report_title",
        "fc_date", "fc_metric", "fc_freq", "fc_periods", "fc_method",
        "cmp_label_a", "cmp_label_b", "ti_freq", "anomaly_method",
        "z_threshold", "pii_mask_mode",
    ]
    recipe = {k: st.session_state.get(k) for k in keep_keys if k in st.session_state}
    recipe["__meta__"] = {"app": APP_NAME, "version": APP_VERSION, "saved": _now()}
    return json.dumps(recipe, indent=2, default=str)


def apply_session_recipe(recipe_text: str) -> Tuple[bool, str]:
    """Load a recipe JSON and write its values back into session_state."""
    try:
        recipe = json.loads(recipe_text)
    except Exception as exc:
        return False, f"Could not parse recipe: {exc}"
    applied = 0
    for k, v in recipe.items():
        if k == "__meta__":
            continue
        st.session_state[k] = v
        applied += 1
    return True, f"Applied {applied} setting(s) from the recipe."


def build_enterprise_readiness_assessment(df: pd.DataFrame) -> pd.DataFrame:
    total_cells = int(df.shape[0] * df.shape[1]) if not df.empty else 0
    missing_cells = int(df.isna().sum().sum()) if total_cells else 0
    duplicate_rows = int(df.duplicated().sum()) if not df.empty else 0
    readiness = [
        ["Authentication", "Available", "Password gate enabled; supports Streamlit Secrets or environment variable APP_PASSWORD.", "Move production password to secrets before public use."],
        ["Auditability", "Available", "Workflow actions and workbook metadata are exported in Audit_Log.", "Review audit log before sharing workbook."],
        ["Data Quality", "Available", f"Missing cells: {missing_cells:,}; duplicate rows: {duplicate_rows:,}.", "Run cleaning, validation, schema contract and duplicate checks."],
        ["Excel Governance", "Available", "Raw data, processed data, profiles, validation and dashboards are separated into sheets.", "Protect workbook externally if sending sensitive data."],
        ["Cost Control", "Available", "No AI API or paid model dependency is used.", "Continue using free Streamlit/pandas/XlsxWriter stack."],
        ["Scalability", "Conditional", f"Current dataset shape: {df.shape[0]:,} rows x {df.shape[1]:,} columns.", "For very large data, sample, aggregate, filter or split before Excel export."],
        ["Privacy", "Conditional", "The app does not intentionally send data to AI APIs.", "Follow hosting provider and institutional data policies."],
    ]
    return pd.DataFrame(readiness, columns=["Enterprise Area", "Status", "Evidence", "Recommended Control"])


def build_access_control_matrix() -> pd.DataFrame:
    return pd.DataFrame([
        ["System Owner", "Full access", "Change password, deploy app, manage repository, approve releases.", "Use GitHub + Streamlit account security."],
        ["Data Analyst", "Workflow access", "Upload data, clean, analyse, export workbook.", "Use app password; avoid sharing confidential files publicly."],
        ["Reviewer / Manager", "Workbook access", "Review exported reports and dashboard.", "Share final workbook only; not raw upload access if unnecessary."],
        ["External Partner", "Limited output access", "Receive selected sheets/reports only.", "Remove Raw_Data or sensitive fields before sharing."],
        ["Public User", "No access", "Should not access institutional data workflow.", "Keep password private and rotate periodically."],
    ], columns=["Role", "Access Level", "Allowed Activities", "Control Recommendation"])


def build_data_governance_matrix() -> pd.DataFrame:
    return pd.DataFrame([
        ["Data Minimisation", "Upload only the columns required for analysis.", "Owner / Analyst", "Before ingestion"],
        ["Sensitive Data Review", "Remove or mask personally identifiable information where not needed.", "Data Owner", "Before upload/export"],
        ["Audit Trail", "Preserve Audit_Log sheet in every exported workbook.", "Analyst", "Every export"],
        ["Validation", "Run validation and schema contract checks for critical reports.", "Analyst", "Before decision-making"],
        ["Version Control", "Commit code changes to GitHub with clear messages.", "Developer", "Every release"],
        ["Password Rotation", "Change production password periodically and after staff changes.", "System Owner", "Monthly/Quarterly"],
        ["No AI API", "Do not connect paid AI/model APIs unless a separate budget and policy are approved.", "System Owner", "Continuous"],
        ["Workbook Sharing", "Share only necessary sheets with external users.", "Manager", "Every external sharing"],
    ], columns=["Control", "Description", "Responsible Role", "Timing"])


def build_enterprise_deployment_checklist() -> pd.DataFrame:
    return pd.DataFrame([
        [1, "Upload extracted files to GitHub", "Required", "app.py, requirements.txt, README.md, FEATURES.md, DEPLOYMENT.md, .streamlit/config.toml"],
        [2, "Confirm requirements.txt", "Required", "streamlit, pandas, numpy, XlsxWriter, openpyxl, plotly"],
        [3, "Deploy on Streamlit Community Cloud", "Required", "Repository main branch; main file path app.py"],
        [4, "Set production password", "Strongly recommended", "Use Streamlit Secrets: APP_PASSWORD = 'strong-password'"],
        [5, "Test demo workflow", "Required", "Load demo data, run clean/analyse/export"],
        [6, "Verify brand sheets", "Required", "HMG_Brand_Profile and HMG_Ecosystem sheets must exist"],
        [7, "Verify enterprise sheets", "Required", "Enterprise_Readiness, Access_Control, Governance_Matrix and Privacy_Notice"],
        [8, "Document release", "Recommended", "Use RELEASE_NOTES.md and GitHub commit message"],
    ], columns=["Step", "Task", "Priority", "Notes"])


def build_free_tools_register() -> pd.DataFrame:
    return pd.DataFrame([
        ["Streamlit", "Web application interface", "Free/community tier", "App may sleep on free hosting."],
        ["pandas", "Data cleaning, transformation and analysis", "Open-source", "Memory depends on hosting limits."],
        ["NumPy", "Numeric computation", "Open-source", "Used for calculations and scoring."],
        ["Plotly", "Interactive charts", "Open-source", "Browser charts only; no paid API."],
        ["XlsxWriter", "Excel workbook generation", "Open-source", "Creates native Excel formatting and charts."],
        ["openpyxl", "Excel file reading support", "Open-source", "Used by pandas for XLSX files."],
        ["sqlite3", "In-memory SQL query engine", "Built into Python", "Only SELECT/WITH queries allowed in app."],
        ["GitHub", "Version control and repository hosting", "Free tier", "Public repo recommended for Streamlit free deployment."],
    ], columns=["Tool", "Purpose", "Cost Model", "Enterprise Note"])


def build_privacy_notice() -> pd.DataFrame:
    return pd.DataFrame([
        ["No AI API", "The platform does not call OpenAI, Gemini, Claude, Grok or paid model APIs."],
        ["Session Processing", "Uploaded files are processed in the active Streamlit session using Python libraries."],
        ["No Intentional External Data Transfer", "The app does not intentionally transmit uploaded datasets to third-party AI services."],
        ["Hosting Responsibility", "If hosted on Streamlit/Render, institutional users must follow the hosting provider's privacy terms."],
        ["Sensitive Data", "Avoid uploading confidential personal data unless the organisation approves the hosting arrangement."],
        ["Workbook Sharing", "Exported Excel files may contain raw data; remove sensitive sheets before external sharing."],
    ], columns=["Privacy Area", "Statement"])


def build_enterprise_release_notes() -> pd.DataFrame:
    return pd.DataFrame([
        [APP_VERSION, "Enterprise folder release", "Consolidates the full Excel workflow feature set plus enterprise controls, governance sheets, secrets-ready password support and deployment pack."],
        ["Core", "No feature removed", "All previous data ingestion, cleaning, analysis, dashboard, SQL, anomaly, RFM, cohort, ABC, basket, scorecard and brand features are retained."],
        ["Enterprise", "New governance workbook sheets", "Enterprise_Readiness, Access_Control, Governance_Matrix, Deployment_Checklist, Free_Tools_Register, Privacy_Notice and Release_Notes."],
        ["Cost", "No AI API", "The release remains free-tools-only and avoids paid AI/model APIs."],
    ], columns=["Version/Area", "Release Item", "Explanation"])

# ============================================================
# Excel writer helpers
# ============================================================

def write_dataframe(writer, df: pd.DataFrame, sheet_name: str, startrow: int = 0, startcol: int = 0, header_format=None, heatmap: bool = True) -> None:
    sheet_name = safe_sheet_name(sheet_name)
    if df is None:
        df = pd.DataFrame()
    if df.empty:
        pd.DataFrame({"Message": ["No records available for this section."]}).to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow, startcol=startcol)
        ws = writer.sheets[sheet_name]
        ws.set_column(0, 0, 48)
        return
    export_df = df.copy()
    if len(export_df) > EXCEL_MAX_ROWS - 1:
        export_df = export_df.head(EXCEL_MAX_ROWS - 1)
    export_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow, startcol=startcol)
    workbook = writer.book
    ws = writer.sheets[sheet_name]
    if header_format is None:
        header_format = workbook.add_format({"bold": True, "bg_color": "#001524", "font_color": "white", "border": 1, "text_wrap": True})
    for i, col in enumerate(export_df.columns):
        ws.write(startrow, startcol + i, col, header_format)
        width = min(max(12, int(export_df[col].astype(str).str.len().quantile(0.9)) if len(export_df) else 12, len(str(col)) + 2), 45)
        ws.set_column(startcol + i, startcol + i, width)
        if heatmap and pd.api.types.is_numeric_dtype(export_df[col]) and len(export_df) > 0:
            col_letter = xl_col_to_name(startcol + i)
            ws.conditional_format(f"{col_letter}{startrow+2}:{col_letter}{startrow+len(export_df)+1}", {"type": "3_color_scale"})
    ws.freeze_panes(startrow + 1, startcol)
    ws.autofilter(startrow, startcol, startrow + len(export_df), startcol + len(export_df.columns) - 1)


def build_excel_workbook(
    df: pd.DataFrame,
    raw_df: pd.DataFrame,
    audit: List[Dict],
    title: str,
    theme_name: str,
    cat_col: Optional[str],
    num_col: Optional[str],
    date_col: Optional[str],
    sec_cat_col: Optional[str],
    num_col_2: Optional[str],
    top_n: int,
    grouped_summary: Optional[pd.DataFrame],
    pivot_table: Optional[pd.DataFrame],
    forecast_df: Optional[pd.DataFrame],
    validation_report: Optional[pd.DataFrame] = None,
    what_if_df: Optional[pd.DataFrame] = None,
    pareto_df: Optional[pd.DataFrame] = None,
    reconciliation_report: Optional[pd.DataFrame] = None,
    time_intelligence_df: Optional[pd.DataFrame] = None,
    kpi_target_df: Optional[pd.DataFrame] = None,
    sql_result_df: Optional[pd.DataFrame] = None,
    anomaly_report_df: Optional[pd.DataFrame] = None,
    rfm_df: Optional[pd.DataFrame] = None,
    cohort_df: Optional[pd.DataFrame] = None,
    schema_contract_df: Optional[pd.DataFrame] = None,
    fuzzy_duplicates_df: Optional[pd.DataFrame] = None,
    abc_df: Optional[pd.DataFrame] = None,
    basket_df: Optional[pd.DataFrame] = None,
    weighted_score_df: Optional[pd.DataFrame] = None,
    goal_seek_df: Optional[pd.DataFrame] = None,
    calendar_df: Optional[pd.DataFrame] = None,
    report_brief_df: Optional[pd.DataFrame] = None,
) -> bytes:
    output = io.BytesIO()
    theme = THEMES[theme_name]
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm", date_format="yyyy-mm-dd") as writer:
        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": theme["header_bg"], "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter"})
        title_format = workbook.add_format({"bold": True, "font_size": 20, "font_color": theme["header_bg"], "bottom": 2, "bottom_color": theme["accent"]})
        subtitle_format = workbook.add_format({"font_color": "#64748B", "italic": True, "font_size": 10})
        kpi_title_format = workbook.add_format({"bold": True, "font_size": 11, "font_color": "#475569", "align": "center", "bg_color": "#E2E8F0", "border": 1})
        kpi_value_format = workbook.add_format({"bold": True, "font_size": 15, "font_color": theme["kpi_val"], "align": "center", "border": 1, "bg_color": "#FFFFFF", "num_format": "#,##0.00"})
        note_format = workbook.add_format({"font_size": 10, "font_color": "#0F172A", "bg_color": theme["soft_bg"], "border": 1, "text_wrap": True, "valign": "top"})
        money_format = workbook.add_format({"num_format": "#,##0.00"})

        # 1. Audit Log
        audit_df = pd.DataFrame(audit) if audit else pd.DataFrame([{"Timestamp": _now(), "Action": "Workbook generated", "Detail": "No prior operations recorded."}])
        metadata = pd.DataFrame([
            ["Workbook Title", title],
            ["Application", APP_NAME],
            ["Version", APP_VERSION],
            ["Generated At", _now()],
            ["Raw Shape", f"{raw_df.shape[0]:,} rows x {raw_df.shape[1]:,} columns"],
            ["Processed Shape", f"{df.shape[0]:,} rows x {df.shape[1]:,} columns"],
            ["Privacy", "No AI API. No external API call. Processing happens in the Streamlit session memory."],
        ], columns=["Item", "Value"])
        write_dataframe(writer, metadata, "Audit_Log", header_format=header_format, heatmap=False)
        audit_df.to_excel(writer, sheet_name="Audit_Log", index=False, startrow=len(metadata) + 3)
        ws = writer.sheets["Audit_Log"]
        ws.write(len(metadata) + 2, 0, "Operation History", title_format)
        for i, c in enumerate(audit_df.columns):
            ws.write(len(metadata) + 3, i, c, header_format)
            ws.set_column(i, i, 24 if c == "Timestamp" else 55)

        # 2. Data profile and governance reports
        overview_df = build_overview(df, raw_df)
        col_profile_df = build_column_profile(df)
        write_dataframe(writer, overview_df, "Data_Profile", header_format=header_format, heatmap=False)
        profile_sheet = writer.sheets["Data_Profile"]
        profile_sheet.write(len(overview_df) + 3, 0, "Column-Level Profile", title_format)
        col_profile_df.to_excel(writer, sheet_name="Data_Profile", index=False, startrow=len(overview_df) + 5)
        for i, c in enumerate(col_profile_df.columns):
            profile_sheet.write(len(overview_df) + 5, i, c, header_format)
            profile_sheet.set_column(i, i, 18)

        write_dataframe(writer, build_missing_report(df), "Missing_Report", header_format=header_format)
        dupes = df[df.duplicated(keep=False)].head(5000) if not df.empty else pd.DataFrame()
        write_dataframe(writer, dupes, "Duplicate_Report", header_format=header_format)
        write_dataframe(writer, build_outlier_report(df), "Outlier_Report", header_format=header_format)

        # 3. Raw and processed data
        write_dataframe(writer, raw_df, "Raw_Data", header_format=header_format)
        write_dataframe(writer, df, "Processed_Data", header_format=header_format)

        # 4. Analysis sheets
        numeric = numeric_cols(df)
        if numeric:
            stats = df[numeric].describe().reset_index().rename(columns={"index": "Statistic"})
            write_dataframe(writer, stats, "Descriptive_Stats", header_format=header_format)
            corr = df[numeric].corr(numeric_only=True).reset_index().rename(columns={"index": "Metric"})
            write_dataframe(writer, corr, "Correlation_Matrix", header_format=header_format)
        else:
            write_dataframe(writer, pd.DataFrame(), "Descriptive_Stats", header_format=header_format)
            write_dataframe(writer, pd.DataFrame(), "Correlation_Matrix", header_format=header_format)

        write_dataframe(writer, grouped_summary if grouped_summary is not None else pd.DataFrame(), "Grouped_Summary", header_format=header_format)
        write_dataframe(writer, pivot_table if pivot_table is not None else pd.DataFrame(), "Pivot_Analysis", header_format=header_format)
        write_dataframe(writer, forecast_df if forecast_df is not None else pd.DataFrame(), "Forecast", header_format=header_format)
        write_dataframe(writer, validation_report if validation_report is not None else pd.DataFrame(), "Validation_Report", header_format=header_format)
        write_dataframe(writer, what_if_df if what_if_df is not None else pd.DataFrame(), "What_If_Analysis", header_format=header_format)
        write_dataframe(writer, pareto_df if pareto_df is not None else pd.DataFrame(), "Pareto_Analysis", header_format=header_format)
        write_dataframe(writer, build_data_dictionary(df), "Data_Dictionary", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_quality_scorecard(df), "Quality_Scorecard", header_format=header_format, heatmap=True)
        write_dataframe(writer, reconciliation_report if reconciliation_report is not None else pd.DataFrame(), "Reconciliation", header_format=header_format)
        write_dataframe(writer, time_intelligence_df if time_intelligence_df is not None else pd.DataFrame(), "Time_Intelligence", header_format=header_format)
        write_dataframe(writer, kpi_target_df if kpi_target_df is not None else pd.DataFrame(), "KPI_Targets", header_format=header_format)
        write_dataframe(writer, sql_result_df if sql_result_df is not None else pd.DataFrame(), "SQL_Query_Result", header_format=header_format)
        write_dataframe(writer, anomaly_report_df if anomaly_report_df is not None else pd.DataFrame(), "Anomaly_Report", header_format=header_format)
        write_dataframe(writer, rfm_df if rfm_df is not None else pd.DataFrame(), "RFM_Analysis", header_format=header_format)
        write_dataframe(writer, cohort_df if cohort_df is not None else pd.DataFrame(), "Cohort_Analysis", header_format=header_format)
        write_dataframe(writer, build_rule_based_executive_summary(df, num_col, cat_col, date_col), "Executive_Summary", header_format=header_format, heatmap=False)
        write_dataframe(writer, schema_contract_df if schema_contract_df is not None else pd.DataFrame(), "Schema_Contract", header_format=header_format)
        write_dataframe(writer, fuzzy_duplicates_df if fuzzy_duplicates_df is not None else pd.DataFrame(), "Fuzzy_Duplicates", header_format=header_format)
        write_dataframe(writer, abc_df if abc_df is not None else pd.DataFrame(), "ABC_Analysis", header_format=header_format)
        write_dataframe(writer, basket_df if basket_df is not None else pd.DataFrame(), "Basket_Analysis", header_format=header_format)
        write_dataframe(writer, weighted_score_df if weighted_score_df is not None else pd.DataFrame(), "Weighted_Scorecard", header_format=header_format)
        write_dataframe(writer, goal_seek_df if goal_seek_df is not None else pd.DataFrame(), "Goal_Seek", header_format=header_format)
        write_dataframe(writer, calendar_df if calendar_df is not None else pd.DataFrame(), "Calendar_Table", header_format=header_format)
        write_dataframe(writer, report_brief_df if report_brief_df is not None else build_report_brief(df, num_col, cat_col, date_col), "Report_Brief", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_hmg_brand_profile(), "HMG_Brand_Profile", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_hmg_ecosystem_map(), "HMG_Ecosystem", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_enterprise_readiness_assessment(df), "Enterprise_Readiness", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_access_control_matrix(), "Access_Control", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_data_governance_matrix(), "Governance_Matrix", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_enterprise_deployment_checklist(), "Deployment_Checklist", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_free_tools_register(), "Free_Tools_Register", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_privacy_notice(), "Privacy_Notice", header_format=header_format, heatmap=False)
        write_dataframe(writer, build_enterprise_release_notes(), "Release_Notes", header_format=header_format, heatmap=False)

        # 5. Formula library / feature explanation inside workbook
        formulas = pd.DataFrame([
            ["SUM", "Adds numbers", "=SUM(B2:B100)", "Used for total revenue, cost, units, scores."],
            ["AVERAGE", "Calculates mean", "=AVERAGE(B2:B100)", "Used for performance averages."],
            ["COUNT/COUNTA", "Counts numeric/non-empty cells", "=COUNTA(A2:A100)", "Used for record volume."],
            ["IF", "Conditional logic", "=IF(B2>=50,\"Pass\",\"Fail\")", "Used for classification."],
            ["XLOOKUP/VLOOKUP", "Lookup matching records", "=XLOOKUP(E2,A:A,B:B)", "Useful after exporting to Excel."],
            ["TEXT", "Formats date/text", "=TEXT(A2,\"mmm yyyy\")", "Used for month grouping."],
            ["PIVOT TABLE", "Summarizes by dimensions", "Insert > PivotTable", "Automated in Pivot_Analysis sheet."],
            ["FILTER/SORT", "Narrows/reorders records", "Data > Filter", "Auto-filter is enabled on data sheets."],
            ["CONDITIONAL FORMATTING", "Highlights patterns", "Home > Conditional Formatting", "Heatmaps are applied to numeric columns."],
        ], columns=["Excel Feature", "Purpose", "Example", "How this platform helps"])
        write_dataframe(writer, formulas, "Formula_Library", header_format=header_format, heatmap=False)

        guide = pd.DataFrame(FEATURE_CATALOG, columns=["Workflow Area", "Explanation"])
        write_dataframe(writer, guide, "Feature_Guide", header_format=header_format, heatmap=False)

        # 6. Hidden calculation engine and dashboard (preserves original dashboard generator concept)
        calc = workbook.add_worksheet("Calc_Engine")
        dashboard = workbook.add_worksheet("Dashboard")
        dashboard.hide_gridlines(2)
        dashboard.set_column("A:A", 2)
        dashboard.set_column("B:N", 14)

        dashboard.write("B2", title, title_format)
        dashboard.write("B3", f"{APP_NAME} • Architected by CSS Adewale • Zero AI API", subtitle_format)

        metric_ready = bool(num_col and num_col in df.columns and pd.api.types.is_numeric_dtype(df[num_col]))
        cat_ready = bool(cat_col and cat_col in df.columns and metric_ready)

        if metric_ready:
            kpis = [
                ("Total", pd.to_numeric(df[num_col], errors="coerce").sum()),
                ("Average", pd.to_numeric(df[num_col], errors="coerce").mean()),
                ("Count", pd.to_numeric(df[num_col], errors="coerce").count()),
                ("Maximum", pd.to_numeric(df[num_col], errors="coerce").max()),
                ("Minimum", pd.to_numeric(df[num_col], errors="coerce").min()),
            ]
        else:
            kpis = [("Rows", len(df)), ("Columns", len(df.columns)), ("Duplicates", int(df.duplicated().sum())), ("Missing Cells", int(df.isna().sum().sum())), ("Numeric Fields", len(numeric))]
        start_cols = ["B", "D", "F", "H", "J"]
        for (label, value), c in zip(kpis, start_cols):
            c2 = chr(ord(c) + 1)
            dashboard.merge_range(f"{c}6:{c2}6", label, kpi_title_format)
            dashboard.merge_range(f"{c}7:{c2}7", value if pd.notna(value) else 0, kpi_value_format)

        insight = f"Workbook generated with {len(df):,} processed rows and {len(df.columns):,} columns. "
        if cat_ready:
            top_summary = df.groupby(cat_col, dropna=False)[num_col].sum().reset_index().sort_values(num_col, ascending=False).head(top_n)
            if not top_summary.empty:
                insight += f"Top {cat_col}: {top_summary.iloc[0][cat_col]} with {top_summary.iloc[0][num_col]:,.2f} total {num_col}. "
        insight += "See Audit_Log, Data_Profile, Missing_Report, Outlier_Report, Pivot_Analysis, and Processed_Data for complete workflow evidence."
        dashboard.merge_range("H2:N4", insight, note_format)

        # Write calc tables and charts
        if cat_ready:
            cat_summary = df.groupby(cat_col, dropna=False)[num_col].sum().reset_index().sort_values(num_col, ascending=False).head(top_n)
            cat_summary.to_excel(writer, sheet_name="Calc_Engine", index=False, startrow=0, startcol=0)
            for i, c in enumerate(cat_summary.columns):
                calc.write(0, i, c, header_format)
            chart = workbook.add_chart({"type": "column"})
            n = len(cat_summary)
            if n:
                chart.add_series({
                    "name": f"{num_col} by {cat_col}",
                    "categories": ["Calc_Engine", 1, 0, n, 0],
                    "values": ["Calc_Engine", 1, 1, n, 1],
                    "fill": {"color": theme["primary_chart"]},
                    "data_labels": {"value": True, "num_format": "#,##0"},
                })
            chart.set_title({"name": f"Top {top_n} by {cat_col}"})
            chart.set_legend({"none": True})
            chart.set_plotarea({"border": {"none": True}})
            dashboard.insert_chart("B10", chart, {"x_scale": 1.55, "y_scale": 1.15})

        if date_col and date_col in df.columns and metric_ready:
            dt = df[[date_col, num_col]].copy()
            dt[date_col] = pd.to_datetime(dt[date_col], errors="coerce")
            dt[num_col] = pd.to_numeric(dt[num_col], errors="coerce")
            dt = dt.dropna()
            if not dt.empty:
                dt["YearMonth"] = dt[date_col].dt.to_period("M").astype(str)
                date_summary = dt.groupby("YearMonth")[num_col].sum().reset_index()
                date_summary.to_excel(writer, sheet_name="Calc_Engine", index=False, startrow=0, startcol=3)
                for i, c in enumerate(date_summary.columns):
                    calc.write(0, 3 + i, c, header_format)
                line = workbook.add_chart({"type": "line"})
                n = len(date_summary)
                line.add_series({
                    "name": f"{num_col} Trend",
                    "categories": ["Calc_Engine", 1, 3, n, 3],
                    "values": ["Calc_Engine", 1, 4, n, 4],
                    "line": {"color": theme["secondary_chart"], "width": 2.25},
                    "marker": {"type": "circle", "size": 5},
                })
                line.set_title({"name": f"Time Trend: {num_col}"})
                line.set_legend({"none": True})
                line.set_plotarea({"border": {"none": True}})
                dashboard.insert_chart("B27", line, {"x_scale": 1.55, "y_scale": 1.15})

        if sec_cat_col and sec_cat_col in df.columns and metric_ready:
            sec = df.groupby(sec_cat_col, dropna=False)[num_col].sum().reset_index().sort_values(num_col, ascending=False).head(8)
            sec.to_excel(writer, sheet_name="Calc_Engine", index=False, startrow=0, startcol=6)
            for i, c in enumerate(sec.columns):
                calc.write(0, 6 + i, c, header_format)
            pie = workbook.add_chart({"type": "pie"})
            n = len(sec)
            if n:
                pie.add_series({
                    "name": f"{num_col} by {sec_cat_col}",
                    "categories": ["Calc_Engine", 1, 6, n, 6],
                    "values": ["Calc_Engine", 1, 7, n, 7],
                    "data_labels": {"percentage": True, "leader_lines": True, "position": "outside_end"},
                })
            pie.set_title({"name": f"Segmentation: {sec_cat_col}"})
            dashboard.insert_chart("J10", pie, {"x_scale": 1.20, "y_scale": 1.15})

        if num_col_2 and num_col_2 in df.columns and metric_ready:
            sc = df[[num_col, num_col_2]].dropna().head(5000)
            sc.to_excel(writer, sheet_name="Calc_Engine", index=False, startrow=0, startcol=9)
            for i, c in enumerate(sc.columns):
                calc.write(0, 9 + i, c, header_format)
            if len(sc) > 1:
                scatter = workbook.add_chart({"type": "scatter"})
                n = len(sc)
                scatter.add_series({
                    "name": "Correlation",
                    "categories": ["Calc_Engine", 1, 9, n, 9],
                    "values": ["Calc_Engine", 1, 10, n, 10],
                    "marker": {"type": "circle", "size": 5, "fill": {"color": theme["accent"]}, "border": {"none": True}},
                })
                scatter.set_title({"name": f"Correlation: {num_col} vs {num_col_2}"})
                scatter.set_x_axis({"name": num_col})
                scatter.set_y_axis({"name": num_col_2})
                scatter.set_legend({"none": True})
                dashboard.insert_chart("J27", scatter, {"x_scale": 1.20, "y_scale": 1.15})

        dashboard.merge_range("B45:N45", "Enterprise Excel Operations Platform by His Marvellous Grace Educational Consult | Developed by CSS Adewale", subtitle_format)
        calc.hide()

        # Ensure number formatting on primary data sheets
        for sname in ["Raw_Data", "Processed_Data"]:
            if sname in writer.sheets:
                ws = writer.sheets[sname]
                for i, col in enumerate((raw_df if sname == "Raw_Data" else df).columns):
                    if pd.api.types.is_numeric_dtype((raw_df if sname == "Raw_Data" else df)[col]):
                        ws.set_column(i, i, 14, money_format)

    return output.getvalue()

# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(
    page_title="HMG Excel Operations Platform — Free Excel Dashboard & Data Analysis (No AI API)",
    page_icon="📊",
    layout="wide",
    menu_items={
        "Get Help": "https://wa.me/2348100866322",
        "Report a bug": "https://github.com/cssadewale",
        "About": (
            "HMG Excel Operations Platform — an HMG Academy product built by HMG Technologies "
            "(HMG Concepts). Free, no-AI-API Excel analytics. Founder: Adewale Samson Adeagbo."
        ),
    },
)

# SEO: inject crawlable meta tags + JSON-LD structured data for search engines.
st.markdown(seo_meta_html(), unsafe_allow_html=True)

# White-label: resolve active brand and apply its CSS (per-tenant re-skin, no AI).
ACTIVE_BRAND = get_active_brand()
st.markdown(brand_css(ACTIVE_BRAND), unsafe_allow_html=True)

# Authentication gate: preserved from original app, improved with optional environment-secret pattern
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "subscription_tier" not in st.session_state:
    st.session_state.subscription_tier = "free"
if "license_info" not in st.session_state:
    st.session_state.license_info = {}

if not st.session_state.authenticated:
    render_brand_hero(ACTIVE_BRAND, ACTIVE_BRAND.get("tagline", "") + " · Secure Data Gateway")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        # Crawlable intro text (helps SEO even before login).
        st.markdown(
            "**Free, no-code Excel & CSV data analysis and dashboard platform.** Clean, profile, pivot, "
            "forecast, visualize and export a complete enterprise Excel workbook — no AI API, no recurring cost. "
            "Part of the **HMG Concepts** ecosystem (Academy · Technologies · Media · Gospel), Lagos, Nigeria."
        )
        locked_for = login_locked_remaining()
        demo_on = _get_secret("ALLOW_DEMO_LOGIN", "true").strip().lower() != "false"
        if demo_on:
            st.info("Demo access keys for testing: `admin` or `HMG2025` (these grant the **Free** tier only). "
                    "Set a strong production password and disable demo login before public use — see DEPLOYMENT.md.")
        else:
            st.caption("Demo login is disabled on this deployment.")

        if locked_for > 0:
            st.error(f"🔒 Too many failed attempts. Try again in {locked_for} seconds.")
            st.stop()

        pwd = st.text_input("Enter Institutional Password", type="password")
        if st.button("Unlock Platform", use_container_width=True, type="primary"):
            tier = verify_password(pwd)
            if tier:
                st.session_state.authenticated = True
                st.session_state.subscription_tier = tier
                reset_login_attempts()
                add_audit("Authentication", f"Platform unlocked at {TIER_LABELS.get(tier, tier)} tier.")
                log_usage_event("Authentication", tier=tier)
                st.rerun()
            else:
                register_failed_login()
                attempts_left = max(0, _MAX_ATTEMPTS - st.session_state.get("login_attempts", 0))
                st.error(f"Invalid credentials. Access denied. Attempts remaining before lockout: {attempts_left}.")

        with st.expander("🔑 Have a subscription license key? Activate Pro / Enterprise"):
            st.caption("Paste a license key issued by HMG Technologies to upgrade your tier. "
                       "Keys are cryptographically signed and cannot be forged or extended. "
                       "For a per-seat licence: device-bound keys verify automatically; named keys need their Seat ID.")
            render_device_fingerprint()
            _dev = get_device_id()
            if _dev:
                st.caption(f"This device's automatic ID: `{_dev}`")
            lic = st.text_input("License key", key="login_license_key")
            seat_in = st.text_input("Seat ID (named per-seat licences only — leave blank for device-bound)", key="login_seat_id",
                                    placeholder="e.g. your work email or device label")
            if st.button("Validate & Continue with License", use_container_width=True):
                # Auto-fallback to the captured device id for device-bound seat keys.
                seat_used = seat_in.strip() or _dev
                info = verify_license_key(lic, seat_used)
                if info["valid"]:
                    st.session_state.authenticated = True
                    st.session_state.subscription_tier = info["tier"]
                    st.session_state.license_info = info
                    reset_login_attempts()
                    add_audit("License activation", f"{info['tier']} ({info['model']}) license for '{info['customer']}' valid until {info['expiry']}.")
                    log_usage_event("License activation", tier=info["tier"], model=info["model"], customer=info["customer"], detail=f"until {info['expiry']}")
                    st.success(f"License valid — {TIER_LABELS.get(info['tier'], info['tier'])} tier unlocked ({info['model']} model).")
                    st.rerun()
                else:
                    st.error(f"License rejected: {info['reason']}")
    st.stop()

# State initialization
if "raw_df" not in st.session_state:
    st.session_state.raw_df = pd.DataFrame()
if "working_df" not in st.session_state:
    st.session_state.working_df = pd.DataFrame()
if "grouped_summary" not in st.session_state:
    st.session_state.grouped_summary = pd.DataFrame()
if "pivot_table" not in st.session_state:
    st.session_state.pivot_table = pd.DataFrame()
if "forecast_df" not in st.session_state:
    st.session_state.forecast_df = pd.DataFrame()
if "validation_report" not in st.session_state:
    st.session_state.validation_report = pd.DataFrame()
if "what_if_df" not in st.session_state:
    st.session_state.what_if_df = pd.DataFrame()
if "pareto_df" not in st.session_state:
    st.session_state.pareto_df = pd.DataFrame()
if "reconciliation_report" not in st.session_state:
    st.session_state.reconciliation_report = pd.DataFrame()
if "time_intelligence_df" not in st.session_state:
    st.session_state.time_intelligence_df = pd.DataFrame()
if "kpi_target_df" not in st.session_state:
    st.session_state.kpi_target_df = pd.DataFrame()
if "sql_result_df" not in st.session_state:
    st.session_state.sql_result_df = pd.DataFrame()
if "anomaly_report_df" not in st.session_state:
    st.session_state.anomaly_report_df = pd.DataFrame()
if "rfm_df" not in st.session_state:
    st.session_state.rfm_df = pd.DataFrame()
if "cohort_df" not in st.session_state:
    st.session_state.cohort_df = pd.DataFrame()
if "schema_contract_df" not in st.session_state:
    st.session_state.schema_contract_df = pd.DataFrame()
if "fuzzy_duplicates_df" not in st.session_state:
    st.session_state.fuzzy_duplicates_df = pd.DataFrame()
if "abc_df" not in st.session_state:
    st.session_state.abc_df = pd.DataFrame()
if "basket_df" not in st.session_state:
    st.session_state.basket_df = pd.DataFrame()
if "weighted_score_df" not in st.session_state:
    st.session_state.weighted_score_df = pd.DataFrame()
if "goal_seek_df" not in st.session_state:
    st.session_state.goal_seek_df = pd.DataFrame()
if "calendar_df" not in st.session_state:
    st.session_state.calendar_df = pd.DataFrame()
if "report_brief_df" not in st.session_state:
    st.session_state.report_brief_df = pd.DataFrame()
if "audit" not in st.session_state:
    st.session_state.audit = []


# Centralized analysis-state keys so a new dataset always resets derived results.
_DERIVED_STATE_KEYS = [
    "grouped_summary", "pivot_table", "forecast_df", "validation_report", "what_if_df",
    "pareto_df", "reconciliation_report", "time_intelligence_df", "kpi_target_df",
    "sql_result_df", "anomaly_report_df", "rfm_df", "cohort_df", "schema_contract_df",
    "fuzzy_duplicates_df", "abc_df", "basket_df", "weighted_score_df", "goal_seek_df",
    "calendar_df", "report_brief_df",
]


def set_master_dataset(df: pd.DataFrame, source_desc: str) -> None:
    """Load a new master dataset and reset all previously derived analysis outputs."""
    if df.shape[0] > EXCEL_MAX_ROWS:
        st.warning("Dataset exceeds Excel row limits. The app will process it, but Excel export will be capped at 1,048,575 data rows per sheet.")
    df = df.copy()
    df.columns = make_unique_columns([str(c) for c in df.columns])
    st.session_state.raw_df = df.copy()
    st.session_state.working_df = df.copy()
    for key in _DERIVED_STATE_KEYS:
        st.session_state[key] = pd.DataFrame()
    add_audit("Dataset loaded", f"{source_desc}; shape {df.shape[0]:,} x {df.shape[1]:,}.")


render_brand_hero(ACTIVE_BRAND, ACTIVE_BRAND.get("tagline", "") + " • Excel Workflows + Governance + Modelling • No AI API")

with st.sidebar:
    st.subheader("HMG Technologies Ecosystem")
    st.caption("Built under HMG Concepts by Adewale Samson Adeagbo — Data Scientist, STEM Educator, AI-Augmented Solutions Developer.")
    st.markdown("""
    **Brand:** HMG Concepts / His Marvellous Grace Educational Consult  
    **Founded:** 2015  
    **Arms:** HMG Academy · HMG Technologies · HMG Media  
    **Motto:** Learning Deliberately. Teaching Authentically.  
    **Contact:** +234 810 086 6322  
    """)
    st.divider()
    # --- Subscription / plan panel ---
    _tier = current_tier()
    _tier_colors = {"free": "#64748B", "pro": "#1A5F7A", "enterprise": "#D4AF37"}
    st.subheader("Your Plan")
    st.markdown(
        f"<span class='tier-badge' style='background:{_tier_colors.get(_tier,'#64748B')}'>"
        f"{TIER_LABELS.get(_tier,'Free')} tier</span>",
        unsafe_allow_html=True,
    )
    if st.session_state.get("license_info", {}).get("valid"):
        _li = st.session_state["license_info"]
        st.caption(f"Licensed to: {_li.get('customer','—')} · {_li.get('model','team')} model · expires {_li.get('expiry','—')}")
    with st.expander("Upgrade with a license key"):
        render_device_fingerprint()
        _dev_sb = get_device_id()
        if _dev_sb:
            st.caption(f"This device's ID: `{_dev_sb}`")
        _lic2 = st.text_input("License key", key="sidebar_license_key")
        _seat2 = st.text_input("Seat ID (named per-seat only — blank = use this device)", key="sidebar_seat_id")
        if st.button("Apply License", use_container_width=True, key="sidebar_apply_license"):
            _seat_used = _seat2.strip() or get_device_id()
            _info = verify_license_key(_lic2, _seat_used)
            if _info["valid"]:
                st.session_state.subscription_tier = _info["tier"]
                st.session_state.license_info = _info
                add_audit("License upgrade", f"{_info['tier']} ({_info['model']}) for '{_info['customer']}' until {_info['expiry']}.")
                log_usage_event("License upgrade", tier=_info["tier"], model=_info["model"], customer=_info["customer"], detail=f"until {_info['expiry']}")
                st.success(f"{TIER_LABELS.get(_info['tier'])} unlocked ({_info['model']} model).")
                st.rerun()
            else:
                st.error(_info["reason"])
    if st.button("Log out", use_container_width=True, key="logout_btn"):
        for k in ["authenticated", "subscription_tier", "license_info"]:
            st.session_state.pop(k, None)
        st.rerun()
    st.divider()
    st.subheader("Workflow Map")
    st.caption("The platform is designed so the analyst clicks actions instead of manually building each Excel workflow.")
    st.markdown("""
    1. Upload / merge files  
    2. Profile data quality  
    3. Clean and standardize  
    4. Transform / engineer columns  
    5. Analyze / pivot / forecast  
    6. Visualize  
    7. Export complete Excel pack
    """)
    st.divider()
    if st.button("Reset Entire Session", type="secondary", use_container_width=True):
        for key in ["raw_df", "working_df", "grouped_summary", "pivot_table", "forecast_df", "validation_report", "what_if_df", "pareto_df", "reconciliation_report", "time_intelligence_df", "kpi_target_df", "sql_result_df", "anomaly_report_df", "rfm_df", "cohort_df", "schema_contract_df", "fuzzy_duplicates_df", "abc_df", "basket_df", "weighted_score_df", "goal_seek_df", "calendar_df", "report_brief_df", "audit"]:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

# Main tabs
tabs = st.tabs([
    "1️⃣ Ingest",
    "2️⃣ Profile",
    "3️⃣ Clean",
    "4️⃣ Transform",
    "5️⃣ Analyze",
    "6️⃣ Visualize",
    "7️⃣ Export",
    "8️⃣ Analyst Tools",
    "📘 Feature Guide",
])

with tabs[0]:
    st.header("1️⃣ Data Ingestion: CSV/XLSX Multi-File Master Upload")
    st.write("Upload one or more files. The system appends them into one master dataset and can retain the source file/sheet for auditability.")
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        uploaded_files = st.file_uploader("Upload Excel or CSV files", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
    with c2:
        read_all_sheets = st.checkbox("Read all Excel sheets", value=False)
        add_source_cols = st.checkbox("Add source file/sheet columns", value=True)
    with c3:
        st.write("Need test data?")
        demo_rows = st.slider("Demo rows", 100, 5000, 500, step=100)
        if st.button("Load Demo Dataset", use_container_width=True):
            demo = make_demo_data(demo_rows)
            set_master_dataset(demo, f"Demo dataset ({len(demo):,} rows)")
            st.success("Demo dataset loaded.")
            st.rerun()

    if uploaded_files and st.button("Ingest / Merge Uploaded Files", type="primary", use_container_width=True):
        try:
            df = load_uploaded_files(uploaded_files, read_all_sheets, add_source_cols)
            set_master_dataset(df, f"{len(uploaded_files)} uploaded file(s): {[f.name for f in uploaded_files]}")
            st.success(f"Ingested {len(uploaded_files)} file(s). Master dataset: {df.shape[0]:,} rows x {df.shape[1]:,} columns.")
        except Exception as exc:
            st.error(f"Could not ingest files: {exc}")

    # NEW (v6): Google Sheets URL import — free, no API key required.
    with st.expander("🔗 Import from a Google Sheets link (no API key, free)"):
        st.caption(
            "Paste a shareable Google Sheets link. The sheet must be viewable by 'Anyone with the link' "
            "or Published to the web (File ▸ Share ▸ Publish to web). No Google API or paid service is used."
        )
        gs_url = st.text_input("Google Sheets URL", key="gs_url", placeholder="https://docs.google.com/spreadsheets/d/.../edit#gid=0")
        if st.button("Import Google Sheet", use_container_width=True):
            try:
                gdf = read_google_sheet(gs_url)
                if gdf.empty:
                    st.warning("The Google Sheet was reached but returned no rows.")
                else:
                    set_master_dataset(gdf, f"Google Sheet import: {gs_url}")
                    st.success(f"Imported Google Sheet: {gdf.shape[0]:,} rows x {gdf.shape[1]:,} columns.")
            except Exception as exc:
                st.error(f"Could not import the Google Sheet. Ensure link sharing is enabled. Details: {exc}")

    if not st.session_state.working_df.empty:
        st.subheader("Current Dataset Preview")
        st.dataframe(st.session_state.working_df.head(20), use_container_width=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Rows", f"{st.session_state.working_df.shape[0]:,}")
        c2.metric("Columns", f"{st.session_state.working_df.shape[1]:,}")
        c3.metric("Missing Cells", f"{int(st.session_state.working_df.isna().sum().sum()):,}")
        c4.metric("Duplicates", f"{int(st.session_state.working_df.duplicated().sum()):,}")
    else:
        st.info("Upload files or load the demo dataset to activate the workflow tabs.")

# Other tabs require data
if st.session_state.working_df.empty:
    with tabs[8]:
        st.header("📘 Feature Guide")
        st.write("The platform activates the full workflow after data is uploaded. Summary of covered Excel analyst operations:")
        for title, desc in FEATURE_CATALOG:
            st.markdown(f"<div class='mini-card'><b>{title}</b><br>{desc}</div>", unsafe_allow_html=True)
    st.stop()

df = st.session_state.working_df
raw_df = st.session_state.raw_df if not st.session_state.raw_df.empty else df.copy()

with tabs[1]:
    st.header("2️⃣ Data Profiling and Quality Diagnostics")
    st.write("This replaces manual Excel inspection with automatic profiling sheets and clickable diagnostics.")
    overview = build_overview(df, raw_df)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Processed Rows", f"{len(df):,}")
    c2.metric("Columns", f"{len(df.columns):,}")
    c3.metric("Numeric", f"{len(numeric_cols(df)):,}")
    c4.metric("Text", f"{len(text_cols(df)):,}")
    c5.metric("Missing Cells", f"{int(df.isna().sum().sum()):,}")

    sub1, sub2, sub3, sub4, sub5 = st.tabs(["Overview", "Column Profile", "Missing Values", "Outliers", "✨ Auto Insights & Report"])
    with sub1:
        st.dataframe(overview, use_container_width=True)
        if numeric_cols(df):
            st.subheader("Numeric Descriptive Statistics")
            st.dataframe(df[numeric_cols(df)].describe().T, use_container_width=True)
    with sub2:
        st.dataframe(build_column_profile(df), use_container_width=True, height=420)
    with sub3:
        miss = build_missing_report(df)
        st.dataframe(miss, use_container_width=True)
        if not miss.empty and miss["Missing Count"].sum() > 0:
            st.bar_chart(miss.set_index("Column")["Missing Count"])
    with sub4:
        out_rep = build_outlier_report(df)
        st.dataframe(out_rep, use_container_width=True)
        if out_rep.empty:
            st.info("No numeric columns available for IQR outlier diagnostics.")
    with sub5:
        # NEW (v6): Automatic rule-based insight cards + downloadable offline HTML report.
        st.write("Deterministic, rule-based insights (no AI API). Pick optional context columns to enrich the analysis.")
        ic1, ic2, ic3 = st.columns(3)
        with ic1:
            ins_metric = st.selectbox("Metric (optional)", ["None"] + numeric_cols(df), key="ins_metric")
        with ic2:
            ins_cat = st.selectbox("Category (optional)", ["None"] + text_cols(df), key="ins_cat")
        with ic3:
            ins_date = st.selectbox("Date (optional)", ["None"] + date_like_cols(df), key="ins_date")
        insights = build_auto_insights(
            df,
            None if ins_metric == "None" else ins_metric,
            None if ins_cat == "None" else ins_cat,
            None if ins_date == "None" else ins_date,
        )
        tone_bg = {"good": "#ECFDF5", "warn": "#FFF7ED", "info": "#EFF6FF"}
        tone_border = {"good": "#16a34a", "warn": "#d97706", "info": "#0B3954"}
        cols_grid = st.columns(2)
        for i, ins in enumerate(insights):
            with cols_grid[i % 2]:
                st.markdown(
                    f"<div style='background:{tone_bg.get(ins['tone'],'#EFF6FF')};border-left:5px solid {tone_border.get(ins['tone'],'#0B3954')};"
                    f"border-radius:10px;padding:12px 14px;margin-bottom:10px;'>"
                    f"<b>{ins['title']}</b><br><span style='color:#334155;font-size:.92rem'>{ins['body']}</span></div>",
                    unsafe_allow_html=True,
                )
        st.divider()
        st.subheader("📄 Download Offline HTML Data Report")
        st.caption("A self-contained, shareable HTML report (KPIs, insights, column profile, data preview). Opens in any browser, no internet needed.")
        report_title = st.text_input("Report title", value="HMG Data Report", key="html_report_title")
        html_report = build_html_eda_report(df, report_title, insights)
        st.download_button(
            "⬇️ Download HTML Report",
            data=html_report.encode("utf-8"),
            file_name=f"{clean_column_name(report_title)}.html",
            mime="text/html",
            use_container_width=True,
        )

with tabs[2]:
    st.header("3️⃣ Data Cleaning Operations")
    st.write("Click cleaning options and apply them to the working dataset. Each action is added to the audit log.")
    c1, c2, c3 = st.columns(3)
    with c1:
        standardize = st.checkbox("Standardize column names")
        trim = st.checkbox("Trim text spaces")
        drop_blank_rows = st.checkbox("Drop fully blank rows")
        drop_blank_cols = st.checkbox("Drop fully blank columns")
    with c2:
        drop_duplicates = st.checkbox("Remove duplicate rows")
        drop_any_na = st.checkbox("Drop rows with any missing value")
        missing_threshold = st.slider("Drop columns with missing % above", 0, 100, 100)
    with c3:
        numeric_strategy = st.selectbox("Fill numeric missing", ["Do nothing", "Zero", "Mean", "Median", "Forward fill", "Backward fill"])
        text_strategy = st.selectbox("Fill text missing", ["Do nothing", "Mode", "Custom value", "Forward fill", "Backward fill"])
        custom_text = st.text_input("Custom text fill", "Unknown")

    st.subheader("Data Type Conversion")
    conv_cols = st.multiselect("Columns to convert", options=df.columns.tolist())
    conv_type = st.selectbox("Target type", ["Number", "Text", "Date/Time", "Category"])

    st.subheader("Outlier Treatment")
    outlier_columns = st.multiselect("Numeric columns for IQR outlier treatment", options=numeric_cols(df))
    outlier_method = st.selectbox("Outlier method", ["Do nothing", "Cap/Winsorize to IQR fences", "Replace with median", "Remove rows with outliers"])

    if st.button("Apply Selected Cleaning Operations", type="primary", use_container_width=True):
        before_shape = df.shape
        work = df.copy()
        actions = []
        if standardize:
            work = standardize_columns(work)
            actions.append("standardized column names")
        if trim:
            work = trim_text(work)
            actions.append("trimmed text columns")
        if drop_blank_rows:
            work = work.dropna(how="all")
            actions.append("dropped fully blank rows")
        if drop_blank_cols:
            work = work.dropna(axis=1, how="all")
            actions.append("dropped fully blank columns")
        if missing_threshold < 100:
            keep_cols = [c for c in work.columns if work[c].isna().mean() * 100 <= missing_threshold]
            dropped = len(work.columns) - len(keep_cols)
            work = work[keep_cols]
            actions.append(f"dropped {dropped} column(s) above {missing_threshold}% missing")
        if drop_duplicates:
            old = len(work)
            work = work.drop_duplicates()
            actions.append(f"removed {old-len(work):,} duplicate row(s)")
        if drop_any_na:
            old = len(work)
            work = work.dropna()
            actions.append(f"dropped {old-len(work):,} row(s) with missing values")
        if numeric_strategy != "Do nothing" or text_strategy != "Do nothing":
            work = apply_missing_strategy(work, numeric_strategy, text_strategy, custom_text)
            actions.append(f"missing fill numeric={numeric_strategy}, text={text_strategy}")
        if conv_cols:
            work = convert_columns(work, conv_cols, conv_type)
            actions.append(f"converted {conv_cols} to {conv_type}")
        if outlier_method != "Do nothing" and outlier_columns:
            work, detail = treat_outliers(work, outlier_columns, outlier_method)
            actions.append(detail)
        st.session_state.working_df = work
        add_audit("Cleaning operations", f"Before {before_shape}, after {work.shape}. Actions: {actions if actions else 'No option selected'}")
        st.success(f"Cleaning completed. Shape changed from {before_shape} to {work.shape}.")
        st.rerun()

    st.subheader("Preview After Latest Cleaning State")
    st.dataframe(df.head(20), use_container_width=True)

with tabs[3]:
    st.header("4️⃣ Transformation and Excel-Style Operations")
    st.write("Perform common worksheet operations without manual formulas.")
    op_tabs = st.tabs(["Filter/Sort", "Calculated Columns", "Text & Date", "Select/Rename"])

    with op_tabs[0]:
        st.subheader("Filter Rows")
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            f_col = st.selectbox("Filter column", df.columns.tolist())
        with fc2:
            f_op = st.selectbox("Operator", ["contains", "not contains", "equals", "not equals", "greater than", "greater or equal", "less than", "less or equal", "starts with", "ends with", "is blank", "is not blank"])
        with fc3:
            f_val = st.text_input("Filter value", "")
        if st.button("Apply Filter", use_container_width=True):
            before = len(df)
            st.session_state.working_df = apply_filter(df, f_col, f_op, f_val)
            add_audit("Filter", f"{f_col} {f_op} {f_val}; rows {before:,} -> {len(st.session_state.working_df):,}")
            st.rerun()

        st.subheader("Sort Rows")
        sc1, sc2 = st.columns(2)
        with sc1:
            sort_cols = st.multiselect("Sort by columns", df.columns.tolist())
        with sc2:
            ascending = st.checkbox("Ascending sort", value=True)
        if st.button("Apply Sort", use_container_width=True) and sort_cols:
            st.session_state.working_df = df.sort_values(sort_cols, ascending=ascending)
            add_audit("Sort", f"Sorted by {sort_cols}, ascending={ascending}")
            st.rerun()

    with op_tabs[1]:
        nums = numeric_cols(df)
        if len(nums) == 0:
            st.warning("No numeric columns detected for calculated columns.")
        else:
            cc1, cc2, cc3 = st.columns(3)
            with cc1:
                new_col = st.text_input("New column name", "new_calculation")
                calc_op = st.selectbox("Calculation", ["Add A + B", "Subtract A - B", "Multiply A * B", "Divide A / B", "Add constant to A", "Multiply A by constant", "A as % of column total", "Running total of A", "Difference from previous row", "Rank A descending", "Z-score of A", "Natural log of A"])
            with cc2:
                col_a = st.selectbox("Column A", nums)
                needs_b = calc_op in ["Add A + B", "Subtract A - B", "Multiply A * B", "Divide A / B"]
                col_b = st.selectbox("Column B", [c for c in nums if c != col_a] or nums) if needs_b else None
            with cc3:
                constant = st.number_input("Constant", value=1.0)
            if st.button("Create Calculated Column", type="primary", use_container_width=True):
                work = create_calculated_column(df, new_col, calc_op, col_a, col_b, constant)
                st.session_state.working_df = work
                add_audit("Calculated column", f"Created {new_col}: {calc_op}; A={col_a}; B={col_b}; constant={constant}")
                st.rerun()

    with op_tabs[2]:
        st.subheader("Text Operations")
        txts = text_cols(df)
        if txts:
            tc1, tc2, tc3 = st.columns(3)
            with tc1:
                t_col = st.selectbox("Text column", txts)
            with tc2:
                text_op = st.selectbox("Text action", ["UPPERCASE", "lowercase", "Title Case", "Replace text", "Split by delimiter"])
            with tc3:
                find_text = st.text_input("Find / delimiter", "-")
                replace_text = st.text_input("Replace with", "")
            if st.button("Apply Text Operation", use_container_width=True):
                work = df.copy()
                if text_op == "UPPERCASE":
                    work[t_col] = work[t_col].astype("string").str.upper()
                elif text_op == "lowercase":
                    work[t_col] = work[t_col].astype("string").str.lower()
                elif text_op == "Title Case":
                    work[t_col] = work[t_col].astype("string").str.title()
                elif text_op == "Replace text":
                    work[t_col] = work[t_col].astype("string").str.replace(find_text, replace_text, regex=False)
                elif text_op == "Split by delimiter":
                    parts = work[t_col].astype("string").str.split(find_text, expand=True)
                    for i in range(min(parts.shape[1], 5)):
                        work[f"{t_col}_part_{i+1}"] = parts[i]
                st.session_state.working_df = work
                add_audit("Text operation", f"{text_op} on {t_col}")
                st.rerun()
        else:
            st.info("No text columns detected.")

        st.subheader("Date Part Extraction")
        date_options = date_like_cols(df) or df.columns.tolist()
        dc1, dc2, dc3 = st.columns(3)
        with dc1:
            date_source = st.selectbox("Date source column", date_options)
        with dc2:
            date_part = st.selectbox("Date part", ["Year", "Quarter", "Month number", "Month name", "Weekday", "Year-Month", "Date only"])
        with dc3:
            date_new_col = st.text_input("Date output column", f"{date_source}_{date_part}".replace(" ", "_"))
        if st.button("Create Date Part Column", use_container_width=True):
            work = create_date_part(df, date_source, date_part, date_new_col)
            st.session_state.working_df = work
            add_audit("Date part", f"Created {date_new_col} from {date_source} ({date_part})")
            st.rerun()

    with op_tabs[3]:
        st.subheader("Select / Drop Columns")
        selected_cols = st.multiselect("Keep only selected columns", df.columns.tolist(), default=df.columns.tolist())
        if st.button("Apply Column Selection", use_container_width=True):
            st.session_state.working_df = df[selected_cols].copy()
            add_audit("Column selection", f"Kept {len(selected_cols)} columns: {selected_cols}")
            st.rerun()

        st.subheader("Rename One Column")
        rc1, rc2 = st.columns(2)
        with rc1:
            old_name = st.selectbox("Column to rename", df.columns.tolist())
        with rc2:
            new_name = st.text_input("New name", old_name)
        if st.button("Rename Column", use_container_width=True):
            work = df.rename(columns={old_name: new_name})
            st.session_state.working_df = work
            add_audit("Rename column", f"{old_name} -> {new_name}")
            st.rerun()

with tabs[4]:
    st.header("5️⃣ Analysis: Grouping, Pivoting, Correlation, Forecast")
    analysis_tabs = st.tabs(["Group-by Summary", "Pivot Table", "Correlation", "Forecast"])

    with analysis_tabs[0]:
        dims = st.multiselect("Dimensions / categories", df.columns.tolist(), default=text_cols(df)[:1])
        measures = st.multiselect("Numeric measures", numeric_cols(df), default=numeric_cols(df)[:1])
        aggs = st.multiselect("Aggregations", ["sum", "mean", "median", "min", "max", "count", "std"], default=["sum", "mean", "count"])
        if st.button("Build Group-by Summary", type="primary", use_container_width=True):
            summary = generate_grouped_summary(df, dims, measures, aggs)
            st.session_state.grouped_summary = summary
            add_audit("Grouped summary", f"dims={dims}, measures={measures}, aggs={aggs}, shape={summary.shape}")
        if not st.session_state.grouped_summary.empty:
            st.dataframe(st.session_state.grouped_summary.head(1000), use_container_width=True)

    with analysis_tabs[1]:
        pc1, pc2, pc3, pc4 = st.columns(4)
        with pc1:
            pivot_rows = st.multiselect("Pivot row fields", df.columns.tolist(), default=text_cols(df)[:1])
        with pc2:
            pivot_col = st.selectbox("Pivot column field", ["None"] + df.columns.tolist())
        with pc3:
            if numeric_cols(df):
                pivot_value = st.selectbox("Pivot value", numeric_cols(df))
            else:
                pivot_value = None
        with pc4:
            pivot_agg = st.selectbox("Pivot aggregation", ["sum", "mean", "count", "min", "max", "median"])
        if st.button("Build Pivot Table", type="primary", use_container_width=True):
            if pivot_value:
                pivot = generate_pivot(df, pivot_rows, pivot_col, pivot_value, pivot_agg)
                st.session_state.pivot_table = pivot
                add_audit("Pivot table", f"rows={pivot_rows}, cols={pivot_col}, value={pivot_value}, agg={pivot_agg}, shape={pivot.shape}")
            else:
                st.warning("A numeric value column is required.")
        if not st.session_state.pivot_table.empty:
            st.dataframe(st.session_state.pivot_table.head(1000), use_container_width=True)

    with analysis_tabs[2]:
        nums = numeric_cols(df)
        if len(nums) >= 2:
            corr = df[nums].corr(numeric_only=True)
            st.dataframe(corr, use_container_width=True)
            fig = px.imshow(corr, text_auto=True, title="Correlation Heatmap", color_continuous_scale="RdBu_r", zmin=-1, zmax=1)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("At least two numeric columns are required for correlation.")

    with analysis_tabs[3]:
        dates = date_like_cols(df) or df.columns.tolist()
        nums = numeric_cols(df)
        if dates and nums:
            fc1, fc2, fc3, fc4 = st.columns(4)
            with fc1:
                fc_date = st.selectbox("Date column", dates)
            with fc2:
                fc_metric = st.selectbox("Metric", nums)
            with fc3:
                fc_freq = st.selectbox("Frequency", ["D", "W", "M", "Q"], index=2, help="D=daily, W=weekly, M=monthly, Q=quarterly")
            with fc4:
                fc_periods = st.slider("Forecast periods", 1, 24, 6)
            fc_method = st.radio("Method", ["Linear trend", "Moving average"], horizontal=True)
            if st.button("Generate Forecast", type="primary", use_container_width=True):
                forecast = generate_forecast(df, fc_date, fc_metric, fc_periods, fc_freq, fc_method)
                st.session_state.forecast_df = forecast
                add_audit("Forecast", f"date={fc_date}, metric={fc_metric}, freq={fc_freq}, periods={fc_periods}, method={fc_method}")
            if not st.session_state.forecast_df.empty:
                st.dataframe(st.session_state.forecast_df, use_container_width=True)
                fdf = st.session_state.forecast_df.copy()
                fig = px.line(fdf, x="Period", y=[c for c in ["Actual", "Forecast"] if c in fdf.columns], title="Actual vs Forecast")
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("A date-like column and a numeric metric are required for forecasting.")

with tabs[5]:
    st.header("6️⃣ Visualization Builder")
    st.write("Create charts for review in the browser. The exported workbook also generates native Excel dashboard charts.")
    # NEW (v6): Smart, rule-based chart recommendations (no AI API).
    with st.expander("💡 Smart Chart Recommendations (auto-suggested from your columns)"):
        st.caption("Deterministic suggestions based on column types and cardinality — a free alternative to the 'AI auto-chart' feature in commercial tools.")
        st.dataframe(smart_chart_recommendations(df), use_container_width=True, hide_index=True)
    chart_type = st.selectbox("Chart type", ["Bar", "Line", "Area", "Scatter", "Histogram", "Box", "Pie", "Treemap", "Correlation Heatmap"])
    cols = df.columns.tolist()
    nums = numeric_cols(df)
    if chart_type == "Correlation Heatmap":
        if len(nums) >= 2:
            fig = px.imshow(df[nums].corr(numeric_only=True), text_auto=True, color_continuous_scale="RdBu_r", zmin=-1, zmax=1, title="Correlation Heatmap")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("At least two numeric columns are required.")
    else:
        vc1, vc2, vc3, vc4 = st.columns(4)
        with vc1:
            x_col = st.selectbox("X / Category", cols)
        with vc2:
            y_col = st.selectbox("Y / Metric", ["None"] + nums, index=1 if nums else 0)
        with vc3:
            color_col = st.selectbox("Color / Segment", ["None"] + cols)
        with vc4:
            chart_limit = st.slider("Max rows for browser chart", 100, 10000, 1000, step=100)
        plot_df = df.head(chart_limit).copy()
        color_arg = None if color_col == "None" else color_col
        y_arg = None if y_col == "None" else y_col
        try:
            if chart_type == "Bar":
                if y_arg:
                    plot_df = plot_df.groupby([x_col] + ([color_arg] if color_arg else []), dropna=False)[y_arg].sum().reset_index().sort_values(y_arg, ascending=False).head(50)
                fig = px.bar(plot_df, x=x_col, y=y_arg, color=color_arg, title=f"{chart_type}: {y_arg or 'Count'} by {x_col}")
            elif chart_type == "Line":
                fig = px.line(plot_df, x=x_col, y=y_arg, color=color_arg, title=f"Line Chart: {y_arg} over {x_col}")
            elif chart_type == "Area":
                fig = px.area(plot_df, x=x_col, y=y_arg, color=color_arg, title=f"Area Chart: {y_arg} over {x_col}")
            elif chart_type == "Scatter":
                y_for_scatter = y_arg or (nums[0] if nums else None)
                fig = px.scatter(plot_df, x=x_col, y=y_for_scatter, color=color_arg, trendline=None, title=f"Scatter: {x_col} vs {y_for_scatter}")
            elif chart_type == "Histogram":
                hist_col = y_arg or x_col
                fig = px.histogram(plot_df, x=hist_col, color=color_arg, title=f"Histogram: {hist_col}")
            elif chart_type == "Box":
                fig = px.box(plot_df, x=x_col, y=y_arg, color=color_arg, title=f"Box Plot: {y_arg} by {x_col}")
            elif chart_type == "Pie":
                if y_arg:
                    plot_df = plot_df.groupby(x_col, dropna=False)[y_arg].sum().reset_index().sort_values(y_arg, ascending=False).head(12)
                    fig = px.pie(plot_df, names=x_col, values=y_arg, title=f"Pie: {y_arg} by {x_col}")
                else:
                    plot_df = plot_df[x_col].value_counts().head(12).reset_index()
                    plot_df.columns = [x_col, "Count"]
                    fig = px.pie(plot_df, names=x_col, values="Count", title=f"Pie: Count by {x_col}")
            elif chart_type == "Treemap":
                fig = px.treemap(plot_df, path=[x_col] + ([color_arg] if color_arg else []), values=y_arg, title="Treemap")
            st.plotly_chart(fig, use_container_width=True)
        except Exception as exc:
            st.error(f"Chart could not be created with the selected fields: {exc}")

with tabs[6]:
    st.header("7️⃣ Export Complete Excel Operations Workbook")
    st.write("This is the main output: a multi-sheet Excel analyst pack, not only a dashboard.")
    nums = numeric_cols(df)
    cats = [c for c in df.columns if c not in nums] or df.columns.tolist()
    dates = ["None"] + date_like_cols(df) + [c for c in df.columns if c not in date_like_cols(df)]
    e1, e2, e3 = st.columns(3)
    with e1:
        title = st.text_input("Workbook / dashboard title", "HMG Excel Operations Analyst Pack")
        theme_name = st.selectbox("Theme", list(THEMES.keys()))
        top_n = st.slider("Top N for dashboard", 5, 50, 15)
    with e2:
        cat_col = st.selectbox("Primary category", cats)
        sec_cat_col = st.selectbox("Secondary category", ["None"] + [c for c in df.columns if c != cat_col])
        date_col = st.selectbox("Date column", dates)
    with e3:
        if nums:
            num_col = st.selectbox("Primary metric", nums)
            num_col_2 = st.selectbox("Second metric / scatter", ["None"] + [c for c in nums if c != num_col])
        else:
            num_col = None
            num_col_2 = "None"
            st.warning("No numeric metric detected. Workbook will still export quality/profile sheets.")

    if st.button("Build Excel Operations Workbook", type="primary", use_container_width=True):
        with st.spinner("Building workbook: audit, profile, cleaning reports, processed data, pivots, statistics, forecast, formula library, and dashboard..."):
            excel_bytes = build_excel_workbook(
                df=df,
                raw_df=raw_df,
                audit=st.session_state.audit,
                title=title,
                theme_name=theme_name,
                cat_col=cat_col,
                num_col=num_col,
                date_col=None if date_col == "None" else date_col,
                sec_cat_col=None if sec_cat_col == "None" else sec_cat_col,
                num_col_2=None if num_col_2 == "None" else num_col_2,
                top_n=top_n,
                grouped_summary=st.session_state.grouped_summary,
                pivot_table=st.session_state.pivot_table,
                forecast_df=st.session_state.forecast_df,
                validation_report=st.session_state.validation_report,
                what_if_df=st.session_state.what_if_df,
                pareto_df=st.session_state.pareto_df,
                reconciliation_report=st.session_state.reconciliation_report,
                time_intelligence_df=st.session_state.time_intelligence_df,
                kpi_target_df=st.session_state.kpi_target_df,
                sql_result_df=st.session_state.sql_result_df,
                anomaly_report_df=st.session_state.anomaly_report_df,
                rfm_df=st.session_state.rfm_df,
                cohort_df=st.session_state.cohort_df,
                schema_contract_df=st.session_state.schema_contract_df,
                fuzzy_duplicates_df=st.session_state.fuzzy_duplicates_df,
                abc_df=st.session_state.abc_df,
                basket_df=st.session_state.basket_df,
                weighted_score_df=st.session_state.weighted_score_df,
                goal_seek_df=st.session_state.goal_seek_df,
                calendar_df=st.session_state.calendar_df,
                report_brief_df=st.session_state.report_brief_df,
            )
            add_audit("Excel workbook built", f"Title={title}; shape={df.shape}")
            st.success("Workbook generated successfully.")
            st.download_button(
                "📥 Download Complete Excel Operations Workbook (.xlsx)",
                data=excel_bytes,
                file_name=f"{clean_column_name(title)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    st.download_button(
        "📄 Download Current Processed Dataset (.csv)",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="processed_dataset.csv",
        mime="text/csv",
        use_container_width=True,
    )

    st.divider()
    # NEW (v3): Saved analysis recipe (session config) — free, reproducible workflows.
    st.subheader("💾 Analysis Recipe (save / load your settings)")
    st.caption("Save your current selections (metric/category/date, forecast and report settings) as a small JSON "
               "file you can reload later or share. No data is stored — only your configuration choices.")
    rc1, rc2 = st.columns(2)
    with rc1:
        st.download_button(
            "⬇️ Save Recipe (.json)",
            data=export_session_recipe().encode("utf-8"),
            file_name="hmg_analysis_recipe.json",
            mime="application/json",
            use_container_width=True,
        )
    with rc2:
        recipe_file = st.file_uploader("Load a recipe (.json)", type=["json"], key="recipe_upload")
        if recipe_file is not None and st.button("Apply Recipe", use_container_width=True):
            ok, msg = apply_session_recipe(recipe_file.read().decode("utf-8"))
            (st.success if ok else st.error)(msg)
            if ok:
                add_audit("Recipe applied", msg)
                st.rerun()

    st.divider()
    # NEW (v3): Tamper-evident audit hash-chain [ENTERPRISE].
    st.subheader("🔗 Tamper-Evident Audit Trail  ·  Enterprise")
    st.caption("Produce a cryptographically chained audit log. Any later edit, deletion or reordering of a past "
               "entry breaks the chain and is detectable on verification.")
    if require_feature("tamper_evident_audit", "Tamper-evident audit trail"):
        chain = build_audit_hash_chain(st.session_state.audit)
        st.dataframe(chain, use_container_width=True, hide_index=True)
        check = verify_audit_hash_chain(chain) if "Full Entry Hash" in chain.columns else {"intact": False, "reason": "No entries."}
        (st.success if check["intact"] else st.warning)(("✅ " if check["intact"] else "⚠️ ") + check["reason"])
        st.download_button(
            "⬇️ Download Audit Hash-Chain (.csv)",
            data=chain.to_csv(index=False).encode("utf-8"),
            file_name="audit_hash_chain.csv",
            mime="text/csv",
            use_container_width=True,
        )


with tabs[7]:
    st.header("8️⃣ Analyst Tools: Validation, What-if, Pareto, Lookup Merge, Sampling")
    st.write("Extra no-API analyst utilities that extend the workflow beyond basic dashboard creation.")
    tool_tabs = st.tabs(["Data Validation", "What-if Scenario", "Pareto 80/20", "Lookup / Merge", "Sampling", "Reconciliation", "Time Intelligence", "KPI Targets", "Binning & IF", "SQL Query", "Anomalies", "RFM", "Cohort", "Split Workbook", "Schema Contract", "Fuzzy Dups", "ABC", "Basket", "Reshape/Scale", "Weighted Score", "Goal Seek", "Calendar", "Report Brief", "Compare Datasets", "🔒 PII Masking", "🔑 License Admin", "🧪 DQ Rules", "📈 Trend Decomp", "🔗 Correlations", "🧹 Clean Advisor", "📊 Survey/Likert", "🕵️ Benford Fraud", "🔀 Cross-tab χ²", "🔤 Text Frequency", "🔢 Number Normalizer", "🏷️ Data Classification", "📡 Usage Analytics", "🎯 Exec Scorecard", "📉 Concentration", "🔁 Snapshot Diff", "📤 Report Delivery", "🎨 White-Label", "HMG Brand"] )

    with tool_tabs[0]:
        st.subheader("Data Validation Rules")
        st.write("Create a rule-based exception report similar to Excel Data Validation and conditional checks.")
        vc1, vc2, vc3 = st.columns(3)
        with vc1:
            val_col = st.selectbox("Column to validate", df.columns.tolist(), key="val_col")
            val_rule = st.selectbox("Validation rule", ["Not blank", "Unique values only", "Numeric range", "Allowed list", "Valid date", "Positive number"], key="val_rule")
        with vc2:
            min_value = st.number_input("Minimum value", value=0.0, key="val_min")
            max_value = st.number_input("Maximum value", value=100.0, key="val_max")
        with vc3:
            allowed_values = st.text_area("Allowed list, comma-separated", "", help="Example: Lagos, Abuja, Kano", key="val_allowed")
        if st.button("Run Validation Rule", type="primary", use_container_width=True):
            report = build_validation_report(df, val_col, val_rule, min_value, max_value, allowed_values)
            st.session_state.validation_report = report
            add_audit("Data validation", f"Rule={val_rule}; column={val_col}; exceptions={len(report) if not report.empty else 0}")
        if not st.session_state.validation_report.empty:
            st.dataframe(st.session_state.validation_report.head(1000), use_container_width=True)
            st.download_button(
                "Download Validation Report CSV",
                data=st.session_state.validation_report.to_csv(index=False).encode("utf-8"),
                file_name="validation_report.csv",
                mime="text/csv",
                use_container_width=True,
            )

    with tool_tabs[1]:
        st.subheader("What-if Scenario Simulator")
        st.write("Model a simple scenario such as increasing revenue by 10% for one segment. This is similar to Excel scenario/what-if thinking without paid tools.")
        nums = numeric_cols(df)
        if nums:
            wc1, wc2, wc3, wc4 = st.columns(4)
            with wc1:
                wi_metric = st.selectbox("Metric", nums, key="wi_metric")
            with wc2:
                wi_change = st.selectbox("Change type", ["Increase by %", "Decrease by %", "Add constant", "Multiply by factor"], key="wi_change")
            with wc3:
                wi_value = st.number_input("Change value", value=10.0, key="wi_value")
            with wc4:
                wi_segment = st.selectbox("Optional segment column", ["None"] + df.columns.tolist(), key="wi_segment")
            wi_segment_value = ""
            if wi_segment != "None":
                choices = df[wi_segment].dropna().astype(str).unique().tolist()[:200]
                wi_segment_value = st.selectbox("Affected segment value", choices, key="wi_segment_value") if choices else ""
            if st.button("Run What-if Scenario", type="primary", use_container_width=True):
                what_if = build_what_if_analysis(df, wi_metric, wi_change, wi_value, wi_segment, wi_segment_value)
                st.session_state.what_if_df = what_if
                add_audit("What-if analysis", f"metric={wi_metric}; scenario={wi_change} {wi_value}; segment={wi_segment}:{wi_segment_value}")
            if not st.session_state.what_if_df.empty:
                st.dataframe(st.session_state.what_if_df, use_container_width=True)
                impact = float(st.session_state.what_if_df["Impact"].iloc[-1]) if "Impact" in st.session_state.what_if_df else 0
                st.metric("Scenario Impact", f"{impact:,.2f}")
        else:
            st.info("A numeric column is required for what-if analysis.")

    with tool_tabs[2]:
        st.subheader("Pareto 80/20 Analysis")
        st.write("Identify the categories that contribute the majority of a metric, such as top products causing 80% of sales or top students causing 80% of absences.")
        nums = numeric_cols(df)
        if nums:
            pc1, pc2 = st.columns(2)
            with pc1:
                pareto_cat = st.selectbox("Category", df.columns.tolist(), key="pareto_cat")
            with pc2:
                pareto_metric = st.selectbox("Metric", nums, key="pareto_metric")
            if st.button("Build Pareto Analysis", type="primary", use_container_width=True):
                pareto = build_pareto_analysis(df, pareto_cat, pareto_metric)
                st.session_state.pareto_df = pareto
                add_audit("Pareto analysis", f"category={pareto_cat}; metric={pareto_metric}; rows={len(pareto)}")
            if not st.session_state.pareto_df.empty:
                st.dataframe(st.session_state.pareto_df.head(200), use_container_width=True)
                fig = px.bar(st.session_state.pareto_df.head(30), x=pareto_cat, y=pareto_metric, color="Pareto Class", title="Pareto Contribution")
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("A numeric metric is required for Pareto analysis.")

    with tool_tabs[3]:
        st.subheader("Lookup / Merge Helper")
        st.write("Upload a lookup table and merge it into the current dataset. This automates common XLOOKUP/VLOOKUP-style enrichment.")
        lookup_file = st.file_uploader("Upload lookup CSV/XLSX", type=["csv", "xlsx", "xls"], key="lookup_file")
        if lookup_file is not None:
            try:
                if lookup_file.name.lower().endswith(".csv"):
                    lookup_df = try_read_csv(lookup_file)
                else:
                    lookup_df = pd.read_excel(lookup_file)
                st.dataframe(lookup_df.head(10), use_container_width=True)
                mc1, mc2, mc3 = st.columns(3)
                with mc1:
                    left_key = st.selectbox("Current dataset key", df.columns.tolist(), key="left_key")
                with mc2:
                    right_key = st.selectbox("Lookup table key", lookup_df.columns.tolist(), key="right_key")
                with mc3:
                    join_type = st.selectbox("Join type", ["left", "inner", "outer", "right"], key="join_type")
                if st.button("Apply Lookup/Merge", type="primary", use_container_width=True):
                    before = df.shape
                    merged = df.merge(lookup_df, how=join_type, left_on=left_key, right_on=right_key, suffixes=("", "_lookup"))
                    st.session_state.working_df = merged
                    add_audit("Lookup/Merge", f"{join_type} join; left_key={left_key}; right_key={right_key}; before={before}; after={merged.shape}")
                    st.success(f"Merge complete. Shape changed from {before} to {merged.shape}.")
                    st.rerun()
            except Exception as exc:
                st.error(f"Lookup file could not be processed: {exc}")

    with tool_tabs[4]:
        st.subheader("Dataset Sampling")
        st.write("Create a manageable sample for quick testing, review, classroom demonstration, or lightweight export.")
        sm1, sm2, sm3 = st.columns(3)
        with sm1:
            sample_mode = st.selectbox("Sample mode", ["Number of rows", "Percentage of rows"], key="sample_mode")
        with sm2:
            sample_size = st.number_input("Rows or percent", min_value=1.0, value=min(100.0, float(len(df))), key="sample_size")
        with sm3:
            random_seed = st.number_input("Random seed", min_value=0, value=42, step=1, key="random_seed")
        if st.button("Create Sample as Working Dataset", type="primary", use_container_width=True):
            if sample_mode == "Number of rows":
                n = int(min(sample_size, len(df)))
                sampled = df.sample(n=n, random_state=int(random_seed))
                detail = f"sampled {n} rows"
            else:
                frac = max(0.01, min(float(sample_size) / 100, 1.0))
                sampled = df.sample(frac=frac, random_state=int(random_seed))
                detail = f"sampled {frac:.2%} of rows"
            st.session_state.working_df = sampled.reset_index(drop=True)
            add_audit("Sampling", detail)
            st.success(f"Sampling complete: {len(sampled):,} rows.")
            st.rerun()


    with tool_tabs[5]:
        st.subheader("Dataset Reconciliation / Compare Two Tables")
        st.write("Upload a second file and compare it against the current working dataset. This automates reconciliation, duplicate checking, missing-record checks, and value mismatch checks often done with VLOOKUP/XLOOKUP.")
        recon_file = st.file_uploader("Upload comparison CSV/XLSX", type=["csv", "xlsx", "xls"], key="recon_file")
        if recon_file is not None:
            try:
                if recon_file.name.lower().endswith(".csv"):
                    comp_df = try_read_csv(recon_file)
                else:
                    comp_df = pd.read_excel(recon_file)
                st.caption("Comparison file preview")
                st.dataframe(comp_df.head(10), use_container_width=True)
                rc1, rc2, rc3 = st.columns(3)
                with rc1:
                    recon_left_key = st.selectbox("Current dataset key", df.columns.tolist(), key="recon_left_key")
                with rc2:
                    recon_right_key = st.selectbox("Comparison file key", comp_df.columns.tolist(), key="recon_right_key")
                common_cols = [c for c in df.columns if c in comp_df.columns and c != recon_left_key]
                with rc3:
                    tolerance = st.number_input("Numeric tolerance", value=0.0, min_value=0.0, key="recon_tolerance")
                compare_cols = st.multiselect("Columns to compare when names match in both files", common_cols, default=common_cols[:3], key="recon_compare_cols")
                if st.button("Run Reconciliation", type="primary", use_container_width=True):
                    report = build_reconciliation_report(df, comp_df, recon_left_key, recon_right_key, compare_cols, tolerance)
                    st.session_state.reconciliation_report = report
                    add_audit("Reconciliation", f"comparison={recon_file.name}; left_key={recon_left_key}; right_key={recon_right_key}; compare_cols={compare_cols}; issues={len(report)}")
                if not st.session_state.reconciliation_report.empty:
                    st.dataframe(st.session_state.reconciliation_report.head(2000), use_container_width=True)
                    st.download_button("Download Reconciliation CSV", st.session_state.reconciliation_report.to_csv(index=False).encode("utf-8"), "reconciliation_report.csv", "text/csv", use_container_width=True)
            except Exception as exc:
                st.error(f"Reconciliation file could not be processed: {exc}")

    with tool_tabs[6]:
        st.subheader("Time Intelligence")
        st.write("Create period summaries with previous-period value, change, percentage change, YTD, and rolling averages. This automates common Excel monthly/quarterly analysis.")
        dates = date_like_cols(df) or df.columns.tolist()
        nums = numeric_cols(df)
        if dates and nums:
            tc1, tc2, tc3, tc4 = st.columns(4)
            with tc1:
                ti_date = st.selectbox("Date column", dates, key="ti_date")
            with tc2:
                ti_metric = st.selectbox("Metric", nums, key="ti_metric")
            with tc3:
                ti_segment = st.selectbox("Optional segment", ["None"] + df.columns.tolist(), key="ti_segment")
            with tc4:
                ti_freq = st.selectbox("Period", ["M", "Q", "W", "D"], index=0, help="M=monthly, Q=quarterly, W=weekly, D=daily", key="ti_freq")
            if st.button("Build Time Intelligence", type="primary", use_container_width=True):
                ti = build_time_intelligence(df, ti_date, ti_metric, ti_segment, ti_freq)
                st.session_state.time_intelligence_df = ti
                add_audit("Time intelligence", f"date={ti_date}; metric={ti_metric}; segment={ti_segment}; freq={ti_freq}; rows={len(ti)}")
            if not st.session_state.time_intelligence_df.empty:
                st.dataframe(st.session_state.time_intelligence_df.head(1000), use_container_width=True)
                y_cols = [ti_metric, "YTD", "Rolling 3 Period Avg"]
                available_y = [c for c in y_cols if c in st.session_state.time_intelligence_df.columns]
                try:
                    fig = px.line(st.session_state.time_intelligence_df, x="Period", y=available_y, color=None if ti_segment == "None" else ti_segment, title="Time Intelligence Trend")
                    st.plotly_chart(fig, use_container_width=True)
                except Exception:
                    pass
        else:
            st.info("A date-like column and numeric metric are required.")

    with tool_tabs[7]:
        st.subheader("KPI Target Tracker")
        st.write("Compare actual performance against a constant target or a target column, with variance, attainment percentage, and status.")
        nums = numeric_cols(df)
        if nums:
            kc1, kc2, kc3, kc4 = st.columns(4)
            with kc1:
                kpi_actual = st.selectbox("Actual metric", nums, key="kpi_actual")
            with kc2:
                target_mode = st.selectbox("Target mode", ["Constant target", "Target column"], key="kpi_target_mode")
            with kc3:
                kpi_target_value = st.number_input("Constant target", value=1000.0, key="kpi_target_value")
            with kc4:
                kpi_segment = st.selectbox("Optional segment", ["None"] + df.columns.tolist(), key="kpi_segment")
            kpi_target_col = None
            if target_mode == "Target column":
                kpi_target_col = st.selectbox("Target column", nums, key="kpi_target_col")
            if st.button("Build KPI Target Analysis", type="primary", use_container_width=True):
                kpi = build_kpi_target_analysis(df, kpi_actual, target_mode, kpi_target_value, kpi_target_col, kpi_segment)
                st.session_state.kpi_target_df = kpi
                add_audit("KPI target analysis", f"actual={kpi_actual}; mode={target_mode}; target_col={kpi_target_col}; segment={kpi_segment}")
            if not st.session_state.kpi_target_df.empty:
                st.dataframe(st.session_state.kpi_target_df, use_container_width=True)
                total_actual = st.session_state.kpi_target_df["Actual"].sum() if "Actual" in st.session_state.kpi_target_df else 0
                total_target = st.session_state.kpi_target_df["Target"].sum() if "Target" in st.session_state.kpi_target_df else 0
                st.metric("Total Variance", f"{total_actual-total_target:,.2f}")
        else:
            st.info("A numeric column is required for KPI target tracking.")

    with tool_tabs[8]:
        st.subheader("Binning and IF-style Classification")
        st.write("Create grouped bands and rule-based labels, similar to Excel bins, nested IF, IFS, and classification formulas.")
        btab1, btab2 = st.tabs(["Numeric Bins", "IF Classification"])
        with btab1:
            nums = numeric_cols(df)
            if nums:
                bc1, bc2, bc3, bc4 = st.columns(4)
                with bc1:
                    bin_col = st.selectbox("Numeric column", nums, key="bin_col")
                with bc2:
                    bin_new_col = st.text_input("New bin column", f"{bin_col}_band", key="bin_new_col")
                with bc3:
                    bin_count = st.slider("Number of bins", 2, 20, 5, key="bin_count")
                with bc4:
                    bin_method = st.selectbox("Method", ["Equal width", "Quantile bands"], key="bin_method")
                if st.button("Create Bin Column", type="primary", use_container_width=True):
                    st.session_state.working_df = create_binned_column(df, bin_col, bin_new_col, bin_count, bin_method)
                    add_audit("Binning", f"Created {bin_new_col} from {bin_col}; bins={bin_count}; method={bin_method}")
                    st.rerun()
            else:
                st.info("A numeric column is required for binning.")
        with btab2:
            ic1, ic2, ic3 = st.columns(3)
            with ic1:
                if_source = st.selectbox("Source column", df.columns.tolist(), key="if_source")
                if_new_col = st.text_input("New classification column", f"{if_source}_class", key="if_new_col")
            with ic2:
                if_operator = st.selectbox("Condition", [">", ">=", "<", "<=", "equals", "not equals", "contains", "is blank", "is not blank"], key="if_operator")
                if_compare = st.text_input("Compare value", "0", key="if_compare")
            with ic3:
                if_true = st.text_input("Value if TRUE", "Yes", key="if_true")
                if_false = st.text_input("Value if FALSE", "No", key="if_false")
            if st.button("Create IF Classification Column", type="primary", use_container_width=True):
                st.session_state.working_df = create_if_column(df, if_source, if_operator, if_compare, if_true, if_false, if_new_col)
                add_audit("IF classification", f"Created {if_new_col}: {if_source} {if_operator} {if_compare}")
                st.rerun()


    with tool_tabs[9]:
        st.subheader("SQL Query Builder  ·  Pro")
        st.write("Run SELECT queries against the working dataset using free in-memory SQLite. The table name is `data`. In SQL mode, column names are standardized to snake_case for reliability.")
        if require_feature("sql_query", "SQL query builder"):
            sql_mapping = pd.DataFrame({
                "Original Column": df.columns.tolist(),
                "SQL Column": make_unique_columns([clean_column_name(c) for c in df.columns]),
            })
            with st.expander("View SQL column name mapping"):
                st.dataframe(sql_mapping, use_container_width=True)
            default_sql = "SELECT * FROM data LIMIT 20"
            sql_text = st.text_area("SQL SELECT query", value=default_sql, height=150, key="sql_text")
            if st.button("Run SQL Query", type="primary", use_container_width=True):
                result = run_sql_query(df, sql_text)
                st.session_state.sql_result_df = result
                add_audit("SQL query", f"Query executed; result shape={result.shape}")
            if not st.session_state.sql_result_df.empty:
                st.dataframe(st.session_state.sql_result_df.head(2000), use_container_width=True)
                st.download_button("Download SQL Result CSV", st.session_state.sql_result_df.to_csv(index=False).encode("utf-8"), "sql_query_result.csv", "text/csv", use_container_width=True)

    with tool_tabs[10]:
        st.subheader("Detailed Anomaly Detection")
        st.write("Generate a row-level anomaly report using deterministic IQR or z-score rules. This is not AI; it is statistical rule-based checking.")
        nums = numeric_cols(df)
        if nums:
            ac1, ac2, ac3 = st.columns(3)
            with ac1:
                anomaly_cols = st.multiselect("Numeric columns", nums, default=nums[:min(3, len(nums))], key="anomaly_cols")
            with ac2:
                anomaly_method = st.selectbox("Method", ["IQR", "Z-score"], key="anomaly_method")
            with ac3:
                z_threshold = st.number_input("Z-score threshold", min_value=1.0, value=3.0, step=0.5, key="z_threshold")
            if st.button("Build Anomaly Report", type="primary", use_container_width=True):
                report = build_anomaly_detail_report(df, anomaly_cols, anomaly_method, z_threshold)
                st.session_state.anomaly_report_df = report
                add_audit("Anomaly report", f"method={anomaly_method}; columns={anomaly_cols}; rows={len(report)}")
            if not st.session_state.anomaly_report_df.empty:
                st.dataframe(st.session_state.anomaly_report_df.head(2000), use_container_width=True)
                st.download_button("Download Anomaly Report CSV", st.session_state.anomaly_report_df.to_csv(index=False).encode("utf-8"), "anomaly_report.csv", "text/csv", use_container_width=True)
        else:
            st.info("A numeric column is required for anomaly detection.")

    with tool_tabs[11]:
        st.subheader("RFM Customer / Member Segmentation")
        st.write("Create Recency-Frequency-Monetary segmentation for customer, student, donor, member, or client activity data.")
        nums = numeric_cols(df)
        dates = date_like_cols(df) or df.columns.tolist()
        if nums and dates:
            r1, r2, r3 = st.columns(3)
            with r1:
                rfm_customer = st.selectbox("Customer / entity column", df.columns.tolist(), key="rfm_customer")
            with r2:
                rfm_date = st.selectbox("Transaction / activity date", dates, key="rfm_date")
            with r3:
                rfm_amount = st.selectbox("Amount / value metric", nums, key="rfm_amount")
            if st.button("Build RFM Analysis", type="primary", use_container_width=True):
                rfm = build_rfm_analysis(df, rfm_customer, rfm_date, rfm_amount)
                st.session_state.rfm_df = rfm
                add_audit("RFM analysis", f"entity={rfm_customer}; date={rfm_date}; amount={rfm_amount}; rows={len(rfm)}")
            if not st.session_state.rfm_df.empty:
                st.dataframe(st.session_state.rfm_df.head(1000), use_container_width=True)
                if "Segment" in st.session_state.rfm_df.columns:
                    fig = px.histogram(st.session_state.rfm_df, x="Segment", title="RFM Segment Distribution")
                    st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("RFM needs a date-like column and a numeric amount/value column.")

    with tool_tabs[12]:
        st.subheader("Cohort Analysis")
        st.write("Build monthly cohorts from an entity ID and date column. This helps analyze retention, repeat activity, and cohort value over time.")
        dates = date_like_cols(df) or df.columns.tolist()
        nums = numeric_cols(df)
        if dates:
            c1, c2, c3 = st.columns(3)
            with c1:
                cohort_entity = st.selectbox("Entity / customer / student ID", df.columns.tolist(), key="cohort_entity")
            with c2:
                cohort_date = st.selectbox("Activity date", dates, key="cohort_date")
            with c3:
                cohort_metric = st.selectbox("Optional value metric", ["None"] + nums, key="cohort_metric")
            if st.button("Build Cohort Analysis", type="primary", use_container_width=True):
                cohort = build_cohort_analysis(df, cohort_entity, cohort_date, cohort_metric)
                st.session_state.cohort_df = cohort
                add_audit("Cohort analysis", f"entity={cohort_entity}; date={cohort_date}; metric={cohort_metric}; shape={cohort.shape}")
            if not st.session_state.cohort_df.empty:
                st.dataframe(st.session_state.cohort_df, use_container_width=True)
        else:
            st.info("A date-like column is required for cohort analysis.")

    with tool_tabs[13]:
        st.subheader("Split Workbook by Segment")
        st.write("Create a separate Excel workbook where each selected segment value becomes its own sheet. This is useful for branch reports, class reports, department reports, or regional packs.")
        split_col = st.selectbox("Split by column", df.columns.tolist(), key="split_col")
        max_sheets = st.slider("Maximum segment sheets", 1, 50, 20, key="max_sheets")
        if st.button("Build Split Workbook", type="primary", use_container_width=True):
            split_bytes = build_split_workbook(df, split_col, max_sheets)
            add_audit("Split workbook", f"split_col={split_col}; max_sheets={max_sheets}")
            st.download_button(
                "Download Split Workbook (.xlsx)",
                data=split_bytes,
                file_name=f"split_by_{clean_column_name(split_col)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


    with tool_tabs[14]:
        st.subheader("Schema Contract / Data Contract Checker")
        st.write("Define expected required fields, key columns, numeric fields, and date fields. The system creates a governance report similar to a data contract checklist.")
        sc1, sc2 = st.columns(2)
        with sc1:
            expected_required_text = st.text_area("Expected required columns, comma-separated", ", ".join(df.columns[:min(3, len(df.columns))]), key="expected_required_text")
            key_cols_contract = st.multiselect("Unique key columns", df.columns.tolist(), key="key_cols_contract")
        with sc2:
            expected_numeric_contract = st.multiselect("Columns expected to be numeric", df.columns.tolist(), default=numeric_cols(df)[:min(3, len(numeric_cols(df)))], key="expected_numeric_contract")
            expected_dates_contract = st.multiselect("Columns expected to be dates", df.columns.tolist(), default=date_like_cols(df)[:min(2, len(date_like_cols(df)))], key="expected_dates_contract")
        if st.button("Run Schema Contract Checks", type="primary", use_container_width=True):
            expected_required = [c.strip() for c in expected_required_text.split(",") if c.strip()]
            report = build_schema_contract_report(df, expected_required, key_cols_contract, expected_numeric_contract, expected_dates_contract)
            st.session_state.schema_contract_df = report
            add_audit("Schema contract", f"required={expected_required}; key_cols={key_cols_contract}; checks={len(report)}")
        if not st.session_state.schema_contract_df.empty:
            st.dataframe(st.session_state.schema_contract_df, use_container_width=True)

    with tool_tabs[15]:
        st.subheader("Fuzzy Duplicate Detection")
        st.write("Find near-duplicate text records using Python's built-in similarity matching. This helps detect spelling variations and near matches without paid tools.")
        fz1, fz2, fz3 = st.columns(3)
        with fz1:
            fuzzy_cols = st.multiselect("Text columns to compare", text_cols(df) or df.columns.tolist(), default=(text_cols(df) or df.columns.tolist())[:1], key="fuzzy_cols")
        with fz2:
            fuzzy_threshold = st.slider("Similarity threshold %", 70, 100, 90, key="fuzzy_threshold")
        with fz3:
            fuzzy_max = st.slider("Max unique candidates", 50, 1000, 300, step=50, key="fuzzy_max")
        if st.button("Find Fuzzy Duplicates", type="primary", use_container_width=True):
            report = build_fuzzy_duplicate_report(df, fuzzy_cols, fuzzy_threshold, fuzzy_max)
            st.session_state.fuzzy_duplicates_df = report
            add_audit("Fuzzy duplicates", f"cols={fuzzy_cols}; threshold={fuzzy_threshold}; rows={len(report)}")
        if not st.session_state.fuzzy_duplicates_df.empty:
            st.dataframe(st.session_state.fuzzy_duplicates_df.head(1000), use_container_width=True)
            st.download_button("Download Fuzzy Duplicate Report CSV", st.session_state.fuzzy_duplicates_df.to_csv(index=False).encode("utf-8"), "fuzzy_duplicates.csv", "text/csv", use_container_width=True)

    with tool_tabs[16]:
        st.subheader("ABC Classification Analysis")
        st.write("Classify categories into A/B/C contribution groups. This is useful for inventory, customer value, product performance, cost drivers, and revenue concentration.")
        nums = numeric_cols(df)
        if nums:
            ab1, ab2, ab3, ab4 = st.columns(4)
            with ab1:
                abc_category = st.selectbox("Category", df.columns.tolist(), key="abc_category")
            with ab2:
                abc_metric = st.selectbox("Metric", nums, key="abc_metric")
            with ab3:
                a_threshold = st.slider("A threshold cumulative %", 50, 90, 80, key="a_threshold")
            with ab4:
                b_threshold = st.slider("B threshold cumulative %", 85, 99, 95, key="b_threshold")
            if st.button("Build ABC Analysis", type="primary", use_container_width=True):
                abc = build_abc_analysis(df, abc_category, abc_metric, a_threshold, b_threshold)
                st.session_state.abc_df = abc
                add_audit("ABC analysis", f"category={abc_category}; metric={abc_metric}; rows={len(abc)}")
            if not st.session_state.abc_df.empty:
                st.dataframe(st.session_state.abc_df.head(500), use_container_width=True)
                if "ABC Class" in st.session_state.abc_df.columns:
                    fig = px.bar(st.session_state.abc_df.head(30), x=abc_category, y=abc_metric, color="ABC Class", title="ABC Classification")
                    st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("ABC analysis requires at least one numeric metric.")

    with tool_tabs[17]:
        st.subheader("Basket / Co-occurrence Analysis")
        st.write("Analyze which items appear together inside the same transaction, invoice, receipt, student activity, or group identifier.")
        ba1, ba2, ba3 = st.columns(3)
        with ba1:
            basket_tx = st.selectbox("Transaction / group ID", df.columns.tolist(), key="basket_tx")
        with ba2:
            basket_item = st.selectbox("Item / activity column", df.columns.tolist(), key="basket_item")
        with ba3:
            basket_min = st.number_input("Minimum pair count", min_value=1, value=1, step=1, key="basket_min")
        if st.button("Build Basket Analysis", type="primary", use_container_width=True):
            basket = build_basket_analysis(df, basket_tx, basket_item, int(basket_min))
            st.session_state.basket_df = basket
            add_audit("Basket analysis", f"transaction={basket_tx}; item={basket_item}; rows={len(basket)}")
        if not st.session_state.basket_df.empty:
            st.dataframe(st.session_state.basket_df.head(1000), use_container_width=True)

    with tool_tabs[18]:
        st.subheader("Reshape and Scale Tools")
        st.write("Additional Power Query-style reshaping and numeric scaling operations.")
        rs1, rs2 = st.tabs(["Scale Numeric Columns", "Unpivot Wide to Long"])
        with rs1:
            nums = numeric_cols(df)
            if nums:
                sg1, sg2, sg3 = st.columns(3)
                with sg1:
                    scale_cols = st.multiselect("Numeric columns to scale", nums, default=nums[:1], key="scale_cols")
                with sg2:
                    scale_method = st.selectbox("Scaling method", ["Min-Max 0-1", "Z-score standardize", "Percent rank"], key="scale_method")
                with sg3:
                    scale_suffix = st.text_input("New column suffix", "scaled", key="scale_suffix")
                if st.button("Add Scaled Columns", type="primary", use_container_width=True):
                    st.session_state.working_df = scale_numeric_columns(df, scale_cols, scale_method, scale_suffix)
                    add_audit("Scale numeric columns", f"cols={scale_cols}; method={scale_method}; suffix={scale_suffix}")
                    st.rerun()
            else:
                st.info("Scaling requires numeric columns.")
        with rs2:
            unpivot_id_cols = st.multiselect("ID columns to keep", df.columns.tolist(), key="unpivot_id_cols")
            unpivot_value_cols = st.multiselect("Value columns to unpivot", [c for c in df.columns if c not in unpivot_id_cols], key="unpivot_value_cols")
            uv1, uv2 = st.columns(2)
            with uv1:
                var_name = st.text_input("Variable column name", "Variable", key="var_name")
            with uv2:
                value_name = st.text_input("Value column name", "Value", key="value_name")
            if st.button("Convert to Long Format", type="primary", use_container_width=True):
                st.session_state.working_df = melt_to_long_format(df, unpivot_id_cols, unpivot_value_cols, var_name, value_name)
                add_audit("Unpivot to long format", f"id_cols={unpivot_id_cols}; value_cols={unpivot_value_cols}; new_shape={st.session_state.working_df.shape}")
                st.rerun()


    with tool_tabs[19]:
        st.subheader("Weighted Scorecard / Decision Matrix")
        st.write("Rank rows using selected numeric criteria and weights. This automates Excel-style weighted scoring for admissions, vendor selection, student ranking, risk scoring, product prioritisation, and performance review.")
        nums = numeric_cols(df)
        if nums:
            ws1, ws2, ws3 = st.columns(3)
            with ws1:
                score_cols = st.multiselect("Scoring metrics", nums, default=nums[:min(3, len(nums))], key="score_cols_v5")
            with ws2:
                weights_text = st.text_input("Weights, comma-separated", ",".join(["1"] * max(len(score_cols), 1)), key="weights_text_v5")
                st.caption("Example: 5,3,2. If invalid, equal weights are used.")
            with ws3:
                cost_cols = st.multiselect("Cost/risk columns where lower is better", score_cols, key="cost_cols_v5")
            if st.button("Build Weighted Scorecard", type="primary", use_container_width=True):
                try:
                    weights = [float(x.strip()) for x in weights_text.split(",") if x.strip()]
                except Exception:
                    weights = [1.0] * len(score_cols)
                scorecard = build_weighted_scorecard(df, score_cols, weights, cost_cols)
                st.session_state.weighted_score_df = scorecard
                add_audit("Weighted scorecard", f"cols={score_cols}; weights={weights}; cost_cols={cost_cols}; rows={len(scorecard)}")
            if not st.session_state.weighted_score_df.empty:
                st.dataframe(st.session_state.weighted_score_df.head(1000), use_container_width=True)
        else:
            st.info("A numeric column is required for weighted scoring.")

    with tool_tabs[20]:
        st.subheader("Goal Seek / Target Gap Calculator")
        st.write("Calculate the gap between current performance and a desired target. This automates simple Excel Goal Seek and target planning.")
        nums = numeric_cols(df)
        if nums:
            gs1, gs2, gs3 = st.columns(3)
            with gs1:
                gs_metric = st.selectbox("Metric", nums, key="gs_metric_v5")
            with gs2:
                current_total = pd.to_numeric(df[gs_metric], errors="coerce").sum()
                gs_target = st.number_input("Target total", value=float(current_total * 1.1 if pd.notna(current_total) else 1000.0), key="gs_target_v5")
            with gs3:
                gs_segment = st.selectbox("Optional segment", ["None"] + df.columns.tolist(), key="gs_segment_v5")
            if st.button("Run Goal Seek", type="primary", use_container_width=True):
                goal = build_goal_seek_analysis(df, gs_metric, gs_target, gs_segment)
                st.session_state.goal_seek_df = goal
                add_audit("Goal seek", f"metric={gs_metric}; target={gs_target}; segment={gs_segment}")
            if not st.session_state.goal_seek_df.empty:
                st.dataframe(st.session_state.goal_seek_df, use_container_width=True)
        else:
            st.info("A numeric column is required for goal seek analysis.")

    with tool_tabs[21]:
        st.subheader("Calendar Table Generator")
        st.write("Create a full date dimension table from a date column. Useful for Excel, Power BI, pivot tables, fiscal year analysis, and time intelligence.")
        date_candidates = date_like_cols(df) or df.columns.tolist()
        cal_col = st.selectbox("Date column", date_candidates, key="cal_col_v5")
        if st.button("Build Calendar Table", type="primary", use_container_width=True):
            calendar = build_calendar_table(df, cal_col)
            st.session_state.calendar_df = calendar
            add_audit("Calendar table", f"date_col={cal_col}; rows={len(calendar)}")
        if not st.session_state.calendar_df.empty:
            st.dataframe(st.session_state.calendar_df.head(1000), use_container_width=True)
            st.download_button("Download Calendar CSV", st.session_state.calendar_df.to_csv(index=False).encode("utf-8"), "calendar_table.csv", "text/csv", use_container_width=True)

    with tool_tabs[22]:
        st.subheader("Rule-Based Report Brief")
        st.write("Generate a practical non-AI analyst brief with findings and recommended actions. This is deterministic and uses no paid model API.")
        nums = numeric_cols(df)
        rb1, rb2, rb3 = st.columns(3)
        with rb1:
            brief_metric = st.selectbox("Metric", ["None"] + nums, key="brief_metric_v5")
        with rb2:
            brief_category = st.selectbox("Category", ["None"] + df.columns.tolist(), key="brief_category_v5")
        with rb3:
            brief_date = st.selectbox("Date", ["None"] + (date_like_cols(df) or df.columns.tolist()), key="brief_date_v5")
        if st.button("Build Report Brief", type="primary", use_container_width=True):
            brief = build_report_brief(
                df,
                None if brief_metric == "None" else brief_metric,
                None if brief_category == "None" else brief_category,
                None if brief_date == "None" else brief_date,
            )
            st.session_state.report_brief_df = brief
            add_audit("Report brief", f"metric={brief_metric}; category={brief_category}; date={brief_date}")
        if not st.session_state.report_brief_df.empty:
            st.dataframe(st.session_state.report_brief_df, use_container_width=True)

    with tool_tabs[23]:
        # NEW (v6): Sweetviz-style dataset comparison (current vs an uploaded file). [PRO]
        st.subheader("Compare Two Datasets  ·  Pro")
        st.write("Compare the current processed dataset against another file (e.g., last month vs this month) — column presence, types, missing % and means side by side. No AI API.")
        if require_feature("dataset_comparison", "Dataset comparison"):
            cmp_file = st.file_uploader("Upload comparison CSV/XLSX", type=["csv", "xlsx", "xls"], key="compare_file")
            cl1, cl2 = st.columns(2)
            with cl1:
                label_a = st.text_input("Label for current dataset", value="Current", key="cmp_label_a")
            with cl2:
                label_b = st.text_input("Label for comparison dataset", value="Comparison", key="cmp_label_b")
            if cmp_file is not None and st.button("Run Dataset Comparison", type="primary", use_container_width=True):
                try:
                    if cmp_file.name.lower().endswith(".csv"):
                        other = try_read_csv(cmp_file)
                    else:
                        other = pd.read_excel(cmp_file)
                    comparison = compare_two_datasets(df, other, label_a or "Current", label_b or "Comparison")
                    st.session_state["compare_df"] = comparison
                    add_audit("Dataset comparison", f"Compared current ({df.shape}) with {cmp_file.name} ({other.shape}).")
                except Exception as exc:
                    st.error(f"Could not compare datasets: {exc}")
            if "compare_df" in st.session_state and not st.session_state["compare_df"].empty:
                st.dataframe(st.session_state["compare_df"], use_container_width=True, hide_index=True)
                st.download_button(
                    "Download Comparison CSV",
                    data=st.session_state["compare_df"].to_csv(index=False).encode("utf-8"),
                    file_name="dataset_comparison.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

    with tool_tabs[24]:
        # NEW (v2): PII detection & masking [PRO]
        st.subheader("Privacy: PII Detection & Masking  ·  Pro")
        st.write("Scan the dataset for likely personal data (emails, phones, IDs, addresses) using rule-based "
                 "detection — no AI — then mask selected columns before exporting or sharing.")
        if require_feature("pii_masking", "PII detection & masking"):
            st.markdown("**Step 1 — Scan for likely PII columns**")
            pii_report = detect_pii_columns(df)
            st.dataframe(pii_report, use_container_width=True, hide_index=True)
            st.markdown("**Step 2 — Mask selected columns**")
            mc1, mc2 = st.columns([2, 1])
            with mc1:
                mask_cols = st.multiselect("Columns to mask", df.columns.tolist(), key="pii_mask_cols")
            with mc2:
                mask_mode = st.selectbox("Mask mode", ["Partial", "Full (hash)"], key="pii_mask_mode")
            if mask_cols and st.button("Apply Masking to Working Dataset", type="primary", use_container_width=True):
                st.session_state.working_df = mask_pii_columns(df, mask_cols, mask_mode)
                add_audit("PII masking", f"Masked {len(mask_cols)} column(s) using {mask_mode} mode: {mask_cols}.")
                st.success(f"Masked {len(mask_cols)} column(s). The working dataset is updated for all downstream exports.")
                st.rerun()
            st.caption("Tip: 'Partial' keeps first/last characters; 'Full (hash)' replaces the value with an irreversible token.")

    with tool_tabs[25]:
        # License key generator [ENTERPRISE / owner only]. v3: selectable licence model.
        st.subheader("License Administration  ·  Enterprise")
        st.write("Generate signed subscription license keys for customers. Keys are HMAC‑signed with your "
                 "`LICENSE_SIGNING_KEY` secret, so they **cannot be forged, re‑tiered, extended, model‑switched "
                 "or seat‑rebound** without that secret. Choose the licence model per client.")
        if require_feature("license_admin", "License administration"):
            using_demo_key = _get_secret("LICENSE_SIGNING_KEY", "") == ""
            if using_demo_key:
                st.error("⚠️ No `LICENSE_SIGNING_KEY` is set — keys are being signed with the PUBLIC demo key and are NOT secure. "
                         "Set `LICENSE_SIGNING_KEY` in Streamlit Secrets before issuing real licenses (see DEPLOYMENT.md).")
            st.markdown("**Licence model**")
            lic_model = st.radio(
                "Choose how this client's licence behaves",
                ["Team (shareable across many devices)",
                 "Per-Seat: Named (bound to a Seat ID you type)",
                 "Per-Seat: Device-bound (auto device fingerprint)"],
                key="adm_model", horizontal=False,
                help="Team = one key for a whole school/team. Named per-seat = key only activates for the exact Seat ID you set. "
                     "Device-bound per-seat = key auto-binds to the customer's specific device — they never type anything.",
            )
            model_code = "seat" if lic_model.startswith("Per-Seat") else "team"
            seat_is_device = lic_model.startswith("Per-Seat: Device")
            lc1, lc2, lc3 = st.columns(3)
            with lc1:
                lic_customer = st.text_input("Customer / organisation", value="Example School", key="adm_cust")
            with lc2:
                lic_tier = st.selectbox("Tier", ["pro", "enterprise"], key="adm_tier")
            with lc3:
                lic_expiry = st.text_input("Expiry (YYYY-MM-DD)", value="2027-12-31", key="adm_exp")
            lic_seat = ""
            if model_code == "seat" and not seat_is_device:
                lic_seat = st.text_input("Seat ID to bind (e.g. staff email or device label)", value="", key="adm_seat",
                                         placeholder="staff@school.edu  or  Lab-PC-01")
            elif model_code == "seat" and seat_is_device:
                st.markdown("**Device-bound binding**")
                st.caption("Ask the customer to open the app, expand the licence box, and send you the **automatic device ID** shown there. "
                           "Paste it below. (Or bind to THIS device using the captured ID.)")
                render_device_fingerprint()
                _this_dev = get_device_id()
                if _this_dev:
                    st.caption(f"This device's ID: `{_this_dev}`")
                lic_seat = st.text_input("Customer device ID to bind", value="", key="adm_seat_dev",
                                         placeholder="DEV-xxxxxxxxyyyyyyyy")
                if st.button("Use THIS device's ID", key="adm_use_this_dev") and _this_dev:
                    st.session_state["adm_seat_dev"] = _this_dev
                    st.rerun()
            if st.button("Generate License Key", type="primary", use_container_width=True):
                if model_code == "seat" and not lic_seat.strip():
                    st.error("Per-Seat model requires a Seat ID to bind the key to.")
                else:
                    key = issue_license_key(lic_customer, lic_tier, lic_expiry, model_code, lic_seat)
                    check = verify_license_key(key, lic_seat)
                    add_audit("License issued", f"{lic_tier} ({model_code}) for '{lic_customer}' until {lic_expiry}" + (f"; seat={lic_seat}" if lic_seat else "") + ".")
                    log_usage_event("License issued", tier=lic_tier, model=model_code, customer=lic_customer, detail=f"until {lic_expiry}")
                    st.success("License key generated. Copy and send it to the customer:")
                    st.code(key, language="text")
                    if model_code == "seat" and seat_is_device:
                        st.info(f"This is a **device-bound** per-seat key (device `{lic_seat}`). It activates "
                                "automatically only on that device — the customer types nothing.")
                    elif model_code == "seat":
                        st.info(f"This is a **named** per-seat key. The customer must enter Seat ID **{lic_seat}** to activate it. "
                                "It will not work for any other seat.")
                    else:
                        st.info("This is a **team** key. It activates without a Seat ID and can be used across the client's devices.")
                    st.caption(f"Self-check: {'✅ valid' if check['valid'] else '❌ ' + check['reason']} · "
                               f"tier={check['tier']} · model={check['model']} · expires {check['expiry']}")
            st.divider()
            st.markdown("**Validate an existing key**")
            test_key = st.text_input("Paste a key to validate", key="adm_test_key")
            test_seat = st.text_input("Seat ID (for per-seat keys)", key="adm_test_seat")
            if st.button("Validate Key", use_container_width=True):
                info = verify_license_key(test_key, test_seat)
                if info["valid"]:
                    st.success(f"Valid · {TIER_LABELS.get(info['tier'])} · {info['model']} · {info['customer']} · expires {info['expiry']}")
                else:
                    st.error(info["reason"])

    with tool_tabs[26]:
        # NEW (v4): Data Quality Rules engine [PRO]
        st.subheader("Data Quality Rules Engine  ·  Pro")
        st.write("Define multiple data-quality rules and run them at once for a pass/fail scorecard. No AI.")
        if require_feature("data_quality_rules", "Data Quality Rules engine"):
            if "dq_rules" not in st.session_state:
                st.session_state.dq_rules = []
            rq1, rq2, rq3, rq4 = st.columns([2, 2, 2, 1])
            with rq1:
                dq_col = st.selectbox("Column", df.columns.tolist(), key="dq_col")
            with rq2:
                dq_check = st.selectbox("Check", ["not_null", "unique", "positive", "min", "max", "allowed", "regex"], key="dq_check")
            with rq3:
                dq_param = st.text_input("Parameter (min/max/list/regex)", key="dq_param")
            with rq4:
                st.write("")
                if st.button("➕ Add", use_container_width=True):
                    st.session_state.dq_rules.append({"column": dq_col, "check": dq_check, "param": dq_param})
            if st.session_state.dq_rules:
                st.caption("Active rules:")
                st.dataframe(pd.DataFrame(st.session_state.dq_rules), use_container_width=True, hide_index=True)
                cc1, cc2 = st.columns(2)
                with cc1:
                    if st.button("Run All Rules", type="primary", use_container_width=True):
                        rep = build_data_quality_rules(df, st.session_state.dq_rules)
                        st.session_state["dq_report"] = rep
                        add_audit("Data quality rules", f"Ran {len(st.session_state.dq_rules)} rule(s).")
                with cc2:
                    if st.button("Clear Rules", use_container_width=True):
                        st.session_state.dq_rules = []
                        st.session_state.pop("dq_report", None)
                        st.rerun()
            if "dq_report" in st.session_state:
                rep = st.session_state["dq_report"]
                st.dataframe(rep, use_container_width=True, hide_index=True)
                if "Status" in rep.columns:
                    passed = int((rep["Status"] == "PASS").sum())
                    st.metric("Rules Passed", f"{passed}/{len(rep)}")
                st.download_button("Download DQ Report CSV", rep.to_csv(index=False).encode("utf-8"), "data_quality_report.csv", "text/csv", use_container_width=True)

    with tool_tabs[27]:
        # NEW (v4): Trend & seasonality decomposition [PRO]
        st.subheader("Trend & Seasonality Decomposition  ·  Pro")
        st.write("Break a time series into trend, seasonality and residual using a lightweight moving-average method (no statsmodels, no AI).")
        if require_feature("trend_decomposition", "Trend decomposition"):
            dates = date_like_cols(df); nums = numeric_cols(df)
            if dates and nums:
                td1, td2, td3 = st.columns(3)
                with td1:
                    td_date = st.selectbox("Date column", dates, key="td_date")
                with td2:
                    td_metric = st.selectbox("Metric", nums, key="td_metric")
                with td3:
                    td_freq = st.selectbox("Frequency", ["M", "Q", "W", "D"], key="td_freq")
                if st.button("Decompose Series", type="primary", use_container_width=True):
                    dec = build_trend_decomposition(df, td_date, td_metric, td_freq)
                    st.session_state["trend_decomp"] = dec
                    add_audit("Trend decomposition", f"{td_metric} over {td_date} @ {td_freq}.")
                if "trend_decomp" in st.session_state and not st.session_state["trend_decomp"].empty:
                    dec = st.session_state["trend_decomp"]
                    st.dataframe(dec, use_container_width=True, hide_index=True)
                    if "Trend" in dec.columns:
                        st.line_chart(dec.set_index("Period")[["Actual", "Trend"]])
                    st.download_button("Download Decomposition CSV", dec.to_csv(index=False).encode("utf-8"), "trend_decomposition.csv", "text/csv", use_container_width=True)
            else:
                st.info("A date-like column and a numeric metric are required.")

    with tool_tabs[28]:
        # NEW (v4): Correlation insights [FREE]
        st.subheader("Correlation Insights")
        st.write("Automatically surface the strongest numeric relationships with plain-English notes. Rule-based, no AI.")
        ci_threshold = st.slider("Minimum |correlation|", 0.3, 0.95, 0.6, 0.05, key="ci_threshold")
        if st.button("Find Strong Correlations", type="primary", use_container_width=True):
            ci = build_correlation_insights(df, ci_threshold)
            st.session_state["corr_insights"] = ci
            add_audit("Correlation insights", f"threshold={ci_threshold}")
        if "corr_insights" in st.session_state:
            st.dataframe(st.session_state["corr_insights"], use_container_width=True, hide_index=True)
        nums = numeric_cols(df)
        if len(nums) >= 2:
            with st.expander("Show full correlation heatmap"):
                fig = px.imshow(df[nums].corr(numeric_only=True), text_auto=".2f", color_continuous_scale="RdBu_r", zmin=-1, zmax=1)
                st.plotly_chart(fig, use_container_width=True)

    with tool_tabs[29]:
        # NEW (v4): Auto clean advisor [FREE]
        st.subheader("Cleaning Advisor")
        st.write("Scan the dataset and get concrete, prioritised cleaning suggestions (missing values, whitespace, casing, numbers-as-text, constant columns). Rule-based, no AI.")
        if st.button("Scan for Cleaning Issues", type="primary", use_container_width=True):
            adv = build_auto_clean_suggestions(df)
            st.session_state["clean_advice"] = adv
            add_audit("Cleaning advisor", "Scan run.")
        if "clean_advice" in st.session_state:
            st.dataframe(st.session_state["clean_advice"], use_container_width=True, hide_index=True)
            st.caption("Apply fixes in the Clean tab (3️⃣) — this advisor tells you what to do; the Clean tab does it.")

    with tool_tabs[30]:
        # NEW (v4): Survey / Likert analysis [FREE]
        st.subheader("Survey / Likert Analysis")
        st.write("Summarise rating-scale questions (1–5 or 1–7): mean, median, top-2-box and bottom-2-box percentages. No AI.")
        nums = numeric_cols(df)
        likert_cols = st.multiselect("Likert/rating columns", nums, key="likert_cols")
        if likert_cols and st.button("Analyse Survey", type="primary", use_container_width=True):
            lk = build_likert_summary(df, likert_cols)
            st.session_state["likert_df"] = lk
            add_audit("Likert analysis", f"{len(likert_cols)} question(s).")
        if "likert_df" in st.session_state:
            st.dataframe(st.session_state["likert_df"], use_container_width=True, hide_index=True)
            st.download_button("Download Survey Summary CSV", st.session_state["likert_df"].to_csv(index=False).encode("utf-8"), "survey_summary.csv", "text/csv", use_container_width=True)

    with tool_tabs[31]:
        # NEW (v4): Benford's Law fraud screen [ENTERPRISE]
        st.subheader("Benford's Law Fraud Screen  ·  Enterprise")
        st.write("Audit-grade first-digit test: compares your numbers' leading-digit distribution to Benford's Law to flag possible manipulation. No AI.")
        if require_feature("benford_fraud", "Benford's Law fraud screen"):
            nums = numeric_cols(df)
            if nums:
                bf_col = st.selectbox("Numeric column to test", nums, key="bf_col")
                if st.button("Run Benford Test", type="primary", use_container_width=True):
                    bf, verdict = build_benford_analysis(df, bf_col)
                    st.session_state["benford_df"] = bf
                    st.session_state["benford_verdict"] = verdict
                    add_audit("Benford test", f"{bf_col}: {verdict}")
                if "benford_df" in st.session_state and not st.session_state["benford_df"].empty:
                    bf = st.session_state["benford_df"]
                    st.dataframe(bf, use_container_width=True, hide_index=True)
                    st.bar_chart(bf.set_index("Leading Digit")[["Observed %", "Benford Expected %"]])
                    st.info(st.session_state.get("benford_verdict", ""))
                    st.download_button("Download Benford CSV", bf.to_csv(index=False).encode("utf-8"), "benford_analysis.csv", "text/csv", use_container_width=True)
                elif "benford_verdict" in st.session_state:
                    st.warning(st.session_state["benford_verdict"])
            else:
                st.info("At least one numeric column is required.")

    with tool_tabs[32]:
        # NEW (v5): Cross-tab + chi-square test of independence [PRO]
        st.subheader("Cross-tab & Chi-Square Association  ·  Pro")
        st.write("Build a contingency table between two categorical columns and test whether they are related "
                 "(chi-square + Cramér's V effect size). Pure numpy — no AI, no scipy.")
        if require_feature("crosstab_chisquare", "Cross-tab & chi-square"):
            cats = text_cols(df) + [c for c in df.columns if df[c].nunique(dropna=True) <= 30 and c not in text_cols(df)]
            cats = list(dict.fromkeys(cats))
            if len(cats) >= 2:
                xc1, xc2 = st.columns(2)
                with xc1:
                    ct_row = st.selectbox("Row column", cats, key="ct_row")
                with xc2:
                    ct_col = st.selectbox("Column column", [c for c in cats if c != ct_row], key="ct_col")
                if st.button("Build Cross-tab & Test", type="primary", use_container_width=True):
                    table, verdict = build_crosstab_chisquare(df, ct_row, ct_col)
                    st.session_state["crosstab_df"] = table
                    st.session_state["crosstab_verdict"] = verdict
                    add_audit("Cross-tab chi-square", f"{ct_row} x {ct_col}: {verdict}")
                if "crosstab_df" in st.session_state and not st.session_state["crosstab_df"].empty:
                    st.dataframe(st.session_state["crosstab_df"], use_container_width=True, hide_index=True)
                    st.info(st.session_state.get("crosstab_verdict", ""))
                    st.download_button("Download Cross-tab CSV", st.session_state["crosstab_df"].to_csv(index=False).encode("utf-8"), "crosstab.csv", "text/csv", use_container_width=True)
            else:
                st.info("Two categorical (or low-cardinality) columns are required.")

    with tool_tabs[33]:
        # NEW (v5): Text/keyword frequency [FREE]
        st.subheader("Text / Keyword Frequency")
        st.write("Count the most common words in a free-text column (e.g. survey open-ends, comments). "
                 "Simple stopword filtering — no AI, no NLP library.")
        txt_cols = text_cols(df)
        if txt_cols:
            tf1, tf2, tf3 = st.columns(3)
            with tf1:
                tf_col = st.selectbox("Text column", txt_cols, key="tf_col")
            with tf2:
                tf_top = st.slider("Top N keywords", 10, 100, 30, key="tf_top")
            with tf3:
                tf_min = st.slider("Min word length", 2, 6, 3, key="tf_min")
            if st.button("Count Keywords", type="primary", use_container_width=True):
                tfreq = build_text_frequency(df, tf_col, tf_top, tf_min)
                st.session_state["text_freq_df"] = tfreq
                add_audit("Text frequency", f"column={tf_col}")
            if "text_freq_df" in st.session_state:
                tfreq = st.session_state["text_freq_df"]
                st.dataframe(tfreq, use_container_width=True, hide_index=True)
                if "Keyword" in tfreq.columns and "Count" in tfreq.columns:
                    st.bar_chart(tfreq.head(20).set_index("Keyword")["Count"])
                    st.download_button("Download Keyword Counts CSV", tfreq.to_csv(index=False).encode("utf-8"), "keyword_frequency.csv", "text/csv", use_container_width=True)
        else:
            st.info("No text columns detected.")

    with tool_tabs[34]:
        # NEW (v5): Number/locale normalizer [FREE]
        st.subheader("Number / Locale Normalizer")
        st.write("Convert locale-formatted number text (e.g. `1.234,56` European or `1,234.56` US) into clean numbers. "
                 "Deterministic parsing — no AI.")
        nn1, nn2, nn3 = st.columns(3)
        with nn1:
            nn_cols = st.multiselect("Columns to normalize", df.columns.tolist(), key="nn_cols")
        with nn2:
            nn_dec = st.selectbox("Decimal separator", [".", ","], key="nn_dec")
        with nn3:
            nn_thou = st.selectbox("Thousands separator", [",", ".", "space", "none"], key="nn_thou")
        thou_map = {"space": " ", "none": ""}
        thou_val = thou_map.get(nn_thou, nn_thou)
        if nn_cols and st.button("Normalize Numbers", type="primary", use_container_width=True):
            new_df, converted = normalize_number_locale(df, nn_cols, nn_dec, thou_val)
            st.session_state.working_df = new_df
            add_audit("Number normalize", f"{len(nn_cols)} column(s); {converted} cell(s) converted; dec='{nn_dec}', thou='{nn_thou}'.")
            st.success(f"Converted {converted} cell(s) across {len(nn_cols)} column(s). Working dataset updated.")
            st.rerun()

    with tool_tabs[35]:
        # NEW (v5): Data classification & watermark [PRO]
        st.subheader("Data Classification & Watermark  ·  Pro")
        st.write("Label this dataset with a confidentiality level. The label is recorded in the audit trail and "
                 "added as a watermark column for exports, supporting governance and information-handling policy.")
        if require_feature("data_classification", "Data classification"):
            cls = st.selectbox("Classification level", CLASSIFICATION_LEVELS, key="cls_level",
                               index=CLASSIFICATION_LEVELS.index(st.session_state.get("data_classification", "Internal")) if st.session_state.get("data_classification") in CLASSIFICATION_LEVELS else 1)
            cls_owner = st.text_input("Data owner / handler", value=st.session_state.get("data_owner", ""), key="cls_owner")
            cc1, cc2 = st.columns(2)
            with cc1:
                if st.button("Apply Classification", type="primary", use_container_width=True):
                    st.session_state["data_classification"] = cls
                    st.session_state["data_owner"] = cls_owner
                    add_audit("Data classification", f"Set to {cls}; owner={cls_owner or 'n/a'}.")
                    log_usage_event("Data classification", detail=cls)
                    st.success(f"Dataset classified as **{cls}**.")
            with cc2:
                if st.button("Add Watermark Column", use_container_width=True):
                    wm = df.copy()
                    wm["__classification"] = st.session_state.get("data_classification", cls)
                    wm["__handled_by"] = st.session_state.get("data_owner", cls_owner)
                    st.session_state.working_df = wm
                    add_audit("Watermark added", f"classification={cls}")
                    st.success("Watermark columns added to the working dataset.")
                    st.rerun()
            current = st.session_state.get("data_classification")
            if current:
                badge = {"Public": "#16a34a", "Internal": "#0B3954", "Confidential": "#d97706", "Restricted": "#b91c1c"}.get(current, "#0B3954")
                st.markdown(f"<span class='tier-badge' style='background:{badge}'>Current: {current}</span>", unsafe_allow_html=True)

    with tool_tabs[36]:
        # NEW (v5): Usage / License analytics dashboard [ENTERPRISE / owner]
        st.subheader("Usage & License Analytics  ·  Enterprise")
        st.write("Owner dashboard of recorded platform events: logins, license activations/upgrades and key issuance. "
                 "Stored in a local SQLite log (no external service, no AI).")
        if require_feature("usage_analytics", "Usage & license analytics"):
            events = load_usage_events()
            if events.empty:
                st.info("No usage events recorded yet. Events accumulate as people log in and activate licences. "
                        "Note: on Streamlit's free tier the local log may reset when the app restarts; set `USAGE_DB_PATH` "
                        "to a persistent location for long-term retention.")
            else:
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total events", f"{len(events):,}")
                m2.metric("Logins", f"{int((events['event']=='Authentication').sum()):,}")
                m3.metric("Activations", f"{int(events['event'].isin(['License activation','License upgrade']).sum()):,}")
                m4.metric("Keys issued", f"{int((events['event']=='License issued').sum()):,}")
                summ = build_usage_summary(events)
                if "by_day" in summ and not summ["by_day"].empty:
                    st.markdown("**Events per day**")
                    st.bar_chart(summ["by_day"].set_index("date")["Events"])
                cda, cdb = st.columns(2)
                with cda:
                    if "by_event" in summ:
                        st.markdown("**By event type**")
                        st.dataframe(summ["by_event"], use_container_width=True, hide_index=True)
                    if "by_tier" in summ and not summ["by_tier"].empty:
                        st.markdown("**By tier**")
                        st.dataframe(summ["by_tier"], use_container_width=True, hide_index=True)
                with cdb:
                    if "by_customer" in summ:
                        st.markdown("**By customer (licences)**")
                        st.dataframe(summ["by_customer"], use_container_width=True, hide_index=True)
                    if "by_model" in summ:
                        st.markdown("**By licence model**")
                        st.dataframe(summ["by_model"], use_container_width=True, hide_index=True)
                st.markdown("**Raw event log**")
                st.dataframe(events.head(1000), use_container_width=True, hide_index=True)
                ec1, ec2 = st.columns(2)
                with ec1:
                    st.download_button("⬇️ Download Usage Log CSV", events.to_csv(index=False).encode("utf-8"), "usage_log.csv", "text/csv", use_container_width=True)
                with ec2:
                    if st.button("Clear Usage Log", use_container_width=True):
                        clear_usage_events()
                        add_audit("Usage log cleared", "Owner cleared usage analytics log.")
                        st.rerun()

    with tool_tabs[37]:
        # NEW (v6): Executive Scorecard [FREE]
        st.subheader("Executive Scorecard")
        st.write("A compact one-look KPI scorecard (totals, average, top contributor, growth). Rule-based, no AI.")
        nums = numeric_cols(df)
        if nums:
            es1, es2, es3 = st.columns(3)
            with es1:
                es_metric = st.selectbox("Metric", nums, key="es_metric")
            with es2:
                es_cat = st.selectbox("Category (optional)", ["None"] + text_cols(df), key="es_cat")
            with es3:
                es_date = st.selectbox("Date (optional)", ["None"] + date_like_cols(df), key="es_date")
            if st.button("Build Scorecard", type="primary", use_container_width=True):
                sc = build_executive_scorecard(df, es_metric,
                                               None if es_cat == "None" else es_cat,
                                               None if es_date == "None" else es_date)
                st.session_state["exec_scorecard"] = sc
                add_audit("Executive scorecard", f"metric={es_metric}")
            if "exec_scorecard" in st.session_state:
                sc = st.session_state["exec_scorecard"]
                st.dataframe(sc, use_container_width=True, hide_index=True)
                st.download_button("Download Scorecard CSV", sc.to_csv(index=False).encode("utf-8"), "executive_scorecard.csv", "text/csv", use_container_width=True)
        else:
            st.info("At least one numeric column is required.")

    with tool_tabs[38]:
        # NEW (v6): Concentration / Gini [PRO]
        st.subheader("Concentration & Gini Index  ·  Pro")
        st.write("Measure how concentrated a metric is across categories (e.g. revenue per customer) using the Gini "
                 "coefficient and a cumulative-share table. Flags dependency/risk. Pure numpy, no AI.")
        if require_feature("gini_concentration", "Concentration & Gini"):
            cats = text_cols(df); nums = numeric_cols(df)
            if cats and nums:
                gc1, gc2 = st.columns(2)
                with gc1:
                    g_cat = st.selectbox("Category", cats, key="gini_cat")
                with gc2:
                    g_metric = st.selectbox("Metric", nums, key="gini_metric")
                if st.button("Compute Concentration", type="primary", use_container_width=True):
                    table, verdict = build_gini_concentration(df, g_cat, g_metric)
                    st.session_state["gini_df"] = table
                    st.session_state["gini_verdict"] = verdict
                    add_audit("Concentration/Gini", f"{g_metric} by {g_cat}: {verdict[:60]}")
                if "gini_df" in st.session_state and not st.session_state["gini_df"].empty:
                    st.dataframe(st.session_state["gini_df"], use_container_width=True, hide_index=True)
                    st.info(st.session_state.get("gini_verdict", ""))
                    st.download_button("Download Concentration CSV", st.session_state["gini_df"].to_csv(index=False).encode("utf-8"), "concentration.csv", "text/csv", use_container_width=True)
            else:
                st.info("A category and a numeric metric are required.")

    with tool_tabs[39]:
        # NEW (v6): Snapshot diff / changelog [PRO]
        st.subheader("Snapshot Diff (row-level change log)  ·  Pro")
        st.write("Compare the current dataset to an earlier snapshot file keyed on an ID column, and list rows "
                 "Added / Removed / Changed. Great for tracking master-data changes over time. No AI.")
        if require_feature("snapshot_diff", "Snapshot diff"):
            diff_file = st.file_uploader("Upload earlier snapshot CSV/XLSX", type=["csv", "xlsx", "xls"], key="diff_file")
            key_col = st.selectbox("Key/ID column", df.columns.tolist(), key="diff_key")
            if diff_file is not None and st.button("Compute Changes", type="primary", use_container_width=True):
                try:
                    old = try_read_csv(diff_file) if diff_file.name.lower().endswith(".csv") else pd.read_excel(diff_file)
                    diff = build_changelog_diff(old, df, key_col)
                    st.session_state["snapshot_diff"] = diff
                    add_audit("Snapshot diff", f"key={key_col}; {len(diff)} change row(s).")
                except Exception as exc:
                    st.error(f"Could not compute diff: {exc}")
            if "snapshot_diff" in st.session_state and not st.session_state["snapshot_diff"].empty:
                st.dataframe(st.session_state["snapshot_diff"], use_container_width=True, hide_index=True)
                st.download_button("Download Change Log CSV", st.session_state["snapshot_diff"].to_csv(index=False).encode("utf-8"), "snapshot_diff.csv", "text/csv", use_container_width=True)

    with tool_tabs[40]:
        # NEW (v6): Report delivery (email + WhatsApp/mailto) [PRO]
        st.subheader("Report Delivery (Email & WhatsApp)  ·  Pro")
        st.write("Email a report attachment via your free SMTP account, or generate WhatsApp / email share links. "
                 "Uses Python's built-in email tools — no paid API, no AI.")
        if require_feature("report_delivery", "Report delivery"):
            st.markdown("**What to send**")
            rep_kind = st.radio("Attachment", ["Processed dataset (CSV)", "Offline HTML report", "No attachment (links only)"], key="rep_kind", horizontal=False)
            rd1, rd2 = st.columns(2)
            with rd1:
                rep_to = st.text_input("Recipient email", key="rep_to", placeholder="client@example.com")
                rep_subject = st.text_input("Subject", value=f"{ACTIVE_BRAND['brand_name']} — Data Report", key="rep_subject")
            with rd2:
                rep_phone = st.text_input("WhatsApp number (for share link)", key="rep_phone", placeholder="2348100866322")
                rep_body = st.text_area("Message body", value="Please find the attached data report generated with the HMG Excel Operations Platform.", key="rep_body", height=90)
            # Build attachment
            attach_bytes, attach_name, sub = None, "report.csv", "csv"
            if rep_kind.startswith("Processed"):
                attach_bytes = df.to_csv(index=False).encode("utf-8"); attach_name = "processed_dataset.csv"; sub = "csv"
            elif rep_kind.startswith("Offline"):
                ins = build_auto_insights(df)
                attach_bytes = build_html_eda_report(df, f"{ACTIVE_BRAND['brand_name']} Report", ins).encode("utf-8")
                attach_name = "data_report.html"; sub = "html"
            cda, cdb = st.columns(2)
            with cda:
                if st.button("✉️ Send via SMTP Email", type="primary", use_container_width=True):
                    ok, msg = send_email_with_attachment(
                        rep_to, rep_subject, rep_body,
                        attach_bytes, attach_name, "text", sub,
                    )
                    (st.success if ok else st.error)(msg)
                    if ok:
                        add_audit("Report emailed", f"to={rep_to}; attach={attach_name}")
                        log_usage_event("Report emailed", customer=rep_to, detail=attach_name)
            with cdb:
                wa = whatsapp_share_link(rep_phone, rep_body)
                ml = mailto_link(rep_to, rep_subject, rep_body)
                st.markdown(f"[📱 Open WhatsApp share link]({wa})")
                st.markdown(f"[📧 Open email client (mailto)]({ml})")
            st.caption("Tip: WhatsApp/mailto links open the user's own app and need no server config. "
                       "SMTP sending requires SMTP_* secrets (see DEPLOYMENT.md §L). Attachments can't ride on wa.me links — "
                       "download the file and attach it in WhatsApp, or use SMTP email for automatic attachment.")

    with tool_tabs[41]:
        # NEW (v6): White-label / multi-tenant branding [ENTERPRISE]
        st.subheader("White-Label Branding  ·  Enterprise")
        st.write("Re-skin the whole platform for a client: brand name, tagline, logo, colours and footer. "
                 "Apply for this session, or copy a JSON to set it permanently per deployment via the "
                 "`WHITE_LABEL_BRAND` secret. Same code, many tenants — no AI.")
        if require_feature("white_label", "White-label branding"):
            b = get_active_brand()
            wl1, wl2 = st.columns(2)
            with wl1:
                wl_name = st.text_input("Brand name", value=b["brand_name"], key="wl_name")
                wl_tag = st.text_input("Tagline", value=b["tagline"], key="wl_tag")
                wl_logo = st.text_input("Logo image URL (optional)", value=b.get("logo_url", ""), key="wl_logo")
                wl_emoji = st.text_input("Logo emoji (fallback)", value=b.get("logo_emoji", "📊"), key="wl_emoji")
            with wl2:
                wl_bg = st.color_picker("Header start colour", value=b["header_bg"], key="wl_bg")
                wl_mid = st.color_picker("Header mid colour", value=b["header_mid"], key="wl_mid")
                wl_to = st.color_picker("Header end colour", value=b["header_to"], key="wl_to")
                wl_accent = st.color_picker("Accent colour", value=b["accent"], key="wl_accent")
            wl_footer = st.text_input("Footer text", value=b["footer"], key="wl_footer")
            override = {
                "brand_name": wl_name, "tagline": wl_tag, "logo_url": wl_logo, "logo_emoji": wl_emoji,
                "header_bg": wl_bg, "header_mid": wl_mid, "header_to": wl_to, "accent": wl_accent,
                "footer": wl_footer,
            }
            ca, cb, cc = st.columns(3)
            with ca:
                if st.button("Apply for this session", type="primary", use_container_width=True):
                    st.session_state["brand_override"] = override
                    add_audit("White-label applied", f"brand={wl_name}")
                    st.rerun()
            with cb:
                if st.button("Reset to HMG default", use_container_width=True):
                    st.session_state.pop("brand_override", None)
                    st.rerun()
            with cc:
                st.download_button("Download brand JSON", json.dumps(override, indent=2).encode("utf-8"), "white_label_brand.json", "application/json", use_container_width=True)
            st.markdown("**Permanent (per deployment):** paste this into the `WHITE_LABEL_BRAND` secret:")
            st.code(json.dumps(override), language="json")

    with tool_tabs[42]:
        st.subheader("HMG Brand and Ecosystem Profile")
        st.write("This platform is part of the HMG Technologies Ecosystem under HMG Concepts. The founder and brand details are embedded in the UI and exported workbook.")
        st.markdown("""
        **Founder / Visioner:** Adewale Samson Adeagbo  
        **Identity:** Data Scientist · STEM Educator · AI-Augmented Solutions Developer  
        **Parent Brand:** HMG Concepts — His Marvellous Grace Educational Consult  
        **Founded:** 2015  
        **Ecosystem:** HMG Academy · HMG Technologies · HMG Media  
        **Philosophy:** Learning Deliberately. Teaching Authentically.  
        **Location:** Lagos / Ogun State, Nigeria  
        **WhatsApp:** +234 810 086 6322  
        **Tech / Partnerships:** buildingmyictcareer@gmail.com  
        """)
        st.dataframe(build_hmg_brand_profile(), use_container_width=True)
        st.dataframe(build_hmg_ecosystem_map(), use_container_width=True)
        st.markdown("""
        - Portfolio: https://cssadewale.pages.dev  
        - HMG Concepts: https://hmgconcepts.pages.dev  
        - HMG Academy: https://hmgacademy.pages.dev  
        - GitHub: https://github.com/cssadewale  
        - YouTube: https://youtube.com/@hmgconcepts
        """)

with tabs[8]:
    st.header("📘 Feature Guide and System Explanation")
    st.write("The platform extends the original dashboard generator into a broader Excel operations automation system while retaining the original dashboard, audit, and pivot features.")
    grid_cols = st.columns(2)
    for idx, (title, desc) in enumerate(FEATURE_CATALOG):
        with grid_cols[idx % 2]:
            st.markdown(f"<div class='mini-card'><b>{title}</b><br>{desc}</div>", unsafe_allow_html=True)
            st.write("")

    st.subheader("Sheets generated in the Excel workbook")
    st.markdown("""
    - **Audit_Log:** who/when/what happened in the workflow.
    - **Data_Profile:** dataset overview plus column-level diagnostics.
    - **Missing_Report:** missing count and percentage by column.
    - **Duplicate_Report:** duplicate rows for review.
    - **Outlier_Report:** IQR outlier fences and counts.
    - **Raw_Data:** original uploaded master data with filters and heatmaps.
    - **Processed_Data:** final cleaned/transformed dataset with filters and heatmaps.
    - **Descriptive_Stats:** numeric summaries similar to Excel Data Analysis ToolPak.
    - **Correlation_Matrix:** relationship matrix across numeric fields.
    - **Grouped_Summary:** analyst-created group-by output.
    - **Pivot_Analysis:** analyst-created pivot/crosstab output.
    - **Forecast:** simple no-API forecast output when generated.
    - **Validation_Report:** rule-based data quality exception report.
    - **What_If_Analysis:** scenario modelling output.
    - **Pareto_Analysis:** 80/20 contribution analysis.
    - **Data_Dictionary:** editable metadata template for every field.
    - **Quality_Scorecard:** rule-based data quality scorecard.
    - **Reconciliation:** missing-record and value-mismatch report.
    - **Time_Intelligence:** period trend, previous-period, YTD, and rolling average output.
    - **KPI_Targets:** actual-versus-target performance tracking.
    - **SQL_Query_Result:** result of safe SQL SELECT/WITH query.
    - **Anomaly_Report:** detailed IQR or z-score row-level anomaly report.
    - **RFM_Analysis:** recency, frequency and monetary segmentation.
    - **Cohort_Analysis:** monthly cohort retention/value analysis.
    - **Executive_Summary:** deterministic non-AI executive summary.
    - **Schema_Contract:** data contract/governance checks.
    - **Fuzzy_Duplicates:** near-duplicate text match report.
    - **ABC_Analysis:** ABC category contribution classification.
    - **Basket_Analysis:** item-pair/co-occurrence report.
    - **Weighted_Scorecard:** weighted decision matrix and ranking.
    - **Goal_Seek:** target gap and required lift calculator.
    - **Calendar_Table:** date dimension table for Excel/Power BI.
    - **Report_Brief:** rule-based findings and recommended actions.
    - **HMG_Brand_Profile:** Adewale Samson Adeagbo and HMG brand details.
    - **HMG_Ecosystem:** HMG Concepts, Academy, Technologies and Media map.
    - **Formula_Library:** common Excel formulas and where they fit.
    - **Feature_Guide:** explanation of platform capabilities.
    - **Dashboard:** native Excel charts and KPI cards.
    - **Calc_Engine:** hidden helper sheet used by Excel charts.
    """)

    st.subheader("Important privacy/cost note")
    st.markdown("<div class='warning-box'>No AI API is used. No OpenAI, Gemini, Claude, or paid model API is called. Processing is performed with free Python libraries: Streamlit, pandas, NumPy, Plotly, openpyxl, and XlsxWriter.</div>", unsafe_allow_html=True)
