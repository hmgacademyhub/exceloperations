# Detailed Feature Explanation — HMG Enterprise Excel Operations Platform

This document explains the system features and enterprise capabilities. The platform is part of the **HMG Technologies Ecosystem** and uses **no AI API**.

---

## 1. Brand and ecosystem features

### HMG Brand Profile

The platform embeds the identity of:

```text
Adewale Samson Adeagbo
Data Scientist · STEM Educator · AI-Augmented Solutions Developer
Founder / Visioner, HMG Concepts
```

Exported sheet:

```text
HMG_Brand_Profile
```

### HMG Ecosystem Map

Documents:

- HMG Concepts.
- HMG Academy.
- HMG Technologies.
- HMG Media.
- CSS Adewale portfolio.

Exported sheet:

```text
HMG_Ecosystem
```

---

## 2. Enterprise access and governance

### Authentication gate

The app includes password access.

Supported password sources:

1. Streamlit Secrets: `APP_PASSWORD`
2. Environment variable: `APP_PASSWORD`
3. Demo defaults: `admin`, `HMG2025`

### Audit trail

Every major workflow action is logged and exported to:

```text
Audit_Log
```

### Enterprise readiness assessment

Evaluates:

- Authentication.
- Auditability.
- Data quality.
- Excel governance.
- Cost control.
- Scalability.
- Privacy.

Exported sheet:

```text
Enterprise_Readiness
```

### Access-control matrix

Defines recommended access roles:

- System Owner.
- Data Analyst.
- Reviewer / Manager.
- External Partner.
- Public User.

Exported sheet:

```text
Access_Control
```

### Governance matrix

Covers:

- Data minimisation.
- Sensitive data review.
- Audit trail.
- Validation.
- Version control.
- Password rotation.
- No AI API control.
- Workbook sharing.

Exported sheet:

```text
Governance_Matrix
```

### Privacy notice

Documents that the system does not intentionally send uploaded data to AI APIs and explains responsible hosting/sharing expectations.

Exported sheet:

```text
Privacy_Notice
```

---

## 3. Data ingestion

- Upload multiple CSV/XLS/XLSX files.
- Read first sheet or all sheets.
- Merge uploads into a master dataset.
- Add source file and source sheet columns.
- Load demo data for testing.

---

## 4. Data profiling and quality

The platform produces:

- Dataset overview.
- Column-level profile.
- Missing-value report.
- Duplicate report.
- Outlier report.
- Data quality scorecard.

Quality checks include:

- Missing counts.
- Missing percentages.
- Unique values.
- Data types.
- Numeric statistics.
- IQR outlier counts.

---

## 5. Cleaning operations

Supported cleaning workflows:

- Standardize column names.
- Trim text spaces.
- Drop blank rows.
- Drop blank columns.
- Remove duplicate rows.
- Drop rows with missing values.
- Drop columns above missing-value threshold.
- Fill numeric missing values.
- Fill text missing values.
- Convert data types.
- Treat outliers.

---

## 6. Transformation operations

Supported transformations:

- Filter rows.
- Sort rows.
- Select/drop columns.
- Rename columns.
- Create calculated columns.
- Text transformations.
- Date part extraction.
- Numeric bins.
- IF-style classification.
- Numeric scaling.
- Unpivot wide data to long format.

---

## 7. Analysis and modelling

### Group-by summary

Builds grouped summaries using selected dimensions, measures and aggregations.

### Pivot table

Builds pivot/crosstab analysis.

### Correlation

Builds numeric correlation table and heatmap.

### Forecast

Creates simple non-AI forecasts using:

- Linear trend.
- Moving average.

### Pareto analysis

Identifies vital few and long-tail contributors.

### ABC analysis

Classifies categories into:

- A - highest impact.
- B - medium impact.
- C - long tail.

### Weighted scorecard

Creates weighted decision rankings using selected numeric criteria.

Exported sheet:

```text
Weighted_Scorecard
```

### Goal seek

Calculates current total, target total, gap, required lift and per-row adjustment.

Exported sheet:

```text
Goal_Seek
```

---

## 8. Analyst tools

### Data validation

Rules:

- Not blank.
- Unique values only.
- Numeric range.
- Allowed list.
- Valid date.
- Positive number.

### What-if analysis

Scenario types:

- Increase by percentage.
- Decrease by percentage.
- Add constant.
- Multiply by factor.

### Lookup / merge helper

Supports left, inner, outer and right joins.

### Sampling

Creates row-count or percentage samples.

### Reconciliation

Compares two datasets and detects missing records, extra records and value mismatches.

### Time intelligence

Creates previous-period, period-change, percentage-change, YTD and rolling-average analysis.

### KPI targets

Compares actual values against target values.

---

## 9. Advanced intelligence features without AI APIs

### SQL Query Builder

Runs safe `SELECT` and `WITH` SQL queries using SQLite.

### Anomaly detection

Creates row-level anomaly reports using:

- IQR.
- Z-score.

### RFM segmentation

Builds Recency-Frequency-Monetary segmentation.

### Cohort analysis

Builds monthly retention or value cohorts.

### Schema contract checker

Validates required columns, key completeness, key uniqueness, numeric expectations and date expectations.

### Fuzzy duplicate detection

Detects near-duplicate text values using Python's built-in similarity matching.

### Basket analysis

Finds item pairs that appear together in the same transaction/group.

### Calendar table

Creates a date dimension table for Excel, Power BI and pivot modelling.

### Report brief

Generates rule-based findings and recommended actions without any AI API.

---

## 10. Excel export

The exported workbook includes raw data, processed data, audit logs, governance, analysis, dashboards, HMG brand sheets and enterprise readiness sheets.

---

## 11. No AI API statement

The platform uses free Python libraries and deterministic rules. It does not call OpenAI, Gemini, Claude, Grok or any paid AI/model API.

---

## v6 — New free, no-AI features

### Google Sheets import (Ingest tab)
Paste a shareable Google Sheets link to load data directly via its CSV export. No Google API key, no OAuth, no cost. The sheet must be shared as "Anyone with the link (Viewer)" or published to the web.

### Auto insight cards (Profile tab → Auto Insights & Report)
Colour-coded, rule-based insight cards: dataset size, completeness, duplicates, metric totals, top-contributor concentration, outliers and latest trend. Fully deterministic — a transparent alternative to "AI insights".

### Offline HTML data report (Profile tab → Auto Insights & Report)
Download a single self-contained `.html` report (inline CSS only) with KPIs, insight cards, a column profile and a 50-row preview. Opens in any browser with no internet and no external assets — Sweetviz-style portability.

### Smart chart recommendations (Visualize tab)
Deterministic chart suggestions based on column types and cardinality (line/bar/pie/scatter/histogram), each with a plain-English reason. Mirrors commercial "AI auto-chart" without any model cost.

### Dataset comparison (Analyst Tools → Compare Datasets)
Upload a second file and compare it to the current dataset at column level: presence, type, missing % and means, plus a totals summary row. Useful for period-over-period and schema-drift checks. Downloadable as CSV.

> All v6 features use only the existing free stack (streamlit/pandas/numpy/plotly/XlsxWriter/openpyxl) and no AI API.

---

## v2 — Security, Subscriptions, Privacy & SEO

### Subscription tiers (Free / Pro / Enterprise)
Feature gating that adds value without removing anything. All original analyst workflow stays Free. Pro adds SQL query builder, dataset comparison, PII masking and split-workbook. Enterprise adds tamper-evident audit, bulk export and license administration. The admin password always maps to Enterprise. See `SUBSCRIPTION_GUIDE.md`.

### Tamper-proof license keys
Owner-issued, HMAC-SHA256-signed keys encode customer, tier and expiry. They activate offline (no payment API needed) and cannot be forged, re-tiered or extended without the secret `LICENSE_SIGNING_KEY` — preventing subscription bypass. Generate/validate in **Analyst Tools ▸ 🔑 License Admin**; customers activate via the login screen or sidebar.

### Hardened authentication
Constant-time comparison, optional hashed passwords (`APP_PASSWORD_HASH`), brute-force lockout, and Free-tier-only demo logins (disableable). See `SECURITY.md`.

### PII detection & masking
Rule-based scan for emails, phones, IDs and addresses with partial or irreversible-hash masking before export. No AI. (**Analyst Tools ▸ 🔒 PII Masking**, Pro.)

### SQL injection hardening
The SQL builder permits only a single read-only SELECT/WITH and blocks multi-statement injection and all write/DDL keywords; queries run against a throwaway in-memory SQLite.

### Search Engine Optimization
Crawlable meta description/keywords, Open Graph, Twitter cards and JSON-LD structured data are injected into the app. A static SEO landing page (`seo/index.html`), `robots.txt` and `sitemap.xml` are included for hosting on GitHub/Cloudflare Pages. See `SEO_GUIDE.md`.

### Brand & ecosystem embedding
The founder profile and all four HMG arms (Academy, Technologies, Media, Gospel) with correct ownership are embedded in the UI and in the workbook's `HMG_Brand_Profile` and `HMG_Ecosystem` sheets.

---

## v3 — Licence Models, Source Protection & More Enterprise Tools

### Selectable licence models (Team vs Per-Seat)
Generate each client either a **Team** key (shareable across devices, activates with no Seat ID) or a **Per-Seat** key (bound to a Seat ID such as a staff email or device label; only activates for that seat). The model and seat are inside the signed payload, so they cannot be switched or rebound without the owner's `LICENSE_SIGNING_KEY`. See `SUBSCRIPTION_GUIDE.md §6`.

### Anti-fork / source protection
Ships under a proprietary HMG licence (forking, copying, redistribution, derivatives and re-branding prohibited) plus a `NOTICE` file, and `ANTI_FORK_GUIDE.md` with concrete steps to keep the GitHub repository private, disable forking, protect branches, and deploy without exposing source.

### Tamper-evident audit trail (Enterprise)
The audit log can be exported as a SHA-256 **hash chain**: each entry is hashed with the previous hash. Any later edit, deletion or reordering breaks the chain, which the app detects and reports. Found in the Export tab.

### Analysis Recipe (free)
Save your current selections (metric/category/date, forecast and report settings) to a small JSON file and reload or share it for reproducible workflows. No data is stored — only configuration. Found in the Export tab.

---

## v4 — Device-Bound Seats & More Free Enterprise Tools

### Device-bound per-seat licences
A third licence model joins Team and Named per-seat. When you choose **Per-Seat: Device-bound** in License Admin, the licence binds to the customer's specific device. A small self-contained JavaScript component computes a stable browser fingerprint locally (canvas drawing + hardware concurrency + screen + timezone, plus a persisted localStorage salt) and writes a `DEV-…` id into the URL, which the app reads back. The fingerprint is then baked (hashed) into the signed key. On activation the app auto-supplies the device id, so the customer types nothing, and the key refuses to activate on any other device. No external service and no AI are used. (See `SUBSCRIPTION_GUIDE.md §7`.)

### Data Quality Rules engine (Pro)
Define multiple rules — `not_null`, `unique`, `positive`, `min`, `max`, `allowed` (comma list), `regex` — add them to a rule set, and run them all at once. You get a pass/fail scorecard per rule with failing-row counts, a "rules passed" metric and a downloadable CSV.

### Trend & Seasonality Decomposition (Pro)
Resamples a metric to your chosen frequency and separates it into **Trend** (centered moving average), **Seasonality** (average period effect) and **Residual**, with a chart and CSV. No statsmodels dependency; pure pandas/numpy; no AI.

### Correlation Insights (Free)
Computes the numeric correlation matrix, then auto-lists every pair at or above your chosen threshold with strength (moderate/strong/very strong), direction and a plain-English note, plus an optional full heatmap.

### Cleaning Advisor (Free)
Scans every column and reports concrete, prioritised issues: missing-value %, leading/trailing whitespace, inconsistent casing that splits categories, numbers stored as text, and constant columns — each with a suggested action you can apply in the Clean tab.

### Survey / Likert Analysis (Free)
For rating-scale questions (auto-detects 1–5 or 1–7), reports responses, mean, median, top-2-box % and bottom-2-box % with an automatic Positive/Mixed/Negative sentiment label, and a CSV export.

### Benford's Law Fraud Screen (Enterprise)
Compares the leading-digit distribution of a numeric column to Benford's Law expected percentages, reports per-digit deviation, a bar chart and a verdict (close / moderate / large deviation) — a classic, free audit and fraud-screening indicator.

---

## v5 — Usage Analytics & More Free Enterprise Tools

### Usage & License Analytics (Enterprise, owner-only)
Every login, license activation/upgrade and key issuance is recorded to a local SQLite event log (`USAGE_DB_PATH`, default `hmg_usage.db`). The owner dashboard (Analyst Tools ▸ 📡 Usage Analytics) shows total events, logins, activations and keys issued; an events-per-day chart; and breakdowns by event type, tier, customer and licence model. You can export the full log to CSV or clear it. Entirely local — no external analytics service and no AI. On Streamlit's free tier the file may reset on app restart, so point `USAGE_DB_PATH` at a persistent path/volume for long-term retention.

### Cross-tab & Chi-Square Association (Pro)
Builds a contingency table between two categorical (or low-cardinality) columns, then computes a chi-square statistic and **Cramér's V** effect size using pure numpy — no scipy, no AI. It reports a plain-English verdict on whether the two variables are associated and how strongly, and exports the table to CSV.

### Text / Keyword Frequency (Free)
Counts the most frequent words in a free-text column (survey open-ends, comments, notes), filtering a small stopword list and short words. Returns counts and share %, with a bar chart and CSV export. No NLP library, no AI.

### Number / Locale Normalizer (Free)
Converts locale-formatted number text into clean numeric values — handles European (`1.234,56`) and US (`1,234.56`) styles, configurable decimal and thousands separators (comma, dot, space, none). Deterministic parsing; updates the working dataset.

### Data Classification & Watermark (Pro)
Tag the dataset with a confidentiality level (Public / Internal / Confidential / Restricted) and an owner/handler. The label is recorded in the audit trail and shown as a colour badge; optionally add `__classification` and `__handled_by` watermark columns so governed exports carry their handling level.

---

## v6 — White-Label, Report Delivery & More Tools

### White-Label / Multi-Tenant Branding (Enterprise)
Re-skin the whole platform for a client without touching the code. In **Analyst Tools ▸ 🎨 White-Label** set the brand name, tagline, logo (image URL or emoji fallback), header gradient (start/mid/end), accent colour and footer. **Apply for this session**, **download the brand JSON**, or paste the JSON into the `WHITE_LABEL_BRAND` secret to make it the permanent default for that deployment. Branding resolves in priority order: session override → secret → HMG default, so one codebase serves many tenants. The login hero, main hero and CSS all follow the active brand. See `WHITE_LABEL_GUIDE.md`.

### Report Delivery — Email & WhatsApp (Pro)
Distribute reports straight from the app (**Analyst Tools ▸ 📤 Report Delivery**):
- **SMTP email with attachment** — choose the processed dataset (CSV) or the offline HTML report and send it via a free mailbox (Gmail App Password, Zoho, Brevo). Uses Python's built-in `smtplib`/`email` — no paid API, no AI. Configure `SMTP_HOST/PORT/USER/PASSWORD/FROM` in Secrets.
- **WhatsApp share link** — a `wa.me` click-to-chat link with a prefilled message.
- **mailto link** — opens the user's own email client; needs no server config.
Email sends are recorded in the audit trail and usage analytics.

### Executive Scorecard (Free)
A compact KPI card: total/average/max of a chosen metric, record count, top category and its share, distinct categories, and latest-period growth — exportable to CSV.

### Concentration & Gini Index (Pro)
Computes the **Gini coefficient** and a cumulative-share (Lorenz-style) table for a metric across categories — e.g. how dependent revenue is on a few customers — with a plain-English risk verdict. Pure numpy, no AI.

### Snapshot Diff / Change Log (Pro)
Upload an earlier snapshot of the data, pick a key/ID column, and get a row-level change log of **Added / Removed / Changed** records (with field-level before→after details) — ideal for master-data governance.
