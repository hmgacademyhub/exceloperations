# 📊 HMG Excel Operations Platform — v6

**An HMG Academy product · Built by HMG Technologies · An HMG Concepts ecosystem brand**

- **Owner:** HMG Academy (subsidiary of HMG Concepts)
- **Built by:** HMG Technologies (innovation arm of HMG Concepts)
- **Founder / Developer:** Adewale Samson Adeagbo — *AI‑Augmented Solutions Developer · Data Scientist · STEM Educator*
- **Brand philosophy:** *Learning Deliberately. Teaching Authentically.*
- **Cost rule:** 100% free/open‑source stack, **no AI API**, no recurring cost.
- **Live demo:** https://cssadewale-exceldashboard-generator.streamlit.app/

> HMG Concepts (His Marvellous Grace, est. 2015, Lagos, Nigeria) is **one brand with four living missions**: 🎓 **HMG Academy** (education), 💻 **HMG Technologies** (innovation), 📢 **HMG Media** (content/visibility), ✝️ **HMG Gospel** (faith outreach). This platform sits inside HMG Academy and is built by HMG Technologies.

---

## 1. What this platform does

Upload one or more Excel/CSV files (or a Google Sheets link) and the platform turns dozens of manual analyst tasks into button clicks: cleaning, profiling, transformation, pivots, forecasting, modelling (RFM, cohort, ABC, Pareto, anomaly), visualization, governance and a complete **40+ sheet enterprise Excel workbook** export — plus an offline HTML report. Every insight is **deterministic and rule‑based; no AI API is used.**

Designed for schools, virtual institutions, SMEs, churches, NGOs, training centres and analysts.

---

## 1zzz. What's new in v6 (white-label, report delivery + 3 new free tools)

v6 builds on v5 and adds, **without removing anything**:

- **🎨 White-Label / Multi-Tenant Branding (Enterprise)** — re-skin the entire platform per client: brand name, tagline, logo (URL or emoji), header gradient colours, accent and footer. Apply for a session from the **White-Label** tab, export a brand JSON, or set it permanently per deployment via the `WHITE_LABEL_BRAND` secret. Same code, unlimited tenants. See `WHITE_LABEL_GUIDE.md`.
- **📤 Report Delivery (Pro)** — email a report attachment (processed CSV or the offline HTML report) through a **free SMTP** account using Python's built-in `smtplib` (Gmail App Password / Zoho / Brevo — no paid API), or generate one-click **WhatsApp** and **mailto** share links.
- **🎯 Executive Scorecard (Free)** — compact one-look KPI card (totals, average, top contributor, distinct count, latest growth).
- **📉 Concentration & Gini Index (Pro)** — measure how concentrated a metric is across categories (revenue-per-customer risk) with the Gini coefficient and cumulative-share table.
- **🔁 Snapshot Diff / Change Log (Pro)** — compare the current data to an earlier snapshot keyed on an ID column and list rows Added / Removed / Changed.

## 1zz. What's new in v5 (usage analytics + 5 new free enterprise tools)

v5 builds on v4 and adds, **without removing anything**:

- **📡 Usage & License Analytics (Enterprise, owner-only)** — a dashboard of recorded events (logins, license activations/upgrades, key issuance) from a local SQLite log: totals, events-per-day, and per-tier / per-customer / per-model breakdowns, with CSV export and a clear-log button. No external service, no AI. *(Free-tier hosting note: the local log can reset on app restart; set `USAGE_DB_PATH` to a persistent location for long-term retention.)*
- **🔀 Cross-tab & Chi-Square (Pro)** — contingency table between two categorical columns + a pure-numpy chi-square test and Cramér's V effect size to judge whether they're related (no scipy).
- **🔤 Text / Keyword Frequency (Free)** — most-common words in a free-text column (survey open-ends, comments) with stopword filtering and a bar chart.
- **🔢 Number / Locale Normalizer (Free)** — convert European `1.234,56` or US `1,234.56` number text into clean numbers.
- **🏷️ Data Classification & Watermark (Pro)** — tag the dataset Public/Internal/Confidential/Restricted, record it in the audit trail, and add watermark columns for governed exports.

## 1z. What's new in v4 (device-bound seats + 6 new free enterprise tools)

v4 builds on v3 and adds, **without removing anything**:

- **🖥️ Device-bound per-seat licences** — per-seat keys can now **auto-bind to a customer's specific device** using a browser fingerprint computed locally (canvas + hardware + timezone signals, stored in `localStorage`; **no external service, no AI**). The customer types nothing; the key simply won't activate on any other device. You still keep the v3 **Team** and **Named per-seat** models — three choices in License Admin.
- **🧪 Data Quality Rules engine (Pro)** — stack many rules (not-null, unique, positive, min, max, allowed-list, regex) and run them at once for a pass/fail scorecard.
- **📈 Trend & Seasonality Decomposition (Pro)** — split a time series into trend, seasonality and residual with a lightweight moving-average method (no statsmodels).
- **🔗 Correlation Insights (Free)** — auto-surface the strongest numeric relationships with plain-English notes + heatmap.
- **🧹 Cleaning Advisor (Free)** — scans for missing values, whitespace, casing, numbers-as-text and constant columns and returns prioritised fixes.
- **📊 Survey / Likert Analysis (Free)** — mean, median, top-2-box/bottom-2-box %, and a sentiment label for rating-scale questions.
- **🕵️ Benford's Law Fraud Screen (Enterprise)** — audit-grade first-digit test to flag possible number manipulation.

## 1a. What's new in v3 (licence models, anti‑fork, audit chain, recipes)

v3 builds on v2 and adds, **without removing anything**:

- **Selectable licence models per client** — issue each customer either a **Team** key (shareable across many devices, activates with no Seat ID) or a **Per‑Seat** key (bound to a Seat ID such as a staff email or device label; only activates for that seat). Both are tamper‑proof — the model and seat are inside the signed key and cannot be switched or rebound without your secret. (See `SUBSCRIPTION_GUIDE.md §6`.)
- **Anti‑fork source protection** — replaced MIT with a **proprietary HMG `LICENSE`** (forking/copying/redistribution/derivatives/re‑branding prohibited) + `NOTICE`, and an **`ANTI_FORK_GUIDE.md`** with concrete steps to keep your GitHub repo **private and non‑forkable** while still deploying on Streamlit. (Deployment now recommends a private repo.)
- **Tamper‑evident audit trail (Enterprise)** — export the audit log as a SHA‑256 hash chain; any later edit/deletion/reordering breaks the chain and is detected.
- **Analysis Recipe (free)** — save/load your UI settings as a JSON file for reproducible, shareable workflows.

## 2. What's new in v2 (security, subscriptions, brand, SEO)

v2 adds **enterprise security, a subscription/licensing system, deeper brand embedding, privacy tooling and full SEO** — **without removing any existing feature**.

### 2.1 🔐 Security hardening (anti‑hacking)
- **Constant‑time password verification** (`hmac.compare_digest`) — resists timing attacks.
- **Hashed passwords:** store a SHA‑256 hash in `APP_PASSWORD_HASH` so plaintext never lives in secrets.
- **Brute‑force lockout:** 5 failed attempts → 5‑minute session lockout.
- **Demo logins restricted:** `admin` / `HMG2025` now grant **Free tier only** and can be fully disabled with `ALLOW_DEMO_LOGIN="false"`.
- **SQL injection hardening:** the SQL tool permits only a single read‑only `SELECT`/`WITH`; multi‑statement injection and all write/DDL keywords are blocked.

### 2.2 💳 Subscription tiers + tamper‑proof licensing (anti‑bypass)
- Three tiers: **Free → Pro → Enterprise** with feature gating. **All original analyst workflow stays on Free** — paid value is added on top, nothing removed.
- **HMAC‑signed offline license keys** carry *customer, tier, expiry*. Because they are signed with your secret `LICENSE_SIGNING_KEY`, a user **cannot forge a key, upgrade their own tier, or extend expiry** — this is what prevents subscription bypass.
- **License Admin** tab (Enterprise/owner) generates and validates keys; customers paste a key in the sidebar or login screen to unlock their tier.

| Tier | Includes |
|------|----------|
| **Free** | Full ingestion, profiling, cleaning, transforms, analysis, forecasting, visualization, Excel export, HTML report, Google Sheets import, auto insights |
| **Pro** | Everything in Free **+** SQL query builder, dataset comparison, split‑workbook, PII detection & masking |
| **Enterprise** | Everything in Pro **+** tamper‑evident audit, license administration, bulk segment export |

### 2.3 🔒 Privacy: PII detection & masking
Rule‑based scan for likely personal data (emails, phones, IDs, addresses) and one‑click **partial** or **irreversible hash** masking before export. No AI.

### 2.4 🏷️ Deeper, accurate brand embedding
The founder profile and the **full four‑arm ecosystem** (Academy, Technologies, Media, Gospel) plus the correct ownership ("an HMG Academy product, built by HMG Technologies, an HMG Concepts subsidiary") are embedded in the UI **and** every exported workbook.

### 2.5 🔎 Search Engine Optimization (SEO)
- Crawlable `<meta>` description/keywords, **Open Graph**, **Twitter cards** and **JSON‑LD** structured data injected into the app.
- A static **SEO landing page** (`seo/index.html`), **`robots.txt`** and **`sitemap.xml`** you can host on GitHub/Cloudflare Pages so search engines can index the product (Streamlit's JS app alone is hard to crawl).

---

## 3. Project structure

```
excel v6/
├── app.py                     # the full Streamlit application (v6)
├── requirements.txt           # pinned free dependencies
├── runtime.txt                # Python version pin (3.11)
├── LICENSE                    # PROPRIETARY HMG licence (no forking/redistribution)
├── NOTICE                     # ownership notice
├── README.md                  # this file
├── FEATURES.md                # detailed feature explanations (every feature)
├── SECURITY.md                # security model & hardening guide
├── SUBSCRIPTION_GUIDE.md      # tiers + licensing + Team/Named/Device-bound models
├── ANTI_FORK_GUIDE.md         # keep the repo private & non-forkable
├── WHITE_LABEL_GUIDE.md       # re-skin per client / multi-tenant deployment
├── SEO_GUIDE.md               # how to get indexed by search engines
├── AUDIT_REPORT.md            # original bug audit & fixes
├── DEPLOYMENT.md              # step-by-step deployment guide (+v2/v3/v4/v5/v6 addenda)
├── RELEASE_NOTES.md           # release history
├── CHANGELOG.md               # version changelog
├── .gitignore
├── .streamlit/
│   ├── config.toml            # theme + server config
│   └── secrets.example.toml   # template for password, demo toggle, license key, URL
├── seo/
│   ├── index.html             # static SEO landing page (host on Pages)
│   ├── robots.txt
│   └── sitemap.xml
└── sample_data/
    └── sample_sales.csv       # demo dataset for testing
```

---

## 4. Quick start (local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

**Default test login:** `admin` or `HMG2025` (Free tier). To test admin/Enterprise locally:

```bash
export APP_PASSWORD="YourStrongPassword"
export LICENSE_SIGNING_KEY="$(python -c 'import secrets;print(secrets.token_urlsafe(48))')"
streamlit run app.py
```

---

## 5. Deployment

Full, unambiguous steps (GitHub → Streamlit Cloud → secrets → security → SEO → verification) are in **`DEPLOYMENT.md`**. See **`SECURITY.md`**, **`SUBSCRIPTION_GUIDE.md`** and **`SEO_GUIDE.md`** for those subsystems.

---

## 6. Brand links

Portfolio: https://cssadewale.pages.dev · HMG Concepts: https://hmgconcepts.pages.dev · HMG Academy: https://hmgacademy.pages.dev · HMG Technologies: https://hmgtechnologies.pages.dev · HMG Media: https://hmgmedia.pages.dev · HMG Gospel: https://hmggospel.pages.dev · GitHub: https://github.com/cssadewale · YouTube: https://youtube.com/@hmgconcepts · WhatsApp: +234 810 086 6322

---

*No pre‑existing feature was removed in v2. All enhancements use the free Streamlit/pandas/XlsxWriter/Plotly stack with no AI API.*
