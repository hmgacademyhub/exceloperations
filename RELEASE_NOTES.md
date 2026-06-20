# Release Notes — HMG Enterprise Excel Operations Platform

## Enterprise Edition

This release consolidates all previous Excel workflow features and adds enterprise governance capabilities while maintaining a free-tools-only architecture.

---

## What is retained

No pre-existing feature was removed. The release retains:

- Authentication gate.
- Multi-file upload.
- Data ingestion.
- Data profiling.
- Cleaning.
- Transformations.
- Pivoting.
- Forecasting.
- Dashboards.
- Validation.
- What-if analysis.
- Pareto and ABC analysis.
- SQL querying.
- Anomaly detection.
- RFM and cohort analysis.
- Reconciliation.
- KPI tracking.
- Weighted scoring.
- Goal seek.
- Calendar table.
- HMG brand embedding.
- Complete Excel export.

---

## New enterprise additions

### 1. Secrets-ready authentication

Production passwords can now be set through:

- Streamlit Secrets.
- Environment variable.

Variable name:

```text
APP_PASSWORD
```

### 2. Enterprise workbook sheets

New enterprise governance sheets:

```text
Enterprise_Readiness
Access_Control
Governance_Matrix
Deployment_Checklist
Free_Tools_Register
Privacy_Notice
Release_Notes
```

### 3. Enterprise documentation

The package includes:

```text
README.md
FEATURES.md
DEPLOYMENT.md
RELEASE_NOTES.md
```

### 4. No AI API assurance

The release continues to avoid paid AI/model APIs.

---

## Brand positioning

The platform is part of:

```text
HMG Technologies Ecosystem
```

Under:

```text
HMG Concepts — His Marvellous Grace Educational Consult
```

Founder:

```text
Adewale Samson Adeagbo
Data Scientist · STEM Educator · AI-Augmented Solutions Developer
```

---

## v6.0 Addendum (2026-06-20)

### Bug fixes
- Renamed config folder `". streamlit"` → `.streamlit` (critical: previously disabled theme + production password).
- Removed stray junk file `a`.
- Consolidated duplicate `deployment.md` into `DEPLOYMENT.md`.
- Fixed deprecated pandas `"M"/"Q"` frequency aliases (future-proofed forecasting & time intelligence).

### New free, no-AI features
- Google Sheets link import (no API key).
- Auto rule-based insight cards (on screen).
- Offline self-contained HTML data report download.
- Smart chart recommendations.
- Column-level dataset comparison.

### Packaging
- Added `runtime.txt`, `LICENSE`, `sample_data/`, `CHANGELOG.md`, `AUDIT_REPORT.md`, `ENHANCEMENTS_v6.md`; rewrote `README.md` and `DEPLOYMENT.md`.

No pre-existing feature was removed.

---

## v2.0 (2026-06-20) — Security, Subscriptions, Brand & SEO

### Security (anti-hacking)
- Constant-time password check, optional SHA-256 password hashing, brute-force lockout, demo logins restricted to Free tier and disableable, SQL injection hardening.

### Subscriptions (anti-bypass)
- Free/Pro/Enterprise tiers (no existing feature paywalled). HMAC-signed offline license keys that cannot be forged/re-tiered/extended without the owner's secret. License Admin tab, sidebar/login activation, plan badge, logout.

### Privacy
- Rule-based PII detection and partial/irreversible masking (Pro).

### Brand
- Accurate founder profile and full four-arm ecosystem (Academy, Technologies, Media, Gospel) with correct ownership embedded in UI and workbook.

### SEO
- In-app meta/Open Graph/Twitter/JSON-LD; static seo/ landing page, robots.txt, sitemap.xml; SECURITY.md, SUBSCRIPTION_GUIDE.md, SEO_GUIDE.md.

No pre-existing feature was removed.

---

## v3.0 (2026-06-20) — Licence models, Anti-fork, Audit chain & Recipes

- Selectable per-client licence models: Team (shareable) vs Per-Seat (bound to a Seat ID); both tamper-proof.
- Proprietary LICENSE + NOTICE + ANTI_FORK_GUIDE.md; deployment recommends a private, non-forkable repo.
- Tamper-evident audit hash-chain (Enterprise) and saveable Analysis Recipes (free).

No pre-existing feature was removed.

---

## v4.0 (2026-06-20) — Device-bound seats + 6 new free enterprise tools

- Device-bound per-seat licences (auto browser fingerprint; no AI, no external service) alongside Team and Named per-seat models.
- New tools: Data Quality Rules engine, Trend & Seasonality Decomposition, Correlation Insights, Cleaning Advisor, Survey/Likert Analysis, Benford's Law Fraud Screen.

No pre-existing feature was removed.

---

## v5.0 (2026-06-20) — Usage analytics + 5 new free tools

- Owner-only Usage & License Analytics dashboard (local SQLite event log; logins, activations, issuance; charts + CSV).
- New tools: Cross-tab & Chi-Square, Text/Keyword Frequency, Number/Locale Normalizer, Data Classification & Watermark.

No pre-existing feature was removed.

---

## v6.0 (2026-06-20) — White-label, report delivery + 3 new tools

- White-label/multi-tenant branding (session, JSON, or WHITE_LABEL_BRAND secret).
- Report delivery: free SMTP email with attachment + WhatsApp/mailto share links.
- New tools: Executive Scorecard, Concentration & Gini Index, Snapshot Diff.

No pre-existing feature was removed.
