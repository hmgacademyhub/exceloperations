# Changelog

All notable changes to the HMG Enterprise Excel Operations Platform.

## [v6.0] — 2026-06-20

### Fixed (bugs & deployment)
- **Critical:** Renamed config folder from `". streamlit"` (with a space — silently ignored by Streamlit) to `.streamlit`, so `config.toml` and `secrets.toml` (`APP_PASSWORD`) now load correctly.
- Removed stray empty junk file `a` from the config folder.
- Consolidated the duplicate, case‑colliding `DEPLOYMENT.md` / `deployment.md` into a single `DEPLOYMENT.md`.
- Resolved pandas `FutureWarning` for deprecated `"M"`/`"Q"` resample aliases via a new `normalize_resample_freq()` helper (forecasting & time intelligence are now future‑proof).

### Added (features — all free, no AI API)
- 🔗 **Google Sheets import** in the Ingest tab (no API key).
- ✨ **Auto insight cards** in the Profile tab (rule‑based).
- 📄 **Offline self‑contained HTML data report** download.
- 💡 **Smart chart recommendations** in the Visualize tab.
- 🔁 **Dataset comparison** tool (column‑level, Sweetviz‑style).

### Added (packaging)
- `runtime.txt` (Python 3.11 pin).
- `LICENSE` (MIT).
- `sample_data/sample_sales.csv` demo dataset.
- Hardened `.streamlit/config.toml` (XSRF on, CORS off, stats off).
- Upper‑bounded dependency versions in `requirements.txt`.
- New docs: `AUDIT_REPORT.md`, `ENHANCEMENTS_v6.md`, `CHANGELOG.md`; rewritten `README.md` and `DEPLOYMENT.md`.

### Unchanged
- No pre‑existing feature was removed. All ingestion, cleaning, transformation, analysis, modelling, governance, branding and Excel‑export capabilities are retained.

---

## [Enterprise Edition] — 2026-05 (previous)
- Consolidated full Excel workflow feature set, enterprise governance sheets, secrets‑ready password support, and deployment pack. (See `RELEASE_NOTES.md`.)

## [v2.0] — 2026-06-20

### Added — Security (anti-hacking)
- Constant-time password verification (`hmac.compare_digest`).
- Optional SHA-256 password hashing via `APP_PASSWORD_HASH` (no plaintext in secrets).
- Brute-force lockout (5 attempts → 5-minute session lock) with attempts counter.
- Demo logins (`admin`/`HMG2025`) restricted to Free tier and disableable via `ALLOW_DEMO_LOGIN="false"`.
- SQL tool hardened: single read-only SELECT/WITH only; multi-statement and write/DDL keywords blocked.

### Added — Subscriptions & licensing (anti-bypass)
- Free / Pro / Enterprise tiers with feature gating (no existing feature paywalled).
- HMAC-signed offline license keys (customer/tier/expiry) that cannot be forged, re-tiered or extended without `LICENSE_SIGNING_KEY`.
- License Admin tab (generate/validate keys) + sidebar & login-screen activation + plan badge + logout.

### Added — Privacy
- Rule-based PII detection (emails/phones/IDs/addresses) and partial/irreversible masking (Pro).

### Added — Brand
- Accurate, expanded founder profile and full four-arm ecosystem (Academy, Technologies, Media, Gospel) embedded in UI and workbook; correct ownership ("HMG Academy product, built by HMG Technologies, an HMG Concepts subsidiary").

### Added — SEO
- In-app meta/Open Graph/Twitter/JSON-LD injection; `APP_PUBLIC_URL` canonical.
- Static `seo/index.html` landing page, `seo/robots.txt`, `seo/sitemap.xml`.
- New docs: `SECURITY.md`, `SUBSCRIPTION_GUIDE.md`, `SEO_GUIDE.md`.

### Unchanged
- No pre-existing feature removed. All v6 fixes and features retained.

## [v3.0] — 2026-06-20

### Added — Selectable licence models
- Per-client choice of **Team (shareable)** or **Per-Seat (bound to a Seat ID)** licences. Model + seat fingerprint are inside the signed key and cannot be switched/rebound without the owner's secret. Activation UIs (login + sidebar) accept a Seat ID; License Admin lets you pick the model and bind a seat.

### Added — Source protection (anti-fork)
- Replaced MIT with a **proprietary HMG LICENSE** forbidding forking/copying/redistribution/derivatives/re-branding; added `NOTICE`.
- New `ANTI_FORK_GUIDE.md`: private-repo, disable-forking, branch protection, secret scanning, and deployment lockdown steps. Deployment guide updated to recommend a **private** repo.

### Added — More free enterprise features
- **Tamper-evident audit hash-chain** (Enterprise): chained SHA-256 audit log with on-screen integrity verification and CSV export.
- **Analysis Recipe** (free): save/load UI settings as JSON for reproducible, shareable workflows.

### Unchanged
- No pre-existing feature removed. All v2/v6 security, subscription, privacy, brand and SEO features retained.

## [v4.0] — 2026-06-20

### Added — Device-bound per-seat licences
- New third licence model: **Per-Seat: Device-bound**. A browser fingerprint (canvas + hardware + timezone, persisted in localStorage) is computed locally and auto-bound into the signed key. The customer types nothing; the key activates only on that device. Team and Named per-seat models retained.
- New helpers: `render_device_fingerprint()` (self-contained JS component, no external library/API) and `get_device_id()`. Activation UIs (login + sidebar) auto-use the captured device id when the Seat ID field is blank. License Admin can bind to a pasted customer device id or "this device".

### Added — More free, no-AI enterprise analytics
- **Data Quality Rules engine (Pro)** — batch rules (not-null, unique, positive, min, max, allowed, regex) → pass/fail scorecard.
- **Trend & Seasonality Decomposition (Pro)** — moving-average trend/seasonality/residual, no statsmodels.
- **Correlation Insights (Free)** — strongest relationships with notes + heatmap.
- **Cleaning Advisor (Free)** — prioritised cleaning suggestions.
- **Survey / Likert Analysis (Free)** — mean/median/top-2-box/bottom-2-box + sentiment.
- **Benford's Law Fraud Screen (Enterprise)** — first-digit audit test.

### Unchanged
- No pre-existing feature removed. All v3/v2/v6 features retained.

## [v5.0] — 2026-06-20

### Added — Usage & License Analytics (Enterprise, owner-only)
- Local SQLite event log (`log_usage_event`) capturing logins, license activations/upgrades and key issuance.
- Owner dashboard: totals, events-per-day chart, and per-event / per-tier / per-customer / per-model breakdowns, with CSV export and clear-log. Configurable via `USAGE_DB_PATH`. No external service, no AI.

### Added — More free, no-AI tools
- **Cross-tab & Chi-Square (Pro)** — contingency table + pure-numpy chi-square and Cramér's V.
- **Text / Keyword Frequency (Free)** — word counts for free-text columns with stopwords + chart.
- **Number / Locale Normalizer (Free)** — parse European/US formatted number text into clean numbers.
- **Data Classification & Watermark (Pro)** — Public/Internal/Confidential/Restricted labelling + watermark columns.

### Unchanged
- No pre-existing feature removed. All v4/v3/v2/v6 features retained.

## [v6.0] — 2026-06-20

### Added — White-label / multi-tenant branding (Enterprise)
- Per-tenant re-skin (brand name, tagline, logo, header colours, accent, footer) via the White-Label tab (session), exportable brand JSON, or the `WHITE_LABEL_BRAND` secret (permanent per deployment). New helpers: `get_active_brand`, `brand_css`, `render_brand_hero`. Hero/CSS now brand-driven.

### Added — Report delivery (Pro)
- Free SMTP email with attachment (`send_email_with_attachment`, built-in smtplib; Gmail App Password/Zoho/Brevo) plus WhatsApp (`whatsapp_share_link`) and mailto (`mailto_link`) share links. Configurable via SMTP_* secrets.

### Added — More free, no-AI tools
- **Executive Scorecard (Free)** — one-look KPI card.
- **Concentration & Gini Index (Pro)** — metric concentration/risk across categories.
- **Snapshot Diff / Change Log (Pro)** — Added/Removed/Changed rows vs an earlier snapshot.

### Unchanged
- No pre-existing feature removed. All v5/v4/v3/v2/v6-base features retained.
