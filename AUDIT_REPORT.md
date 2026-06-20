# 🔍 Diagnostic & Audit Report — HMG Enterprise Excel Operations Platform

**Audited:** Live site `https://cssadewale-exceldashboard-generator.streamlit.app/` and repo `hmgtechnologies/exceldashboard-generator`.
**Method:** Repository clone, static analysis (`py_compile`, `pyflakes`), runtime smoke‑testing of all 25+ analytical functions and the 40+‑sheet Excel workbook builder, a headless Streamlit boot test, and a competitor feature scan.

The application code is, overall, **well‑structured and functional** — every analytical function and the workbook export run correctly. The most serious problems were in **packaging/deployment**, where they silently disable security and configuration.

---

## A. Bugs found and fixed

### 🔴 A1 — Config directory misnamed `". streamlit"` (CRITICAL)

**Symptom.** In the repo the folder is literally named `". streamlit"` (a leading dot, then a **space**, then `streamlit`). Streamlit only loads configuration from a folder named exactly **`.streamlit`**.

**Impact.**
- `config.toml` (theme, `maxUploadSize`) was **never applied**.
- `secrets.toml` would **never be read**, so the documented production password mechanism (`APP_PASSWORD` via Streamlit Secrets) **silently failed** — the app would always fall back to the public demo passwords `admin` / `HMG2025`. This is a real security exposure for an "enterprise" deployment.

**Fix.** Renamed the directory to `.streamlit/`. Verified `get_allowed_passwords()` now actually receives `st.secrets["APP_PASSWORD"]`.

---

### 🟠 A2 — Stray junk file `a` committed in the config folder

**Symptom.** A 1‑byte empty file named `a` was committed inside `". streamlit"`.

**Impact.** Repository noise; signals an accidental commit; can confuse contributors.

**Fix.** Removed.

---

### 🟠 A3 — Duplicate, case‑colliding deployment docs

**Symptom.** The repo contains both `DEPLOYMENT.md` (418 lines) and `deployment.md` (64 lines).

**Impact.** On case‑insensitive filesystems (Windows, macOS default) these two files **collide**, causing checkout/Git problems and ambiguity about which guide is authoritative.

**Fix.** Consolidated into a single authoritative `DEPLOYMENT.md`; the lowercase variant is dropped from the package.

---

### 🟡 A4 — Deprecated pandas frequency aliases (future breakage)

**Symptom.** Forecasting (`generate_forecast`) and time intelligence (`build_time_intelligence`) pass `"M"`, `"Q"` to `resample()` / `date_range()`. Pandas ≥ 2.2 emits:
`FutureWarning: 'M' is deprecated ... please use 'ME' instead.` These aliases are scheduled for **removal**, which will eventually raise errors and break forecasting on free hosts that auto‑upgrade pandas.

**Fix.** Added `normalize_resample_freq()` which maps the UI codes (`D/W/M/Q`) to the future‑proof offset aliases (`D/W/ME/QE`) for resampling/date ranges, while keeping the classic single‑letter codes for `to_period()` (which still requires them). Verified **0 `FutureWarning`s** remain via `python -W error::FutureWarning`.

---

### 🟡 A5 — Packaging / repository hygiene gaps

**Symptoms & fixes.**
- No Python version pin → added **`runtime.txt`** (`python-3.11`) for reproducible Streamlit Cloud builds.
- No licence file → added **`LICENSE`** (MIT).
- No sample/demo data file → added **`sample_data/sample_sales.csv`**.
- Dependencies unbounded on the upper side → added safe upper bounds in `requirements.txt` (e.g. `pandas>=2.0,<3.0`) to avoid surprise breaking upgrades.
- Hardened `.streamlit/config.toml` with `enableXsrfProtection`, `enableCORS=false`, and `gatherUsageStats=false`.

---

## B. Potential bugs / robustness notes (reviewed)

These were examined; the existing code already handles them acceptably, but they are documented for maintainers:

1. **Authentication is client‑gate only.** The password gate uses `st.session_state`; it prevents casual access but is not a substitute for network‑level auth on truly sensitive data. *Recommendation kept in `Privacy_Notice` / `Access_Control` sheets.*
2. **Google Sheets import (new)** depends on the sheet being shared "anyone with the link" or published; failures are caught and reported to the user with a clear hint.
3. **Large datasets.** Excel export is correctly capped to `1,048,575` rows per sheet; the app warns the user. Browser charts are row‑capped via a slider. Good.
4. **`run_sql_query`** standardizes column names to snake_case and only the in‑memory `data` table is exposed via `sqlite3`; errors return a friendly hint row rather than crashing. Good defensive design.
5. **Empty / all‑missing columns** across profiling, outlier, RFM, cohort and anomaly functions are guarded with `dropna()`/empty checks and `try/except` fallbacks. Verified via smoke tests on the demo dataset.

---

## C. Verification performed

| Check | Result |
|-------|--------|
| `python -m py_compile app.py` | ✅ Pass |
| `pyflakes app.py` | ✅ No undefined names / unused‑import errors |
| Smoke test of 25+ analytical functions on demo data | ✅ All return valid DataFrames |
| Full 46‑sheet workbook build + reopen with openpyxl | ✅ 46 sheets, opens cleanly |
| `python -W error::FutureWarning` on time/forecast paths | ✅ 0 warnings after fix |
| Headless `streamlit run app.py` boot + HTTP 200 | ✅ Boots, serves |
| New v6 functions (Google import parse, insights, HTML report, recommendations, compare) | ✅ All pass |

---

*Prepared as part of the v6 enhancement pass. No existing feature was removed; all fixes are backward‑compatible.*
