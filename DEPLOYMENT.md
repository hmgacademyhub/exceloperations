# 🚀 Deployment Guide — HMG Enterprise Excel Operations Platform (v6)

This guide gives **clear, unambiguous, step‑by‑step** instructions to deploy the platform for **free**. It covers:

1. Prerequisites
2. Preparing the files
3. Deploy on **Streamlit Community Cloud** (recommended, free)
4. Setting the production password (Secrets)
5. Deploy on **Render** (alternative free tier)
6. Run **locally**
7. Post‑deployment verification checklist
8. Updating the app
9. Troubleshooting

> The platform uses free/open‑source tools only and uses **no AI API**.

---

## 1. Prerequisites

| You need | Where to get it | Cost |
|----------|-----------------|------|
| A GitHub account | https://github.com/join | Free |
| A Streamlit Community Cloud account | https://share.streamlit.io (sign in with GitHub) | Free |
| (Optional) Git installed locally | https://git-scm.com/downloads | Free |
| The files in this `excel/` folder | This package | — |

---

## 2. Prepare the files

Make sure the `excel/` folder contains **at least** these files (this package already does):

```
app.py
requirements.txt
runtime.txt
README.md
.gitignore
.streamlit/config.toml
.streamlit/secrets.example.toml     ← template only; do NOT put real secrets here
sample_data/sample_sales.csv         ← optional test data
```

> ⚠️ **Critical:** the configuration folder must be named exactly **`.streamlit`** (dot, then the word, **no space**). In the previous version it was wrongly named `". streamlit"`, which silently disabled the theme and the production password. This package fixes that.

> 🔒 **Never commit a real `secrets.toml`.** The `.gitignore` already excludes `.streamlit/secrets.toml`. You will add the real password in the Streamlit Cloud dashboard (Step 4), not in the repo.

---

## 3. Deploy on Streamlit Community Cloud (recommended)

### Step 3.1 — Create a GitHub repository

**Option A — via the GitHub website (no Git needed):**
1. Go to https://github.com/new
2. Repository name: e.g. `exceldashboard-generator`
3. Visibility: **Private** (recommended — protects your source; Streamlit Community Cloud deploys private repos when you connect your GitHub account). Public also works but allows anyone to view/fork — see `ANTI_FORK_GUIDE.md`.
4. Click **Create repository**.
5. On the next page click **uploading an existing file**.
6. Drag in **all the files inside the `excel/` folder** (including the `.streamlit` folder and `sample_data` folder). Commit.

**Option B — via Git on your computer:**
```bash
cd excel
git init
git add .
git commit -m "HMG Excel Operations Platform v6 (bug fixes + new features)"
git branch -M main
git remote add origin https://github.com/<your-username>/exceldashboard-generator.git
git push -u origin main
```

> Tip: confirm on GitHub that the folder appears as **`.streamlit`** (not `". streamlit"`). If GitHub hides dot‑folders, click into the repo file list — you should see `.streamlit/config.toml`.

### Step 3.2 — Create the Streamlit app

1. Go to https://share.streamlit.io and sign in with GitHub.
2. Click **Create app** → **Deploy a public app from GitHub**.
3. Fill in:
   - **Repository:** `<your-username>/exceldashboard-generator`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **Deploy**. The first build takes a few minutes while it installs `requirements.txt`.

When it finishes you'll get a public URL like:
`https://<your-app-name>.streamlit.app/`

---

## 4. Set the production password (Secrets) — strongly recommended

By default the app keeps the demo logins `admin` / `HMG2025` for easy testing. **Change this before sharing publicly.**

1. In the Streamlit Cloud dashboard open your app → **⋮ menu → Settings → Secrets**.
2. Paste:
   ```toml
   APP_PASSWORD = "YourStrongPasswordHere"
   ```
   You can allow several passwords (e.g. per team member) comma‑separated:
   ```toml
   APP_PASSWORD = "PasswordOne,PasswordTwo,PasswordThree"
   ```
3. Click **Save**. The app reboots automatically.

> The app loads passwords in this order: **Streamlit Secrets → `APP_PASSWORD` environment variable → demo defaults**. To fully disable the demo logins, edit `get_allowed_passwords()` in `app.py` and remove the `["admin", "HMG2025"]` line, then redeploy.

---

## 5. Deploy on Render (free alternative)

1. Create an account at https://render.com (free).
2. Push the same files to GitHub (Step 3.1).
3. On Render: **New → Web Service → Build and deploy from a Git repository** → select your repo.
4. Settings:
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:**
     ```bash
     streamlit run app.py --server.port $PORT --server.address 0.0.0.0
     ```
5. Add an **Environment Variable**: `APP_PASSWORD = YourStrongPassword`.
6. Click **Create Web Service**.

> Free Render web services sleep when idle and cold‑start on the next visit — same trade‑off as Streamlit's free tier.

---

## 6. Run locally (for development/testing)

```bash
# from inside the excel/ folder
pip install -r requirements.txt
streamlit run app.py
```

Open the URL Streamlit prints (default http://localhost:8501).

To test the production password locally:
```bash
# macOS / Linux
export APP_PASSWORD="YourStrongPassword"
streamlit run app.py

# Windows (PowerShell)
$env:APP_PASSWORD="YourStrongPassword"; streamlit run app.py
```

---

## 7. Post‑deployment verification checklist

Run through this once after deploying:

- [ ] App loads and shows the password gate.
- [ ] Your production password unlocks it (and the old demo password does **not**, if you removed it).
- [ ] **Ingest tab:** click **Load Demo Dataset** → preview appears with row/column metrics.
- [ ] **Ingest tab:** (optional) import a shared Google Sheet via link.
- [ ] **Profile tab → Auto Insights & Report:** insight cards render; **Download HTML Report** works and the file opens in a browser.
- [ ] **Visualize tab:** *Smart Chart Recommendations* expander shows suggestions; a chart renders.
- [ ] **Export tab:** **Build Excel Operations Workbook** → download the `.xlsx` → it opens in Excel/LibreOffice with the Dashboard and governance sheets.
- [ ] **Analyst Tools → Compare Datasets:** upload a second file → comparison table renders.
- [ ] Confirm the theme (gold/navy) is applied — proves `.streamlit/config.toml` is being read.

---

## 8. Updating the app

1. Edit files locally (or on GitHub).
2. Commit & push to `main`:
   ```bash
   git add .
   git commit -m "Describe your change"
   git push
   ```
3. Streamlit Cloud auto‑redeploys on push. (On Render, redeploy is also automatic, or click **Manual Deploy**.)

---

## 9. Troubleshooting

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| Theme/colors not applied | Config folder misnamed | Ensure folder is exactly `.streamlit` (no space). |
| Production password ignored | `secrets.toml` not set in dashboard, or config folder misnamed | Add `APP_PASSWORD` under **Settings → Secrets**; verify `.streamlit` name. |
| Build fails on `pip install` | Dependency/Python mismatch | Keep `runtime.txt` = `python-3.11`; don't pin versions tighter than `requirements.txt`. |
| App is slow / sleeps | Free tier idles | Expected; it wakes on the next visit. For always‑on, use a paid tier. |
| Google Sheet won't import | Sheet not shared | Set link sharing to "Anyone with the link (Viewer)" or File ▸ Share ▸ Publish to web. |
| `FutureWarning` about `'M'` | Old code | v6 already fixes this via `normalize_resample_freq()`. |
| Excel export missing rows | > 1,048,575 rows | Excel's hard limit; sample/aggregate/split first (the app warns you). |

---

*Need help? WhatsApp +234 810 086 6322 · buildingmyictcareer@gmail.com — HMG Technologies.*

---

# 🔐 v2 ADDENDUM — Securing, Licensing & SEO

The base steps above still apply. This addendum adds the v2 production hardening, subscription setup and SEO publishing. Do these after Step 3 (app deployed).

## A. Harden authentication (do this before going public)

In **Streamlit Cloud ▸ your app ▸ Settings ▸ Secrets**, paste:

```toml
# Strong admin password as a SHA-256 hash (preferred).
# Generate locally:
#   python -c "import hashlib;print(hashlib.sha256(('HMG-STATIC-SALT-v2'+'YourPassword').encode()).hexdigest())"
APP_PASSWORD_HASH = "your-sha256-hash"

# Turn OFF the public demo logins (admin / HMG2025).
ALLOW_DEMO_LOGIN = "false"

# Secret used to sign subscription license keys (long & random):
#   python -c "import secrets;print(secrets.token_urlsafe(48))"
LICENSE_SIGNING_KEY = "your-long-random-secret"

# Your real public URL (used for the SEO canonical tag):
APP_PUBLIC_URL = "https://your-app-name.streamlit.app/"
```

Click **Save**; the app reboots. Verify:
- Logging in with `admin` now fails (demo disabled).
- Your strong password logs you in as **Enterprise** (badge in sidebar).
- 5 wrong attempts triggers a 5-minute lockout.

> Prefer not to hash? You can use `APP_PASSWORD = "YourStrongPassword"` instead, but a hash is safer.

## B. Set up subscriptions (selling Pro/Enterprise)

1. Confirm `LICENSE_SIGNING_KEY` is set (Step A) — otherwise keys are insecure and the License Admin tab will warn you.
2. Log in as admin → **8️⃣ Analyst Tools ▸ 🔑 License Admin**.
3. Enter customer name, choose tier (`pro`/`enterprise`), set an expiry (`YYYY-MM-DD`), click **Generate License Key**.
4. Send the key to your paying customer. They paste it on the login screen ("Have a subscription license key?") or in the **sidebar ▸ Upgrade**.
5. Full details and FAQs: `SUBSCRIPTION_GUIDE.md`. Security rationale: `SECURITY.md`.

## C. Publish SEO assets (get found on Google)

The app already injects meta + JSON-LD. To be fully crawlable, also publish the static `seo/` folder on a domain you control (free):

1. Edit URLs in `seo/index.html`, `seo/robots.txt`, `seo/sitemap.xml` — replace the placeholder with your real app URL.
2. Deploy the `seo/` folder to **Cloudflare Pages** or **GitHub Pages**:
   - Cloudflare Pages → Create project → connect repo → **Build output directory:** `seo` (no build command).
3. In **Google Search Console** add the Pages URL, verify, and submit `sitemap.xml`. Repeat in **Bing Webmaster Tools**.
4. Add a link to the platform from your existing HMG sites (Academy, Technologies, portfolio) to speed up indexing.
5. Full guide: `SEO_GUIDE.md`.

## D. v2 verification checklist
- [ ] Demo logins disabled; strong password works (Enterprise badge shows).
- [ ] Brute-force lockout triggers after repeated wrong passwords.
- [ ] License Admin generates a key; pasting it as a fresh visitor unlocks the right tier.
- [ ] Tampering with one character of a key makes it invalid.
- [ ] PII Masking tab flags email/phone columns and masks them.
- [ ] SQL tool rejects `SELECT ...; DROP TABLE data`.
- [ ] Page source shows `<meta name="description">` and `application/ld+json`.
- [ ] `seo/` landing page is live on Pages and submitted to Search Console.

*Support: WhatsApp +234 810 086 6322 · buildingmyictcareer@gmail.com — HMG Technologies.*

---

# 🧩 v3 ADDENDUM — Anti-Fork, Licence Models & New Features

## E. Protect your source (make the repo non-forkable)
Do this so platform users cannot fork or copy your code. Full detail: `ANTI_FORK_GUIDE.md`.

1. **Make the repo Private:** GitHub ▸ repo ▸ Settings ▸ General ▸ Danger Zone ▸ **Change visibility ▸ Private**.
   - Streamlit Community Cloud can still deploy it: when creating the app, sign in with GitHub and grant access to the private repo.
2. **Disable forking:** Settings ▸ General ▸ untick **Allow forking**.
3. **Least-privilege collaborators** + **branch protection** on `main`.
4. **Enable secret scanning & push protection:** Settings ▸ Code security.
5. The shipped **proprietary `LICENSE`** and **`NOTICE`** legally forbid forking/redistribution.

> Result: end users only ever see the running app URL. They cannot view, fork, or clone the source.

## F. Choose a licence model per client
In **Analyst Tools ▸ 🔑 License Admin**, pick:
- **Team (shareable)** — one key for a whole client/team, activates with no Seat ID.
- **Per-Seat (bound)** — set a Seat ID (staff email/device label); the customer must enter that exact Seat ID to activate.

Send the key (and, for per-seat, the Seat ID) to the customer. They activate on the login screen or sidebar. See `SUBSCRIPTION_GUIDE.md §6`.

## G. New v3 features to verify
- [ ] **Analysis Recipe:** Export tab → Save Recipe (.json); reload it → settings restored.
- [ ] **Tamper-evident audit (Enterprise):** Export tab → audit hash-chain shows "intact"; editing a row would break it.
- [ ] **Team key** activates with blank Seat ID; **Per-seat key** rejects a wrong Seat ID and accepts the right one.

*Support: WhatsApp +234 810 086 6322 · buildingmyictcareer@gmail.com — HMG Technologies.*

---

# 🆕 v4 ADDENDUM — Device-Bound Seats & New Tools

All earlier steps still apply. v4 adds a device-bound licence option and six new tools.

## H. Issue a device-bound per-seat licence
1. Ask the customer to open the app, expand the licence box (login screen or sidebar ▸ Upgrade), and copy the **"This device's automatic ID" (DEV-…)** shown there. They send it to you.
2. In **Analyst Tools ▸ 🔑 License Admin**, choose **Per-Seat: Device-bound**, paste their device id (or click **Use THIS device's ID**), set tier + expiry, click **Generate License Key**.
3. Send the key. The customer pastes it and it activates **automatically on that device only** — no Seat ID typing. (Full detail: `SUBSCRIPTION_GUIDE.md §7`.)

> Note: the device id is computed in the browser (canvas/hardware/timezone + localStorage salt). A different browser/profile or cleared storage yields a new id and needs a re-issued key — the normal trade-off for a free, no-account method.

## I. New v4 tools to verify
Under **Analyst Tools**:
- [ ] **🧪 DQ Rules** — add a couple of rules, **Run All Rules**, see the pass/fail scorecard.
- [ ] **📈 Trend Decomp** — pick date + metric, decompose, see the trend chart.
- [ ] **🔗 Correlations** — find strong correlations; heatmap renders.
- [ ] **🧹 Clean Advisor** — scan returns suggestions.
- [ ] **📊 Survey/Likert** — select rating columns; summary renders.
- [ ] **🕵️ Benford Fraud** (Enterprise) — pick a numeric column; chart + verdict render.

*Support: WhatsApp +234 810 086 6322 · buildingmyictcareer@gmail.com — HMG Technologies.*

---

# 📡 v5 ADDENDUM — Usage Analytics & New Tools

All earlier steps still apply.

## J. (Optional) Persist the usage analytics log
The owner-only **Usage Analytics** dashboard reads a local SQLite log of logins, activations and key issuance.
- It works out of the box at the default path `hmg_usage.db`.
- On Streamlit Community Cloud's free tier the container storage is ephemeral, so the log can reset when the app sleeps/redeploys. For long-term history, set a persistent location in **Settings ▸ Secrets**:
  ```toml
  USAGE_DB_PATH = "/mount/data/hmg_usage.db"   # or any persistent volume path your host provides
  ```
- The log file is git-ignored (`*.db`) so it is never committed.

## K. New v5 tools to verify
Under **Analyst Tools**:
- [ ] **🔀 Cross-tab χ²** — pick two categorical columns → table + association verdict.
- [ ] **🔤 Text Frequency** — pick a text column → keyword counts + chart.
- [ ] **🔢 Number Normalizer** — pick columns + separators → numbers cleaned.
- [ ] **🏷️ Data Classification** — set a level → badge + audit entry; add watermark columns.
- [ ] **📡 Usage Analytics** (Enterprise) — after a couple of logins/activations, see metrics, per-day chart and breakdowns; CSV export works.

*Support: WhatsApp +234 810 086 6322 · buildingmyictcareer@gmail.com — HMG Technologies.*

---

# 🎨📤 v6 ADDENDUM — White-Label & Report Delivery

All earlier steps still apply.

## L. Configure free SMTP for emailed reports (optional)
To let the **Report Delivery** tab email attachments, add a free mailbox's SMTP details in **Settings ▸ Secrets**:
```toml
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = "587"          # 587 STARTTLS, or 465 for SSL
SMTP_USER = "you@gmail.com"
SMTP_PASSWORD = "your-app-password"   # Gmail: create an App Password (not your login password)
SMTP_FROM = "you@gmail.com"
```
- **Gmail:** enable 2-Step Verification, then create an **App Password** (Google Account ▸ Security ▸ App passwords) and use it as `SMTP_PASSWORD`.
- **Zoho / Brevo (Sendinblue) / Mailjet** free tiers also work — use their SMTP host/port/credentials.
- WhatsApp and mailto links need **no** configuration (they open the user's own app).

## M. White-label a deployment for a client (optional)
1. Log in as Enterprise → **Analyst Tools ▸ 🎨 White-Label**.
2. Set brand name, tagline, logo, colours, footer → **Apply for this session** to preview.
3. Click **Download brand JSON** (or copy the JSON block).
4. To make it permanent for that deployment, paste the JSON into the `WHITE_LABEL_BRAND` secret:
   ```toml
   WHITE_LABEL_BRAND = '{"brand_name":"Acme Analytics","tagline":"Powered by Acme","accent":"#FF5A00","header_bg":"#101820","header_mid":"#1b2a3a","header_to":"#27496d","footer":"© Acme"}'
   ```
5. For **multiple clients**, deploy the same repo multiple times (one Streamlit app per client) and give each a different `WHITE_LABEL_BRAND` secret. Full guide: `WHITE_LABEL_GUIDE.md`.

## N. New v6 tools to verify
- [ ] **🎯 Exec Scorecard** — pick a metric → KPI card renders.
- [ ] **📉 Concentration** — pick category + metric → Gini verdict + table.
- [ ] **🔁 Snapshot Diff** — upload an older file + key column → Added/Removed/Changed list.
- [ ] **📤 Report Delivery** — WhatsApp/mailto links open; SMTP email sends when configured.
- [ ] **🎨 White-Label** — apply a brand → hero/colours change; reset returns to HMG.

*Support: WhatsApp +234 810 086 6322 · buildingmyictcareer@gmail.com — HMG Technologies.*
