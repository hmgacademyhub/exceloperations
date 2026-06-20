# 💳 Subscription & Licensing Guide — HMG Excel Operations Platform v2

This explains the tier system and how to sell/grant access using offline license keys. The whole system is **free to run** (no payment gateway or AI API required) — you collect payment however you prefer (bank transfer, WhatsApp, Paystack link, etc.) and then issue a key.

---

## 1. The three tiers

| Feature area | Free | Pro | Enterprise |
|---|:---:|:---:|:---:|
| Ingest (multi‑file + Google Sheets), profiling, quality scorecard | ✅ | ✅ | ✅ |
| Cleaning, transforms, calculated columns | ✅ | ✅ | ✅ |
| Pivots, group‑by, correlations, forecasting | ✅ | ✅ | ✅ |
| Visualization + smart chart recommendations | ✅ | ✅ | ✅ |
| Auto insight cards + offline HTML report | ✅ | ✅ | ✅ |
| RFM, cohort, ABC, Pareto, anomaly, KPI, reconciliation, etc. | ✅ | ✅ | ✅ |
| Full 40+ sheet Excel workbook export | ✅ | ✅ | ✅ |
| **SQL query builder** | 🔒 | ✅ | ✅ |
| **Dataset comparison** | 🔒 | ✅ | ✅ |
| **PII detection & masking** | 🔒 | ✅ | ✅ |
| **Split workbook by segment** | 🔒 | ✅ | ✅ |
| **Tamper‑evident audit / bulk export** | 🔒 | 🔒 | ✅ |
| **License administration** | 🔒 | 🔒 | ✅ |

> Design rule: **nothing from the original platform was moved behind a paywall.** Free remains a complete analyst tool. Pro/Enterprise add *new* convenience and governance features.

The admin/owner password always maps to **Enterprise**, so you keep full access.

---

## 2. How customers get upgraded

```
You (owner)                      Customer
-----------                      --------
1. Receive payment      ──▶
2. License Admin tab:
   enter customer, tier,
   expiry → Generate key
3. Send the key         ──▶      4. Paste key on login screen
                                    OR sidebar ▸ Upgrade with a license key
                                 5. Tier unlocks instantly (offline)
```

### Step‑by‑step (owner)
1. Log in with your admin password (Enterprise).
2. Go to **8️⃣ Analyst Tools ▸ 🔑 License Admin**.
3. Enter the **customer/organisation name**, choose **tier** (`pro` or `enterprise`), and an **expiry date** (`YYYY-MM-DD`).
4. Click **Generate License Key** and copy the key.
5. Send it to the customer (WhatsApp/email).

### Step‑by‑step (customer)
1. Open the app.
2. Either on the **login screen** ("Have a subscription license key?") or, once inside, in the **sidebar ▸ Upgrade with a license key**, paste the key.
3. Click **Validate / Apply**. The plan badge updates to Pro/Enterprise.

---

## 3. Why the key is safe to hand out
The key is signed with your secret `LICENSE_SIGNING_KEY` (HMAC‑SHA256). A customer **cannot**:
- change `pro` → `enterprise`,
- push the expiry further out,
- fabricate a key for a new customer,

…without your secret, because any change breaks the signature and the app rejects it. See `SECURITY.md §3` for the cryptographic detail.

---

## 4. Setting it up
In Streamlit Cloud ▸ **Settings ▸ Secrets** add a long random signing key:
```toml
LICENSE_SIGNING_KEY = "paste-output-of: python -c 'import secrets;print(secrets.token_urlsafe(48))'"
```
Do this **before** issuing any real keys. If you ever rotate this key, all previously issued keys stop working (issue replacements).

---

## 5. Common questions
- **Can one key be used by many people?** Yes — it's not device‑locked (intentional, for schools/teams). Scope it by setting a sensible expiry and per‑customer naming. For stricter control, issue short‑lived keys.
- **What happens when a key expires?** The user silently drops to Free tier; nothing breaks. Issue a renewal key.
- **Do I need a payment API?** No. Collect payment any way you like, then issue a key. This keeps the platform free to operate.
- **Can I add more tiers/features?** Yes — edit `TIER_ORDER`, `TIER_LABELS` and `FEATURE_TIER` in `app.py`, then gate a feature with `require_feature("key","Name")`.

*HMG Technologies · WhatsApp +234 810 086 6322 · buildingmyictcareer@gmail.com*

---

## 6. NEW in v3 — Selectable licence models (Team vs Per-Seat)

When generating a key in **License Admin** you now choose the model per client:

### Team (shareable)
- Activates **without** a Seat ID.
- One key works across the client's many devices/staff.
- Best for: a school, NGO, or business that wants site-wide access from one key.

### Per-Seat (named / bound)
- You set a **Seat ID** when generating the key (e.g. `teacher@school.edu` or `Lab-PC-01`).
- The Seat ID is hashed and **baked into the signed key**. The user must type the matching Seat ID to activate; the key **fails for any other seat**.
- Best for: selling individual seats, or limiting a key to one named person/device.
- Seat matching is case- and space-insensitive (`Staff@School.edu ` == `staff@school.edu`).

> Because the model and seat are inside the signed payload, a user **cannot** convert a per-seat key into a team key, rebind it to another seat, or remove the binding — any change breaks the signature (verified in testing).

### How customers activate
- **Team key:** paste the key on the login screen or sidebar; leave Seat ID blank.
- **Per-seat key:** paste the key **and** enter the exact Seat ID it was issued to.

### Choosing per client
Use **Team** for institutions buying one organisational licence; use **Per-Seat** when you want to count/limit individual users or tie a licence to a specific machine. You can mix models across your customer base freely.

---

## 7. NEW in v4 — Device-bound per-seat licences (auto fingerprint)

You now have **three** licence models to choose from in **License Admin**:

| Model | How it activates | Best for |
|-------|------------------|----------|
| **Team (shareable)** | No Seat ID needed | A whole school/team/org under one key |
| **Per-Seat: Named** | Customer types the Seat ID you set (email/label) | Selling to a specific named person |
| **Per-Seat: Device-bound** | Auto-binds to the customer's device — they type nothing | Locking a seat to one physical device |

### How device-bound works (no AI, no external service)
1. A tiny JavaScript snippet runs in the customer's browser and computes a **stable device fingerprint** from canvas rendering, hardware concurrency, screen, timezone and a persisted `localStorage` salt.
2. That becomes a `DEV-…` id which the app reads back through the URL.
3. The id is **hashed and baked into the signed licence key**, so it cannot be changed without your `LICENSE_SIGNING_KEY`.
4. On activation the app **auto-supplies the device id** — the key works only on that device and silently fails elsewhere.

### Issuing a device-bound key
1. Ask the customer to open the app → expand **"Have a subscription license key?"** (login) or **sidebar → Upgrade** → they'll see **"This device's automatic ID: DEV-…"**. They send you that id.
2. In **License Admin** choose **Per-Seat: Device-bound**, paste the customer's device id (or click **Use THIS device's ID** to bind your own), set tier + expiry, and **Generate**.
3. Send them the key. They paste it and it activates automatically on that device — no Seat ID typing.

### Notes & honest limitations
- Browser fingerprints are **per browser/profile**. A different browser or a cleared site-storage may produce a new id (the customer would need a re-issued key). This is the normal trade-off for a free, privacy-friendly, no-account approach.
- For the strongest binding to a person across devices, use **Named per-seat** (their email). For one-device locking, use **Device-bound**. For org-wide simplicity, use **Team**.
