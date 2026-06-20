# 🔐 Security Model & Hardening Guide — HMG Excel Operations Platform v2

This document explains how the platform is protected against unauthorized access and against users trying to bypass the subscription, and how to configure it securely. Everything here uses Python's standard library — **no paid security service or AI API**.

---

## 1. Threats addressed

| Threat | Control in v2 |
|--------|---------------|
| Password guessing / brute force | Per‑session lockout (5 attempts → 5‑minute lock) + attempts counter shown to user |
| Timing attacks on password compare | Constant‑time comparison via `hmac.compare_digest` |
| Plaintext password leaking from secrets | Optional SHA‑256 **password hash** (`APP_PASSWORD_HASH`) so plaintext is never stored |
| Public demo logins reaching paid features | Demo keys (`admin`/`HMG2025`) grant **Free tier only**, and can be disabled entirely |
| Subscription bypass / tier forgery | **HMAC‑signed license keys** that cannot be forged, re‑tiered, or extended without the owner's secret |
| License sharing past expiry | Expiry date is embedded **inside the signed payload** and checked on every activation |
| SQL injection / data exfiltration via SQL tool | Only single read‑only `SELECT`/`WITH`; multi‑statement and write/DDL keywords blocked; in‑memory DB only |
| Accidental PII exposure in exports | Rule‑based PII detection + partial/irreversible masking before export |
| Cross‑site request issues | `enableXsrfProtection = true`, `enableCORS = false` in `config.toml` |

---

## 2. Authentication

### 2.1 How a login is verified
1. The user enters a password.
2. `verify_password()` compares it (constant‑time) against every configured record:
   - hashes from `APP_PASSWORD_HASH`,
   - plaintext from `APP_PASSWORD`,
   - demo defaults (only if `ALLOW_DEMO_LOGIN` ≠ `false`).
3. On success it returns the **tier** for that record (admin/hash/plain → `enterprise`; demo → `free`).
4. On failure the attempt counter increments; after 5 failures the session is locked for 5 minutes.

### 2.2 Recommended production setup
```toml
# .streamlit/secrets.toml  (set in Streamlit Cloud ▸ Settings ▸ Secrets)
APP_PASSWORD_HASH = "ab12…"   # SHA-256 hash of your strong password
ALLOW_DEMO_LOGIN  = "false"   # turn off admin/HMG2025
LICENSE_SIGNING_KEY = "a-long-random-secret"
```

Generate the hash locally (never commit the plaintext):
```bash
python -c "import hashlib;print(hashlib.sha256(('HMG-STATIC-SALT-v2'+'YourPassword').encode()).hexdigest())"
```

> To rotate the static salt, change `HMG-STATIC-SALT-v2` in `_hash_password()` and regenerate hashes. For most deployments the default salt is fine because the hash still requires knowing your password.

---

## 3. Subscription licensing (anti‑bypass)

### 3.1 Why it cannot be bypassed
A license key looks like `base64(payload).base64(signature)` where:
- **payload** = JSON `{customer, tier, expiry, issued}`,
- **signature** = `HMAC‑SHA256(payload, LICENSE_SIGNING_KEY)`.

To upgrade themselves a user would have to produce a valid signature for a modified payload (e.g. tier=`enterprise`, expiry=2099). That requires the secret `LICENSE_SIGNING_KEY`, which only you hold. Any change to the payload invalidates the signature, and the check uses constant‑time comparison. Expired keys are rejected because expiry is *inside* the signed payload.

### 3.2 Issuing keys
- Log in as admin/Enterprise → **Analyst Tools ▸ 🔑 License Admin** → enter customer, tier, expiry → **Generate**.
- Send the key to the customer. They paste it on the login screen or in the sidebar **Upgrade** box.

> ⚠️ Set `LICENSE_SIGNING_KEY` before issuing real keys. Without it, keys are signed with a **public demo key** and provide no protection (the app warns you in the License Admin tab).

---

## 4. SQL tool safety
`run_sql_query()` enforces:
- query must start with `select` or `with`;
- no `;` (blocks stacked/multi statements);
- blocks `attach, detach, pragma, insert, update, delete, drop, alter, create, replace, vacuum, reindex, load_extension`;
- runs against a throwaway **in‑memory** SQLite built from the current dataframe only.

---

## 5. Privacy / data handling
- Uploaded data is processed **in the active session**; the app does not intentionally transmit datasets to any third party or AI service.
- **PII tools** (Pro): scan for emails/phones/IDs/addresses and mask them (partial or irreversible hash) before you export or share.
- Exported workbooks may contain raw data — remove sensitive sheets before external sharing (the workbook's `Privacy_Notice` and `Access_Control` sheets remind reviewers of this).

---

## 6. Operational checklist
- [ ] `APP_PASSWORD_HASH` (or `APP_PASSWORD`) set to a strong value.
- [ ] `ALLOW_DEMO_LOGIN = "false"` in production.
- [ ] `LICENSE_SIGNING_KEY` set to a long random secret.
- [ ] Real `.streamlit/secrets.toml` **never** committed (already in `.gitignore`).
- [ ] Rotate password and signing key after staff changes.
- [ ] Review the `Audit_Log` sheet before sharing any workbook.

---

## 7. Honest limitations
This is an application‑level gate suitable for institutional and commercial use on free hosting. It is **not** a replacement for network‑level controls (VPN, SSO/IdP) when handling highly sensitive regulated data. For such cases, deploy behind your organisation's identity provider and restrict hosting access accordingly.

*HMG Technologies · buildingmyictcareer@gmail.com · WhatsApp +234 810 086 6322*
