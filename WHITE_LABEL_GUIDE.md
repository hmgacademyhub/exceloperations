# 🎨 White-Label & Multi-Tenant Guide

Run the **same codebase** for many clients, each with its own brand name, logo, colours and footer — with no source changes and no AI. This is the basis for reselling the platform as "their" tool.

---

## 1. How branding resolves
The active brand is chosen in this priority order (highest wins):
1. **Session override** — set live in **Analyst Tools ▸ 🎨 White-Label** (great for previewing/demos).
2. **`WHITE_LABEL_BRAND` secret** — a JSON string set per deployment (the permanent default for that app).
3. **HMG default** — the built-in HMG branding.

Brand fields: `brand_name`, `tagline`, `logo_url`, `logo_emoji`, `header_bg`, `header_mid`, `header_to`, `accent`, `footer`.

---

## 2. Brand one client (session preview)
1. Log in as Enterprise.
2. Go to **Analyst Tools ▸ 🎨 White-Label**.
3. Fill brand name, tagline, logo (image URL or emoji), header start/mid/end colours, accent, footer.
4. Click **Apply for this session** → the hero, colours and header update immediately.
5. Click **Download brand JSON** (or copy the JSON code block) to save the configuration.

---

## 3. Make it permanent for a deployment
1. Copy the JSON from the White-Label tab.
2. In **Streamlit Cloud ▸ your app ▸ Settings ▸ Secrets**, add:
   ```toml
   WHITE_LABEL_BRAND = '{"brand_name":"Acme Analytics","tagline":"Powered by Acme","logo_url":"https://acme.com/logo.png","accent":"#FF5A00","header_bg":"#101820","header_mid":"#1b2a3a","header_to":"#27496d","footer":"© Acme Analytics"}'
   ```
3. Save → the app reboots with that brand as the default everyone sees.

---

## 4. Serve MANY clients (multi-tenant patterns)

### Pattern A — one deployment per client (simplest, recommended)
- Deploy the **same private repo** as several Streamlit apps (e.g. `acme-analytics`, `school-x-data`).
- Give each app a different `WHITE_LABEL_BRAND` secret, its own `APP_PASSWORD_HASH`, and its own `LICENSE_SIGNING_KEY`.
- Each client gets a clean URL (front with your own domain via Cloudflare if desired). Fully isolated branding, passwords and licences.

### Pattern B — one deployment, branding via the tab
- A single app; each user/admin applies their brand in the session. Good for demos and consultants who re-skin on the fly. Branding is per session (not persisted across users).

> Tip: combine with the per-seat/device licences (v3/v4) and Usage Analytics (v5) so each tenant is separately licensed and measured.

---

## 5. Logos
- **Logo URL:** any public image URL (PNG/SVG/JPG). Host it on the client's site or a free image host / Cloudflare Pages.
- **Logo emoji:** a fallback used when no URL is given (e.g. `📊`, `📈`, `🏫`).

> In the in-app preview, external images load when the app runs in a real browser. (Sandboxed previews may not fetch external images — that's expected.)

---

## 6. What stays HMG
The proprietary `LICENSE`, `NOTICE` and the embedded authorship/ecosystem (workbook `HMG_Brand_Profile` / `HMG_Ecosystem` sheets and the HMG Brand tab) remain, asserting that the underlying engine is built by HMG Technologies. White-labelling changes the **client-facing skin**, not the ownership of the software. Removing HMG attribution from the source is prohibited by the licence — see `LICENSE` / `ANTI_FORK_GUIDE.md`.

*HMG Technologies · buildingmyictcareer@gmail.com · WhatsApp +234 810 086 6322*
