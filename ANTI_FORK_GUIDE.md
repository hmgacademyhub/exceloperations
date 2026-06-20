# 🚫 Anti-Fork & Source Protection Guide

Goal: prevent **users of the platform** (and the public) from forking, cloning, or copying your source code, while still letting you deploy it. This combines **legal**, **repository**, and **deployment** controls. None of it costs money.

> Reality check: any code that runs in a browser/app can theoretically be inspected, and any repo a host can read could in principle be copied by that host. What you *can* do — and what this guide does — is make forking **legally prohibited** and **practically blocked** for normal users, while keeping your hosted app fully functional.

---

## 1. Legal layer (already done)
- The project ships with a **proprietary `LICENSE`** that explicitly forbids forking, copying, redistribution, derivative works and re-branding (replaces the old MIT licence).
- A **`NOTICE`** file restates ownership.
- The app footer and the `HMG_Brand_Profile` workbook sheet assert authorship.

This means even if someone obtains the code, redistribution is a licence breach and (with your branding embedded) easy to prove.

---

## 2. Repository layer — the most important control

### 2.1 Keep the repository PRIVATE
A **private GitHub repo cannot be forked or viewed** by the public or by your platform's end users. End users only ever see the *running app*, never the repo.

- GitHub ▸ your repo ▸ **Settings ▸ General ▸ Danger Zone ▸ Change visibility ▸ Private**.

> Streamlit Community Cloud **can deploy from a private repo** when you connect your GitHub account and grant access. So you do **not** need a public repo to deploy. (This removes the old "public repo required" assumption.)

### 2.2 Disable forking explicitly
Even while private, turn forking off so collaborators can't fork:
- GitHub ▸ repo ▸ **Settings ▸ General** ▸ untick **Allow forking**.
- For Organisation-owned repos: **Org ▸ Settings ▸ Member privileges ▸ Repository forking ▸ Disabled**.

### 2.3 Lock down collaborator access
- Add collaborators only with the **minimum role** needed (Read/Triage, not Write/Admin).
- Protect the `main` branch: **Settings ▸ Branches ▸ Add rule** → require PR reviews, restrict who can push.
- Enable **Settings ▸ Code security** ▸ secret scanning & push protection (free) so secrets never leak.

### 2.4 Never commit secrets
- `.gitignore` already excludes `.streamlit/secrets.toml`. Keep all secrets (password hash, `LICENSE_SIGNING_KEY`) in the host's Secrets manager, never in code. A leaked signing key would let others mint licences.

---

## 3. Deployment layer

### 3.1 Streamlit Community Cloud
- Deploy from your **private** repo (Step 2.1).
- Keep the app behind the built-in **password gate** (this platform's login) so only licensed users reach it.
- App settings ▸ Sharing: do not make the source viewable; share only the app URL.

### 3.2 Optional: render/host without exposing source
If you ever host elsewhere (e.g., Render), connect the **private** repo with a deploy key; the public only gets the running URL.

### 3.3 Custom domain (optional)
Front the app with your own domain via Cloudflare so the public never sees a generic repo-linked URL.

---

## 4. Runtime hardening (already in the app)
- **Login gate** with brute-force lockout — non-licensed users can't even open the workflow.
- **Tier gating + signed licences** — premium features are unreachable without a valid key.
- **No client-side secrets** — the signing key lives only in server-side Secrets.

---

## 5. What end users of the platform can and cannot do

| Action | Possible for a platform user? |
|--------|-------------------------------|
| Use the hosted app (per their tier) | ✅ Yes (intended) |
| See your GitHub source | ❌ No — repo is private |
| Fork your repo | ❌ No — private + forking disabled |
| Clone/redistribute the code | ❌ No — prohibited by LICENSE |
| Self-upgrade their tier | ❌ No — signed licences |
| Re-brand and resell | ❌ No — prohibited by LICENSE |

---

## 6. Quick checklist
- [ ] Repo visibility = **Private**.
- [ ] **Allow forking** unticked.
- [ ] Branch protection on `main`.
- [ ] Collaborators on least-privilege roles.
- [ ] Secret scanning / push protection on.
- [ ] No secrets in code (only in host Secrets).
- [ ] Proprietary `LICENSE` + `NOTICE` present (shipped).
- [ ] App reachable only via password/licence.

*Questions: HMG Technologies — buildingmyictcareer@gmail.com · WhatsApp +234 810 086 6322*
