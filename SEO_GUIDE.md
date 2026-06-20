# 🔎 SEO Guide — Making the Platform Searchable

Streamlit apps render their content with JavaScript, which search engines index poorly. v2 ships two layers of SEO so the product can be found on Google/Bing and previews nicely when shared.

---

## 1. What is built in

### 1.1 In‑app meta tags (automatic)
`seo_meta_html()` injects, on every page load:
- `<meta name="description">` and `<meta name="keywords">`
- `<link rel="canonical">`
- **Open Graph** tags (rich previews on WhatsApp, Facebook, LinkedIn)
- **Twitter card** tags
- **JSON‑LD** structured data describing the app as a free `SoftwareApplication`, its creator (Adewale Samson Adeagbo) and publisher (HMG Academy / HMG Concepts).

Set your real URL so the canonical tag is correct — in Streamlit Secrets:
```toml
APP_PUBLIC_URL = "https://your-app-name.streamlit.app/"
```

### 1.2 Static SEO assets (in `seo/`)
- **`index.html`** — a fully crawlable landing page describing the product, with the same meta/JSON‑LD and a big "Launch the App" button linking to the Streamlit URL.
- **`robots.txt`** — allows all crawlers and points to the sitemap.
- **`sitemap.xml`** — lists the app + ecosystem URLs.

---

## 2. Recommended setup (best ranking)

Because you cannot upload `robots.txt`/`sitemap.xml` to the root of a `*.streamlit.app` domain, host the `seo/` folder on a domain you control — **GitHub Pages** or **Cloudflare Pages** (both free, and you already use Cloudflare Pages for the HMG sites):

1. **Edit the URLs** in `seo/index.html`, `seo/robots.txt`, `seo/sitemap.xml` — replace `https://your-app-name.streamlit.app/` with your real app URL.
2. **Deploy `seo/` to Cloudflare Pages** (or GitHub Pages):
   - Cloudflare Pages → Create project → connect repo → set build output directory to `seo` (no build command needed).
   - You'll get a URL like `https://hmg-excel.pages.dev/`.
3. **Submit to search engines:**
   - Google Search Console (https://search.google.com/search-console) → add your Pages URL → verify → submit `sitemap.xml`.
   - Bing Webmaster Tools (https://www.bing.com/webmasters) → import from Google or submit sitemap.
4. **Link from your existing HMG sites** (hmgacademy / hmgtechnologies / cssadewale) to the landing page — internal links from already‑indexed sites speed up discovery a lot.

---

## 3. Linking strategy (free authority)
Add a link to the platform from each relevant HMG site you already run:
- `hmgacademy.pages.dev` → "Free Excel Data Tool" (owner)
- `hmgtechnologies.pages.dev/products` → as a live product (builder)
- `cssadewale.pages.dev/projects` → as a portfolio project

These backbones already rank, so internal links pass discovery and authority to the new platform.

---

## 4. Keyword focus
The copy targets high‑intent, low‑competition phrases:
`free Excel dashboard generator`, `CSV to dashboard no code`, `Excel data analysis tool no AI`, `pivot table online free`, `data profiling tool`, plus brand terms `HMG Academy`, `HMG Technologies`, `Adewale Samson Adeagbo`.

Edit these in `seo_meta_html()` (in `app.py`) and in `seo/index.html` if you want to retarget.

---

## 5. Verification
- View page source of the running app → confirm the `<meta>` and `application/ld+json` block are present.
- Paste the landing page URL into the [Facebook Sharing Debugger](https://developers.facebook.com/tools/debug/) and Google's [Rich Results Test](https://search.google.com/test/rich-results) to confirm the preview and structured data parse.

*HMG Media supports content/visibility for the ecosystem — https://hmgmedia.pages.dev*
