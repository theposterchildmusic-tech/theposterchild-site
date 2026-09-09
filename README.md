# The Poster Child — site

Static single-page site. No build step.

- `index.html` — the whole site (markup, styles, script)
- `img/` — web-optimized photos (jpg + webp)
- `apps-script/Code.gs` — Google Apps Script that receives fan-capture submissions and writes them to a Google Sheet

## Wiring the fan capture

1. Follow the setup notes at the top of `apps-script/Code.gs`.
2. Paste the Web app URL into `index.html` → `const SHEET_ENDPOINT = "..."`.
3. Commit. Vercel redeploys automatically.
