# Togrul's Finance Portfolio — Claude Code Commands

**Workflow:**
1. Install Claude Code if you haven't: https://docs.claude.com/en/docs/claude-code
2. Create an empty folder on your machine: `mkdir finance-portfolio && cd finance-portfolio`
3. Run `claude` in that folder.
4. Paste **Command 1** (Bootstrap). Answer the questions it asks.
5. Push to GitHub, enable GitHub Pages → your site is live.
6. Run **Command 2** (Industry Deep-Dive) five times — once per industry.

**Have ready before you start:**
- Your GitHub username
- Your CV as a PDF (you'll drop it in `/cv/cv.pdf`)
- A 2–3 sentence professional bio
- A headshot (optional, but recommended — square, ~400×400px)
- Your email + LinkedIn URL

---

## COMMAND 1 — Bootstrap (run once)

```
You are helping me build a public-facing finance portfolio site that will be hosted on GitHub Pages and linked from my LinkedIn and CV. My name is Togrul Mirzayev, finance professional (4+ years, ACCA F3, ex-EY auditor, currently Lead Reporting & Budgeting Specialist at AzIntelecom). My edge is combining traditional finance with Python/SQL/ML — the portfolio must reflect that, not look like a generic Excel-dump.

GOAL
Build a clean, professional, static portfolio site that hosts five industry deep-dive financial models (Telecom/IT, Oil & Gas, Banking, Agriculture, Retail/Consumer). This first command bootstraps the scaffolding and landing page only — the industry deep-dives will be filled in by a separate command, one at a time.

TECH STACK — STRICT
- Vanilla HTML, CSS, JavaScript only. No React, Vue, Svelte, or any framework. No build step.
- Chart.js via CDN for charts (when needed later).
- Site must work by opening index.html directly in a browser (file://) AND on GitHub Pages.
- All assets self-contained in the repo. No external image hosts.

DESIGN BAR — this is a finance professional's site
- Aesthetic: think FT.com / Stratechery / Bloomberg Terminal modernized. Editorial, serious, restrained.
- Color palette: deep navy (#0a1828 or similar), off-white background (#fafaf7), one accent color (warm copper/amber). Dark mode toggle.
- Typography: serif headlines (Source Serif Pro or similar via Google Fonts), sans-serif body (Inter), tabular-nums for all numbers.
- Generous whitespace. No drop shadows. No gradients. No emoji. No stock photos.
- Mobile responsive — finance recruiters check on phones.

SITE STRUCTURE
Create this exact structure:
```
/
├── index.html                  Landing page
├── about.html                  Extended bio + CV download
├── industries/
│   ├── index.html              Overview of all 5 industries
│   ├── telecom-it/
│   │   └── index.html          PLACEHOLDER ("Coming soon")
│   ├── oil-gas/
│   │   └── index.html          PLACEHOLDER
│   ├── banking/
│   │   └── index.html          PLACEHOLDER
│   ├── agriculture/
│   │   └── index.html          PLACEHOLDER
│   └── retail-consumer/
│       └── index.html          PLACEHOLDER
├── cv/
│   └── README.md               Tells me to drop cv.pdf here
├── assets/
│   ├── css/styles.css          Single shared stylesheet
│   ├── js/main.js              Shared JS (dark mode toggle, nav)
│   └── img/                    For headshot, etc.
├── .gitignore
├── README.md                   Repo-level README
└── CNAME                       Empty for now, ready for custom domain later
```

LANDING PAGE (index.html) MUST CONTAIN
1. Header: my name, a one-line professional headline (ask me to draft this or propose 3 options), nav links to About / Industries / CV / LinkedIn / Email.
2. Hero section: my bio in 2–3 sentences (ask me for it), a clear value proposition explaining what the portfolio is: "Industry operating models built from first principles — covering Telecom & IT, Oil & Gas, Banking, Agriculture, and Retail/Consumer."
3. A 5-card grid linking to each industry deep-dive page. Each card shows: industry name, a one-line description of what the model covers, status badge ("Live" or "In progress"). Initially all five show "In progress."
4. Footer: copyright, links to LinkedIn / GitHub / Email, last-updated date.

ABOUT PAGE
Longer bio, education, certifications, technical stack (Excel, Power BI, SAP, Python, SQL), languages, contact, prominent "Download CV" button linking to /cv/cv.pdf.

ASK ME BEFORE WRITING CODE
1. My GitHub username (needed for the eventual URL and repo)
2. My LinkedIn URL
3. My email address
4. My exact professional headline (give me 3 options to choose from, finance-professional tone)
5. My 2–3 sentence bio (give me a draft I can edit)
6. Custom domain plan: do I have one, or use the default username.github.io URL?

AFTER ASKING, BUILD
- Generate all files above with real content (not Lorem Ipsum).
- Initialize git: `git init`, sensible .gitignore (node_modules unnecessary but include OS junk: .DS_Store, Thumbs.db).
- First commit with message "Bootstrap finance portfolio site".
- Print exact next steps for me to: (a) create a new GitHub repo, (b) push, (c) enable GitHub Pages in repo settings (which branch + folder), (d) confirm site is live at the URL.

QUALITY CHECKS BEFORE FINISHING
- Open index.html in a browser and verify it renders correctly. Use a headless browser or curl to confirm.
- Validate HTML (no broken tags).
- Check all internal links resolve.
- Confirm mobile view doesn't break (responsive at 375px width).
- Show me a screenshot or describe what the rendered page looks like.

DO NOT
- Don't add fake testimonials, fake achievements, or invented credentials.
- Don't add a contact form (needs backend; just use mailto:).
- Don't add analytics, cookie banners, or any third-party scripts beyond the Chart.js CDN.
- Don't make industry deep-dive content yet — those pages must remain "Coming soon" placeholders.

Begin by asking me the questions in the "ASK ME BEFORE WRITING CODE" section.
```

---

## COMMAND 2 — Industry Deep-Dive (run once per industry)

Run this five times — change the `INDUSTRY` value each time (telecom-it, oil-gas, banking, agriculture, retail-consumer). Do them one at a time so you can review and edit between runs.

```
You are building one industry deep-dive page for my finance portfolio site (the site bootstrapped by my previous command). I am Togrul Mirzayev, finance professional. The portfolio is hosted on GitHub Pages, vanilla HTML/CSS/JS, no build step, Chart.js via CDN for visualizations.

INDUSTRY TO BUILD: [INSERT ONE: telecom-it | oil-gas | banking | agriculture | retail-consumer]

PAGE LOCATION: /industries/[INDUSTRY]/index.html
Plus supporting files in /industries/[INDUSTRY]/ as needed (an .xlsx download, any data files).

GOAL
Replace the "Coming soon" placeholder with a complete, professional industry deep-dive. The page must demonstrate that I genuinely understand the economics of this industry, not just that I can copy a DCF template. This page is what a CFO or recruiter sees when deciding whether to take me seriously.

PAGE STRUCTURE — exact four sections, in this order

SECTION 1: INDUSTRY PRIMER (~600–900 words)
- How the business model works (what customers pay for, why, how often).
- Revenue drivers: name the top 3–5 with units (e.g., for telecom: subscribers × ARPU × (1 - churn)).
- Cost structure: split fixed vs. variable, name the top 5 cost lines.
- Capex cycle: typical intensity (% of revenue), what triggers spending.
- Working capital characteristics (DSO, DIO, DPO patterns).
- 5–7 key KPIs analysts actually watch, with typical ranges.
- Common valuation methods used (DCF, EV/EBITDA, P/B for banks, EV/Production for E&P, etc.) and when each applies.
- Two or three common analyst pitfalls / things that get modeled wrong.
- Cite real data points where possible (e.g., "industry average EBITDA margin: X–Y%").

SECTION 2: INTERACTIVE OPERATING MODEL
- Build it directly in the HTML page using vanilla JS + Chart.js.
- Inputs: sliders or number inputs for the top 5–8 drivers identified in Section 1. Pre-populated with realistic defaults.
- Outputs in real-time:
  - 5-year forecast table: Revenue, EBITDA, EBIT, Net Income, FCF
  - Key ratios: EBITDA margin, FCF margin, ROE, leverage (or industry-appropriate ratios)
  - One bar chart (revenue & EBITDA over 5 years)
  - One line chart (margins over 5 years)
  - DCF valuation output: enterprise value, equity value, implied per-share value, with WACC and terminal growth as adjustable inputs
- Make tabular numbers right-aligned, comma-separated, with negative numbers in red.
- Include a "Reset to defaults" button.
- Code must be clean, commented, and editable later.

SECTION 3: REAL COMPANY CASE STUDY (~400–600 words)
- Apply the model to one real listed company.
- Suggested companies (pick the one with cleanest public data unless I tell you otherwise):
  - telecom-it → Saudi Telecom (STC), Vodafone Group, or Singtel
  - oil-gas → Shell plc, Equinor, or BP
  - banking → HSBC, JPMorgan, or Emirates NBD
  - agriculture → Archer Daniels Midland (ADM), Bunge, or Olam Agri
  - retail-consumer → Inditex (Zara), Costco, or Walmart
- Use FY2024 reported figures as the starting point. Cite the annual report / 10-K page.
- Walk through: how I'd set the model's drivers based on this company, what 5-year forecast falls out, what valuation I get, how it compares to current market cap. Be honest about uncertainties.
- Include a clean comparison table: My Model vs. Market.

SECTION 4: DOWNLOADABLE EXCEL
- Generate a properly-built .xlsx version of the same model using a Python script (run it in the session, ship the output file alongside the page).
- Structure: Inputs sheet → Forecast sheet → Valuation sheet → Cover sheet with my name and date.
- All inputs colored blue, all formulas black, all outputs bolded — standard finance modeling color convention.
- Embed a "Download Excel Model" button on the page linking to the file.

DESIGN
Match the existing site's CSS exactly (same fonts, colors, spacing). Page should feel like part of the same publication, not a bolt-on. Add a breadcrumb at the top: Home → Industries → [Industry Name]. Add prev/next industry navigation at the bottom.

ASK ME BEFORE WRITING
1. Which real company should I use for the case study (or accept your suggestion)?
2. Any specific KPIs I want highlighted that you wouldn't include by default?
3. Should the case study assume long-position bias, neutral, or include a short thesis?

AFTER BUILDING
- Update the homepage's industry card for this industry: change status from "In progress" to "Live".
- Update the /industries/index.html overview page.
- Commit with message: "Add [Industry name] deep-dive".
- Tell me exactly what changed and what I should review before pushing.

QUALITY CHECKS
- The interactive model must produce sensible outputs at default inputs (sanity-check the math).
- Numbers must tie between the HTML model and the Excel model.
- All citations must be real and verifiable — no hallucinated annual report pages.
- Page must render correctly on mobile (test at 375px).

DO NOT
- Don't invent financial data. If you can't find a real figure, mark it clearly as "assumption" and explain the basis.
- Don't write generic "this industry is important because..." filler. Be specific or cut it.
- Don't add features I didn't ask for (chatbots, contact forms, animations).
```

---

## After all 5 deep-dives are live

Share each one separately on LinkedIn as its own post — five posts over 5 weeks beats one mega-post. Link the post to the specific industry page, not just the homepage. Each post should be a 200-word teaser ending with "Full model here: [link]".

Update your CV to include the portfolio URL right under your name/contact line. Done.
