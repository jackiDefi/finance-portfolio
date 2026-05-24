# Togrul Mirzayev — Finance Portfolio

Public-facing finance portfolio site. Five industry deep-dives (Telecom & IT, Oil & Gas, Banking, Agriculture, Retail/Consumer), each with a primer, an interactive operating model, a real-company case study, and a downloadable Excel.

**Live site:** https://jackidefi.github.io/finance-portfolio/ (after GitHub Pages is enabled)

## Stack

- Vanilla HTML, CSS, JavaScript — no framework, no build step
- Chart.js via CDN (added on industry pages as they go live)
- Google Fonts: Source Serif 4 (headlines) and Inter (body)

## Run locally

Open `index.html` directly in a browser. That's it.

For a proper local server (useful if testing relative paths):

```sh
python3 -m http.server 8000
# then visit http://localhost:8000
```

## Structure

```
/
├── index.html                  Landing page
├── about.html                  Bio, technical stack, CV download
├── industries/
│   ├── index.html              Overview of all five industries
│   ├── telecom-it/index.html
│   ├── oil-gas/index.html
│   ├── banking/index.html
│   ├── agriculture/index.html
│   └── retail-consumer/index.html
├── cv/
│   ├── README.md               (drop cv.pdf here)
│   └── cv.pdf                  (you add this)
├── assets/
│   ├── css/styles.css
│   ├── js/main.js
│   └── img/
├── .gitignore
├── README.md
└── CNAME                       (empty; for future custom domain)
```

## Adding content

Each industry page starts as a "Coming soon" placeholder. Replace one at a time with the full deep-dive (primer + interactive model + case study + Excel). Each replacement flips the homepage status badge from **In progress** to **Live**.

## Contact

- Email: mutaliboglutogrul@gmail.com
- LinkedIn: https://www.linkedin.com/in/togrul-mirzayev/
- GitHub: https://github.com/jackiDefi
