# NEVAGS Eco Brick & Construction — DEC Presentation Website

> *Building Tomorrow Sustainably* — Musewu, Mulanje District, Malawi

A professional, single-page presentation website built for the District Executive Committee (DEC) presentation. Covers the full company story: products, investment, environmental impact, employment, workforce plan, community benefits, and partnership invitation.

---

## Live Site

Deploy via **Cloudflare Pages** (see below) — no build step required.

---

## Structure

```
/
├── index.html          # Full presentation (all 16 sections)
├── assets/
│   └── js/
│       └── main.js     # Charts (Chart.js), counters, animations, nav
├── images/             # Brand assets and facility photography
│   ├── logo.jpeg
│   ├── facility.jpeg
│   ├── hollow-bricks.jpg
│   ├── face-bricks.jpg
│   ├── vsk-kiln.jpeg
│   └── vsk-kiln2.jpeg
├── _headers            # Cloudflare Pages security headers
├── _redirects          # Cloudflare Pages redirect rules
└── README.md
```

---

## Sections

| # | Section | Content |
|---|---------|---------|
| 1 | **Hero** | Company intro, key stats, CTA buttons |
| 2 | **About** | Why NEVAGS exists, firewood ban context |
| 3 | **Mission / Vision** | Strategic foundation, founder's quote |
| 4 | **Location & Ownership** | Charles Billy Nasala (MD), Chancy Tsonga (BD), facility location |
| 5 | **Products** | Ordinary VSK bricks, Face bricks, Biomass briquettes |
| 6 | **Investment** | K450M+ capital base, Atmosfair EUR 175K loan, founder K175M+ equity |
| 7 | **Pricing** | QS rates + cost-per-m² comparison vs cement blocks & illegal bricks |
| 8 | **Environment** | SDG 3·5·13·15, zero-firewood, ESG alignment |
| 9 | **Employment** | 51 staff now → 8,000+ jobs in 3 years |
| 10 | **Workforce Plan** | KPI-driven phased scaling, community moulding programme, zero-waste model |
| 11 | **Budget** | 3-month operational budget breakdown |
| 12 | **Partners** | GIZ, Atmosfair, CCODE, TERA, MUBAS, CIRA |
| 13 | **Roadmap** | 4 phases + biogas 2nd kiln shaft (14 months) |
| 14 | **Community** | Jobs, skills, forest protection, SME chain, biogas innovation |
| 15 | **Why District / Why NEVAGS** | Mutual benefit framing for DEC |
| 16 | **Contact / CTA** | Full contact details for MD, BD Manager, and HR |

---

## Tech Stack

- **HTML5** — semantic, accessible
- **Tailwind CSS** — via CDN (no build step)
- **Chart.js** — 5 animated charts (revenue, growth, gender, inclusion, staffing, budget, cost comparison)
- **Vanilla JS** — counters, scroll animations, nav spy, progress bars
- **Cloudflare Pages** — static hosting with `_headers` and `_redirects`

---

## Deploy to Cloudflare Pages

1. Go to [dash.cloudflare.com](https://dash.cloudflare.com) → **Pages** → **Create a project**
2. Connect to GitHub → select **Tausi95/nevags**
3. Build settings:
   - **Framework preset:** None
   - **Build command:** *(leave empty)*
   - **Build output directory:** `/` *(root)*
4. Click **Save and Deploy**

Every push to `master` will auto-deploy. The site will be live at `https://nevags.pages.dev` (or a custom domain you configure).

---

## Contacts

| Role | Name | Phone | Email |
|------|------|-------|-------|
| Managing Director | Charles Billy Nasala | +265 888 34 75 75 / +265 99 751 0160 | nasalacharles.b@gmail.com |
| BD & Marketing Manager | Chancy Tausi Tsonga | +265 984 000 366 / WhatsApp +27 764 998 4601 | chancy.tsonga@yahoo.com |
| Careers / HR | — | — | careers.nevags@gmail.com |

**Location:** Musewu, Mulanje District, Malawi · P.O. Box 90, Mulanje
