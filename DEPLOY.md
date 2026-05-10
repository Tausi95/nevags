# NEVAGS Website — Cloudflare Pages Deployment

## File Structure
```
website/
├── index.html          ← Main presentation (all sections)
├── assets/
│   └── js/
│       └── main.js     ← Charts, counters, animations
├── images/             ← All brand images (copy before deploying)
│   ├── logo.jpeg
│   ├── facility.jpeg
│   ├── hollow-bricks.jpg
│   ├── face-bricks.jpg
│   ├── vsk-kiln.jpeg
│   └── vsk-kiln2.jpeg
├── _headers            ← Cloudflare security headers
├── _redirects          ← Cloudflare redirect rules
└── DEPLOY.md           ← This file
```

## Deploy to Cloudflare Pages

### Option A — Drag & Drop (Fastest, no git needed)
1. Go to https://dash.cloudflare.com → Pages → Create a project
2. Choose **"Direct Upload"**
3. Drag the entire `website/` folder into the upload window
4. Click **Deploy site**
5. Your site will be live at `https://nevags.pages.dev` (or custom domain)

### Option B — GitHub (Recommended for updates)
1. Create a new GitHub repository (e.g. `nevags-website`)
2. Copy all files from `website/` into the repo root
3. In Cloudflare Pages: Connect to Git → Select your repo
4. Build settings: **None required** (static site — no build command)
5. Deploy. Every push to main auto-deploys.

## Performance Notes
- No build step required — plain HTML/CSS/JS
- TailwindCSS loaded via CDN (cached by browsers)
- Chart.js loaded via CDN (cached by browsers)
- Images are optimised; total page weight ~1.5MB
- Works on slow connections (Malawi rural networks)

## Custom Domain (Optional)
In Cloudflare Pages → Custom domains → Add `nevags.mw` or similar.

## Print / PDF Export
Open site in browser → File → Print → Save as PDF
Or press the "Download Proposal" button which triggers window.print()
