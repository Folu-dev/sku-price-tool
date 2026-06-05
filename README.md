# SKU Price Breakdown Tool

A browser-based tool that automatically expands your SKU price list by channel and Nigerian state.

## What it does

Upload any Excel price list and get back a fully expanded breakdown:

- **All Channels** → expanded into Retail, Horeca, Modern Trade
- **All Location** → expanded into all 36 states + FCT Abuja (37 rows)
- **"Other region except X"** → all states minus the excluded ones
- **Comma-separated states** → each state gets its own row
- **Latest price logic** → specific-state date wins over All Location; most recent always takes priority

Output is a 3-sheet Excel file: Summary, Full Expanded, Latest Prices.

---

## How to deploy (takes ~5 minutes)

### Step 1 — Create a free Vercel account
Go to https://vercel.com and sign up with your GitHub account (or email).

### Step 2 — Upload this project to GitHub
1. Go to https://github.com and create a new repository (call it `sku-price-tool`)
2. Upload all the files in this folder to that repository

### Step 3 — Deploy on Vercel
1. On Vercel, click **"Add New Project"**
2. Import your `sku-price-tool` GitHub repository
3. Vercel will auto-detect it as a Vite project
4. Click **Deploy**
5. In ~60 seconds you'll get a permanent URL like `https://sku-price-tool.vercel.app`

### Step 4 — Share the link
Send the URL to your team. Anyone can visit it, upload a file, and download the breakdown.

---

## Running locally (optional)

If you want to run it on your own computer first:

```bash
# Install Node.js from https://nodejs.org if you don't have it

# In this folder, run:
npm install
npm run dev

# Open http://localhost:5173 in your browser
```

---

## File structure

```
sku-price-tool/
├── index.html          # Entry point
├── package.json        # Dependencies
├── vite.config.js      # Build config
├── vercel.json         # Deployment config
├── public/
│   └── favicon.svg
└── src/
    ├── main.jsx        # React entry
    ├── App.jsx         # UI and state
    ├── App.css         # Styles
    ├── index.css       # Global styles
    └── breakdown.js    # All processing logic
```

## Updating the tool

To change the logic (e.g. add new states, change channel names):
- Edit `src/breakdown.js` — `ALL_STATES` and `CHANNELS` are at the top
- Push to GitHub → Vercel auto-redeploys in ~60 seconds
