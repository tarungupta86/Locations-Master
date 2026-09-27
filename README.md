# POL Locations PWA

App covering the 124 POL locations (from POL LOCS.xlsx, sheet 124 LOCS) with material-wise tankage (from tankstk 01.06.2026.xlsx, Material Name column J, Tankage column L).

## Files

- `index.html` – the complete app (data embedded, works even without internet)
- `manifest.json` – app manifest (name, icons, standalone display)
- `sw.js` – service worker for offline caching
- `icon-192.png`, `icon-512.png`, `apple-touch-icon.png` – app icons
- `data.json` + `app_template.html` – source data and template, used only to regenerate the app when data changes

## How to use on iPhone 16

A PWA must be served over HTTPS for install and offline mode to work. Easiest options:

1. Upload this folder to any free static host (GitHub Pages, Netlify Drop at https://app.netlify.com/drop, or Vercel).
2. Open the link in Safari on the iPhone.
3. Tap Share, then "Add to Home Screen". It installs like a normal app with its own icon and works offline afterwards.

You can also just open `index.html` directly in any browser to use it without installing.

## Refreshing with new tankage data

Replace the stock card file, re-run the extraction to update `data.json`, and rebuild `index.html` from `app_template.html` (ask Claude to "refresh the POL locations app with the new stock card").
