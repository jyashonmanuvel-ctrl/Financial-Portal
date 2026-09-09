# Yash Finance

A personal finance dashboard — Asset Allocation, Expense & Income, and Budget tracking — built as a static site (no backend, no build step).

## Files

- `index.html` — page structure
- `styles.css` — all styling
- `script.js` — all app logic (tables, modals, Excel/CSV/JSON import-export, Google Sign-In & Drive save)

## Run it locally

No build tools needed. Either:

- Double-click `index.html` to open it directly in a browser (everything works except Google Sign-In — see below), or
- Serve the folder locally, e.g. `npx serve .` or `python3 -m http.server`, then open the printed `http://localhost` address.

## Deploy on GitHub Pages

1. Push these three files to a GitHub repository (root of the repo, or a `/docs` folder).
2. In the repo, go to **Settings → Pages**.
3. Under **Source**, pick the branch (and folder) these files are in, then **Save**.
4. GitHub gives you a URL like `https://yourname.github.io/your-repo/` — that's your live site.

## Enabling Google Sign-In & Drive save

This is optional — the app works fully without it, using local Excel/CSV/JSON export-import instead.

Google Sign-In needs a one-time OAuth Client ID, and only works once the site is hosted at a real `http://`/`https://` address (not opened as a local file). In `script.js`, find:

```js
const GOOGLE_CLIENT_ID = 'PASTE_YOUR_GOOGLE_CLIENT_ID_HERE.apps.googleusercontent.com';
```

Replace the placeholder with your own Client ID. Click the ⓘ icon next to "Sign in with Google" in the app itself for the full step-by-step setup guide (Google Cloud Console → OAuth consent screen → enable Drive API → create Client ID → add your GitHub Pages URL as an authorized origin).
