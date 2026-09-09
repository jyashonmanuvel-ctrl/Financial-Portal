# Yash Finance

A personal finance dashboard — Asset Allocation, Expense & Income, and Budget tracking — built as a static site (no backend, no build step).

## Files

- `index.html` — page structure
- `styles.css` — all styling
- `script.js` — all app logic (tables, modals, Excel/CSV/JSON import-export, Firebase sign-in & cloud save)

## Run it locally

No build tools needed. Either:

- Double-click `index.html` to open it directly in a browser (everything works except Sign-In — see below), or
- Serve the folder locally, e.g. `npx serve .` or `python3 -m http.server`, then open the printed `http://localhost` address.

## Deploy on GitHub Pages

1. Push these files to a GitHub repository (root of the repo, or a `/docs` folder).
2. In the repo, go to **Settings → Pages**.
3. Under **Source**, pick the branch (and folder) these files are in, then **Save**.
4. GitHub gives you a URL like `https://yourname.github.io/your-repo/` — that's your live site.

## Enabling Google Sign-In & Firebase cloud save

This is optional — the app works fully without it, using local Excel/CSV/JSON export-import instead.

It's built on **Firebase Authentication** (Google sign-in provider) + **Cloud Firestore**, and only works once the site is hosted at a real `http://`/`https://` address (not opened as a local file). In `script.js`, find:

```js
const firebaseConfig = {
  apiKey: "PASTE_YOUR_FIREBASE_API_KEY_HERE",
  authDomain: "PASTE_YOUR_PROJECT_ID.firebaseapp.com",
  projectId: "PASTE_YOUR_PROJECT_ID",
  storageBucket: "PASTE_YOUR_PROJECT_ID.appspot.com",
  messagingSenderId: "PASTE_YOUR_SENDER_ID",
  appId: "PASTE_YOUR_APP_ID"
};
```

Replace the placeholders with your own Firebase project's config. Click the ⓘ icon next to "Sign in with Google" in the app itself for the full 8-step setup guide (create a Firebase project → register a web app → enable Google sign-in → create a Firestore database → set a security rule → add your hosted URL as an authorized domain).

The security rule restricts each signed-in person to their own saved data:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /yashFinancePortals/{uid} {
      allow read, write: if request.auth != null && request.auth.uid == uid;
    }
  }
}
```

