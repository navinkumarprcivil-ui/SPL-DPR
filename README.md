# SPL DPR — Daily Progress Report System

A mobile-first Daily Progress Report (DPR) app for SPL Infrastructure Pvt. Ltd.
Built with **React + Vite**, backed by **Firebase Realtime Database**, and deployed on **Vercel**.

Engineers submit daily site reports (work activities, material moves, plant & machinery, labour); incharges/management approve them; admins manage projects, users, roles and dropdown lists. Reports export to Excel.

---

## Tech stack

- **React 18** + **Vite 5** (build tool / dev server)
- **Firebase Realtime Database** (data store — accessed via both the JS SDK and the REST API for offline resilience)
- **xlsx** (Excel export)
- **Tabler Icons** (web font, via CDN)
- **Vercel** (hosting)

---

## Project structure

```
index.html            App shell + base styles + icon font
main.jsx              React entry point
App.jsx               The entire application (UI + logic)
firebase.js           Firebase app init + RTDB handle
vite.config.js        Vite + chunking config
vercel.json           Vercel build settings
package.json          Dependencies & scripts
database.rules.json   Reference Realtime Database rules
logo.jpg              Company logo (also inlined as base64 in App.jsx)
```

---

## Run locally

Requires **Node.js 18+**.

```bash
npm install
npm run dev
```

Open the printed local URL (default http://localhost:5173).

Other scripts:

```bash
npm run build     # production build → dist/
npm run preview   # preview the production build locally
```

---

## Firebase setup

The Firebase config lives in `firebase.js`. It currently points at the
`spl-project-dpr` project. Firebase **web API keys are public by
design** — security is enforced by database rules, not by hiding the key.

If you fork this to a new Firebase project:

1. Create a project at https://console.firebase.google.com and add a **Web app**.
2. Enable **Realtime Database**.
3. Replace the `firebaseConfig` object in `firebase.js` with your project's values,
   and update `RTDB_URL` near the top of `App.jsx` to match your database URL.
4. Publish database rules. A permissive starter set is in `database.rules.json`
   (open read/write — tighten before real production use).

### Data model (top-level RTDB keys)

- `users/` — accounts (name, role, pin, caps, project assignment)
- `projects/{id}/` — per-project `submissions`, `engineers`, `config`
- `config/globalLists` — roles, required fields, date-lock setting

---

## Deploy to Vercel

1. Push this folder to the GitHub repository (`navinkumarprcivil-ui/SPL-DPR`).
2. In Vercel, **Add New → Project** and import the repo.
3. Vercel auto-detects Vite. Settings (also declared in `vercel.json`):
   - Framework preset: **Vite**
   - Build command: `npm run build`
   - Output directory: `dist`
   - Install command: `npm install`
4. Deploy. No environment variables are required — Firebase config is in the client.

Every push to the connected branch triggers an automatic redeploy.
