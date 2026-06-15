# AircraftLog

A single-page React app for tracking aircraft flight times (Hobbs/Tach/fuel/oil), dispatching flights with airborne-time reminders, and managing maintenance compliance — required inspections (Annual, Transponder, Altimeter, ELT, optional 100-hour), ADs, squawks, and an aircraft registry.

All data is stored in the browser's `localStorage`, so it persists across sessions on the same device/browser but is not synced anywhere else.

## Run locally

Requires [Node.js](https://nodejs.org/) 18+.

```bash
npm install
npm run dev
```

Then open the printed local URL (typically `http://localhost:5173`).

## Build for production

```bash
npm run build
```

This outputs a static site to `dist/`. Preview it locally with:

```bash
npm run preview
```

## Deploy

`dist/` is a fully static site — any static host works. A few easy options:

### Vercel
```bash
npm i -g vercel
vercel
```
Vercel auto-detects the Vite config; just confirm the build command (`npm run build`) and output directory (`dist`).

### Netlify
```bash
npm i -g netlify-cli
npm run build
netlify deploy --prod --dir=dist
```

### GitHub Pages
1. `npm run build`
2. Push the contents of `dist/` to a `gh-pages` branch (or use a GitHub Action with `actions/deploy-pages`).
3. If deploying to a subpath (e.g. `username.github.io/repo-name`), set `base: "/repo-name/"` in `vite.config.js`.

## Notes

- Data lives in `localStorage` under keys prefixed `acft_*`. Clearing browser storage/site data will erase all flight logs, compliance records, squawks, and profile info.
- Since each browser/device has its own storage, this is best suited to a single user or a shared device (e.g. a tablet in the hangar) rather than multi-device sync. Adding a backend (e.g. Supabase, Firebase) would be the natural next step for multi-user/multi-device support.
- The Excel export uses [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs), bundled as a dependency.
