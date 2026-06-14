# Takshashila — Attendance

> *Inspired by the Past, Empowering the Future*

A **zero-install, offline-first Progressive Web App** for tracking class attendance — built with vanilla HTML/JS and Google Sheets as the database. No frameworks. No backend server. No app store.

---

## What it does

Open the app, tap each student who is present, press **Save**. That's it.

- Every save writes directly into a Google Sheet — students as columns, dates as rows, **P / A** colour-coded green and red
- Works completely **offline** — saves queue locally and sync automatically when connectivity returns
- Installable on any phone or desktop from the browser — no app store required

---

## Features

| | |
|---|---|
| **Mark Attendance** | Tap cards to cycle: Unmarked → Present → Absent. Bulk mark all / reset. Edit any past date. |
| **Records Dashboard** | Per-student attendance %, at-risk filter (<75%), full session history. Tap any row to jump back and edit. |
| **Alumni Management** | Archive graduated students. Their historical records are preserved. Re-admit with one tap. |
| **Offline Queue** | Saves made without internet are stored locally and retried automatically when back online. |
| **Historical Import** | One-time bulk import from `historical_data.json` — chunked in batches of 200 to avoid timeouts. |
| **CSV Export** | Download all records as a CSV file for further analysis. |
| **PWA Install** | Add to home screen on iOS, Android, or desktop — works like a native app. |

---

## Tech Stack

```
Frontend      Vanilla HTML + CSS + JavaScript  (single file, no build step)
Backend       Google Apps Script (doGet / JSONP)
Database      Google Sheets  (one tab per year, dates as rows, students as columns)
Offline       Service Worker  +  localStorage
Install       Web App Manifest  (PWA)
Typography    Playfair Display  +  Inter  (Google Fonts)
```

---

## Architecture

```
┌─────────────────────────────┐        JSONP (GET)        ┌──────────────────────┐
│      index.html             │ ◄─────────────────────── │  Google Apps Script  │
│  (UI + all JS logic)        │                           │     (Code.gs)        │
│                             │ ──────────────────────── ►│                      │
└──────────┬──────────────────┘       save / fetch        └──────────┬───────────┘
           │                                                          │
     Service Worker                                           Google Sheets
     (sw.js — cache,                                     (one tab per year,
      offline queue)                                      P/A heatmap grid)
           │
     localStorage
     (students, records,
      sync queue, config)
```

The app talks to Google Sheets via **JSONP** — this lets a static HTML file (hosted anywhere, even a local file) call the Apps Script endpoint without CORS issues. There is no server, no database to maintain, no cost.

---

## File Structure

```
Attendance-App/
├── index.html            — The entire frontend (UI, styles, all JavaScript)
├── sw.js                 — Service Worker: asset caching + offline support
├── manifest.json         — PWA manifest: name, icons, theme, install behaviour
├── Code.gs               — Google Apps Script backend (paste into Apps Script editor)
├── historical_data.json  — Seed file for one-time historical import
└── Attendance.xlsx       — Original spreadsheet reference
```

---

## Setup

**1. Create a Google Sheet**

Make a new Google Sheet. Copy the Sheet ID from the URL:
```
https://docs.google.com/spreadsheets/d/SHEET_ID_IS_HERE/edit
```

**2. Deploy the Apps Script backend**

- In the Sheet, go to **Extensions → Apps Script**
- Replace all the code with the contents of `Code.gs`
- Click **Deploy → New deployment → Web App**
- Set *Execute as*: **Me**, *Who has access*: **Anyone**
- Copy the deployment URL

**3. Open the app**

- Open `index.html` in a browser (or host it anywhere — GitHub Pages, a CDN, etc.)
- Go to **Settings**, paste the Apps Script URL and your Sheet URL
- Add your student names (one per line)
- Click **Save & Test Connection**

That's it. The app is ready.

---

## How the Google Sheet looks

| Date | Aarav Shah | Priya Nair | Rohan Mehta |
|------|-----------|-----------|-------------|
| 1/6/25 | **P** | **P** | **A** |
| 2/6/25 | **A** | **P** | **P** |
| 3/6/25 | **P** | **A** | **P** |

Cells are green for Present, red for Absent. Each year gets its own tab (2025, 2026 …) auto-created on first use.

---

## PWA — Install on any device

Because of `manifest.json` and `sw.js`, any modern browser will offer an **"Add to Home Screen"** prompt:

- **Android / Chrome** — install banner appears automatically
- **iOS / Safari** — Share → Add to Home Screen
- **Desktop / Chrome or Edge** — install icon appears in the address bar

Once installed, the app launches full-screen with a splash screen, no browser chrome, and loads instantly even with no internet.

---

## Offline behaviour

The Service Worker caches `index.html`, `manifest.json`, icons, and fonts on first load. After that:

1. All reads come from cache — the app opens instantly offline
2. Attendance saves are written to `localStorage` first, then queued for sync
3. When the device comes back online, the sync queue drains automatically
4. Google API calls are never intercepted — they always go live

---

## Customisation

| What | Where |
|------|-------|
| Secret key | `SECRET_KEY` in `Code.gs` and matching field in Settings |
| App name / colours | `manifest.json` and CSS `:root` variables in `index.html` |
| At-risk threshold | `setFilter()` function (currently 75%) |
| Import chunk size | `CHUNK` constant in `importHistoricalData()` (currently 200) |

---

## Why "Takshashila"?

[Takshashila](https://en.wikipedia.org/wiki/Taxila) (also Taxila) was one of the world's earliest universities, founded around 700 BCE in ancient India. It drew students from across Asia to study philosophy, medicine, grammar, and statecraft. This app is named in that spirit — a small tool for teachers, carrying the idea forward.

---

*Built with no dependencies. Runs anywhere. Costs nothing.*
