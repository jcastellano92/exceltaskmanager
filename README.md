# Product Management Tool

A lightweight task board (Kanban, list, roadmap, and workstream views) that runs
as an Office Add-in inside Excel. The Excel workbook is the data store — the
add-in renders an interactive UI over a set of structured tables and writes
changes back to them.

## Setup (overview)

1. **Host the static files** (`board.*`, `taskpane*.*`, `data.js`, icons) on any
   static HTTPS host.
2. In each manifest (`manifest-dialog.xml` recommended), replace the
   `https://YOUR-HOST` placeholders with your host and set a support URL.
3. **Sideload** the manifest in Excel (Add-ins → Upload My Add-in), or deploy it
   centrally via your Microsoft 365 admin center.
4. Open the backing workbook (it must contain the expected named tables, e.g.
   `TasksTable`, `WorkstreamsTable`, `GoalsTable`, `SubtasksTable`, `UpdatesTable`,
   `MilestonesTable`, `Config`). See **[SETUP-TABLES.md](SETUP-TABLES.md)** for the
   full list of tables and required columns.

No build step — it's plain HTML/CSS/JS.

## Usage

**Original author's use only.** Use, copying, modification, deployment, or distribution
by anyone else requires the original author's explicit approval. All rights reserved. It
is **not** an official product and is **not affiliated with, endorsed by, or
representative of any company or organization**.

## How it handles data

This is a **visual front end only.** The static host (e.g. GitHub Pages) serves **only
the front-end files** (HTML/CSS/JS) — it never receives or stores any of your data.
There is **no backend, no database, and no analytics.**

Your data lives in **two places only: the Excel workbook and the browser session that
renders the UI** (in memory and the browser's local storage on your own device) — and it
**never leaves either of those locations.** Nothing is transmitted to any external
service. It is engineered so your data stays on your system at all times, never leaving
the Excel workbook or the browser serving it.

Provided **as-is, without warranty of any kind. Use at your own risk.**
