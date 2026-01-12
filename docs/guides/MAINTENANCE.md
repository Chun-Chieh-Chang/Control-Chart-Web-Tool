# Project Reorganization & MECE Maintenance Guide

This document outlines the organized structure of the **Mouldex Control Chart Builder** project following the **MECE (Mutually Exclusive, Collectively Exhaustive)** principle.

## 1. Project Directory Structure

### root/ (Entry & Core)
- `index.html`: Main application entry point.
- `README.md`: Project overview and quick start.
- `AI數字人.png`: UI Author avatar.
- `favicon.svg`: Site icon.
- `package.json`: Project metadata and dependencies.
- `docs/`: Documentation, logs, and guides.
- `js/`: Application logic.
- `css/`: Stylesheets.
- `_archive/`: Legacy files and sample exports (Keep clean).

### js/ (Source Code)
- `app.js`: Main application controller and UI bridging.
- `engine.js`: Pure statistical logic for SPC (Nelson Rules, Limits).
- `excel-builder.js`: Manual Excel report generation logic (Manual Mode).
- `qip-app.js`: QIP Analysis entry point.
- `qip/`: Sub-modules for QIP parsing and extraction.
- `template-store.js`: Storage for Excel templates (Base64).

### docs/ (Documentation System)
- `logs/`: Contains development history, fix summaries, and bug patches.
- `guides/`: User manuals, deployment SOPs, and technical deep-dives.

---

## 2. MECE File Management Guidelines

To maintain a clean and efficient workspace, follow these rules:

### Mutually Exclusive (ME)
- **Source Code vs. History**: Never leave `.bak`, `_old`, or date-suffixed scripts in `js/`. Move them to `_archive/` immediately.
- **Root vs. Docs**: The project root should ONLY contain entry points and essential configuration. All descriptive documents belong in `docs/`.
- **Assets vs. Logic**: Keep images in the root or an `assets/` folder, and library scripts in a separate `lib/` folder if added.

### Collectively Exhaustive (CE)
- **Traceability**: Every major bug fix or technical pivot must be recorded in `docs/logs/`.
- **Instructional**: Every module (SPC, QIP) must have an associated guide or clear section in `README.md`.
- **Build Readiness**: Ensure all scripts referenced in `index.html` are present and correctly paths.

## 3. Deployment (GitHub Pages)

The project is configured for GitHub Pages via the root `index.html`.
- Always verify that paths in `index.html` are relative (e.g., `js/app.js` instead of `/js/app.js`).
- Pushing to the `main` branch will automatically update the live version.
