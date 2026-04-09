# Development Log (DEV_LOG.md)

## [2026-04-09] Stabilization: Localization & Branch Merge

### 1. Issue Description
-   Persistent `[object Object]` rendering errors across AI diagnosis cards.
-   Typography inconsistency and low sidebar contrast in Light mode.
-   Merge conflicts between `main` and `master` preventing stable deployment.

### 2. Root Cause Analysis (RCA)
-   **Localization Regression**: The `SPCEngine` results were recently upgraded to return objects `{zh, en}`, but the UI rendering logic in `app.js` was still concatenating them as raw objects, leading to string coercion.
-   **Contrast Violation**: The sidebar used a pure white background with light borders, failing WCAG AA standards in light mode.
-   **Merge Desync**: Divergence between functional fixes in `master` and UI refinements in `main` caused syntax errors and layout breaks after improper merging.

### 3. Corrective and Preventive Action (CAPA)
-   **Granular Localization**: Refactored `renderAnalysisView` to explicitly check `this.settings.language` before accessing `.zh` or `.en` properties of diagnostic results.
-   **Visual Excellence**: Transitioned to the `Inter` font stack and implemented a `slate-50` sidebar background with reinforced borders for better hierarchy.
-   **Branch Finalization**: Performed a hard reset of `main` to `a01977a`, merged `master` with manual conflict resolution, and validated the syntax integrity.

---


## [2026-04-08] Bug Fix: History Deletion Reliability (RCA/CAPA)

### 1. Issue Description
Users reported that deleting items from the "Recent Files" (Dashboard) or the History Tab often resulted in the wrong item being deleted or the UI failing to update correctly. This was a recurring regression.

### 2. Root Cause Analysis (RCA)
- **Index-Based Dependency**: The deletion logic was tightly coupled to the array index of the `spc_history` array.
- **View Slicing**: The Dashboard only displays a slice of the history (`.slice(0, 5)`). If a user deleted the first item on the dashboard (UI index 0), the application would remove the 0th item of the *entire* history array, which might not be the same record if the list was sorted or filtered.
- **State Desync**: Manual updates to `localStorage` were scattered across various functions, leading to inconsistent persistence states between the Dashboard and the History view.

### 3. Corrective and Preventive Action (CAPA)
- **Architecture Shift (Unique IDs)**:
    - Implemented a unique ID generator for every history item: `h_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`.
    - Migrated all history operations (`delete`, `loadDetail`, `render`) to use these unique IDs instead of integers.
- **Backward Compatibility**:
    - Updated `loadFromHistory()` to detect legacy records (missing `id`) and automatically assign them a unique ID upon first load.
- **Centralized Persistence**:
    - Created `saveHistoryState()` as the single authority for committing history changes to `localStorage` and synchronizing internal state.
- **Event Delegation**:
    - Refactored `renderRecentFiles` and `renderHistoryTable` to attach the unique ID to the DOM elements, ensuring the correct reference is passed to the deletion handler.

### 4. Verification Results
- **Dashboard Deletion**: Verified that deleting any item from the "Top 5" list correctly removes that specific item from the database.
- **History List Consistency**: Verified that sorting and then deleting from the full history tab correctly identifies the target record via ID.
- **Legacy Migration**: Confirmed that old browser data is successfully upgraded with IDs without data loss.

---

## [2026-04-08] Bug Fix: Event Listener Stacking in renderHistoryView (RCA/CAPA)

### 1. Issue Description
Despite the ID-based migration, history deletion continued to fail intermittently. Users reported multiple `confirm()` dialogs appearing or items not disappearing from the UI.

### 2. Root Cause Analysis (RCA)
- **Listener Stacking**: `renderHistoryView()` called `tbody.addEventListener('click', ...)` on every render without removing previous listeners.
- **Cascade Effect**: After N delete+re-render cycles, N identical listeners were active on the same `<tbody>`. Each click fired all N handlers, causing N `confirm()` dialogs and N calls to `deleteHistoryItem()`.
- **Global Reference**: Dashboard inline `onclick` handlers used `SPCApp.xxx()` instead of `window.SPCApp.xxx()`, which could fail in strict mode or when the global scope chain resolves differently.

### 3. Corrective and Preventive Action (CAPA)
- **Clone-Replace Pattern**: On each `renderHistoryView()` call, the `<tbody>` is cloned (shallow, stripping all listeners) and replaced in the DOM. This guarantees exactly one click listener per render cycle.
- **window.SPCApp**: Dashboard inline handlers now use `window.SPCApp` for unambiguous global resolution.
- **Diagnostic Logging**: Added `console.log` traces to `deleteHistoryItem`, `clearHistory`, and `saveHistoryState` for future debugging.

---

## [2026-04-08] UI/UX: Bilingual Localization Completion

### 1. Objective
Complete the localization audit to ensure 100% bilingual support (Simplified Chinese & English) across all charts and detailed data reports.

### 2. Implementation
- **Chart Engines**: Wrapped hardcoded control chart titles, series names (X-Bar, UCL, CL, LCL, Range), and axis labels in `this.t(zh, en)`.
- **Detailed Tables**: Localized metadata fields including 'Quality Dept.', 'Avg/Range', and statistical headers.
- **Expert System**: Refactored `nelsonExpertise` injection to ensure Nelson Rules advice and technical labels are consistently translated.
- **Refinement**: Standardized on system-wide `data-en`/`data-zh` attributes for HTML elements and `t()` helper for JS-driven content.

### 3. Verification
- Verified English mode displays "Extended Shewhart X̄ Chart" vs "X̄ Control Chart" correctly.
- Verified Tooltips in both X̄ and R charts show correct translations for violations and advice.
