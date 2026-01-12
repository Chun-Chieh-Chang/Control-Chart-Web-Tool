# Development Log & Lessons Learned
Date: 2026-01-12

## Recent Feature: Excel Export Optimization

### 1. Objective
Enable high-fidelity Excel export with native charts and support for varying cavity counts (pagination).

### 2. Attempts & Challenges
- **Attempt 1: ExcelJS Template Loading**
  - **Issue:** ExcelJS (browser version) has limitations in preserving complex chart objects when reading/writing templates. Charts often disappeared or became uneditable.
- **Attempt 2: Xlsx-Populate Library**
  - **Issue:** Introduced `xlsx-populate` to preserve template binary structure.
  - **Failure:** The library faced severe loading/compatibility issues in the browser environment (`ReferenceError`, `Library not loaded`), causing the entire export function to crash.
- **Attempt 3: Hybrid/Manual Mode (Current Stable Solution)**
  - **Decision:** Reverted to a robust "Manual Generation" approach using `ExcelJS`.
  - **Implementation:** 
    - Construct the Excel file structure programmatically (Rows, Columns, Styles).
    - **Charts:** Removed static image insertion to avoid duplication/confusion across pages.
    - **Pagination:** Implemented correct "Local Statistics" (Mean, Range) calculation for each page (25 batches per page).
    - **UX:** Added Auto-fit column width logic to adapt to content length (especially for Chinese characters).

### 3. Key Lessons
- **Browser Compatibility:** Heavy Excel processing libraries (like `xlsx-populate`) can be unstable in client-side bundling without a proper build step (Webpack/Vite). Pure CDN inclusion is risky.
- **Stability First:** "Native Editable Charts" is a nice-to-have, but "Working Data Export" is critical. Improving the manual data presentation (Auto-fit, correct stats) provided a better immediate value than broken native charts.
- **Complexity Management:** Separating the Excel logic into `js/excel-builder.js` was a good MECE move, keeping `app.js` cleaner.

### 4. Current Status
- **Excel Export:** Stable (Manual Mode).
- **UI:** Anomaly Sidebar toggle fixed.
- **Structure:** Code is modularized (`excel-builder.js`).

### 5. Next Steps
- Validate the Auto-fit column width in actual usage.
- Consider server-side generation if "Native Editable Charts" becomes a hard requirement in the future.
