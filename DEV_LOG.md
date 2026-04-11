# Development Log (DEV_LOG.md)

## [2026-04-10] Milestone: Industrial SPC Diagnostic Architecture & ISO 80002-2 Validation

### 1. Issue Description
- **Linguistic Ambiguity**: Users often confused "Batch Analysis" with "Multi-Cavity Analysis," leading to incorrect process adjustment decisions.
- **LaTeX Rendering Fragility**: Math formulas in the validation report and modals were inconsistently rendered due to CDN loading timing and delimiter conflicts.
- **Technical Gaps**: The previous "Math Principles" modal was too academic and lacked practical "Management Action" guidance.

### 2. Root Cause Analysis (RCA)
- **Conceptual Overlap**: The tool initially presented data without explicit context on whether variance was "Within-Shot" (Mold) or "Shot-to-Shot" (Machine).
- **DOM Life-cycle Issues**: `renderMathInElement` was firing before modal content was fully visible or dynamic content was injected, leaving raw `$$` or `\[` strings.
- **Artifact Pathing Error**: During development, an attempt was made to write artifacts to the user workspace instead of the App Data Directory, leading to tool failures.

### 3. Corrective and Preventive Action (CAPA)
- **Diagnostic Tiering**:
    - Implemented **"Expert Selection Matrix"** in `index.html`.
    - Defined **Imbalance Ratio ($Ratio$)** as a key KPI for mold health (>5% threshold).
- **Hardened Rendering Engine**:
    - Standardized on `\[ ... \]` and `\( ... \)` delimiters for the most robust parsing.
    - Implemented a **Double-Trigger Pattern**: `window.onload` for the static report and a manual `renderMathInElement(modal)` trigger in `SPCApp.showMathPrinciples()` for the app view.
- **ISO 80002-2 Compliance**:
    - Reconstructed `Consolidated_Technical_Validation_Report.html` as a standalone audit file.
    - Integrated a **Raw Data Appendix** (Audit Trail) showing all 81 test data points for 100% traceability.

### 4. Technical Achievements
- Successfully bridged the gap between raw statistical math and **Injection Molding Management Logic**.
- Achieved "Zero-Latency" mathematical rendering for complex formulas- [x] 統計符號全球化統一 ($\sigma_{within}$, $\sigma_{overall}$)
- [x] 詳細數據表格標頭換行優化 (Long labels auto-wrap)
- [x] 圖表標籤字體微縮 (7.5px) 與 Scale 補償
- [x] `file:///` 協議兼容性優化 (移除路徑 Query String)

### 2026-04-10 | 技術故障排除報告 (UI & Rendering)

#### 1. 現象分析 (Root Cause Analysis)
- **問題 A**：JS 修改字體大小無效。
  - **診斷結果**：Chrome 瀏覽器存在 12px 最小顯示字體限制；且全域 CSS 中存在 `!important` 鎖死字體數值。
- **問題 B**：JS 邏輯（換行）完全沒冒出來。
  - **診斷結果**：`file:///` 協議下，路徑中的版本號（`?v=1.2.1`）觸發了安全性攔截，導致瀏覽器讀取舊版快取。

#### 2. 矯正措施 (CAPA)
- **策略 A**：實施內聯 CSS 補強，利用 `transform: scale(0.8)` 搭配 `transform-box: fill-box` 強制實施視覺縮放，成功繞過 12px 限制。
- **策略 B**：淨化所有靜態資源路徑，移除 `?v=`。改用內聯 CSS 確保高優先級樣式第一時間生效。
- **策略 C**：為表格數據與標頭手動加上 `!important` 樣式。

#### 3. 目前狀態
- **已儲存並檢驗**：所有核心檔案已對齊「極致壓縮」標準。建議用戶在本地查看時執行 **Ctrl + F5**。

## [2026-04-10] Full Audit: AI Diagnosis Stability & UI Dynamic Sync

### 1. Issue Description
-   **Statistical Crash**: During bulk import (e.g., Project B), the AI diagnosis card displayed `[object Object]` instead of human-readable text.
-   **UI/Code Desync**: The SPC Constants reference table was incomplete (8 rows) compared to actual engine support (47 rows).
-   **Runtime Hotfix**: Resolved a syntax error introduced during logic merging.

### 2. Root Cause Analysis (RCA)
-   **Data Skewness**: Project B's distribution triggered a `distWarning` object. The UI logic used string concatenation on this object, causing default `[object Object]` coercion.
-   **Static hardcoding**: The informational modal was built with legacy static HTML, failing to reflect modern high-cavity logic in `engine.js`.
-   **Syntax Oversight**: A missing comma delimiter after `resetSystem()` in `SPCApp` during object insertion.

### 3. Corrective and Preventive Action (CAPA)
-   **Bilingual Protocol**: Updated `renderAnalysisView` to use `this.t()` for all diagnostic objects, securing bilingual output (zh/en).
-   **Dynamic Table Engine**: 
    -   Implemented `SPCApp.renderConstantsTable()` to fetch and generate rows from `SPCEngine` data.
    -   Added **Sticky Header** and **Scroll Area** to the help modal for $n=2\sim48$ data management.
-   **Technical Transparency**: 
    -   Added `docs/SPC_Math_Principles.html` with KaTeX formulas for engineering verification.
-   **Robustness Patch**: Verified and corrected comma delimiters to restore stable runtime.
-   **Stability Fix (Runtime)**: Resolved "Cannot read properties of null (reading 'toFixed')" in `renderCharts()` by implementing null checks for all statistical annotations.

### 4. Next Steps
-   **Regression Audit**: Perform a full scan of Nelson Rules #1-8 to ensure similar bilingual object patterns are implemented.
-   **Performance**: Monitor DOM performance when rendering the 47-row table on lower-end devices.

---

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
- Verified Tooltips in both X̄ and R charts show correct translations for violations and advice.
 
---
 
## [2026-04-10] Perfection: Symbol Standardization & Signature Compliance
 
### 1. Issue Description
- **Subscript Invisibility**: Unicode subscripts (e.g., A₂, d₂) were too thin and small, especially in Dark Mode, making statistical constants unreadable.
- **Regulatory Gaps**: Exported reports pre-filled "Dept" and "Inspector" fields, which violated the requirement for physical handwritten signatures in quality audits.
- **Legend Misalignment**: Data point markers (Circles) were slightly offset vertically compared to line markers (UCL/LCL) in chart legends, creating a "bumpy" visual effect.
 
### 2. Root Cause Analysis (RCA)
- **Font Rendering**: Browsers render Unicode subscript glyphs as a separate font style which often ignores parent `font-weight` and `line-height` settings, leading to poor contrast.
- **Input Over-automation**: The system was designed to be too "smart" by pulling metadata into signature fields, forgetting that some fields MUST be blank for legal traceability.
- **SVG Bounding Box Variance**: ApexCharts calculates vertical positions differently for `scatter` markers vs `line` stroke markers in the legend SVG, leading to sub-pixel misalignment.
 
### 3. Corrective and Preventive Action (CAPA)
- **Semantic & Document Standardization**:
    - Replaced all Unicode characters with HTML `<sub>` tags in `index.html`.
    - **Docs Sync**: Standardized all technical specifications (`README.md`, `SPC_Calculation_Logic.md`) using ASCII subscripts and unified nomenclature ($\sigma_{within}$/$\sigma_{overall}$).
    - Applied `font-weight: 700` and `font-size: 11px` to `sub` elements to ensure readability.
- **Signature Compliance (Blank-out)**:
    - Updated `renderDetailedDataTable` in `app.js` to set signature fields to `''`.
    - **Parallel Expansion**: Synced this change into `js/excel-builder.js` for both template and manual build paths.
- **Universal Legend Alignment**:
    - Implemented a CSS Flex-Center patch for `.apexcharts-legend-series`.
    - Standardized all chart `legend.markers` to **12x12px** width/height across `xOpt`, `rOpt`, `meanOpt`, `stdOpt`, `gOpt`, and `vOpt`.
    - Set `offsetY: 0` to override internal SVG offsets.
 
### 4. Technical Achievements
- Successfully aligned the digital tool with **ISO/Industrial Quality Standards** for handwritten audit trails.
- Achieved "Pixel-Perfect" alignment across all analytical charts.
- Pushed final stable builds to GitHub for production deployment.
## [2026-04-11] Robustness: Nelson Rules Global Normalization & Pagination Sync

### 1. Issue Description
- **Pagination Fragility**: Nelson Rules and Control Limits were being recalculated per分頁 (25 batches), causing:
    - **Limit Jumps**: UCL/CL/LCL shifting when navigating between pages.
    - **Detection Gaps**: Patterns (e.g., Rule 2: 9 points in a row) failing to trigger if they spanned across page boundaries.
    - **Visual Desync**: Charts showing different scales for the same process depending on the current view range.

### 2. Root Cause Analysis (RCA)
- **Local Context Trap**: The `renderCharts()` function called `calculateXBarRLimits()` using `pageData` instead of reusing the global results from `executeAnalysis()`.
- **State Isolation**: The rendering engine treated each page as a new independent dataset for statistical calculation.

### 3. Corrective and Preventive Action (CAPA)
- **Global-First Rendering**:
    - Refactored `renderCharts` to pull limits ($UCL$, $CL$, $LCL$) and violations from the global `analysisResults` object.
    - Implemented **Violation Context Mapping**: Global violation indices are now mapped to page-relative indices during rendering.
- **Statistical Consistency**:
    - Guaranteed identical control limits across all pages for a single analysis session.
    - Enabled cross-page pattern recognition for all 8 Nelson Rules.

### 4. Verification Plan
- [X] Load dataset with >25 batches.
- [X] Verify UCL/LCL labels remain identical on Page 1 and Page 2.
- [X] Construct a Rule 2 violation (9 points same side) crossing index 25 and verify it is correctly marked on both pages.

---

## 2026-04-11: Phase II Fixed Baseline Monitoring (Lock Baseline)

### 1. User Requirement
- In an evolving process, recalculating limits based on the full dataset (Rolling Global Limits) can cause "historical shifting," where previously safe points are re-judged as outliers because the process improved or drifted.
- Need a way to simulate real-world Phase II monitoring by "locking" a stable baseline.

### 2. Implementation Scope
- **UI Enhancements**: Added "Advanced Configuration" in Step 3 with a "Fixed Baseline Monitoring (Phase II)" toggle and baseline point count input (default: 25).
- **Statistical Engine Upgrade**:
    - Modified `SPCEngine.calculateXBarRLimits` to accept `fixedCL` and `fixedSigma`.
    - Implemented capacity for Phase I (Baseline Calculation) vs. Phase II (Monitoring) logic.
- **Workflow Integration**:
    - `SPCApp.executeAnalysis` now checks the lock status.
    - If active, it slices the first N points to calculate a master baseline, then applies these limits to evaluate the entire dataset.
- **Visual Feedback**:
    - Added a blue "Phase II" badge in the analysis view when limits are locked to inform the user of the monitoring context.

### 3. Benefits
- **Consistency**: Historical points will NEVER change their "In Control" status regardless of how many new points are added.
- **Industrial Standards Alignment**: Bridges the gap between "Batch Analysis" and "Real-time Monitoring" (Phase II).

---

## 2026-04-11: Normal Distribution Analysis & Robustness Fixes

### 1. New Feature: Distribution Analysis
- Added a dedicated "Distribution Analysis" model card in Step 3.
- Implemented **Histogram Visualization** using ApexCharts.
- Added **Normality Check** based on Skewness and Kurtosis.
- Integrated comprehensive statistical summary (Mean, Median, StdDev, Cpk/Ppk) into the distribution view.

### 2. Bug Fixes & Robustness
- **TypeError Prevention**: Resolved `Cannot read properties of undefined (reading 'type')` by adding safety guards in `renderAnalysisView`.
- **Cache Breaking**: Added `?v=2.1.1` to script tags to force browser updates.
- **Error Handling**: Added explicit error reporting for unknown analysis types.


## [2026-04-11] Milestone: Industrial-Grade Distribution Analysis (Minitab/Excel High-Fidelity)

### 1. Issue Description & Diagnostics
- **Problem A: Interpolation Wobble (Bitten Curve)**: The Gaussian curve appeared "bitten" or wavy (Runge's phenomenon) especially for narrow distributions (Sigma < 0.001).
- **Problem B: Histogram Gaps & "Needle" Bars**: On a numeric X-axis, ApexCharts would shrink bars to thin needles if curve resolution was high, or leave gaps between bins.
- **Problem C: Scoping Crash**: Runtime error `this.t is not a function` during analysis execution.
- **Problem D: UI Desync**: The cavity selection dropdown was missing or failed to reflect the current file's structure.

### 2. Root Cause Analysis (RCA)
- **RCA-A (Smoothing Artifacts)**: Using `curve: 'smooth'` with only 50 sampling points on a sharp bell curve caused spline-based overshooting.
- **RCA-B (Numeric Axis Constraint)**: ApexCharts determines bar width based on the *minimum* distance between X-points across all series. A high-res line series (200 pts) forces the bar series to shrink to a tiny fraction of the bin width.
- **RCA-C (Context Loss)**: Inside `setTimeout` or nested callbacks, `this` no longer referred to the `SPCApp` instance.
- **RCA-D (State Decoupling)**: The `currentDataInput` object was not being globally updated during analysis execution, causing the UI renderer to see `undefined` metadata.

### 3. Corrective and Preventive Action (CAPA)
- **Algorithm Revolution (Area-Step Histogram)**:
    - Abandoned the native `bar` chart type for histograms.
    - Implemented a **Step-Coordinate Area Engine** to manually draw histogram boxes. 
    - **Result**: Perfect 100% width adjacency with 1px borders, regardless of the fitting curve's resolution.
- **High-Fidelity Fitting (200-pt Bell Curve)**:
    - Increased sampling density from 50 to **200 points**.
    - Switched to `curve: 'straight'` with high density to simulate a perfect vector arc without cubic-spline artifacts.
- **Standardized Scoping**:
    - Enforced `var self = this;` at the start of all major controller functions.
    - Replaced `this.t` with `self.t` in all analysis logic to prevent runtime exceptions.
- **Embedded Control UI**:
    - Moved the **Cavity Selector** directly into the Chart Card's header.
    - Added a mandatory sync `this.currentDataInput = dataInput` in `executeAnalysis` before rendering.

### 4. Technical Achievements
- [x] **Minitab-Style Aesthetics**: 100% bar width, high-resolution smooth bell curve overlay.
- [x] **Statistical Clarity**: Glassmorphism callout bubbles for $\mu$ and $\pm1\sigma, \pm3\sigma$ markers.
- [x] **Interactive Re-analysis**: Seamlessly switching between "Global Data" and "Specific Cavity" with zero-latency recalculation.
- [x] **Robustness**: Eliminated all known context-related runtime crashes in the analysis module.
- [x] **Universal Legend Normalization**: Globally replaced circular legend markers with **25x5px Rounded Line Segments** to accurately represent control limits and fit curves.

### 5. Recurring Issue Audit: Legend Semantic Mismatch (RCA/CAPA)
- **Problem**: Reference lines (UCL/LCL/Fit Curves) repeatedly rendered as "Dots" in chart legends in previous iterations.
- **Root Cause Analysis (RCA)**: 
    1. **ApexCharts Defaulting**: The engine defaults to circular markers for all series if not explicitly styled, leading to a "scatter point" visual language for constant reference lines.
    2. **Configuration Inertia**: Previous fixes used `radius: 12` to solve vertical alignment, which accidentally locked the shape to a circle.
    3. **Visual Semantic Gap**: Failure to distinguish between "Observation Points" (Discrete) and "Process Limits" (Continuous) at the legend metadata level.
- **Corrective and Preventive Action (CAPA)**:
    1. **Standardized UI Token**: Established a mandatory standard for line-based series: `markers: { width: 25, height: 5, radius: 2 }`.
    2. **Global Elimination**: Performed a full codebase audit (grep) to remove all remaining `radius: 12` instances.
    3. **Implementation SOP**: Added a rule to check legend marker shapes during any new chart implementation to ensure adherence to the "Line Segment" aesthetic.

---

