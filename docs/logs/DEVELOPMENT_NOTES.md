# Development Notes

## Architecture Overview
This is a strictly client-side Single Page Application (SPA) for Statistical Process Control (SPC) analysis.
- **Frontend**: HTML5, Vanilla JavaScript.
- **Styling**: Tailwind CSS (CDN).
- **Data Processing**: SheetJS (xlsx.full.min.js) for interpreting Excel files.
- **Visualization**: ApexCharts for all control charts and histograms.

## Data Flow
1.  **Input**: User uploads an `.xlsx` file or extracts data via the QIP module.
2.  **Parsing**:
    - `js/input.js` (`DataInput` class) normalizes the raw Excel data.
    - **Crucial**: If the Excel file lacks a "Product Name" cell, the system matches the filename as `productInfo.item`.
3.  **Analysis**:
    - `js/engine.js` (`SPCAnalyser` class) performs statistical calculations (Mean, StdDev, Cp, Cpk, Pp, Ppk).
    - Checks for Nelson Rules (1-8) violations.
4.  **Rendering**:
    - `js/app.js` orchestrates the DOM updates and Chart initialization.

## Key Technical Decisions

### Chart Interaction Logic
- **Scroll Zoom**: globally **DISABLED** to prevent page scroll interference.
- **Control Charts (X-Bar, R)**: Zooming is locked to preserve the view of the process limits.
- **Analytical Charts (Mean, StdDev, Cpk)**:
    - Implemented a custom **Selection Zoom**. Users can drag to select an area to inspect.
    - **Double-click** anywhere on the chart to reset the view.

### Nelson Rules Implementation
- Rules are checked sequentially in `js/engine.js`.
- Violations are stored as an array of rule IDs.
- The Sidebar renders expert advice by mapping these IDs to the `nelsonExpertise` dictionary in `app.js`.

### Data Consistency (QIP vs Direct Upload)
To ensure the "History" and "Analysis" modules treat all data sources equally:
- When data comes from QIP Extraction, a **Mock File Object** is created.
- This mock object mimics the standard `File` API structure, ensuring that `saveToHistory` and `DataInput` have access to a valid "filename" property.

## Deployment
- The project is deployed via **GitHub Pages**.
- No build step is required (static assets only).
- Ensure `.nojekyll` exists in the root to prevent GitHub Pages from ignoring folders starting with `_` (though not strictly used here, it's best practice).
