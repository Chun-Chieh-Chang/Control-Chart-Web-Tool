# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased] - 2026-01-11

### Added
- **Data Import Fallback**: Implemented logic to use the source filename as the "Item P/N" (商品料號) when the uploaded Excel file or QIP extraction lacks explicit product metadata.
- **QIP Data Consistency**: Added a mock `selectedFile` object generation for QIP extracted data, ensuring strictly consistent behavior between direct file uploads and internal QIP extractions.

### Fixed
- **StdDev Chart Scale**: Fixed the Standard Deviation chart Y-axis to correctly display 6 decimal places and force a minimum value of 0 for better visibility of small variances.
- **Chart Interaction**: 
    - Disabled accidental scroll zooming on all charts.
    - Implemented manual "Selection Zoom" (drag to zoom, double-click to reset) for non-control charts (Cpk, Mean, StdDev, Group, Variation).
- **Chart Styling**:
    - Forced all Chart Axis labels to `12px` 'Inter' font.
    - Changed Control Limit lines (UCL, CL, LCL) to solid lines (removed dash style) and removed data point markers for a cleaner look.
- **Text Corrections**: Corrected the typo "標差比較" to "標準差比較".
- **Nelson Rules Sidebar**: Fixed a rendering issue where multiple expert advice messages weren't displaying correctly (missing HTML element creation).

### Changed
- **UI Typography**:
    - **Inspection Item Cards**: Increased font sizes for better readability (Title: 18px, Labels: 12px, Values: 14px).
    - **Author Info**: Adjusted footer author text to 15px bold and removed truncation to ensure the full name and year are visible.
