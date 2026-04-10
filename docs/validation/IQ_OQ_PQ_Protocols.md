# IQ/OQ/PQ Validation Protocols

## 1. Installation Qualification (IQ)
| Test ID | Objective | Procedure | Expected Result |
| :--- | :--- | :--- | :--- |
| **IQ-01** | Browser Compatibility | Open tool in Chrome/Edge latest version. | UI renders correctly with glassmorphism effects. |
| **IQ-02** | Library Loading | Check console for script loading status. | No "File Not Found" (404) or "Access Denied" errors. |
| **IQ-03** | Local Storage Init | Check if history logic initializes. | `localStorage` entry for `spc_history` is created or read. |

## 2. Operational Qualification (OQ)
| Test ID | Objective | Procedure | Expected Result |
| :--- | :--- | :--- | :--- |
| **OQ-01** | Algorithm (Mean) | Input [10, 20, 30]. | Mean = 20.000. |
| **OQ-02** | Algorithm (SD) | Input [10, 20, 30]. | Sample SD = 10.000. |
| **OQ-03** | Nelson Rule #1 | Input one point > X̄ + 3σ. | Rule #1 warning triggered in sidebar. |
| **OQ-04** | Language Toggle | Click EN/ZH toggle. | `renderConstantsTable` is re-called with correct labels. |

## 3. Performance Qualification (PQ)
| Test ID | Objective | Procedure | Expected Result |
| :--- | :--- | :--- | :--- |
| **PQ-01** | Bulk Import Stress | Load "Project B" (Batch) files. | No `[object Object]` error; `distWarning` displayed correctly. |
| **PQ-02** | High-Cavity Sync | Select n=48. | `SPCEngine.getConstants(48)` returns $A_2=0.109, D_4=1.437$. |
| **PQ-03** | Gold Standard | Compare results with [Validation_Gold_Standard.xlsx](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/Validation_Gold_Standard.xlsx). | Variance < 1e-4 for UCL/LCL. |
