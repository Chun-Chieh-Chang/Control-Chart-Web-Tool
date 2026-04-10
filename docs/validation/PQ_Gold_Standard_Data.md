# Performance Qualification (PQ) - Gold Standard Data Specifications

> [!NOTE]
> This document specifies the standard inputs and expected outputs to be used during Performance Qualification (PQ).

## Dataset 1: Standard Subgroup (n=5)
This test verifies the default Shewhart calculation accuracy.

### Input Data Points
- **Subgroup 1**: [1.52, 1.54, 1.51, 1.53, 1.52]
- **Subgroup 2**: [1.55, 1.53, 1.52, 1.54, 1.53]
- **Subgroup 3**: [1.50, 1.51, 1.52, 1.53, 1.51]

### Expected Output (Verification values)
| Metric | Expected Result (4 Decimals) |
| :--- | :--- |
| **CL (X-double-bar)** | 1.5260 |
| **R-bar** | 0.0300 |
| **UCL (X-bar)** | 1.5433 (使用 $A_2=0.577, \sigma_{within}$ 推估) |
| **LCL (X-bar)** | 1.5087 |
| **UCL (R)** | 0.0634 (使用 $D_4=2.114$) |
| **$\sigma_{within}$** | 0.0129 (推估值) |
| **$\sigma_{overall}$** | 0.0135 (實測值) |

---

## Dataset 2: High-Cavity Subgroup (n=16)
This test verifies the accuracy of the extended constants table ($n=16$).

### Input Data (Sample of n=16)
- [10.2, 10.5, 10.3, 10.4, 10.2, 10.1, 10.6, 10.3, 10.4, 10.2, 10.1, 10.5, 10.3, 10.4, 10.2, 10.3] (Single shot)

### Expected Output
| Metric | Expected Result |
| :--- | :--- |
| **X-bar (Shot 1)** | 10.3000 |
| **Range (Shot 1)** | 0.5000 |
| **$A_2$ (n=16)** | 0.2120 |
| **$d_2$ (n=16)** | 3.5320 |

---

## Acceptance Criteria
The software results must match the values above with a variance of **no more than 0.0001 (0.01%)**. If the software correctly identifies these values, the PQ-03 test is considered **PASSED**.
