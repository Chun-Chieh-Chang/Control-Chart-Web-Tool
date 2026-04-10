# User & Software Requirements Specification (URS/SRS)

## 1. User Requirements (URS)
| Req ID | Requirement Description | Priority |
| :--- | :--- | :--- |
| **UR-01** | System must calculate UCL, CL, LCL for X-Bar charts. | Critical |
| **UR-02** | System must calculate UCL, CL, LCL for R charts. | Critical |
| **UR-03** | System must identify Nelson Rule #1-8 violations in real-time. | High |
| **UR-04** | System must support subgroup sizes (n) from 2 to 48. | High |
| **UR-05** | System must output Cpk and Ppk indices. | Critical |
| **UR-06** | UI must be switchable between Traditional Chinese and English. | Medium |

## 2. Software Requirements (SRS)
| Req ID | Technical Specification | Category |
| :--- | :--- | :--- |
| **SR-01** | Mathematical precision: Results must align with ISO 22514 logic. | Quality |
| **SR-02** | Browser Compatibility: Must run on Chrome (v100+) and Edge stably. | Environment |
| **SR-03** | Data Privacy: All calculations must occur browser-side (No Cloud). | Security |
| **SR-04** | Dynamic Table: Statistical constants table must auto-render based for n. | Functional |
| **SR-05** | Error Handling: Must display warnings for non-normal distributions. | Robustness |

## 3. Performance Metrics
- **Load Time**: Analysis results must render within < 2 seconds for a 1000-point dataset.
- **Accuracy**: Calculation variance against "Gold Standard" must be < 1e-4.
