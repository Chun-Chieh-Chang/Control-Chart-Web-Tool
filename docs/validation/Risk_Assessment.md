# Software Risk Assessment (FMEA)

## 1. Risk Methodology
- **S (Severity)**: Impact on product quality/patient safety (1-5).
- **P (Probability)**: Likelihood of failure (1-5).
- **D (Detectability)**: Ability of system to detect/warn of failure (1-5).
- **RPN (Risk Priority Number)**: S × P × D.

## 2. Risk Matrix
| ID | Failure Mode | S | P | D | RPN | Mitigation (CAPA) |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **RA-01** | Incorrect $A_2, d_2$ constants used for rare $n$. | 3 | 2 | 2 | **12** | **Dynamic Table Validation**: Automated unit tests for all $n=2\sim48$ entries. |
| **RA-02** | Nelson Rule misfire (False Positive). | 2 | 3 | 2 | **12** | **Bilingual Logic Review**: Double check logic with `nelsonExpertise` strings. |
| **RA-03** | Skewed data causes invalid Cpk. | 2 | 3 | 1 | **6** | **Auto-Warning System**: Trigger `distWarning` object if Skewness > 1. |
| **RA-04** | Browser caching old script (Syntax Error). | 1 | 4 | 2 | **8** | **Version Tagging**: use `?v=x.x.x` in `<script>` tags (Implemented). |
| **RA-05** | Bulk import memory overflow. | 2 | 2 | 2 | **8** | **Chunk Processing**: ensure UI remains responsive during bulk ops. |

## 3. Residual Risk Statement
With the implemented auto-warnings (RA-03) and dynamic table logic (RA-01), the highest residual risks are categorized as **Low**, which is acceptable for QMS software under ISO 80002-2 guidelines (Category: Non-release decision support).
