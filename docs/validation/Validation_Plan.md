# Software Validation Plan (SVP) - SPC Analysis Tool

> [!IMPORTANT]
> **Document ID**: SVP-SPC-2026-001  
> **Reference Standard**: ISO 80002-2:2017  
> **Status**: DRAFT  

## 1. Introduction
This document defines the validation strategy for the **SPC Analysis Tool** (Web-based). The validation ensures that the software consistently performs according to its intended use in the medical device quality management system (QMS).

## 2. Scope
- **Included**: X-Bar Chart calculations, R Chart calculations, Nelson Rules #1-8 detection, Cpk/Ppk indices, and Bilingual UI display.
- **Excluded**: Infrastructure validation (OS, Browser core engine - these are qualified by manufacturers).

## 3. Validation Strategy (Risk-Based)
Based on ISO 80002-2:2017, we apply the following lifecycle:
1. **Definition**: Establish URS (User Requirements Specification).
2. **Risk Assessment**: Identify failure modes (FMEA).
3. **Qualification**:
    - **IQ**: Installation in designated browser environments.
    - **OQ**: verification of algorithms using edge cases.
    - **PQ**: Validation of consistency with production-like data (Gold Standard).

## 4. Roles and Responsibilities
| Role | Responsibility |
| :--- | :--- |
| **Project Lead** | Overall validation management and approval. |
| **QA/RA** | Review and audit for compliance. |
| **Developer (Antigravity)** | Creation of software and validation documentation. |

## 5. Acceptance Criteria
- 100% of "Critical" requirements in URS must be tested and passed.
- No "High" risk items remain without mitigation in the Risk Assessment.
- PQ results must match the "Gold Standard" datasets to 4 decimal places.

## 6. Schedule
- **Phase 1**: Requirements & Risk Analysis (Today)
- **Phase 2**: Protocol Execution (Today)
- **Phase 3**: Final Report (Today)
