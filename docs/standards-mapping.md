# Standards mapping

How the workbook supports each standard, and a cheat sheet for the **Control Mapping** column of the risk register.

> This is a working aid, not an official interpretation. Always check the current text of each standard.

## 1. Template element → standard

| Template element | ISO/IEC 27001:2022 | NIST CSF 2.0 | ISO/IEC 42001:2023 |
|------------------|--------------------|--------------|--------------------|
| `02_Scales` (criteria, acceptance) | 6.1.2 a) risk criteria | GV.RM — risk management strategy | 6.1.2 AI risk criteria |
| `03_Application_Inventory` | 6.1.2 (scope of assessment); Annex A.5.9 inventory of assets | ID.AM — asset management | 4.3 scope; `Contains AI/ML?` flag marks AI systems |
| `04_Risk_Register` — scenario, threat, vulnerability | 6.1.2 c) identify risks | ID.RA — risk assessment | 6.1.2 AI risk assessment; 6.1.4 AI system impact assessment |
| Inherent / residual scoring | 6.1.2 d) analyse and evaluate | ID.RA-04, ID.RA-05 | 6.1.2 |
| Treatment decision, controls, mapping | 6.1.3 risk treatment; input to the Statement of Applicability | GV.RM-06; PR / DE / RS / RC categories | 6.1.3 AI risk treatment |
| Risk owner, target date, status | 6.1.2 c) 2) risk owners; 6.1.3 e) treatment plan | GV.RR — roles and responsibilities | 5.3 roles, responsibilities |
| Last review | 8.2 regular assessment; 10.1 continual improvement | ID.IM — improvement | 8.2, 8.3, 10.1 |
| `05_Dashboard`, `06_Risk_Matrix` | 9.3 management review inputs | GV.OV — oversight | 9.3 management review |

## 2. Control-mapping cheat sheet

Frequently used references for the *Control Mapping* column. Identifiers follow ISO/IEC 27001:2022 Annex A and NIST CSF 2.0 categories.

| Risk theme | ISO/IEC 27001:2022 Annex A | NIST CSF 2.0 |
|------------|----------------------------|--------------|
| Account takeover, weak authentication | A.5.17 Authentication information · A.8.5 Secure authentication | PR.AA Identity management, authentication, access control |
| Excessive or stale access | A.5.15 Access control · A.5.18 Access rights · A.8.2 Privileged access rights | PR.AA |
| Unpatched vulnerabilities | A.8.8 Management of technical vulnerabilities · A.8.32 Change management | ID.RA · PR.PS Platform security |
| Insecure configuration | A.8.9 Configuration management | PR.PS |
| Data leakage, insider exfiltration | A.8.12 Data leakage prevention · A.5.12 Classification of information | PR.DS Data security |
| Personal-data exposure | A.5.34 Privacy and protection of PII | PR.DS · GV.OC |
| Cryptography gaps | A.8.24 Use of cryptography | PR.DS |
| Insufficient logging / detection | A.8.15 Logging · A.8.16 Monitoring activities | DE.CM Continuous monitoring · DE.AE Adverse event analysis |
| Network exposure, bot traffic | A.8.20 Networks security | PR.IR Technology infrastructure resilience · DE.CM |
| Supplier / cloud / third-party AI risk | A.5.19 Supplier relationships · A.5.20 Supplier agreements · A.5.23 Cloud services | GV.SC Supply chain risk management |
| Backup and recovery failure | A.8.13 Information backup · A.5.30 ICT readiness for business continuity | PR.DS · RC.RP Incident recovery plan execution |
| Availability, resilience | A.5.29 Security during disruption · A.8.14 Redundancy | PR.IR · RC.RP |
| Secure development flaws | A.8.25 Secure development life cycle · A.8.28 Secure coding | PR.PS |

## 3. ISO/IEC 42001:2023 — Annex A groups for AI-related risks

| Annex A group | Typical risk scenario |
|---------------|-----------------------|
| A.2 Policies related to AI | No AI policy or acceptable-use rules |
| A.3 Internal organization | Unclear accountability for AI systems |
| A.4 Resources for AI systems | Insufficient data, tooling or competence |
| A.5 Assessing impacts of AI systems | Harm to individuals or groups; unassessed societal impact |
| A.6 AI system life cycle | No testing, validation or monitoring; model drift |
| A.7 Data for AI systems | Poor provenance or quality; data poisoning; bias in training data |
| A.8 Information for interested parties | Users not told how AI is used or what its limits are |
| A.9 Use of AI systems | Misuse, over-reliance, missing human oversight |
| A.10 Third-party and customer relationships | Data leakage to, or unmanaged dependency on, a third-party model provider |
