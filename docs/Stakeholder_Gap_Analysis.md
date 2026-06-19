# Stakeholder & Gap Analysis Log — FairFarm BD (Matir Doctor)

A business analyst work-product: the documented gaps, stakeholders, and
recommendations behind this project's KPIs. (Machine-readable version:
`data/csv/stakeholder_gap_analysis.csv`.)

| Area | Stakeholder | Current-State Gap | Recommended Solution | Priority | Status | Expected Impact |
|---|---|---|---|---|---|---|
| IoT Device Onboarding | Field Operations Team | No standardized BRD for Matir Doctor setup; install steps varied by region, causing setup support tickets. | Documented functional requirements & a standardized device-activation flow in the app. | High | Implemented | Lower Device Setup ticket volume & resolution time |
| Reporting Workflow | Management / Founders | Manual weekly reporting from spreadsheets across teams; high effort, delayed insight. | Mapped reporting process, proposed a single source-of-truth dashboard (this project) replacing manual reports. | High | Implemented | Reduced manual reporting effort; faster decisions |
| Website / App UX | Marketing Team | High bounce rate and unclear CTA on product pages; visitors dropping before checkout. | Translated UX research into business recommendations: simplified checkout, clearer device comparison. | High | Implemented | +25% conversion rate, +30% visitor traffic |
| Customer Support | Customer Support Team | No category tagging on tickets; recurring sensor calibration issues went untracked. | Introduced ticket categorization & CSAT tracking to spot recurring product issues. | Medium | Implemented | Faster issue detection, improved CSAT |
| Sales & Inventory | Sales / Dealer Network | Limited regional visibility into which divisions/devices were driving revenue. | Built a sales & regional performance view (Power BI/Excel) for stock & dealer planning. | Medium | Implemented | Better regional stock allocation |

## Methodology notes

- **Requirements elicitation**: stakeholder interviews (simulated for this
  portfolio project, modeled on realistic IoT/agri-tech product gaps) mapped
  against functional requirements for the Matir Doctor device and the
  reporting workflows around it.
- **Gap analysis**: current-state vs. desired-state comparison per business
  area, scored by priority and expected business impact — the standard BA
  technique referenced on the resume.
- **Validation**: each "Expected Impact" above is checked against the
  dataset in this repo (see `docs/eda_kpi_summary.json` and
  `sql/fairfarm_analysis_queries.sql`) so the numbers aren't just claimed —
  they're demonstrably present in the before/after data.
