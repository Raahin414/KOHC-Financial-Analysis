# KOHC Financial Analysis

Introduction to Business Finance project comparing **Kohat Cement (KOHC)** with **Lucky Cement (LUCK), D.G. Khan Cement (DGKC), Maple Leaf Cement (MLCF) and Fauji Cement (FCCL)** for FY2021-FY2025.

## Start here

- [Latest submission workbook](workbooks/KOHC_Financial_Analysis_FINAL_SUBMISSION.xlsx)
- [Internal audit workbook and reconciliation evidence](workbooks/KOHC_Financial_Analysis_Audit.xlsx)
- [Updated ten-page viva guide](docs/KOHC_Viva_Guide_Updated_10_Pages.pdf)
- [Viva guide as readable Markdown](docs/Viva_Guide.md)
- [Complete financial analysis report - PDF](docs/IBF_KOHC_Ratio_Analysis_Explanation_and_Findings.pdf)
- [Editable report - Word](docs/IBF_KOHC_Ratio_Analysis_Explanation_and_Findings.docx)
- [Detailed company performance comparison](docs/Five_Companies_Detailed_Ratio_Performance_Comparison.pdf)
- [Methodology and reconciliation](docs/Methodology_and_Reconciliation.md)
- [Official-source evidence](evidence/SOURCES.md)
- [Project milestone history](docs/PROJECT_HISTORY.md)

## Latest presentation update

28 October 2026: Basis headers and explanations were removed from all five company ratio tables. Verified no-dividend notes remain beside the Dividend Yield formula. Year addresses and hidden FY2020 column I are unchanged. No financial formula, input, saved result, sample mean or chart changed. The internal Audit is unchanged.

## Current scope and methodology

The latest model retains **27 selected measures, including all 19 mandatory ratios**, 675 company-year ratio calculations, 135 equal-weight sample averages and **54 charts** (a five-company chart and a KOHC-versus-calculated-average chart for every measure). Audited unconsolidated accounts are used consistently. Inventory/receivables use opening-closing averages; asset and return denominators follow the disclosed closing convention. FY2020 openings remain in hidden column I. Vertical and Horizontal Analysis are absent from the submission.

Market measures use official PSX financial-year-end closes or the last prior session with compatible share bases. LUCK split trading began 28-Apr-2025; KOHC began 25-Aug-2025, after FY2025. Genuine zero Dividend Yields remain numeric and count in means. Negative/non-meaningful displays follow the project's n/m convention; signed inputs remain intact. The sample average includes KOHC and is not the whole cement industry.

## Qualifications - read before drawing conclusions

The current Audit contains **455 published-value comparisons: 437 numerical reconstructions and 18 qualified differences**. The original 63-case investigation is preserved. Standardized results were not forced to match differently defined or unresolved published values.

A known implementation/classification issue remains: KOHC short-term accrued mark-up of **PKR3.704884m in FY2025** (FY2024 comparative **PKR0.061253m**) remains in interest-bearing debt despite nil principal utilization. D/E, debt-to-capital and ROIC remain qualified. The Overview and comprehensive report disclose this issue. This repository update does not change financial calculations or certify the model as financially error-free.

The 41-page report and historical methodology documents reflect their saved stage. The latest workbook Overview and updated viva guide take precedence for the shortened layout and current caveats. Earlier 36- and 33-measure workbooks are retained through commit history, not as the current submission.

## History and verification

Commits import actual saved milestones now, in development order. **They are not backdated historical commits.** File changes document market verification, full audit, 27-measure scope, dividend presentation, structural cleanup, chart expansion, decluttering, shortened Overview and the updated viva discussion.

The uploaded submission and Audit are byte-for-byte copies of their saved sources. SHA-256 hashes, read-only validation and source evidence accompany the files. Native Excel recalculation checks in the evidence directory were performed at the documented earlier stage; repository packaging does not claim a new Excel recalculation.

## Rebuild the updated guide

With Python and `openpyxl`, `reportlab`, `pypdf` installed:

```sh
python scripts/build_viva_guide.py --workbook workbooks/KOHC_Financial_Analysis_FINAL_SUBMISSION.xlsx --output docs/KOHC_Viva_Guide_Updated_10_Pages.pdf --markdown docs/Viva_Guide.md
```

Run `python scripts/verify_submission.py` for read-only arithmetic, average, chart, dividend-record and file-integrity checks. These checks preserve the stated financial qualifications.

The guide script reads Excel without saving it. Calibri is used on Windows, with a bundled ReportLab font fallback elsewhere.

## Group

| Member | ERP |
| --- | --- |
| Abdullah Khan Sherwani | 27880 |
| Immaduddin Durrani | 29204 |
| Raahin Tajuddin | 29207 |
| Zuhair Merchant | 29021 |
| Arba | 32571 |

Official annual reports and PSX evidence remain third-party source material; no ownership or reuse license is asserted over them. This public upload, including names and ERP numbers, was expressly authorized by the project owner.
