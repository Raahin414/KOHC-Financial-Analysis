# Historical share-price and market-basis audit

Completed 22 October 2026. Scope: the user's mandatory PSX historical-price verification and the directly affected market calculations. The broader financial-audit instructions in the pasted attachment were treated as background. Non-market financial inputs, ratio selection and existing non-market findings were preserved and are not newly certified.

## Verified unadjusted closing prices

All 25 observations were independently retrieved from the official [PSX Historical Data service](https://dps.psx.com.pk/historical). The CLOSE column was used, rather than LDCP, OPEN or LOW. No current, estimated or third-party price was used.

| FY | Actual trading date | KOHC (Rs) | LUCK (Rs) | DGKC (Rs) | MLCF (Rs) | FCCL (Rs) |
|---|---|---:|---:|---:|---:|---:|
| 2021 | 2021-06-30 | 206.49 | 863.44 | 117.92 | 46.98 | 23.00 |
| 2022 | 2022-06-30 | 130.13 | 459.04 | 62.50 | 27.35 | 14.17 |
| 2023 | 2023-06-27 | 173.47 | 522.09 | 51.30 | 28.33 | 11.76 |
| 2024 | 2024-06-28 | 250.43 | 906.73 | 90.27 | 38.00 | 22.91 |
| 2025 | 2025-06-30 | 379.36 | 355.24 | 165.56 | 84.28 | 44.67 |

The 2023 last session was 27 June. PSX date searches returned no records on 28, 29 or 30 June, and the June symbol history ends on 27 June. The 2024 fiscal year-end was a Sunday, following Saturday 29 June; the selected session is Friday 28 June. Each observation's request parameters, raw PSX row, official endpoint and response hash are retained in the audit workbook.

## Corrections

- Nine price cells changed in each workbook: five KOHC years and four pre-split LUCK years. All 25 prices were checked, including the 16 unchanged inputs.
- KOHC's five-for-one split was executed on 23 August 2025; first split trading was 25 August 2025. This is after FY2025. The June 2025 close is **Rs379.36**, paired with pre-split weighted shares of **193.428420 million** and calculated EPS of approximately **Rs59.8421**. The report's retrospectively adjusted EPS of Rs11.97 is preserved as source evidence; it is not paired with the raw June price. [Official PSX notice](https://www.psx.com.pk/psx/files/?file=257828-1.pdf).
- LUCK first traded on its five-for-one split basis on 28 April 2025. FY2021–24 therefore use Rs10 par, own-year pre-split weighted shares and DPS. FY2025 uses Rs2 par. Historical PSX prices remain unadjusted. [Official PSX notice](https://www.psx.com.pk/psx/files/?file=250915-1.pdf).
- MLCF FY2022 outstanding shares now exclude the 18.545 million treasury shares disclosed in note 6.3. The source discloses that quantity to 0.001 million; no unsupported extra precision was added. The remaining purchases and cancellation of the full 25 million occurred after year-end and were not backdated.
- FCCL's 12.5% bonus issue occurred after FY2022 and before FY2023's close. FY2022 uses its own-year, pre-bonus EPS denominator. FY2023 standardized EPS uses the full-year bonus-adjusted denominator, equal to closing shares, because the bonus was the only share issue in that year. This applies [IAS 33 paragraphs 26–28](https://www.ifrs.org/content/dam/ifrs/publications/pdf-standards/english/2023/issued/part-a/ias-33-earnings-per-share.pdf?bypass=on). The conflicting published count is retained and flagged. FCCL FY2024 uses the corrected comparative weighted shares printed in its FY2025 EPS note.
- All six market measures (BVPS, market-to-book, EPS, P/E, dividend yield and dividend payout) contain Excel formulas for all five companies and five years. Industry formulas and editable chart references remain linked.

## Annual-report differences retained

**23 of 25 price comparisons agree exactly or within the rounding introduced by the reports' five-for-one retrospective presentation.** Two KOHC price comparisons disagree beyond that rounding:

| FY | PSX actual close | Report price × confirmed split factor | Difference |
|---|---:|---:|---:|
| 2024 | 250.43 | 245.00 | 5.43 |
| 2025 | 379.36 | 378.35 | 1.01 |

The actual PSX closes govern. The reason for these report differences remains unresolved; the audit does not infer a cause or alter prices to force agreement.

Other source discrepancies remain visible:

- KOHC FY2025's comparative EPS note prints conflicting FY2024 weighted-share counts. The calculation uses the independently inspected FY2024 note's 196,435,067 pre-split shares.
- LUCK's FY2025 six-year summary prints FY2023 split-basis EPS of 8.49, inconsistent with its own-year audited EPS of 43.06 divided by five. The standardized calculation uses the audited own-year weighted shares.
- FCCL's published FY2023 weighted shares of 2,353.539 million cannot be reconstructed from its opening shares and sole 12.5% bonus issue. The workbook shows the standardized full-year bonus treatment separately from that source disclosure.
- Of the 50 existing published market-ratio comparisons rechecked, 40 reconstruct to displayed precision. Ten remain unresolved: five KOHC published market-to-book values and five FCCL published P/E values. All reconstructed and standardized market values now use formulas; source disagreements are not marked verified.

## Validation and scope

Independently recomputed all **150 company-year market results** in each workbook using Python and compared them with the saved calculated values. DGKC FY2023 P/E and dividend payout remain `n/m` because earnings are negative. All 150 market cells contain formulas; no spreadsheet errors were found in the saved files. Price and treasury-share input changes were tested in a disposable imported copy, and the dependent ratios responded correctly. Chart caches were refreshed from the recalculated cells, avoiding stale initial chart displays.

Both workbooks were reopened for structural inspection and rendered for visual checks. The submission retains its eight original worksheets and 14 editable charts. The audit retains its original sheets and 31 charts and adds four evidence schedules. No unrelated values or formulas changed. The originals were not overwritten.

Recalculation was verified with the bundled spreadsheet engine, including input-sensitivity tests. A separate Microsoft Excel desktop recalculation was not performed.

No ratios were removed in this focused revision. This work does **not** certify all financial-statement inputs, all 33 measures mentioned in the pasted attachment, FY2020 opening balances, taxation or unrelated published-ratio reconciliations.
