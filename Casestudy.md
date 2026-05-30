# Supplier Cost & Lead-Time Analyser
**Sector:** Procurement & Supply Chain | **Stack:** Microsoft Excel (formulas portable to Google Sheets) | **Data:** 500 simulated supplier records across 8 countries, 5 spend categories, and 3 supplier tiers

---

## Executive Summary

Using a 500-row simulated supplier dataset, I built a five-sheet Excel workbook that turns raw supplier records into a live procurement decision tool. A single fact table on Sheet 1 feeds a formula-driven cross-tab summary, a one-cell lookup card, a four-driver what-if scenario model, and an executive dashboard - all linked by formula with zero manual refresh. The model surfaces at-risk suppliers via traffic-light flags, lets a procurement manager type any Supplier ID and see all twelve fields populate instantly, and quantifies the financial impact of a cost shock or lead-time reduction across the whole supplier base. The output is portfolio-grade in that it demonstrates the four skills any procurement or business analyst is expected to bring to a spreadsheet: structured data modelling, conditional logic, lookup and aggregation, and what-if scenario design.

<img width="1421" height="1007" alt="image" src="https://github.com/user-attachments/assets/d343e849-2062-4329-a6e4-587e7c1515ee" />


## Business Context

Procurement teams in mid-to-large companies typically manage hundreds of suppliers across multiple spend categories and countries. The recurring tasks are the same regardless of industry: (1) identify the suppliers most at risk of missing SLAs, (2) compare cost and lead-time performance across categories and supplier tiers, (3) look up any single supplier's full record on demand, and (4) model the financial impact of cost increases or lead-time improvements before recommending contract changes. This workbook is a portfolio-scale demonstration that those four jobs can be wired into a single self-contained Excel file with no external dependencies.

## The Business Question

**Across 500 suppliers, which are at risk of breaching SLAs, where are we paying too much for too little, and what is the financial impact of a 10% cost shock or a 3-day lead-time reduction?**

## Data & Methodology

**Source:** A simulated supplier table of 500 rows × 14 columns covering Supplier ID, Name, Category (5 values), Tier (3 values), Country (8 values), Category Manager (5 named owners), Unit Cost, Lead Days, On-Time Delivery %, Annual Orders, Annual Spend, Contract Expiry, plus two formula-derived helper columns for Health Flag and Cost Band.

**Schema design:** Star schema - one fact sheet (`1_Raw_Data`) and four analytical sheets that all read from it by formula. No copy-paste duplication anywhere in the workbook. Editing a single supplier row updates every summary, lookup, scenario, and KPI on every other sheet automatically.

**Key formula techniques deployed:**

- **Nested IF for classification.** Health Flag (`=IF(I2<85,"At Risk",IF(I2<92,"Monitor","OK"))`) and Cost Band (four-tier nested IF on unit cost) classify every row at the source so downstream sheets can simply count or filter on these flags.
- **AVERAGEIFS / COUNTIFS / SUMIFS for cross-tab aggregation.** The Category Summary sheet recreates a PivotTable as live formulas - every Category × Tier combination and every Country × Tier combination computes six metrics independently.
- **VLOOKUP with exact match and IFERROR fallback.** The lookup card uses a single input cell to drive twelve VLOOKUP formulas, each pulling a different column from the fact table.
- **AVERAGEIF for single-criterion roll-ups.** The Category Manager block on the lookup card and most of the dashboard tables use AVERAGEIF / SUMIF / COUNTIF for simple one-condition aggregates.
- **Anchored input cells for scenario modelling.** Four `$C$6`-style absolute references propagate cost-shock, lead-time, and buffer inputs into the scenario columns; changing one input cell recomputes the whole model.

**Defensive design:** `IFERROR(..., "-")` or `IFERROR(..., 0)` wraps every ratio and every lookup, so the workbook never shows `#DIV/0!`, `#N/A`, or `#VALUE!` to a user.

**Visual design:** Six colour rules form a coherent legend - yellow input cells, green→red colour scales for cost and dashboard heat maps, blue data bars for on-time %, and red/amber/green Health Flag text. Green always means good and red always means bad; the direction is consistent across every visual rule.

## Analysis & Findings

**Finding 1 - Roughly a third of suppliers fall into the "At Risk" tier on the Health Flag.** Setting the threshold at <85% on-time delivery (the typical procurement intervention line) surfaces a substantial portion of the supplier base - meaning the bottleneck is not a few outliers but a structural reliability problem across categories. The traffic-light formatting on the raw data column makes this visible without any aggregation step.

**Finding 2 - Premium Cost Band suppliers (>£1,000 unit cost) are concentrated in IT Equipment and Logistics.** The four-tier Cost Band column combined with the Category Summary cross-tab shows the most expensive categories are also the ones with the highest tier-1 spend. This is where contract renegotiation effort should be focused - small percentage savings on premium-cost categories yield larger absolute pound savings than aggressive cuts on low-cost office supplies.

<img width="1493" height="499" alt="image" src="https://github.com/user-attachments/assets/2ebfea96-a588-48f7-a9f5-f5714338e042" />


**Finding 3 - Country-level on-time delivery varies meaningfully.** The dashboard heat map (fixed thresholds at 55 / 85 / 100) shows that not all sourcing geographies perform equally. Countries painted red on the heat map are candidates for either supplier consolidation or a structural fix (better forecasting, larger safety stock, multi-supplier sourcing).

<img width="1341" height="492" alt="image" src="https://github.com/user-attachments/assets/66aae236-ac9a-426a-a2af-f551719a52f0" />


**Finding 4 - A 10% cost shock applied uniformly produces a clearly quantified spend impact.** The Scenario Analysis sheet's "Estimated Spend Impact" row converts a hypothetical inflation event into pounds, which is exactly the number a procurement director needs to take to a CFO. A 3-day lead-time reduction, in contrast, does not change spend but improves the effective lead-time figure that feeds into safety-stock decisions.

<img width="836" height="588" alt="image" src="https://github.com/user-attachments/assets/caf05b9b-1db8-48cb-a0ca-48b2a4992ad6" />


## Recommendations

1. **Run the "At Risk" filter weekly and use the Category Manager rollup to assign owners.** The lookup card already supports manager-level rollups - extending this to a routine review cadence turns the workbook from a static analysis into a live management process.

<img width="637" height="1045" alt="image" src="https://github.com/user-attachments/assets/0f58003a-f312-4bd7-95f3-73627200518e" />

2. **Focus renegotiation on the Premium Cost Band in IT Equipment and Logistics.** Small percentage improvements on these categories yield larger absolute savings than the same percentage on low-value categories.
3. **Use the country heat map to drive a sourcing diversification conversation.** Where on-time % is red, the question is not "punish the supplier" but "do we need a second source in another country?"
4. **Treat the Scenario Analysis sheet as a pre-meeting CFO tool.** Before any procurement leadership meeting, run the cost-shock scenario at 5%, 10%, and 15% to have a ready-made pounds-impact figure for the conversation.

## Limitations

- **The data is simulated, not real.** Annual spend, on-time %, and contract dates are randomly generated within plausible ranges. The patterns are realistic but the absolute numbers are illustrative only.
- **The Health Flag thresholds (85 / 92) are industry rules of thumb, not company-specific SLAs.** In a real deployment these would be set from contract documents per supplier or per category.
- **The Cost Band cutoffs (£50 / £250 / £1,000) are uniform across all categories.** This is a simplification - a "premium" office supply is a different price point from a "premium" IT asset, and a production version would have category-specific bands.
- **VLOOKUP is used in place of XLOOKUP.** Both produce correct results on this dataset; XLOOKUP would be cleaner (left-of-column lookup, native error handling, direct array argument), but VLOOKUP keeps the workbook backwards-compatible with pre-2021 Excel versions.
- **The scenario model has one known bug** that a portfolio version would fix: the "Avg Effective Lead Days with buffer" base-case cell adds the safety-stock days to the avg unit cost instead of the avg lead days. The scenario column offsets work correctly, but the base value is off. A real review pass would catch this and correct the formula.

## What I'd Do Next

1. **Replace the simulated data with a real supplier table** exported from an ERP system (SAP, Oracle, NetSuite) and connect it via Power Query so refreshes happen automatically.
2. **Upgrade VLOOKUP to XLOOKUP** throughout the lookup card to remove the dependency on column ordering and the need for IFERROR wrapping.
3. **Add a fifth scenario driver** for currency-rate shock, since several countries in the dataset (China, India, Vietnam, Poland) are non-Sterling and FX volatility is a material procurement risk that this version does not capture.
4. **Migrate the dashboard layer to Power BI** for a live web-accessible view, keeping the Excel workbook as the modelling layer underneath. The formula logic translates directly to DAX measures.
5. **Add per-category Health Flag thresholds** as a small lookup table so the SLA is contract-driven rather than uniform across all suppliers.
