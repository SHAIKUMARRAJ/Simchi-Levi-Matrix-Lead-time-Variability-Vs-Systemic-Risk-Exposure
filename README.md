# Integrated Procurement Risk Analytics Suite (David Simchi-Levi Framework)

An advanced, data-driven procurement risk architecture implementing **MIT Professor David Simchi-Levi's** core philosophy: *treating the supply chain as a single, integrated system of choices*. 

This project moves beyond backward-looking spend analytics to model structural operational trade-offs, mapping delivery unpredictability and financial liabilities across 10,000 corporate procurement transactions.

## Strategic Frameworks Implemented

### 1. Lead-Time Variability Matrix ($\sigma$)
Average lead times are inherently deceptive. This module isolates the standard deviation ($\sigma$) of supplier fulfillment timelines to capture delivery unpredictability. High variability forces downstream operations to accumulate expensive, defensive safety stock buffers.

### 2. Financial Risk Exposure Index (REI) & Time-to-Recovery (TTR)
Rather than relying on static financial health scores, this suite maps actual transaction velocity against simulated **Time-to-Recovery (TTR)** timelines (the duration required to find alternative capacity or clear a bottleneck during a catastrophic disruption). 
$$\text{REI} = \text{Daily Spend Velocity} \times \text{TTR (Days)}$$
This explicitly calculates the total dollar exposure the enterprise faces if a specific node goes completely offline.

### 3. Prescriptive "Chained Flexibility" (The Power of 2)
To mitigate multi-million dollar vulnerabilities without the massive cost of full supply network customization, this suite applies Simchi-Levi's **Power of 2** theorem. By structuring contracts so strategic suppliers can pivot across just *two* key asset categories, we build an interconnected resilient chain that eliminates single points of failure.

---

## Repository Architecture

*   `IT_Procurement_Project.xlsx`: The baseline enterprise database (Purchase Orders, Suppliers, Items).
*   `simchi_levi_analysis.py`: Production-grade Python engine calculating pipeline velocities, $\sigma$ metrics, and REI tiers.
*   `Simchi_Levi_Risk_Report_With_Chart.xlsx`: Final executive spreadsheet report featuring data formatting, conditional risk highlights, and an **embedded, interactive native Excel scatter chart**.
*   `simchi_levi_matrix_chart.png`: High-resolution visualization plot mapping supplier chaos vs. total financial liability.

---

## Key Analytical Insights
