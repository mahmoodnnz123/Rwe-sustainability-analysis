RWE Group — Sustainability Analysis
Carbon Footprint · Coal-to-Renewables · ESG KPIs · CSRD/ESRS Gap Analysis (FY 2025)

A comprehensive Python-based sustainability analysis of RWE Group using real audited data from the RWE Annual Report 2025. The project covers four distinct analytical sections — each generating publication-quality figures and a structured Excel export.

Data extracted directly from the RWE Annual Report 2025, Group Sustainability Statement (p. 72–160). Audited by KPMG AG (limited assurance, ISAE 3000 revised).
Why RWE?

RWE is one of Europe's largest electricity producers and is undergoing one of the most significant energy transitions in the sector — phasing out all coal by 2030 while targeting carbon neutrality by 2040. This makes it an ideal case study for combining GHG accounting, renewable energy analysis, EU Taxonomy reporting, and CSRD/ESRS compliance in a single project.
Project Structure

rwe-sustainability-analysis/
├── rwe_full_analysis.py        ← runs all 4 sections
├── requirements.txt
├── README.md
├── .gitignore
├── fig1_carbon_footprint.png
├── fig2_coal_to_renewables.png
├── fig3_esg_dashboard.png
└── fig4_csrd_gap.png

The Four Sections
Section 1 — Carbon Footprint & Decarbonization Pathway

What it does: Tracks RWE's Scope 1+2+3 emissions from 2022 to 2025 and models three possible trajectories to carbon neutrality by 2040 — a linear path, the current SBTi power sector path (−7%/yr), and the coal phaseout scenario.

Key data:
Metric 	2022 	2025 	Change
Scope 1 (Mt CO₂e) 	64.8 	51.9 	−20%
Carbon Intensity (tCO₂/MWh) 	0.55 	0.43 	−22%
Coal/Lignite Emissions (Mt) 	64.8 	37.3 	−42%
Scope 2 LB (Mt CO₂e) 	— 	0.22 	—
Scope 3 (Mt CO₂e) 	— 	19.5 	—

Target: Carbon neutral by 2040 (carbon intensity: 0.01 tCO₂/MWh)
Section 2 — Coal-to-Renewables Transition Dashboard

What it does: Visualizes RWE's shift from a coal-heavy to a renewables-dominated generation portfolio. Shows installed capacity by technology (offshore wind, onshore wind, solar, gas, coal) from 2021 through 2025 to the 2031 strategic target, alongside the coal emissions phaseout curve.

Key data:
Metric 	2021 	2025 	2031 Target
Total Capacity (GW) 	25 	48.9 	65
Renewables (GW) 	9.7 	21.3 	34+
Renewables Share (%) 	39% 	45% 	55%+
Coal/Lignite (GW) 	4.0 	7.5 	0 (phaseout)
Battery Storage (GW) 	0.3 	1.7 	4.0

Investment plan: €35 billion net investment 2026–2031 in renewables, storage, and hydrogen-capable gas.
Section 3 — Full ESG KPI Dashboard

What it does: Side-by-side comparison of 6 key ESG metrics across 2024 and 2025, covering environmental performance, safety, employee engagement and EU Taxonomy alignment. Each panel shows year-on-year change and direction vs. target.

Key data:
KPI 	2024 	2025 	Target 	Result
Scope 1 (Mt CO₂e) 	53.3 	51.9 	↓ 	✅
Carbon Intensity (tCO₂/MWh) 	0.46 	0.43 	↓ 	✅
Renewable Capacity (GW) 	19.5 	21.3 	↑ 	✅
LTIF (per M hrs) 	1.5 	1.2 	≤1.7 	✅
Employee Engagement (%) 	~80 	84 	≥80% 	✅
CapEx Taxonomy-Aligned (%) 	~90 	94 	95% by 2030 	✅
Section 4 — CSRD / ESRS Gap Analysis

What it does: Maps RWE's current disclosures against the 10 ESRS topical standards (E1–E5, S1–S4, G1) and scores each topic on a 0–10 compliance scale. Identifies remaining gaps and flags where RWE applied Quick Fix easements (EU Commission amendment reducing reporting volume by ~20%).

Key findings:
ESRS Topic 	Score 	Gap 	Note
E1 — Climate Change 	8.5/10 	1.5 	Strongest topic
E4 — Biodiversity 	5.0/10 	5.0 	⚡ Quick Fix applied
S2 — Value Chain Workers 	5.0/10 	5.0 	⚡ Quick Fix applied
S4 — Consumers & End-Users 	4.0/10 	6.0 	Weakest topic
Overall average 	6.0/10 		2nd year ESRS filing

RWE is filing ESRS for the second year (since FY 2024) — making this a more mature disclosure than most peers.
Data Source
Field 	Detail
Company 	RWE AG
Report 	RWE Annual Report 2025
Section 	Group Sustainability Statement, p. 72–160
GHG standard 	GHG Protocol Corporate Accounting Standard
CSRD standard 	ESRS Set 1 (2023)
EU Taxonomy 	Regulation (EU) 2020/852
Auditor 	KPMG AG — limited assurance (ISAE 3000 revised)
ESRS filing 	FY 2025 (2nd year mandatory ESRS)
How to Run

pip install pandas matplotlib openpyxl numpy

# Open in Spyder and press F5
# or from terminal:
python rwe_full_analysis.py

All 4 figures and the Excel workbook are saved to output/.
Methodology
Item 	Detail
GHG method 	GHG Protocol, location-based + market-based Scope 2
Carbon intensity 	Scope 1+2 per MWh electricity generated
SBTi path 	Sectoral Decarbonisation Approach — power sector (−7%/yr)
ESRS scoring 	0 = not disclosed, 5 = partial, 10 = full compliance
EU Taxonomy 	Revenue, CapEx alignment per Regulation (EU) 2020/852
Quick Fix 	EU Commission easements — reduced E4 and S2 reporting in 2025
