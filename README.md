# Agentic AI / GenAI Revenue Impact Dashboard

An interactive Streamlit dashboard and Excel-based scenario model to assess how Agentic AI / GenAI could impact revenue growth, revenue mix, monetization upside, and cannibalization risk across selected US-listed software companies.

## Live App

**Streamlit App:** https://aiimpact-analysis-companies.streamlit.app/

## What this project covers

This project evaluates the potential business impact of GenAI / Agentic AI for a selected set of software companies:

- Microsoft
- Salesforce
- Adobe
- ServiceNow
- Intuit
- Palantir

The solution combines:

- an **Excel model** for scenario-based revenue analysis
- a **Streamlit dashboard** for interactive visualization
- a **one-page brief** for executive summary
- a **prompt / workflow log** for traceability
- a **source register** for reference tracking

## Files in this repository

- `app.py` — Streamlit application
- `Ai_Impact_Analysis.xlsx` — Excel model with all tabs and calculations
- `Ai_Impact_Analysis_Brief.docx` — one-page summary brief
- `requirements.txt` — Python dependencies needed to run the app

## How to use the solution

### 1. Open the live app
Use the Streamlit app link above to view the interactive dashboard.

### 2. Review the Excel model
Open `Ai_Impact_Analysis.xlsx` to understand the assumptions, calculations, scenario model, revenue mix logic, sources, and workflow log.

### 3. Read the brief
Open `Ai_Impact_Analysis_Brief.docx` for a concise business summary of the model, key findings, assumptions, and suggested next steps.

## Excel workbook structure

The Excel workbook is organized tab by tab so the model can be understood from inputs to outputs.

### 1. Overview
This is the guide sheet for the workbook.

It explains:
- the purpose of the model
- how to use the workbook
- the color logic for inputs, formulas, and assumptions
- the overall modeling flow
- the core formula:
  - **Net Revenue = Baseline Revenue + AI Uplift − Cannibalization**

### 2. Company_Selection
This tab defines the companies included in the analysis.

It captures:
- company name
- ticker
- category / rationale for selection
- main AI monetization angle
- priority / inclusion flag

This is the scope-definition tab.

### 3. Raw_Data
This is the factual input layer of the workbook.

It contains:
- company name and ticker
- fiscal year used
- revenue and operating income
- existing AI revenue share proxy for 2025 where available
- business segment names and segment revenues
- major AI products
- AI monetization model
- recent AI traction / proxy statements
- primary and secondary source URLs
- notes and caveats

This tab answers: **What do we know today from reported data and public company disclosures?**

### 4. Assumptions
This is the forward-looking assumptions layer.

It contains company-wise Low / Base / High scenario assumptions for:
- baseline growth 2026 and 2027
- AI-exposed revenue share
- AI adoption 2026 and 2027
- monetization yield 2026 and 2027
- cannibalization 2026 and 2027
- assumption notes

This tab answers: **How are we modeling the future impact of AI?**

### 5. Scenario_Model
This is the calculation engine of the workbook.

It combines Raw_Data + Assumptions and calculates:
- Revenue 2025
- Revenue 2026 baseline
- AI base 2026
- AI uplift 2026
- Cannibalization 2026
- Net revenue 2026
- Revenue 2027 baseline
- AI base 2027
- AI uplift 2027
- Cannibalization 2027
- Net revenue 2027

This tab answers: **What is the modeled revenue outcome under the selected scenario?**

### 6. Revenue_Mix
This tab shows how the revenue composition changes over time.

It contains:
- AI share 2025
- Legacy/Core share 2025
- AI share 2026
- Legacy/Core share 2026
- AI share 2027
- Legacy/Core share 2027
- shift in AI share from 2025 to 2027
- interpretation notes

This tab answers: **How much of the business becomes AI-driven over time?**

### 7. Charts
This is the helper / staging tab for visuals.

It organizes chart-ready data for:
- 2025 vs 2027 revenue comparison
- Low / Base / High scenario comparison
- AI uplift vs cannibalization comparison
- AI share / revenue mix comparison

This tab supports both Excel visuals and the Streamlit dashboard views.

### 8. Sources
This is the source register.

It tracks:
- company
- data point / statement
- value or quote
- source type
- URL
- page / note
- verification status

This tab answers: **What evidence supports the model inputs?**

### 9. Prompt_Log
This tab documents the AI-assisted workflow used during the project.

It includes:
- step number
- prompt / task
- tool used
- output summary
- manual validation status
- used in final or not

This tab answers: **How was AI used in building the solution?**

## Dashboard structure

The Streamlit app is organized into three main dashboard sections.

### 1. Portfolio Overview
This is the overall business view.

It shows:
- portfolio revenue in 2025
- portfolio revenue in 2027 under the selected scenario
- total AI uplift in 2027
- winner highlight
- high-risk highlight
- company ranking cards
- 2025 vs 2027 revenue comparison chart
- AI uplift vs cannibalization chart
- ranking table
- AI revenue mix shift chart

How to read it:
- start with the KPI cards at the top
- identify the portfolio growth and AI uplift
- review the winner and risk highlights
- compare company growth and downside in the charts and ranking table
- use the mix shift chart to see which companies become more AI-led

### 2. Company Deep Dive
This is the detailed single-company view.

It shows:
- 2025 revenue
- 2026 net revenue
- 2027 net revenue
- business snapshot
- major AI products
- monetization approach
- latest segment revenue chart
- AI impact profile
- scenario comparison for the selected company
- revenue mix visuals
- interpretation text

How to read it:
- select a company from the sidebar
- observe its revenue progression from 2025 to 2027
- review the business snapshot and AI monetization logic
- compare AI uplift vs cannibalization
- check how the company’s revenue mix shifts toward AI

### 3. Sources
This tab lets the user inspect the source/reference layer.

It can be viewed for:
- the selected company only, or
- all companies

How to read it:
- use this section when validating assumptions, proxies, and source references
- it is especially useful for auditability and transparency

## How to interpret the dashboard

A good reading sequence is:

1. **Start with Portfolio Overview**
   - Understand the portfolio-level growth and AI effect
2. **Check the winner and risk highlights**
   - See who has the highest modeled upside and who has the highest relative downside pressure
3. **Review the ranking table**
   - Compare companies across growth, AI uplift, cannibalization, and risk score
4. **Study revenue mix shift**
   - Understand which companies become more AI-led by 2027
5. **Move to Company Deep Dive**
   - Analyze one company at a time in detail
6. **Use Sources for validation**
   - Review source references and supporting notes

## Scenario logic

The model supports three scenarios:

- **Low** — conservative adoption and monetization case
- **Base** — central / most realistic case
- **High** — optimistic adoption and monetization case

These scenarios are selected in the dashboard sidebar and are driven by the values maintained in the `Assumptions` tab.

## Methodology summary

The model follows a structured logic:

1. Take reported / sourced baseline revenue from `Raw_Data`
2. Apply baseline growth assumptions from `Assumptions`
3. Estimate the AI-exposed revenue pool
4. Apply AI adoption rates
5. Apply monetization yield
6. Subtract cannibalization impact
7. Calculate net revenue for 2026 and 2027
8. Derive AI share and legacy/core share in `Revenue_Mix`

Core formula:

**Net Revenue = Baseline Revenue + AI Uplift − Cannibalization**

## How to run locally

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the app:

```bash
streamlit run app.py
```

Then open the local Streamlit URL shown in your terminal.

## Notes

- This is a **scenario-based analytical model**, not company guidance or investment advice.
- Some AI revenue values are **proxy-based** where direct disclosures were unavailable.
- Results are sensitive to assumptions around adoption, monetization, and cannibalization.

## Deliverables summary

This repository collectively provides:
- Excel solution output
- 1-page business brief
- Streamlit dashboard
- source/reference list
- prompts / AI workflow log

---

For quick review, use the live app first and then refer to the Excel workbook for full calculation logic and source traceability.
