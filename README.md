# Fleet Fuel & Efficiency Analysis — Excel VBA

## Overview

An Excel-based fleet fuel and efficiency analysis project combining data cleaning, KPI analysis, dashboard development and VBA automation.

The project analyses fleet fuel consumption and vehicle efficiency for the period **April 2024 to March 2025**, using fleet fuel-consumption data and UK average diesel prices.
## Excel VBA Workbook

[Download the Excel VBA Fleet Fuel & Efficiency Analysis](Fleet_Fuel_Efficiency_Analysis.xlsm)

## Key Objectives

- Analyse fleet fuel consumption and distance travelled
- Identify vehicles with high estimated fuel costs
- Identify vehicles with lower MPG
- Calculate fleet-level efficiency KPIs
- Estimate fuel expenditure using average diesel prices
- Validate and flag questionable records
- Build an interactive Excel dashboard
- Automate dashboard refresh using VBA

## Key KPIs

| KPI | Result |
|---|---:|
| Total Valid Fuel | 2,035,068.58 litres |
| Total Valid Distance | 7,469,725 miles |
| Overall Fleet MPG | 16.69 |
| Estimated Fuel Cost | £2,989,922.76 |
| Average Diesel Price | £1.47/L |
| Total Records | 1,138 |
| Valid Records | 1,126 |
| Records Requiring Review | 12 |
| Data Quality Review Rate | 1.05% |
| Fleet Average MPG | 23.74 |
| Fuel Cost per Mile | £0.40 |
| Fuel Cost per 100 Miles | £40.03 |

## Dashboard

![Fleet Fuel & Efficiency Dashboard](screenshots%3Adashboard.png)

The dashboard includes:

- Fleet-level KPI summary
- Top 10 vehicles by estimated fuel cost
- Top 10 lowest-MPG vehicles
- Fuel cost versus distance analysis
- Data quality status
- Fleet efficiency status
- Automated dashboard refresh

## Data Preparation

The workbook contains separate sheets for:

- `Raw_Data` — original fleet fuel data
- `Clean_Data` — validation, cleaning and calculated metrics
- `Fuel_Price` — diesel price assumption used for estimated cost
- `Dashboard` — management dashboard and visual analysis
- `VBA_Automation` — VBA automation component

Data-quality checks identify records with:

- Negative distance
- Invalid MPG
- Suspiciously high MPG
- Records requiring review

Questionable records are excluded from fuel-efficiency and estimated-cost calculations.

## Fuel Cost Methodology

Estimated fuel cost is calculated using:

**Valid Fuel Consumption × Average Diesel Price**

The diesel price used is an average UK pump price for the project period.

Therefore, the fuel expenditure figure represents an **estimated fuel cost**, not actual supplier invoice expenditure.

## Excel & VBA Techniques

### Excel

- Data cleaning
- Conditional logic
- SUMIF / COUNTIF calculations
- KPI calculations
- Data validation
- Charts
- Trendline analysis
- Dashboard design
- Conditional formatting

### VBA

The workbook includes a `RefreshFleetDashboard` macro that:

- Refreshes workbook calculations
- Recalculates formulas
- Updates the dashboard
- Records the dashboard update timestamp
- Provides a confirmation message

A **Refresh Dashboard** button is available directly on the dashboard.

## Business Insights

The analysis allows users to investigate:

- Which vehicles generate the highest estimated fuel costs
- Which vehicles have lower MPG
- The relationship between distance travelled and estimated fuel cost
- Overall fleet fuel efficiency
- Data-quality issues requiring review

## Limitations

The estimated fuel cost uses an average diesel pump price rather than vehicle-level fuel transaction prices or supplier invoices.

The analysis should therefore be interpreted as an **estimated cost model**, rather than an accounting record of actual fuel expenditure.

## Tools

- Microsoft Excel
- Excel VBA
- Data cleaning and validation
- Dashboard development
- Data visualisation
- Business KPI analysis


## Source Data & References

### Fleet Fuel Consumption Data

**Source:** UK Government / National Data Library  
**Dataset:** Fleet Vehicles Fuel Consumption  
**Period:** April 2024 – March 2025

[View and download the original fleet fuel-consumption dataset](https://www.data.gov.uk/dataset/2766ab07-2714-43ea-926b-7cc93ef72aa5/fleet-vehicles-fuel-consumption)

### Diesel Price Data

**Source:** UK Government  
**Dataset:** Weekly Road Fuel Prices

[View the official UK Government weekly road fuel prices](https://www.gov.uk/government/statistics/weekly-road-fuel-prices)

The analysis uses the average UK diesel pump price for the project period. Fuel expenditure shown in the dashboard is therefore an **estimated fuel cost**, rather than actual supplier invoice expenditure.

## Project Structure

```text
fleet-fuel-efficiency-vba-analysis/
│
├── README.md
├── Fleet_Fuel_Efficiency_Analysis.xlsm
│
└── screenshots/
    └── dashboard.png
