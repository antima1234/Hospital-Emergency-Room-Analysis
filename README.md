# 🏥 Hospital Emergency Room Analysis — Excel Data Analytics Project

An end-to-end Excel analytics project simulating one year (FY 2024) of Emergency Room operations
across a 10-department, multi-city hospital network — built as a Data Analyst portfolio piece.

![Excel](https://img.shields.io/badge/Tool-Microsoft%20Excel-217346?logo=microsoft-excel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)

---

## 📌 Project Overview

Hospitals generate huge volumes of ER intake data every day, but raw data rarely tells a clean
story. This project takes a realistic, intentionally-messy ER intake export, cleans it,
engineers analytical fields with native Excel formulas, summarizes it into pivot-style reports,
and presents it in a single-screen executive dashboard — the kind of deliverable a hospital
operations or analytics team would actually use to track patient flow, department load, revenue,
and care quality.

**Business questions answered:**
- Which departments carry the most patient volume and revenue?
- How does severity level affect waiting time, treatment time, and billing?
- What is the network-wide recovery rate, and how does it vary by severity/department?
- Where are the operational risk points (critical patients without a bed)?
- How satisfied are patients, and what drives satisfaction?

---

## 🗂️ Dataset Information

| | |
|---|---|
| **Raw records** | 5,120 (includes intentional duplicates & data-quality issues) |
| **Cleaned records** | 5,000 unique patients |
| **Time period** | Jan 2024 – Dec 2024 |
| **Departments** | 10 (Cardiology, Neurology, Orthopedics, Pediatrics, General Medicine, Emergency Trauma, Gynecology, ENT, Oncology, Urology) |
| **Doctors** | 30 (3 per department) |
| **Cities** | 10 |

**Columns:** Patient ID, Patient Name, Age, Gender, Admission Date, Admission Time, Discharge
Date, Department, Doctor Name, Diagnosis, Severity Level, Waiting Time (Minutes), Treatment Time
(Minutes), Bed Assigned, Insurance Type, Payment Method, Total Bill, Patient Outcome, City,
Patient Satisfaction Score — plus 14 formula-driven calculated columns (see below).

Data is synthetic (generated with seeded Python for reproducibility) but modeled on realistic
ER distributions: waiting/treatment time scale with severity, bed-assignment probability rises
with severity, and outcome mix shifts toward Admitted/Referred/Deceased as severity increases.

---

## 🧹 Data Cleaning

Full step-by-step log with issue, action, rows affected, and rationale lives in the
**`Cleaning_Log`** sheet. Summary of what was fixed going from `Raw_Data` → `Cleaned_Data`:

1. Removed duplicate patient records
2. Handled missing/invalid Age values (negative or impossible ages)
3. Handled missing Doctor Name
4. Standardized inconsistent Gender text (`male`, `MALE`, `M` → `Male`)
5. Trimmed whitespace and standardized casing in Patient Name / City
6. Handled missing Total Bill and missing Patient Satisfaction Score (with different strategies — see log)
7. Converted blank-string Insurance Type to an explicit `None` category
8. Standardized date formats for reliable `YEAR()` / `MONTH()` / `WEEKDAY()` use

---

## 🧮 Excel Skills Used

| Category | Functions / Features |
|---|---|
| **Logic** | `IF`, `IFS`, `IFERROR` |
| **Lookup** | `INDEX` + `MATCH` (department budget/head, doctor → employee ID) |
| **Aggregation** | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `COUNTIF`, `SUMIF` |
| **Date/Text** | `YEAR`, `MONTH`, `WEEKDAY`, `TEXT`, `ROUND` |
| **Data quality** | Remove Duplicates, `TRIM`, `PROPER`, Data Validation dropdowns |
| **Structuring** | Excel Tables (structured ranges), frozen panes, named calculated columns |
| **Visualization** | Line, Bar, Column, Pie/Histogram charts; conditional formatting (color scale + icon set + rule-based highlighting) |
| **Formatting** | Custom number formats, KPI cards, cell comments documenting formula logic |

> **Note on `XLOOKUP`:** modern Excel's `XLOOKUP` is the intended production formula for
> single-value lookups, and is demonstrated conceptually in `Lookup_Tables`; the workbook's live
> lookups use `INDEX`/`MATCH` instead, which is functionally equivalent and guarantees the file
> recalculates identically across all Excel/LibreOffice versions.

**14 calculated columns in `Cleaned_Data`:** Admission Year, Admission Month, Admission Month
No., Weekday, Age Group, Length of Stay (Days), Total ER Time (Minutes), Bill per ER Minute,
Critical Alert, Department Budget, Department Head, Doctor Employee ID, Outcome Score,
Satisfaction Category.

---

## 📊 Pivot-Style Summary Tables (`Pivot_Summary`)

Built with `SUMIFS` / `COUNTIFS` / `AVERAGEIFS` so they recalculate live and need no manual
"Refresh":

1. Patients by Department
2. Patients by Severity
3. Average Waiting Time by Department
4. Monthly Patient Trend
5. Doctor-wise Patient Count
6. Gender Distribution
7. Age Group Analysis
8. Revenue by Department
9. Patient Outcomes
10. Satisfaction Score Analysis
11. Age Distribution (10-year bins, feeds the dashboard histogram)

> **Want native Excel PivotTables + Slicers?** `Cleaned_Data` is already formatted as an Excel
> Table, so in Excel you can select any cell in it → **Insert → PivotTable**, drop in the fields
> above, then **Insert → Slicer** for Department / Gender / Month / Severity / City. That's a
> 2-click Excel-UI step once the file is open — the formula-based summaries here give you the
> same numbers immediately and update automatically without a manual "Refresh."

---

## 📈 Dashboard Features (`Dashboard`)

- **KPI cards:** Total Patients · Total Revenue · Avg Waiting Time · Avg Treatment Time ·
  Critical Cases · Recovery Rate · Avg Patient Satisfaction
- **Monthly Trend Line Chart**
- **Department-wise Bar Chart**
- **Severity Pie Chart** (with % labels)
- **Revenue Column Chart**
- **Age Distribution Histogram**
- **Quick Filter Panel** — dropdown selectors for Department, Gender, Month, Severity, City
  (data-validation driven reference panel; pair with a native PivotTable/Slicer for live
  cross-filtering, as noted above)
- Blue-and-white healthcare color theme, fit-to-one-page layout, conditional formatting
  (severity color-coding, satisfaction icon sets, red "URGENT – No Bed" alert highlighting)

---

## 💡 Key Insights (`Insights`)

1. Emergency Trauma and Cardiology carry the highest Critical-severity share — prioritize
   capacity planning there.
2. Waiting time is inversely related to severity, confirming triage is working as intended.
3. Critical cases are under 10% of volume but drive a disproportionate share of revenue and
   treatment time.
4. The "Critical Alert" flag (Critical + no bed) surfaces a small, operationally urgent set of
   at-risk cases.
5. Network-wide recovery rate sits in the high-70%/low-80% range; Referred/Admitted outcomes
   rise with severity as expected.
6. Patient volume shows mild seasonality rather than a flat distribution across the year.
7. Government + Private insurance cover roughly two-thirds of patients; self-pay patients skew
   toward lower bills.
8. Satisfaction correlates with Outcome and drops with longer Waiting Time — a strong lever for
   perceived quality.
9. Bill per ER Minute is highest for Critical cases — acuity, not just time, drives cost.
10. A small cluster of doctors carry a visibly higher patient load than department peers — worth
    reviewing for staffing balance.

*(Full narrative version in the `Insights` sheet.)*

---

## 📁 Files Included

```
hospital-er-analysis/
├── data/
│   ├── raw_data_export.csv          # optional: raw export if you want a plain CSV alongside the workbook
│   └── cleaned_data_export.csv
├── excel/
│   └── Hospital_ER_Analysis.xlsx    # the full workbook (Read Me, Raw_Data, Cleaning_Log,
│                                     #   Cleaned_Data, Lookup_Tables, Pivot_Summary, Dashboard, Insights)
├── images/
│   └── dashboard_preview.png        # screenshot of the Dashboard sheet for the README
├── README.md
└── LICENSE
```

---

## 📁 Suggested Repository Folder Structure

```
hospital-er-analysis/
│
├── README.md
├── LICENSE
│
├── excel/
│   └── Hospital_ER_Analysis.xlsx
│
├── data/
│   ├── raw_data_export.csv
│   └── cleaned_data_export.csv
│
├── images/
│   ├── dashboard_preview.png
│   ├── cleaning_log_preview.png
│   └── pivot_summary_preview.png
│
└── docs/
    └── data_dictionary.md
```

---

## 🛠️ How to Use

1. Download `Hospital_ER_Analysis.xlsx` and open it in Excel (2016+ recommended).
2. Start at **Read Me** for a sheet-by-sheet guide.
3. Explore **Dashboard** for the executive summary.
4. Dig into **Cleaned_Data**, **Pivot_Summary**, and **Cleaning_Log** for the underlying work.
5. Everything recalculates automatically — try editing a row in `Cleaned_Data` and watch the
   Dashboard KPIs and charts update.

---

## 👤 Author & Contact 
Name:     Antima Pandey
Email:    antimapandey600@gmail.com
Linkdin:  https://github.com/antima1234
Github:   https://github.com/antima1234

Built as a portfolio project demonstrating end-to-end Excel data analysis: data generation/
cleaning documentation, formula engineering, pivot-style reporting, dashboard design, and
business storytelling.
