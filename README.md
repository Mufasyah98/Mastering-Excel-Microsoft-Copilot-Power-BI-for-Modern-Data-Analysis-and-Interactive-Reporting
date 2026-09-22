# Mastering Excel, Microsoft Copilot & Power BI for Modern Data Analysis and Interactive Reporting

A complete 2-day corporate training programme that teaches **Microsoft Excel**, **Microsoft Copilot** and **Microsoft Power BI** as one integrated workflow, from raw business data to interactive dashboards and business decisions.

```
Raw Data → Clean & Prepare → Analyse → Use AI / Copilot → Model → Visualise → Insight & Decision
```

The package includes a participant handbook, a slide deck with speaker notes, practice datasets, and a capstone case study based on **Maju Niaga Sdn Bhd**, a fictional Malaysian distributor of electronics, office furniture and office supplies.

---

## Contents

- [Who This Is For](#who-this-is-for)
- [Learning Outcomes](#learning-outcomes)
- [What's Included](#whats-included)
- [Repository Structure](#repository-structure)
- [Getting Started](#getting-started)
- [Programme Outline](#programme-outline)
- [The Datasets](#the-datasets)
- [Capstone Case Study](#capstone-case-study)
- [Using This Material](#using-this-material)
- [Software Requirements](#software-requirements)
- [About Copilot Features](#about-copilot-features)
- [Author](#author)
- [Licence](#licence)

---

## Who This Is For

Corporate executives, officers, analysts, finance and administrative staff, operations teams and managers who work with business data and reporting. The material suits **beginner to intermediate** users with mixed technical backgrounds. It is practical and business-oriented rather than academic.

## Learning Outcomes

By the end of the programme, participants will be able to:

- Clean and structure messy business data in Excel and Power Query
- Analyse data with essential, lookup and dynamic array functions and PivotTables
- Automate repetitive data preparation with Power Query
- Write effective prompts for Microsoft Copilot and verify its output
- Design a star schema data model and write core DAX measures in Power BI
- Build interactive dashboards that answer real business questions
- Turn analysis into clear insights, actions and decisions

## What's Included

| Material | Format | Description |
|---|---|---|
| Participant Handbook | `.docx` | Reference handbook covering all 19 modules, with syntax, worked examples, common mistakes, checklists, a prompt library and appendices |
| Training Slides | `.pptx` | 56 widescreen slides with speaker notes, hands-on exercise slides and native PowerPoint charts |
| Day 1 Practice Dataset | `.xlsx` | Clean practice workbook with named Excel Tables for the Day 1 exercises |
| Case Study Data | `.csv`, `.xlsx` | 18 monthly sales CSV files with built-in data quality problems, plus master data and 2026 targets |
| Case Study Participant Brief | `.docx` | Scenario, data dictionary, business rules, 25 Day 1 exercises, 9 capstone tasks and an assessment rubric |
| Trainer Answer Key | `.docx` | Expected results, a data cleaning reconciliation and facilitation notes (**trainers only**; see [below](#a-note-on-the-answer-key)) |

## Repository Structure

```
.
├── README.md
├── handbook/
│   └── Participant_Handbook_Excel_Copilot_PowerBI.docx
├── slides/
│   └── Excel_Copilot_PowerBI_Training_Slides.pptx
├── case-study/
│   └── MajuNiaga_Case_Study_Participant_Brief.docx
└── data/
    ├── Day1_Excel_Practice_Dataset.xlsx
    └── CaseStudy_Data/
        ├── Monthly_Sales/
        │   ├── Sales_2025_01.csv
        │   ├── ...
        │   └── Sales_2026_06.csv
        ├── MajuNiaga_MasterData.xlsx
        └── Sales_Targets_2026.xlsx
```

## Getting Started

1. **Download the repository.** Click **Code → Download ZIP**, or clone it:
   ```bash
   git clone https://github.com/<your-username>/<repository-name>.git
   ```
2. **Open the handbook** in `handbook/` and read *How to Use This Handbook*.
3. **Open the Day 1 workbook** in `data/` to follow the Day 1 exercises.
4. **For the capstone,** keep the `Monthly_Sales` folder intact. Power BI connects to the whole folder, and dropping a new month's file into it tests the automatic refresh.

> **Tip:** Keep a copy of the original data files. Some exercises use destructive tools such as Remove Duplicates.

## Programme Outline

### Day 1: Excel & Microsoft Copilot for Data Analysis

| # | Module | Key Topics |
|---|---|---|
| 1 | Modern Data Analysis in the AI Era | Four types of analytics, analysis workflow, data quality |
| 2 | Preparing Business Data in Excel | Excel Tables, data cleaning tools, Data Validation, cleaning checklist |
| 3 | Essential Excel Functions | SUMIFS, COUNTIFS, IF, IFS, text, date and rounding functions |
| 4 | Modern Lookup Functions | XLOOKUP, XMATCH, INDEX + MATCH, troubleshooting #N/A |
| 5 | Dynamic Array Functions | FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, #SPILL! |
| 6 | PivotTables | Grouping, Show Values As, Pivot Charts, Slicers, Timelines |
| 7 | Introduction to Power Query | Extract-Transform-Load, folder combine, Merge, Append, Unpivot |
| 8 | Microsoft Copilot for Data Analysis | Copilot in Excel, capabilities, limitations, responsible use |
| 9 | Prompt Engineering | ROLE–CONTEXT–TASK–DATA–OUTPUT–CONSTRAINTS framework, prompt library |
| 10 | Data Storytelling with Excel | Chart selection, conditional formatting, sparklines, KPI indicators |

### Day 2: Power BI for Interactive Reporting

| # | Module | Key Topics |
|---|---|---|
| 11 | Introduction to Power BI | Desktop vs Service, workflow, Excel vs Power BI |
| 12 | Connecting Data to Power BI | Common sources, Import vs DirectQuery |
| 13 | Data Transformation in Power Query | Standard cleaning workflow, custom and conditional columns |
| 14 | Data Modelling Fundamentals | Fact and dimension tables, star schema, relationships, date table |
| 15 | Introduction to DAX | Columns vs measures, CALCULATE, filter context, time intelligence |
| 16 | Visualisation & Report Design | Visual selection, layout, interactivity, design principles |
| 17 | Copilot in Power BI | Report generation, summaries, DAX assistance, prerequisites |
| 18 | Publishing, Sharing & Refresh | Workspaces, apps, scheduled refresh, row-level security |
| 19 | Capstone Case Study | End-to-end project applying the full workflow |

## The Datasets

All data is **fictional** and generated for training purposes. Any resemblance to real organisations or people is coincidental. Currency is Malaysian Ringgit (RM).

### Day 1 Practice Dataset (`Day1_Excel_Practice_Dataset.xlsx`)

| Sheet | Excel Table | Rows | Used In |
|---|---|---|---|
| Sales | `tblSales` | 5,511 | Modules 3, 5, 6, 8–10 |
| Products | `tblProducts` | 24 | Module 4 |
| Customers | `tblCustomers` | 80 | Modules 4, 5 |
| Staff | `tblStaff` | 15 | Modules 3, 4, 5 |
| Employees | `tblEmployees` | 128 | Modules 3, 6 |
| Expenses | `tblExpenses` | 756 | Modules 8, 9 |
| Budget_Grid | `tblBudgetGrid` | 7 | Module 7 (Unpivot) |
| CommissionTiers | `tblTier` | 5 | Module 4 (approximate match) |
| Regions | `tblRegion` | 5 | Module 2 (Data Validation) |
| Messy_Data | *(plain range)* | 37 | Module 2 (cleaning exercise) |

Table and column names match the formulas in the handbook and slides, so examples work as written. For example, `=SUMIFS(tblSales[Amount], tblSales[Region], "Central")`.

### Case Study Data (`CaseStudy_Data/`)

| File | Contents |
|---|---|
| `Monthly_Sales/Sales_YYYY_MM.csv` | 18 monthly ERP exports (Jan 2025 – Jun 2026), about 5,500 transaction lines in total |
| `MajuNiaga_MasterData.xlsx` | Products, Customers, Salespeople and Regions tables (the dimension tables) |
| `Sales_Targets_2026.xlsx` | Monthly targets by region, in a report-style layout for unpivot practice |

The CSV files **intentionally contain data quality problems** for participants to fix in Power Query: duplicate lines, test transactions, cancelled orders, invalid product codes, blank quantities, inconsistent region spellings and extra spaces.

## Capstone Case Study

Maju Niaga's Finance team spends three days each month combining branch CSV files into a static report. The Managing Director wants a report that refreshes itself and explains what is driving first-half 2026 performance.

Participants answer six business questions:

1. How did H1 2026 sales compare with H1 2025, in total and by region?
2. How is each region performing against its 2026 target?
3. What is driving the change in the weakest-performing region?
4. Which products, categories and customer segments contribute most to sales and margin?
5. How is each sales representative performing against target?
6. Where is the company growing, and what should management do next?

The dataset has deliberate patterns built in, so thorough analysis leads to a clear root cause and actionable recommendations. The full scenario, business rules, task steps and rubric are in the **Case Study Participant Brief**.

## Using This Material

### For Trainers

- The slides carry speaker notes with timing and demonstration cues for every slide.
- Hands-on exercise slides point to the matching exercises in the Case Study Brief.
- Suggested pacing: Modules 1–10 on Day 1, and Modules 11–18 plus the capstone on Day 2.
- Use the Trainer Answer Key for debriefs. It includes a table that maps common wrong totals to the specific business rule a participant missed.

### For Participants and Self-Paced Learners

- Work through the handbook module by module, then complete the matching Part A exercises in the Case Study Brief.
- Attempt the capstone after finishing Module 18.
- Check your totals against a PivotTable of your cleaned data before trusting any visual.

### A Note on the Answer Key

The Trainer Answer Key is **not included in this repository** so that exercise and capstone answers are not public. Trainers who need it can contact the author (see [Author](#author)).

## Software Requirements

| Software | Version | Notes |
|---|---|---|
| Microsoft Excel | Microsoft 365 or Excel 2021+ | Needed for XLOOKUP, FILTER, SORT, UNIQUE and other dynamic arrays. Excel 2016/2019 users can use INDEX + MATCH instead |
| Power BI Desktop | Latest version | Free download from Microsoft; Windows only |
| Power BI Service | Pro or Premium Per User | Optional; needed only for publishing and sharing (Module 18) |
| Microsoft Copilot | Organisation-licensed | Optional; see [About Copilot Features](#about-copilot-features) |

## About Copilot Features

Copilot capabilities change frequently. What participants see may differ from the handbook or from a colleague's screen. Availability depends on:

- Microsoft 365 subscription and Copilot licence
- Organisation configuration and tenant policies
- Application version and update channel

**Copilot in Power BI** generally requires a workspace on a paid Microsoft Fabric capacity (F2 or higher) or Power BI Premium capacity, with Copilot enabled by the tenant administrator. Confirm current requirements with your IT or Power BI administrator.

Every Copilot exercise can also be completed manually, so participants without Copilot access can still finish the programme.

## Author

**Fakhrul Syahmi**
Data Consultant & HRD Corp Certified Trainer
Mufasyah Consultant

- Master of Data Science, Universiti Kebangsaan Malaysia (UKM)
- Microsoft Certified: Azure AI Fundamentals (AI-900), Azure Data Fundamentals (DP-900), Power Platform Fundamentals (PL-900)

📧 mufasyahcons@gmail.com

For corporate training, customised programmes or consulting on Excel, Power BI, Power Platform and AI productivity, please get in touch by email.

## Licence

© 2026 Fakhrul Syahmi. All rights reserved.

This material is made available for viewing and personal learning. Redistribution, resale, or use in commercial training delivery requires written permission from the author.

Microsoft, Excel, Power BI, Copilot and Microsoft 365 are trademarks of the Microsoft group of companies. This material is independent and is not affiliated with or endorsed by Microsoft.
