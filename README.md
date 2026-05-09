# 📊 Excel Salary Dashboard — Data Jobs (2023)

[![Live Demo](https://img.shields.io/badge/Live%20Demo-View%20Here-4f9cf7?style=for-the-badge)](https://KrishnaSai315.github.io/Data-Science-Salary-Dashboard/excel_salary_dashboard_demo.html)
[![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)]()
[![Dataset](https://img.shields.io/badge/Dataset-32K%20Records-orange?style=for-the-badge)]()

> An interactive Excel dashboard for exploring real-world data science job salaries across job titles, countries, and schedule types.

![Dashboard Preview](0_Resources/Images/1_Salary_Dashboard_Final_Dashboard.gif)

---

## 🧭 Table of Contents

- [Overview](#overview)
- [Live Demo](#live-demo)
- [Dashboard Preview](#dashboard-preview)
- [Dataset](#dataset)
- [Excel Skills Used](#excel-skills-used)
- [Dashboard Features](#dashboard-features)
- [Key Formulas Explained](#key-formulas-explained)
- [How to Use the Dashboard](#how-to-use-the-dashboard)
- [Key Insights](#key-insights)
- [File Structure](#file-structure)
- [About](#about)

---

## Overview

This **Salary Dashboard** was built to help data professionals and job seekers:

- 🔍 Investigate median salaries for specific data roles
- 🌍 Compare compensation across countries and schedule types
- 💡 Quickly identify whether a job offer reflects market rates

The dataset contains **real-world data science job listings from 2023**, covering job titles, salary figures, locations, and required skills — all presented through an interactive, no-code Excel interface.

**Data scope:** 2023 · Global · ~32,000 job listings across 10 major data roles

---

## Live Demo

👉 **[View Interactive Demo](./excel_salary_dashboard_demo.html)** — open in any browser, no Excel required.

---

## Dashboard Preview

### 🖥️ Full Dashboard — All Filters Applied

![Full Dashboard](0_Resources/Images/1_Salary_Dashboard.png)

> *Dashboard shown with Job Title: Data Analyst · Country: United States · Type: Full-time — Median Salary: $90,000 · Job Count: 5,849*

---

## Dataset

| Field | Description |
|-------|-------------|
| `job_title_short` | Standardized role (e.g., Data Analyst, Data Scientist) |
| `salary_year_avg` | Average annual salary (USD) |
| `salary_hour_avg` | Average hourly rate |
| `job_country` | Country where the role is based |
| `job_schedule_type` | Full-time, Part-time, Contractor, etc. |
| `job_skills` | Required technical skills |
| `company_name` | Hiring company |
| `job_work_from_home` | Whether remote work is available |

---

## Excel Skills Used

| Skill | Purpose |
|-------|---------|
| **Bar Chart** | Compare median salaries across all job titles |
| **Map Chart** | Visualize geographic salary distribution |
| **MEDIAN + IF (Array Formula)** | Compute filtered median salary |
| **FILTER Function** | Generate a clean unique list of schedule types |
| **Data Validation** | Restrict dropdowns to valid, filtered options |
| **Named Ranges** | Power the dynamic filtering logic |

---

## Dashboard Features

### 1. 📊 Salary Bar Chart — Job Titles Ranked by Median Salary

![Salary Bar Chart](0_Resources/Images/1_Salary_Dashboard_Chart1.png)

A horizontal bar chart ranks all job titles by descending **median salary**. The selected job title (Data Analyst, shown in dark blue) is highlighted against all other roles, making direct salary comparison effortless.

> **Key finding:** Senior roles and Engineers command significantly higher salaries than Analyst-level roles — Senior Data Scientist leads at ~$155K vs Data Analyst at $90K.

---

### 2. 🗺️ Country Median Salary Map

![Country Map](0_Resources/Images/1_Salary_Dashboard_Country_Map.gif)

An Excel map chart plots median salary by country using color intensity. Darker blue = higher salary concentration. Countries with no data remain gray.

> **Key finding:** The United States, Australia, and parts of Western Europe show the highest salary concentration. Asia and South America trail significantly.

---

### 3. 🎛️ Three Interactive Filters — Data Validation in Action

![Data Validation](0_Resources/Images/1_Salary_Dashboard_Data_Validation.gif)

Users can slice salary data across **three dimensions simultaneously**:

| Filter | Options |
|--------|---------|
| **Job Title** | 10 standardized roles (Data Analyst → Senior Data Scientist) |
| **Country** | 50+ countries with sufficient data |
| **Schedule Type** | Full-time, Part-time, Contractor, Temp work, Internship |

All three filters drive a single **Median Salary** result card and **Job Count** that update instantly — no macros, no VBA. Pure Excel formulas.

#### Job Title Filter Panel

<img src="0_Resources/Images/1_Salary_Dashboard_Job_Title.png" width="400" alt="Job Title Dropdown">

#### Schedule Type Filter Panel

<img src="0_Resources/Images/1_Salary_Dashboard_Type.png" width="400" alt="Schedule Type Dropdown">

---

## Key Formulas Explained

### 💰 Median Salary — Multi-Criteria Array Formula

```excel
=MEDIAN(
  IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
  )
)
```

**What it does:**
- Filters rows where job title, country, and schedule type all match the selected values
- Excludes records with zero or blank salary
- Returns the **median** (not average) of the matching salary values
- The `*` between conditions acts as AND — all must be TRUE

**Background helper table** (this table feeds all the chart and stat card calculations):

![Screenshot1](0_Resources/Images/1_Salary_Dashboard_Screenshot1.png)

> ⚠️ **Array formula:** In older Excel versions, enter with `Ctrl + Shift + Enter`.

---

### ⏰ Unique Schedule Type List — FILTER Formula

```excel
=FILTER(
  J2#,
  (NOT(ISNUMBER(SEARCH("and",J2#))+ISNUMBER(SEARCH(",",J2#))))*(J2#<>0)
)
```

**What it does:**
- Removes combined entries like "Full-time and Part-time" or "Full-time, Contractor"
- Removes blank/zero entries
- Returns a **clean, validated list** used to populate the Schedule Type dropdown

**FILTER formula output table:**

![Screenshot2](0_Resources/Images/1_Salary_Dashboard_Screenshot2.png)

---

### ❎ Data Validation — Dropdown Restriction

The filtered clean list is applied as a **Data Validation rule** under `Data → Data Validation → List`. This ensures users can only select valid schedule types — preventing formula errors and keeping the dashboard reliable.

![Type Dropdown](0_Resources/Images/1_Salary_Dashboard_Screenshot3.png)

---

## How to Use the Dashboard

| Step | Action | Result |
|------|--------|--------|
| **1** | Open `Salary_Dashboard.xlsx` | Dashboard loads with all jobs visible |
| **2** | Click the **Job Title** dropdown | Select any of the 10 data roles |
| **3** | Click the **Country** dropdown | Filter to a specific country or keep "All" |
| **4** | Click the **Schedule Type** dropdown | Choose Full-time, Part-time, Contractor, etc. |
| **5** | Read the **Median Salary** card | Result updates automatically |
| **6** | Read the **Bar Chart** | Highlighted bar shows your selected role vs all others |

> 💡 **Tip:** Start with "All" in Country and Type to see the broadest market view. Then narrow down to your target country and schedule type to get a precise benchmark.

---

## Key Insights

- 🏆 **Senior Data Scientists** earn the highest median salary (~$155K/yr in the US)
- 🇺🇸 **Full-time US roles** consistently pay 15–25% above the global median for the same title
- 📊 **Data Analysts** have the widest salary spread — from ~$50K entry-level to $150K+ senior
- 🌐 **Remote roles** show a slightly higher median, skewing toward senior-level hires
- 📋 **5,849 job listings** match the US · Data Analyst · Full-time filter — a strong, statistically reliable sample
- ⏱️ **Contractor roles** at top rates ($65–$85/hr) annualize to over $130K

---

## File Structure

```
📁 Data-Science-Salary-Dashboard/
│
├── 📊 Salary_Dashboard.xlsx              ← Main dashboard file (open this)
├── 📄 README.md                          ← This file
├── 🌐 excel_salary_dashboard_demo.html   ← Browser-based interactive demo
│
└── 📁 0_Resources/
    └── 📁 Images/
        ├── 1_Salary_Dashboard_Final_Dashboard.gif   ← Hero animation
        ├── 1_Salary_Dashboard.png                   ← Full dashboard screenshot
        ├── 1_Salary_Dashboard_Chart1.png            ← Salary bar chart
        ├── 1_Salary_Dashboard_Chart2.png            ← Map chart static
        ├── 1_Salary_Dashboard_Country_Map.gif       ← Map animation
        ├── 1_Salary_Dashboard_Data_Validation.gif   ← Dropdown interaction
        ├── 1_Salary_Dashboard_Job_Title.png         ← Job title filter panel
        ├── 1_Salary_Dashboard_Type.png              ← Schedule type filter panel
        ├── 1_Salary_Dashboard_Screenshot1.png       ← Median salary table
        ├── 1_Salary_Dashboard_Screenshot2.png       ← FILTER formula output
        └── 1_Salary_Dashboard_Screenshot3.png       ← Data validation panel
```

---

## About

**Built by:** Loknadh Venkata Krishna Sai Kona  
**GitHub:** [KrishnaSai315](https://github.com/KrishnaSai315)  
**LinkedIn:** [lvkrishna3](https://linkedin.com/in/lvkrishna3/)  

**Tools used:** Microsoft Excel 2021 / Microsoft 365  
**Dataset:** 2023 Data Science Job Listings (real-world, ~32,000 entries)  
**Project type:** Data Analytics Portfolio Project  

---

> ⭐ If this dashboard helped you understand data job salaries, consider **starring** the repository!
