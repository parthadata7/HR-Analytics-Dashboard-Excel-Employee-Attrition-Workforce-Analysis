# HR Analytics Dashboard (Excel)

## Screenshots

---

### HR Analytics Dashboard
![HR Analytics Dashboard](https://github.com/parthadata7/HR-Analytics-Dashboard-Excel-Employee-Attrition-Workforce-Analysis/blob/main/screenshots/hr_analytics_dashboard.PNG)

---

## Quick Links
- 📊 [Dashboard Analysis](#dashboard-analysis)
- 🧠 [Key Insights](#key-insights)
- 🗂️ [Project Structure](#project-structure)
- 🖼️ [Screenshots](#screenshots)
- 📬 [Contact](#contact)

---

## Table of Contents
- [Project Overview](#project-overview)
- [Business Problem](#business-problem)
- [Objectives](#objectives)
- [Dataset Description](#dataset-description)
- [Tools Used](#tools-used)
- [Dashboard Analysis](#dashboard-analysis)
- [Key Insights](#key-insights)
- [Business Impact](#business-impact)
- [Challenges Faced](#challenges-faced)
- [Key Learnings](#key-learnings)
- [Project Structure](#project-structure)
- [Screenshots](#screenshots)
- [Contact](#contact)

---

## Project Overview

The **HR Analytics Dashboard** is an interactive Excel-based workforce analytics solution designed to help Human Resource professionals monitor employee attrition, workforce demographics, compensation trends, and employee satisfaction.

Built entirely in **Microsoft Excel**, the dashboard leverages Pivot Tables, Pivot Charts, formulas, and slicers to transform HR data into meaningful business insights. It provides a centralized view of workforce health by highlighting key HR metrics such as employee count, attrition rate, average salary, average tenure, and attrition patterns across departments, job roles, age groups, salary bands, and overtime status.

This dashboard serves as a **descriptive analytics** tool that enables HR managers and business leaders to identify employee retention challenges, evaluate workforce trends, and support strategic HR decision-making.

---

## Business Problem

Organizations often face challenges in understanding why employees leave and which workforce segments experience the highest turnover.

Common business questions include:

- What is the overall employee attrition rate?
- Which departments experience the highest employee turnover?
- Which job roles are most affected by attrition?
- Does overtime contribute to employee resignation?
- How does attrition vary by age, salary, tenure, and job satisfaction?

Analyzing raw HR data manually is time-consuming and makes it difficult to identify workforce trends.

This project addresses these challenges by presenting critical HR metrics in a **single, interactive Excel dashboard**.

---

## Objectives

- Monitor overall workforce size and attrition
- Analyze employee turnover across departments
- Identify high-risk job roles with elevated attrition
- Evaluate attrition by age group, salary slab, and years at the company
- Examine the relationship between overtime and employee attrition
- Understand employee satisfaction levels associated with attrition
- Enable department-wise interactive analysis through slicers

---

## Dataset Description

- **Domain:** Human Resources / Workforce Analytics
- **Dataset:** Employee HR Dataset
- **Granularity:** Employee-level records

### Key Fields

- Employee ID
- Department
- Job Role
- Gender
- Age
- Attrition Status
- Monthly/Annual Salary
- Years at Company
- Overtime
- Job Satisfaction
- Education
- Marital Status

---

## Tools Used

- **Microsoft Excel**
  - Data Cleaning
  - Pivot Tables
  - Pivot Charts
  - Excel Formulas
  - KPI Cards
  - Interactive Slicers
  - Dashboard Design & Formatting

> ⚠️ **Note:** This project was developed entirely in **Microsoft Excel** without using Power BI, SQL, Python, or Tableau.

---

## DAX Measures Used

The following DAX measures were created in Power BI to calculate key HR analytics metrics:

| Measure | DAX Formula | Purpose |
|---------|-------------|---------|
| **Attrition Rate** | `attribution_rate = DIVIDE(CALCULATE(COUNT([EmpID]), Range[Attrition] = "YES"), COUNTROWS(Range))` | Calculates the percentage of employees who left the company. |
| **Total Employees** | `total_employ = COUNTROWS(Range)` | Counts the total number of employees in the dataset. |
| **Average Yearly Salary** | `average_salary_yearly = AVERAGE(Range[YearlyIncome])` | Calculates the average yearly income of employees. |
| **Total Attrition** | `attribution_total = CALCULATE(COUNT([EmpID]), Range[Attrition] = "YES")` | Counts the total number of employees who left the organization. |
| **Total Salary** | `total_salary = SUM(Range[YearlyIncome])` | Calculates the total yearly salary expenditure across all employees. |

### Key Metrics Generated
- 📊 Attrition Rate (%)
- 👥 Total Employees
- 💰 Average Yearly Salary
- 🚪 Total Employee Attrition
- 💵 Total Salary Expense

---

## Dashboard Analysis

The dashboard provides a comprehensive overview of workforce performance through multiple visualizations and KPI cards.

### Key Performance Indicators (KPIs)

- **Total Employees:** 1,480
- **Attrition Rate:** 16%
- **Employees Left:** 238
- **Average Yearly Salary:** $78,060
- **Average Years at Company:** 7 Years

### Department Analysis

Displays employee attrition percentage across:

- Human Resources
- Research & Development
- Sales

This helps identify departments with comparatively higher employee turnover.

### Age Group Analysis

Shows attrition across different employee age groups:

- 18–25
- 26–35
- 36–45
- 46–55
- 55+

The visualization highlights which age segments contribute most to employee exits.

### Job Role Analysis

Ranks employee attrition by job role, allowing HR teams to identify positions experiencing the highest turnover.

Examples include:

- Sales Representative
- Laboratory Technician
- Research Scientist
- Sales Executive
- Manufacturing Director
- Manager

### Attrition by Years at Company

A trend chart illustrates how attrition changes with employee tenure, helping identify critical periods when employees are most likely to leave.

### Overtime Analysis

A donut chart compares attrition between employees who:

- Work overtime
- Do not work overtime

This visualization helps evaluate the impact of overtime on employee retention.

### Salary Slab Analysis

Employee attrition is analyzed across salary ranges such as:

- Up to $5K
- $5K–10K
- $10K–15K
- $15K+

This helps determine whether compensation influences employee turnover.

### Job Satisfaction Analysis

Attrition is segmented by job satisfaction ratings, providing insight into employee engagement and satisfaction levels.

### Interactive Filters

Users can filter dashboard insights using:

- Department
- Gender

These filters enable focused workforce analysis for specific employee groups.

---

## Key Insights

- Overall employee attrition stands at **16%**, indicating a moderate turnover rate.
- The **Sales department** records the highest employee attrition among all departments.
- Employees aged **18–25** show the highest attrition, suggesting early-career turnover.
- Certain job roles, particularly customer-facing and operational positions, experience significantly higher employee exits.
- Employees working **overtime** are considerably more likely to leave the organization than those who do not.
- Lower salary bands contribute a larger share of employee attrition.
- Employee turnover varies with years of service, indicating critical retention periods during an employee's lifecycle.
- Job satisfaction levels reveal opportunities to improve employee engagement and retention.

---

## Business Impact

- Enables HR teams to monitor workforce health through real-time KPIs.
- Identifies departments and job roles requiring targeted retention strategies.
- Supports compensation and employee engagement initiatives.
- Helps management recognize workforce patterns before turnover becomes critical.
- Reduces manual HR reporting through an interactive Excel dashboard.
- Facilitates data-driven HR planning and strategic decision-making.

---

## Challenges Faced

- Designing an informative dashboard without visual clutter.
- Integrating multiple HR metrics into a single reporting interface.
- Creating meaningful KPI cards and consistent chart formatting.
- Ensuring accurate attrition calculations across multiple employee dimensions.
- Building interactive filtering while maintaining dashboard performance.

---

## Key Learnings

- Designing professional HR dashboards using Microsoft Excel.
- Building dynamic KPI cards and interactive reports.
- Using Pivot Tables and Pivot Charts for workforce analysis.
- Transforming raw HR data into actionable business insights.
- Applying effective dashboard design principles for executive reporting.

---

## Project Structure

```text
├── data/
│   └── hr_analytics_data.xlsx
├── screenshots/
│   └── hr_analytics_dashboard.png
└── README.md
```

---

## Contact

**Name:** Partha Pratim Das
**Email:** [parthadataanalyst@gmail.com](mailto:parthadataanalyst@gmail.com)
**LinkedIn:** [LinkedIn](linkedin.com/in/partha-pratim-das-01a579423) 
**Portfolio:** []()
