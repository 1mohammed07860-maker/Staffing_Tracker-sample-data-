# Staffing_Tracker-sample-data-
# Staffing Tracker Dashboard (Excel)

An Excel workbook that tracks who is assigned to which project, how busy each person is, and which projects still need people. It is modeled on the kind of staffing and scheduling data a resource management team handles day to day.

> **All data in this workbook is fictional sample data.** Names, projects and hours were made up for practice and do not describe any real person or organization.

![Dashboard](dashboard.png)

## What it answers

- How many hours are assigned in total, and across which business units?
- Which employees are over capacity (above 100%) or under-utilized (below 70%)?
- Which projects still have unfilled demand, and how many hours?

## Key results from the sample data

| Metric | Result |
|---|---|
| Employees / projects / assignments | 17 / 6 / 30 |
| Total hours assigned | 520 |
| Average utilization | 76% |
| Over capacity (above 100%) | 2 employees (E05 and E12, both at 113%) |
| Under-utilized (below 70%) | 6 employees |
| Open demand | 35 hours across 3 projects (P02: 10, P05: 5, P06: 20) |

Average utilization by business unit: Audit 81%, Consulting 73%, Tax 75%.

## Workbook structure

| Sheet | Purpose |
|---|---|
| `employees` | Employee master list (ID, name, business unit, grade, location, 40-hour capacity) plus calculated hours assigned and utilization |
| `Projects` | Project list with hours needed, plus calculated hours assigned, gap and status |
| `Assignments` | Which employee works on which project, and for how many hours per week |
| `dashboards` | KPI cards, four pivot tables, a chart and a business unit slicer |

## Excel features used

- **VLOOKUP** pulls employee name, business unit, grade and location, and project name, into the Assignments sheet, so only IDs and hours need to be typed.
  `=VLOOKUP(B2,employees!$A:$F,2,FALSE)`
- **SUMIF** totals hours per employee and per project.
  `=SUMIF(Assignments!$B:$B,A2,Assignments!$D:$D)`
- **Utilization** is hours assigned divided by weekly capacity.
  `=G2/F2`
- **Automatic status** marks each project Open or Staffed based on its gap.
  `=IF(G2>0,"Open","Staffed")`
- **Data validation** limits entries on the Assignments sheet: dropdown lists for Employee ID and Project ID, and a whole-number rule (0 to 60) for hours per week.
- **Conditional formatting** highlights utilization above 100% in red and below 70% in yellow, and flags projects with open demand.
- **Pivot tables and pivot chart** summarize hours by employee, project and business unit.
- **Slicer** filters the dashboard by business unit.
- **KPI cards** use SUM and COUNTIF formulas for total hours, average utilization, over-capacity count, under-utilized count and open demand.

## How to use it

1. Download `Staffing_Tracker_Sample_Data.xlsx` and open it in Excel (built in Excel 2010, so it also works in newer versions).
2. To add an assignment, go to the `Assignments` sheet and add a row. Pick the Employee ID and Project ID from the dropdowns and enter the hours. The name, unit, grade, location and project name fill in automatically (copy the formulas down in columns E to I).
3. Go to the `dashboards` sheet and click **Data > Refresh All** (or right-click a pivot and choose Refresh) so the pivot tables and chart update.

## What I practiced

- Linking several tables together with lookups instead of typing repeated information
- Building checks into a sheet (validation) so bad data can't get in
- Turning raw assignment records into numbers a manager can act on
- Presenting the results in a dashboard that can be filtered

## Author

**Sayed Mohammed Ali**, BBA (Information Technology) student, Hyderabad, India
Email: 1mohammed07860@gmail.com
