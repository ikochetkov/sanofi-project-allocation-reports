
# PDF Generator Enhancement Instructions – Context Blocks

## Goal
Improve the PDF generator so that each **Summary** and **Project** page clearly explains:
- What data is shown
- Which projects are included
- What date range the data represents
- How key metrics (Allocated, Actual, Variance, Effort Utilized) should be interpreted

This context must be rendered **before charts and tables**, immediately after the page title.

---

## Source of Truth (Payload)
All contextual text must be derived from the **payload metadata**, not recalculated in the PDF generator.

### Summary Page
Use:
```
workbook.meta
workbook.sheets[0].payload.meta
```

### Project Pages
For each project tab:
```
sheet.sheetName
sheet.payload.meta
sheet.payload.meta.context   (if present)
```

---

## Where to Render Context in PDF

### Page Structure (Project Page)

1. **Title**
   - Example:
     ```
     PRJ0100697 – Sanofi – Azure Arc Deployment
     ```

2. **Context Block (NEW – REQUIRED)**
   - Render directly under the title
   - Before charts and tables

3. Charts

4. Tables

---

## Context Block – Project Page

### Content (render as short paragraph or bullet-style lines)

Use the following wording template:

```
Project date span: {project_start} → {project_end}
Reporting period: {report_start} → {report_end}

Allocated Hours (Total): Total allocated hours for this project (lifetime allocation).
Actual Hours (Period): Actual hours recorded during the reporting period.
Variance: Remaining hours calculated as Allocated − Actual.
Effort Utilized: Actual Hours ÷ Allocated Hours × 100.
```

### Field Mapping
| Text Placeholder | Payload Source |
|-----------------|---------------|
| project_start | sheet.payload.meta.projectStart |
| project_end | sheet.payload.meta.projectEnd |
| report_start | sheet.payload.meta.startDate |
| report_end | sheet.payload.meta.endDate |

If `projectStart` / `projectEnd` are not present, fall back to:
```
sheet.payload.meta.startDate / endDate
```

---

## Context Block – Summary Page

### Content Template

```
This Summary aggregates allocated and actual effort across {project_count} active projects.

Portfolio date span: {earliest_project_start} → {latest_project_end}
Reporting period: {report_start} → {report_end}

Allocated Hours (Total): Sum of all allocations across included projects.
Actual Hours (Period): Sum of actual hours recorded within the reporting period.
Variance: Allocated − Actual.
Effort Utilized: Actual ÷ Allocated × 100.
```

### Field Mapping
| Text Placeholder | Payload Source |
|-----------------|---------------|
| project_count | workbook.meta.projectCount |
| earliest_project_start | workbook.meta.portfolioStart |
| latest_project_end | workbook.meta.portfolioEnd |
| report_start | workbook.meta.summary.startDate |
| report_end | workbook.meta.summary.endDate |

---

## Formatting Rules

- Place context inside a **light gray box** or subtle container
- Font size: **9–10 pt**
- Line spacing: slightly tighter than main body
- Max height: **6–7 lines**
- Left-aligned text
- No charts or tables should appear above this block

---

## Important Notes

- PDF generator **must not compute metrics**
- PDF generator **must not infer dates**
- If context fields are missing, gracefully hide only that line
- Do NOT truncate project names in PDF (already handled in payload)

---

## Result

After this change:
- Every PDF page explains exactly what the reader is seeing
- Customers understand effort utilization without verbal explanation
- Summary vs Project pages are clearly differentiated
