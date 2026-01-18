# Resource Allocation Report Generator

A Node.js API service for generating Resource Allocation reports in Excel and PDF formats with SVG charts.

## Features

- **Excel Generation**: Multi-sheet Excel workbooks with role/user allocation data
- **PDF Generation**: Professional PDF reports with:
  - Cover page with Table of Contents
  - SVG semi-circular gauge charts showing utilization
  - SVG bar charts for Hours Planned vs Consumed by Period
  - SVG stacked bar charts for Projects comparison (Summary page)
  - Data tables with role/user breakdowns
  - Context blocks with metric definitions

## Installation

```bash
npm install
```

## Running the Server

```bash
node server.js
```

Server runs at `http://localhost:3000` by default.

## API Endpoints

### 1. Generate Excel Report

```
POST /api/generate-excel
Content-Type: application/json
```

Generates an Excel workbook (.xlsx) from the payload.

**Response**: Binary Excel file download

### 2. Generate PDF Report (Simplified)

```
POST /api/generate-pdf
Content-Type: application/json
```

Generates a PDF report with **user-level allocation data hidden** (Allocated Hours, Variance, and Effort Utilized columns show blanks for user rows). Role-level data is fully visible.

**Response**: Binary PDF file download

### 3. Generate PDF Report (Extended)

```
POST /api/generate-pdf-extended
Content-Type: application/json
```

Generates a PDF report with **full data for all rows** including user-level allocation details.

**Response**: Binary PDF file download

### 4. Health Check

```
GET /health
```

**Response**:
```json
{
  "status": "ok",
  "timestamp": "2026-01-17T12:00:00.000Z"
}
```

## Payload Format

The API accepts two payload formats:

### Multi-Tab Format (Recommended)

Used for reports with multiple projects/sheets:

```json
{
  "sheets": [
    {
      "sheetName": "Summary",
      "payload": {
        "meta": {
          "generatedOn": "2026-01-17 21:59:43",
          "startDate": "2025-03-01",
          "endDate": "2026-01-31",
          "bucket": "month",
          "months": [
            {"key": "2025-03", "label": "Mar 2025"},
            {"key": "2025-04", "label": "Apr 2025"}
          ],
          "context": {
            "type": "summary",
            "title": "Resource Allocation Summary",
            "description": "This summary aggregates resource allocation...",
            "date_context": {
              "portfolio_span": {
                "start": "2025-03-16",
                "end": "2027-10-24"
              },
              "reporting_period": {
                "start": "2025-03-01",
                "end": "2026-01-31"
              }
            },
            "metric_definitions": {
              "allocated_hours_total": "Allocated Hours (Total) is...",
              "actual_hours_period": "Actual Hours (Period) is...",
              "variance_hours": "Variance (Hours) = ...",
              "effort_utilized_pct": "Effort Utilized % = ..."
            }
          }
        },
        "rows": [
          {
            "level": "role",
            "label": "Senior Developer",
            "plannedTotal": 2000,
            "actualTotal": 1500,
            "effortPct": 75.0,
            "months": {
              "2025-03": {"planned": 200, "actual": 150},
              "2025-04": {"planned": 200, "actual": 180}
            },
            "roleKey": "abc123"
          },
          {
            "level": "user",
            "label": "John Smith",
            "plannedTotal": 1000,
            "actualTotal": 800,
            "effortPct": 80.0,
            "months": {
              "2025-03": {"planned": 100, "actual": 80},
              "2025-04": {"planned": 100, "actual": 90}
            },
            "parentRoleKey": "abc123",
            "userSysId": "user123"
          }
        ]
      }
    },
    {
      "sheetName": "PRJ001 - Project Name",
      "payload": {
        "meta": { ... },
        "rows": [ ... ]
      }
    }
  ]
}
```

### Single-Sheet Format

For simple single-project reports:

```json
{
  "meta": {
    "months": [...],
    "context": {...}
  },
  "rows": [...]
}
```

## Row Structure

Each row in the `rows` array should have:

| Field | Type | Description |
|-------|------|-------------|
| `level` | string | "role" or "user" |
| `label` | string | Display name for the role or user |
| `plannedTotal` | number | Total allocated hours |
| `actualTotal` | number | Total actual hours worked |
| `effortPct` | number | Utilization percentage |
| `months` | object | Monthly breakdown with planned/actual |
| `roleKey` | string | (roles only) Unique identifier for the role |
| `parentRoleKey` | string | (users only) Reference to parent role |
| `userSysId` | string | (users only) Unique user identifier |

## Example cURL Commands

### Generate Excel
```bash
curl -X POST http://localhost:3000/api/generate-excel \
  -H "Content-Type: application/json" \
  -d @payload_sample.json \
  -o report.xlsx
```

### Generate PDF (Simplified)
```bash
curl -X POST http://localhost:3000/api/generate-pdf \
  -H "Content-Type: application/json" \
  -d @payload_sample.json \
  -o report.pdf
```

### Generate PDF (Extended)
```bash
curl -X POST http://localhost:3000/api/generate-pdf-extended \
  -H "Content-Type: application/json" \
  -d @payload_sample.json \
  -o report_extended.pdf
```

## Minimal Sample Payload

```json
{
  "sheets": [
    {
      "sheetName": "Summary",
      "payload": {
        "meta": {
          "generatedOn": "2026-01-17",
          "startDate": "2026-01-01",
          "endDate": "2026-01-31",
          "bucket": "month",
          "months": [
            {"key": "2026-01", "label": "Jan 2026"}
          ],
          "context": {
            "type": "summary",
            "title": "Resource Allocation Summary"
          }
        },
        "rows": [
          {
            "level": "role",
            "label": "Developer",
            "plannedTotal": 160,
            "actualTotal": 120,
            "effortPct": 75,
            "months": {
              "2026-01": {"planned": 160, "actual": 120}
            },
            "roleKey": "role1"
          },
          {
            "level": "user",
            "label": "Jane Doe",
            "plannedTotal": 160,
            "actualTotal": 120,
            "effortPct": 75,
            "months": {
              "2026-01": {"planned": 160, "actual": 120}
            },
            "parentRoleKey": "role1",
            "userSysId": "user1"
          }
        ]
      }
    }
  ]
}
```

## Dependencies

- **express**: Web server framework
- **exceljs**: Excel file generation
- **puppeteer**: Headless Chrome for PDF generation

## License

Private - Sanofi Project
