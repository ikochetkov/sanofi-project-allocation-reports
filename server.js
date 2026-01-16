const express = require('express');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

const app = express();
app.use(express.json({ 
  limit: '10mb',
  verify: (req, res, buf) => {
    req.rawBody = buf.toString();
  }
}));

// Custom JSON error handler
app.use((err, req, res, next) => {
  if (err instanceof SyntaxError && err.status === 400 && 'body' in err) {
    const position = err.message.match(/position (\d+)/)?.[1];
    let context = '';
    if (position && req.rawBody) {
      const pos = parseInt(position);
      const start = Math.max(0, pos - 50);
      const end = Math.min(req.rawBody.length, pos + 50);
      context = `\n\nContext around position ${pos}:\n...${req.rawBody.substring(start, end)}...\n${'─'.repeat(pos - start)}^`;
    }
    console.error(`JSON Parse Error: ${err.message}${context}`);
    return res.status(400).json({ 
      error: 'Invalid JSON in request body',
      details: err.message,
      hint: 'Check for trailing commas, unquoted keys, or invalid characters'
    });
  }
  next(err);
});

/**
 * Resource Allocation Excel Export
 * Generates Excel file matching ServiceNow RMW visual structure
 */

// Styling constants
const ROLE_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFE0E0E0' } // Light gray
};

const HEADER_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF333333' } // Dark #333
};

const HEADER_FONT = {
  bold: true,
  color: { argb: 'FFFFFFFF' } // White text
};

const BORDER_STYLE = {
  top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  right: { style: 'thin', color: { argb: 'FFB0B0B0' } }
};

const BORDER_STYLE_COL_E = {
  top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  right: { style: 'medium', color: { argb: 'FF333333' } }
};

const GRAND_TOTAL_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF333333' } // Dark background like header
};

const GRAND_TOTAL_FONT = {
  bold: true,
  color: { argb: 'FFFFFFFF' } // White text
};

/**
 * Generate a single worksheet with the standard styling
 * @param {ExcelJS.Workbook} workbook - The workbook to add the sheet to
 * @param {string} sheetName - Name for the worksheet
 * @param {Object} sheetPayload - The payload with meta.months and rows
 */
function generateSheet(workbook, sheetName, sheetPayload) {
  const sheet = workbook.addWorksheet(sheetName);

  const months = sheetPayload.meta?.months || [];
  const rows = sheetPayload.rows || [];

  // Calculate total columns: A (label) + B (Allocated) + C (Actual) + D (Variance) + E (%) + 2 per month
  const monthStartCol = 6; // F

  // Set column widths
  sheet.getColumn(1).width = 32; // A - Role/Resource
  sheet.getColumn(2).width = 16; // B - Allocated Hours
  sheet.getColumn(3).width = 16; // C - Actual Hours
  sheet.getColumn(4).width = 16; // D - Variance
  sheet.getColumn(5).width = 16; // E - Effort %
  
  for (let i = 0; i < months.length * 2; i++) {
    sheet.getColumn(monthStartCol + i).width = 14;
  }

  // ============ ROW 1 & 2: Fixed Headers (Merged Vertically) ============
  const headerRow1 = sheet.getRow(1);
  const headerRow2 = sheet.getRow(2);
  
  // Merge A1:A2 - Role/User
  sheet.mergeCells(1, 1, 2, 1);
  headerRow1.getCell(1).value = 'Role/User';
  headerRow1.getCell(1).font = HEADER_FONT;
  headerRow1.getCell(1).fill = HEADER_FILL;
  headerRow1.getCell(1).border = BORDER_STYLE;
  headerRow1.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow2.getCell(1).border = BORDER_STYLE;

  // Merge B1:B2 - Allocated Hours
  sheet.mergeCells(1, 2, 2, 2);
  headerRow1.getCell(2).value = 'Allocated Hours';
  headerRow1.getCell(2).font = HEADER_FONT;
  headerRow1.getCell(2).fill = HEADER_FILL;
  headerRow1.getCell(2).border = BORDER_STYLE;
  headerRow1.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow2.getCell(2).border = BORDER_STYLE;

  // Merge C1:C2 - Actual Hours
  sheet.mergeCells(1, 3, 2, 3);
  headerRow1.getCell(3).value = 'Actual Hours';
  headerRow1.getCell(3).font = HEADER_FONT;
  headerRow1.getCell(3).fill = HEADER_FILL;
  headerRow1.getCell(3).border = BORDER_STYLE;
  headerRow1.getCell(3).alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow2.getCell(3).border = BORDER_STYLE;

  // Merge D1:D2 - Variance
  sheet.mergeCells(1, 4, 2, 4);
  headerRow1.getCell(4).value = 'Variance';
  headerRow1.getCell(4).font = HEADER_FONT;
  headerRow1.getCell(4).fill = HEADER_FILL;
  headerRow1.getCell(4).border = BORDER_STYLE;
  headerRow1.getCell(4).alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow2.getCell(4).border = BORDER_STYLE;

  // Merge E1:E2 - Effort Utilized
  sheet.mergeCells(1, 5, 2, 5);
  headerRow1.getCell(5).value = 'Effort Utilized';
  headerRow1.getCell(5).font = HEADER_FONT;
  headerRow1.getCell(5).fill = HEADER_FILL;
  headerRow1.getCell(5).border = BORDER_STYLE_COL_E;
  headerRow1.getCell(5).alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow2.getCell(5).border = BORDER_STYLE_COL_E;

  // Month headers (merged across 2 columns each) - Row 1 only
  months.forEach((month, idx) => {
    const startCol = monthStartCol + idx * 2;
    const endCol = startCol + 1;
    
    // Merge cells for month header
    sheet.mergeCells(1, startCol, 1, endCol);
    
    const cell = headerRow1.getCell(startCol);
    cell.value = month.label;
    cell.font = HEADER_FONT;
    cell.fill = HEADER_FILL;
    cell.border = BORDER_STYLE;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    
    // Apply border to merged cell end
    headerRow1.getCell(endCol).border = BORDER_STYLE;
  });

  headerRow1.height = 20;
  headerRow2.height = 20;

  // Month sub-headers (Allocated / Actual)
  months.forEach((month, idx) => {
    const allocCol = monthStartCol + idx * 2;
    const actualCol = allocCol + 1;
    
    const allocCell = headerRow2.getCell(allocCol);
    allocCell.value = 'Allocated';
    allocCell.font = HEADER_FONT;
    allocCell.fill = HEADER_FILL;
    allocCell.border = BORDER_STYLE;
    allocCell.alignment = { vertical: 'middle', horizontal: 'center' };
    
    const actualCell = headerRow2.getCell(actualCol);
    actualCell.value = 'Actual';
    actualCell.font = HEADER_FONT;
    actualCell.fill = HEADER_FILL;
    actualCell.border = BORDER_STYLE;
    actualCell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  headerRow2.height = 20;

  // ============ DATA ROWS ============
  let currentRow = 3;

  rows.forEach((rowData) => {
    const excelRow = sheet.getRow(currentRow);
    const isRole = rowData.level === 'role';

    // Column A: Label
    const labelCell = excelRow.getCell(1);
    labelCell.value = rowData.label;
    labelCell.border = BORDER_STYLE;
    
    if (isRole) {
      labelCell.font = { bold: true };
      labelCell.fill = ROLE_FILL;
    } else {
      labelCell.font = { bold: false };
      labelCell.alignment = { indent: 2 }; // Indent for user rows
    }

    // Column B: Allocated Hours (plannedTotal)
    const allocTotalCell = excelRow.getCell(2);
    allocTotalCell.value = rowData.plannedTotal;
    allocTotalCell.numFmt = '#,##0.00';
    allocTotalCell.border = BORDER_STYLE;
    allocTotalCell.alignment = { horizontal: 'center' };
    if (isRole) {
      allocTotalCell.font = { bold: true };
      allocTotalCell.fill = ROLE_FILL;
    }

    // Column C: Actual Hours (actualTotal)
    const actualTotalCell = excelRow.getCell(3);
    actualTotalCell.value = rowData.actualTotal;
    actualTotalCell.numFmt = '#,##0.00';
    actualTotalCell.border = BORDER_STYLE;
    actualTotalCell.alignment = { horizontal: 'center' };
    if (isRole) {
      actualTotalCell.font = { bold: true };
      actualTotalCell.fill = ROLE_FILL;
    }

    // Column D: Variance (Allocated - Actual) as Excel formula
    const varianceCell = excelRow.getCell(4);
    varianceCell.value = { formula: `B${currentRow}-C${currentRow}` };
    varianceCell.numFmt = '#,##0.00';
    varianceCell.border = BORDER_STYLE;
    varianceCell.alignment = { horizontal: 'center' };
    if (isRole) {
      varianceCell.font = { bold: true };
      varianceCell.fill = ROLE_FILL;
    }

    // Column E: Effort Utilized % (Actual / Allocated) as Excel formula
    const effortCell = excelRow.getCell(5);
    effortCell.value = { formula: `IF(B${currentRow}=0,0,C${currentRow}/B${currentRow})` };
    effortCell.numFmt = '0.00%';
    effortCell.border = BORDER_STYLE_COL_E;
    effortCell.alignment = { horizontal: 'center' };
    if (isRole) {
      effortCell.font = { bold: true };
      effortCell.fill = ROLE_FILL;
    }

    // Monthly columns
    months.forEach((month, idx) => {
      const allocCol = monthStartCol + idx * 2;
      const actualCol = allocCol + 1;
      
      const monthData = rowData.months?.[month.key] || { planned: 0, actual: 0 };
      
      // Allocated
      const monthAllocCell = excelRow.getCell(allocCol);
      monthAllocCell.value = monthData.planned;
      monthAllocCell.numFmt = '#,##0.00';
      monthAllocCell.border = BORDER_STYLE;
      monthAllocCell.alignment = { horizontal: 'center' };
      if (isRole) {
        monthAllocCell.font = { bold: true };
        monthAllocCell.fill = ROLE_FILL;
      }
      
      // Actual
      const monthActualCell = excelRow.getCell(actualCol);
      monthActualCell.value = monthData.actual;
      monthActualCell.numFmt = '#,##0.00';
      monthActualCell.border = BORDER_STYLE;
      monthActualCell.alignment = { horizontal: 'center' };
      if (isRole) {
        monthActualCell.font = { bold: true };
        monthActualCell.fill = ROLE_FILL;
      }
    });

    currentRow++;
  });

  // ============ GRAND TOTAL ROW ============
  // Calculate totals from role-level rows only
  let grandTotalAllocated = 0;
  let grandTotalActual = 0;
  const grandTotalMonths = {};

  rows.filter(r => r.level === 'role').forEach((roleRow) => {
    grandTotalAllocated += roleRow.plannedTotal || 0;
    grandTotalActual += roleRow.actualTotal || 0;
    
    // Sum monthly data
    months.forEach((month) => {
      const monthData = roleRow.months?.[month.key] || { planned: 0, actual: 0 };
      if (!grandTotalMonths[month.key]) {
        grandTotalMonths[month.key] = { planned: 0, actual: 0 };
      }
      grandTotalMonths[month.key].planned += monthData.planned || 0;
      grandTotalMonths[month.key].actual += monthData.actual || 0;
    });
  });

  // Calculate effort utilized % (Actual / Allocated)
  const grandTotalEffort = grandTotalAllocated > 0 ? grandTotalActual / grandTotalAllocated : 0;

  // Add GRAND TOTAL row
  const grandTotalRow = sheet.getRow(currentRow);

  // Column A: Label
  const gtLabelCell = grandTotalRow.getCell(1);
  gtLabelCell.value = 'GRAND TOTAL';
  gtLabelCell.font = GRAND_TOTAL_FONT;
  gtLabelCell.fill = GRAND_TOTAL_FILL;
  gtLabelCell.border = BORDER_STYLE;
  gtLabelCell.alignment = { horizontal: 'left', vertical: 'middle' };

  // Column B: Allocated Hours
  const gtAllocCell = grandTotalRow.getCell(2);
  gtAllocCell.value = grandTotalAllocated;
  gtAllocCell.numFmt = '#,##0.00';
  gtAllocCell.font = GRAND_TOTAL_FONT;
  gtAllocCell.fill = GRAND_TOTAL_FILL;
  gtAllocCell.border = BORDER_STYLE;
  gtAllocCell.alignment = { horizontal: 'center' };

  // Column C: Actual Hours
  const gtActualCell = grandTotalRow.getCell(3);
  gtActualCell.value = grandTotalActual;
  gtActualCell.numFmt = '#,##0.00';
  gtActualCell.font = GRAND_TOTAL_FONT;
  gtActualCell.fill = GRAND_TOTAL_FILL;
  gtActualCell.border = BORDER_STYLE;
  gtActualCell.alignment = { horizontal: 'center' };

  // Column D: Variance (Allocated - Actual) as Excel formula
  const gtVarianceCell = grandTotalRow.getCell(4);
  gtVarianceCell.value = { formula: `B${currentRow}-C${currentRow}` };
  gtVarianceCell.numFmt = '#,##0.00';
  gtVarianceCell.font = GRAND_TOTAL_FONT;
  gtVarianceCell.fill = GRAND_TOTAL_FILL;
  gtVarianceCell.border = BORDER_STYLE;
  gtVarianceCell.alignment = { horizontal: 'center' };

  // Column E: Effort Utilized % (Actual / Allocated) as Excel formula
  const gtEffortCell = grandTotalRow.getCell(5);
  gtEffortCell.value = { formula: `IF(B${currentRow}=0,0,C${currentRow}/B${currentRow})` };
  gtEffortCell.numFmt = '0.00%';
  gtEffortCell.font = GRAND_TOTAL_FONT;
  gtEffortCell.fill = GRAND_TOTAL_FILL;
  gtEffortCell.border = BORDER_STYLE_COL_E;
  gtEffortCell.alignment = { horizontal: 'center' };

  // Monthly columns for GRAND TOTAL
  months.forEach((month, idx) => {
    const allocCol = monthStartCol + idx * 2;
    const actualCol = allocCol + 1;
    const monthData = grandTotalMonths[month.key] || { planned: 0, actual: 0 };

    // Allocated
    const gtMonthAllocCell = grandTotalRow.getCell(allocCol);
    gtMonthAllocCell.value = monthData.planned;
    gtMonthAllocCell.numFmt = '#,##0.00';
    gtMonthAllocCell.font = GRAND_TOTAL_FONT;
    gtMonthAllocCell.fill = GRAND_TOTAL_FILL;
    gtMonthAllocCell.border = BORDER_STYLE;
    gtMonthAllocCell.alignment = { horizontal: 'center' };

    // Actual
    const gtMonthActualCell = grandTotalRow.getCell(actualCol);
    gtMonthActualCell.value = monthData.actual;
    gtMonthActualCell.numFmt = '#,##0.00';
    gtMonthActualCell.font = GRAND_TOTAL_FONT;
    gtMonthActualCell.fill = GRAND_TOTAL_FILL;
    gtMonthActualCell.border = BORDER_STYLE;
    gtMonthActualCell.alignment = { horizontal: 'center' };
  });

  // ============ FREEZE PANES ============
  // Freeze first 5 columns (A-E) and first 2 rows (headers)
  sheet.views = [
    { state: 'frozen', xSplit: 5, ySplit: 2, topLeftCell: 'F3', activeCell: 'F3' }
  ];
}

/**
 * Sanitize sheet name for Excel (31 char limit, no special chars, unique)
 * @param {string} rawName - The raw sheet name from payload
 * @param {Object} usedNames - Object tracking used names for uniqueness
 * @returns {string} Safe, unique sheet name
 */
function makeSafeUniqueSheetName(rawName, usedNames = {}) {
  let name = (rawName || 'Sheet').toString();

  // Remove invalid Excel sheet characters: : \ / ? * [ ]
  name = name.replace(/[:\\\/\?\*\[\]]/g, ' ');

  // Normalize whitespace (collapse multiple spaces)
  name = name.replace(/\s+/g, ' ').trim();

  if (!name) name = 'Sheet';

  // Enforce Excel 31-character limit
  if (name.length > 31) {
    name = name.substring(0, 31).trim();
  }

  // Ensure uniqueness
  const base = name;
  let counter = 2;

  while (usedNames[name]) {
    const suffix = ` (${counter})`;
    const maxBaseLength = 31 - suffix.length;
    const trimmedBase = base.length > maxBaseLength
      ? base.substring(0, maxBaseLength).trim()
      : base;

    name = trimmedBase + suffix;
    counter++;

    if (counter > 99) break; // safety guard
  }

  usedNames[name] = true;
  return name;
}

/**
 * Generate Excel workbook from payload
 * Supports both single-sheet (legacy) and multi-tab (new) payload formats
 * @param {Object} payload - The resource allocation payload
 * @returns {ExcelJS.Workbook}
 */
async function generateExcel(payload) {
  const workbook = new ExcelJS.Workbook();
  const usedNames = {}; // Track used sheet names for uniqueness

  // Check if this is a multi-tab payload (has sheets array)
  if (payload.sheets && Array.isArray(payload.sheets)) {
    // Multi-tab format: iterate over each sheet entry
    payload.sheets.forEach((sheetEntry) => {
      const rawName = sheetEntry.sheetName || 'Sheet';
      const safeName = makeSafeUniqueSheetName(rawName, usedNames);
      const sheetPayload = sheetEntry.payload || { meta: { months: [] }, rows: [] };
      generateSheet(workbook, safeName, sheetPayload);
    });
  } else {
    // Legacy single-sheet format
    generateSheet(workbook, 'Resource Allocation (Monthly)', payload);
  }

  return workbook;
}

// ============ API ENDPOINT ============
app.post('/api/generate-excel', async (req, res) => {
  try {
    const payload = req.body;

    // Validate payload - support both single-sheet and multi-tab formats
    const isMultiTab = payload.sheets && Array.isArray(payload.sheets);
    const isSingleSheet = payload.meta && payload.meta.months && payload.rows;

    if (!isMultiTab && !isSingleSheet) {
      return res.status(400).json({
        error: 'Invalid payload. Required: either (meta.months + rows) for single sheet, or (sheets[]) for multi-tab format'
      });
    }

    const workbook = await generateExcel(payload);

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const filename = `Resource_Allocation_${timestamp}.xlsx`;

    // Set response headers for file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    // Write workbook to response
    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Error generating Excel:', error);
    res.status(500).json({ error: 'Failed to generate Excel file', details: error.message });
  }
});

// ============ PDF GENERATION ============

/**
 * Process payload into PDF-ready data structure
 * @param {Object} payload - The resource allocation payload
 * @returns {Object} Processed data for PDF template
 */
function processPayloadForPDF(payload) {
  const sheets = [];
  const chartsData = [];

  // Normalize to array of sheets
  const rawSheets = payload.sheets && Array.isArray(payload.sheets)
    ? payload.sheets
    : [{ sheetName: 'Resource Allocation', payload: payload }];

  // Count projects (excluding Summary)
  const projectCount = rawSheets.filter(s => s.sheetName !== 'Summary').length;

  rawSheets.forEach((sheetEntry) => {
    let sheetName = sheetEntry.sheetName || 'Sheet';
    const sheetPayload = sheetEntry.payload || payload;
    const rows = sheetPayload.rows || [];
    const months = sheetPayload.meta?.months || [];

    // Add project count to Summary title
    if (sheetName === 'Summary') {
      sheetName = `Summary - ${projectCount} projects`;
    }

    // Calculate grand totals from role-level rows
    let grandTotalAllocated = 0;
    let grandTotalActual = 0;
    const monthlyTotals = {};

    rows.filter(r => r.level === 'role').forEach((roleRow) => {
      grandTotalAllocated += roleRow.plannedTotal || 0;
      grandTotalActual += roleRow.actualTotal || 0;

      months.forEach((month) => {
        const monthData = roleRow.months?.[month.key] || { planned: 0, actual: 0 };
        if (!monthlyTotals[month.key]) {
          monthlyTotals[month.key] = { label: month.label, planned: 0, actual: 0 };
        }
        monthlyTotals[month.key].planned += monthData.planned || 0;
        monthlyTotals[month.key].actual += monthData.actual || 0;
      });
    });

    const grandTotalVariance = grandTotalAllocated - grandTotalActual;
    const grandTotalEffortPct = grandTotalAllocated > 0 
      ? (grandTotalActual / grandTotalAllocated) * 100 
      : 0;

    // Process rows for table
    const processedRows = rows.map(row => ({
      label: row.label,
      plannedTotal: row.plannedTotal || 0,
      actualTotal: row.actualTotal || 0,
      variance: (row.plannedTotal || 0) - (row.actualTotal || 0),
      effortPct: row.effortPct || 0,
      isRole: row.level === 'role'
    }));

    sheets.push({
      sheetName,
      rows: processedRows,
      grandTotal: {
        allocated: grandTotalAllocated,
        actual: grandTotalActual,
        variance: grandTotalVariance,
        effortPct: grandTotalEffortPct
      }
    });

    // Chart data for this sheet
    const chartLabels = months.map(m => m.label);
    const chartAllocated = months.map(m => monthlyTotals[m.key]?.planned || 0);
    const chartActual = months.map(m => monthlyTotals[m.key]?.actual || 0);

    chartsData.push({
      labels: chartLabels,
      allocated: chartAllocated,
      actual: chartActual
    });
  });

  return {
    generatedOn: payload.meta?.generatedOn || new Date().toISOString().slice(0, 19).replace('T', ' '),
    sheets,
    chartsData
  };
}

/**
 * Render HTML from data (replaces template engine)
 */
function renderPDFHtml(data) {
  const formatNum = (v) => (v || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const formatPct = (v) => formatNum(v) + '%';

  let sheetsHtml = '';
  
  data.sheets.forEach((sheet, idx) => {
    // Build rows HTML
    let rowsHtml = '';
    sheet.rows.forEach(row => {
      const rowClass = row.isRole ? 'role-row' : 'user-row';
      rowsHtml += `
        <tr class="${rowClass}">
          <td>${row.label}</td>
          <td>${formatNum(row.plannedTotal)}</td>
          <td>${formatNum(row.actualTotal)}</td>
          <td>${formatNum(row.variance)}</td>
          <td>${formatPct(row.effortPct)}</td>
        </tr>`;
    });

    sheetsHtml += `
    <div class="sheet-section">
      <div class="sheet-title">${sheet.sheetName}</div>
      
      <div class="chart-container">
        <canvas id="chart-${idx}"></canvas>
      </div>
      
      <table>
        <thead>
          <tr>
            <th style="width: 35%">Role / User</th>
            <th style="width: 16%">Allocated Hours</th>
            <th style="width: 16%">Actual Hours</th>
            <th style="width: 16%">Variance</th>
            <th style="width: 17%">Effort Utilized</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
          <tr class="grand-total">
            <td>GRAND TOTAL</td>
            <td>${formatNum(sheet.grandTotal.allocated)}</td>
            <td>${formatNum(sheet.grandTotal.actual)}</td>
            <td>${formatNum(sheet.grandTotal.variance)}</td>
            <td>${formatPct(sheet.grandTotal.effortPct)}</td>
          </tr>
        </tbody>
      </table>
    </div>`;
  });

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Resource Allocation Report</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Segoe UI', Arial, sans-serif; padding: 30px; color: #333; background: #fff; }
    .header { text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 2px solid #333; }
    .header h1 { font-size: 24px; color: #333; margin-bottom: 5px; }
    .header .subtitle { font-size: 14px; color: #666; }
    .sheet-section { margin-bottom: 40px; }
    .sheet-section:not(:first-of-type) { page-break-before: always; }
    .sheet-title { font-size: 16px; font-weight: bold; color: #333; margin-bottom: 20px; padding: 10px; background: #f5f5f5; border-left: 4px solid #333; }
    .chart-container { width: 100%; height: 300px; margin-bottom: 30px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 11px; }
    th { background: #333; color: #fff; font-weight: bold; padding: 10px 8px; text-align: center; border: 1px solid #333; }
    th:first-child { text-align: left; }
    td { padding: 8px; border: 1px solid #ccc; text-align: center; }
    td:first-child { text-align: left; }
    tr.role-row { background: #e0e0e0; font-weight: bold; }
    tr.user-row td:first-child { padding-left: 20px; }
    tr.grand-total { background: #333; color: #fff; font-weight: bold; }
    tr.grand-total td { border-color: #333; }
    .footer { margin-top: 30px; padding-top: 15px; border-top: 1px solid #ccc; font-size: 10px; color: #999; text-align: center; }
  </style>
</head>
<body>
  <div class="header">
    <h1>Resource Allocation Report</h1>
    <div class="subtitle">Generated on ${data.generatedOn}</div>
  </div>

  ${sheetsHtml}

  <div class="footer">
    Resource Management Workbench Report • Confidential
  </div>

  <script>
    Chart.register(ChartDataLabels);
    const chartsData = ${JSON.stringify(data.chartsData)};
    
    chartsData.forEach((chartData, index) => {
      const ctx = document.getElementById('chart-' + index);
      if (!ctx) return;

      new Chart(ctx, {
        type: 'bar',
        data: {
          labels: chartData.labels,
          datasets: [
            { label: 'Allocated', data: chartData.allocated, backgroundColor: 'rgba(51, 51, 51, 0.85)', borderColor: '#333', borderWidth: 1 },
            { label: 'Actual', data: chartData.actual, backgroundColor: 'rgba(76, 175, 80, 0.85)', borderColor: '#4CAF50', borderWidth: 1 }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'top', labels: { font: { size: 11 } } },
            title: { display: true, text: 'Hours Planned vs Consumed by Period', font: { size: 14, weight: 'bold' } },
            datalabels: {
              anchor: 'end', align: 'top', offset: 2,
              font: { size: 10, weight: 'bold' }, color: '#333',
              formatter: (value) => value === 0 ? '' : value.toLocaleString('en-US', { maximumFractionDigits: 0 })
            }
          },
          scales: {
            x: { title: { display: true, text: 'Allocated/Actual Utilization', font: { size: 12, weight: 'bold' } } },
            y: { beginAtZero: true, title: { display: true, text: 'Hours', font: { size: 12, weight: 'bold' } }, ticks: { callback: function(v) { return v.toLocaleString(); } } }
          }
        }
      });
    });
  </script>
</body>
</html>`;
}

/**
 * Generate PDF from payload
 * @param {Object} payload - The resource allocation payload
 * @returns {Buffer} PDF buffer
 */
async function generatePDF(payload) {
  // Process payload
  const data = processPayloadForPDF(payload);

  // Render HTML directly (no template file needed now)
  const html = renderPDFHtml(data);

  // Launch Puppeteer and generate PDF
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0' });

  // Wait for charts to render
  await page.waitForFunction(() => {
    const canvases = document.querySelectorAll('canvas');
    return canvases.length === 0 || Array.from(canvases).every(c => c.getContext('2d'));
  });
  await new Promise(resolve => setTimeout(resolve, 500)); // Extra time for Chart.js animations

  const pdfBuffer = await page.pdf({
    format: 'A4',
    printBackground: true,
    margin: { top: '20mm', right: '15mm', bottom: '20mm', left: '15mm' }
  });

  await browser.close();

  return pdfBuffer;
}

// PDF Generation Endpoint
app.post('/api/generate-pdf', async (req, res) => {
  try {
    const payload = req.body;

    // Validate payload
    const isMultiTab = payload.sheets && Array.isArray(payload.sheets);
    const isSingleSheet = payload.meta && payload.meta.months && payload.rows;

    if (!isMultiTab && !isSingleSheet) {
      return res.status(400).json({
        error: 'Invalid payload. Required: either (meta.months + rows) for single sheet, or (sheets[]) for multi-tab format'
      });
    }

    const pdfBuffer = await generatePDF(payload);

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const filename = `Resource_Allocation_${timestamp}.pdf`;

    // Set response headers for file download
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', pdfBuffer.length);

    res.end(pdfBuffer);

  } catch (error) {
    console.error('Error generating PDF:', error);
    res.status(500).json({ error: 'Failed to generate PDF file', details: error.message });
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Simple HTML interface
app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Resource Allocation Excel Generator</title>
      <style>
        body { font-family: Arial, sans-serif; max-width: 900px; margin: 40px auto; padding: 20px; }
        h1 { color: #333; }
        textarea { width: 100%; height: 300px; font-family: monospace; font-size: 12px; }
        button { background: #0066cc; color: white; padding: 12px 24px; border: none; cursor: pointer; font-size: 16px; margin-top: 10px; }
        button:hover { background: #0052a3; }
        .info { background: #f0f0f0; padding: 15px; margin: 20px 0; border-radius: 5px; }
        code { background: #e0e0e0; padding: 2px 6px; border-radius: 3px; }
      </style>
    </head>
    <body>
      <h1>📊 Resource Allocation Excel Generator</h1>
      
      <div class="info">
        <strong>API Endpoint:</strong> <code>POST /api/generate-excel</code><br>
        <strong>Content-Type:</strong> <code>application/json</code><br>
        <strong>Response:</strong> Binary Excel file (.xlsx)
      </div>

      <h3>Paste JSON Payload:</h3>
      <textarea id="payload" placeholder='{"meta": {"months": [...]}, "rows": [...]}'></textarea>
      
      <button onclick="generateExcel()">Generate Excel</button>

      <script>
        async function generateExcel() {
          const payload = document.getElementById('payload').value;
          
          try {
            const parsed = JSON.parse(payload);
            
            const response = await fetch('/api/generate-excel', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: payload
            });
            
            if (!response.ok) {
              const error = await response.json();
              alert('Error: ' + error.error);
              return;
            }
            
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Resource_Allocation.xlsx';
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);
            
          } catch (e) {
            alert('Invalid JSON: ' + e.message);
          }
        }
      </script>
    </body>
    </html>
  `);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Resource Allocation API running at http://localhost:${PORT}`);
  console.log(`📋 POST /api/generate-excel - Generate Excel from JSON payload`);
  console.log(`📄 POST /api/generate-pdf - Generate PDF report with charts`);
  console.log(`💊 GET /health - Health check`);
});
