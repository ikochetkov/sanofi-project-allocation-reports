const express = require('express');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

const app = express();

// API Key for authentication (set via environment variable)
const API_KEY = process.env.API_KEY || 'dev-key-change-me';

// API Key authentication middleware
function requireApiKey(req, res, next) {
  const apiKey = req.headers['x-api-key'] || req.headers['authorization']?.replace('Bearer ', '');

  if (!apiKey) {
    return res.status(401).json({ error: 'Missing API key. Provide via X-API-Key header.' });
  }

  if (apiKey !== API_KEY) {
    return res.status(403).json({ error: 'Invalid API key.' });
  }

  next();
}

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
 * Validate payload for required to-date effort fields
 * @param {Object} payload - The resource allocation payload
 * @throws {Error} If required fields are missing
 */
function validatePayloadFields(payload) {
  const errors = [];
  
  // Normalize to array of sheets
  const sheets = payload.sheets && Array.isArray(payload.sheets)
    ? payload.sheets
    : [{ sheetName: 'Sheet', payload: payload }];
  
  sheets.forEach((sheetEntry, sheetIdx) => {
    const sheetPayload = sheetEntry.payload || payload;
    const rows = sheetPayload.rows || [];
    const sheetName = sheetEntry.sheetName || `Sheet ${sheetIdx + 1}`;
    
    rows.forEach((row, rowIdx) => {
      if (row.allocated_effort_to_date === undefined) {
        errors.push(`Sheet "${sheetName}", row ${rowIdx + 1} (${row.label || 'unnamed'}): missing required field 'allocated_effort_to_date'`);
      }
      if (row.actual_effort_to_date === undefined) {
        errors.push(`Sheet "${sheetName}", row ${rowIdx + 1} (${row.label || 'unnamed'}): missing required field 'actual_effort_to_date'`);
      }
    });
  });
  
  if (errors.length > 0) {
    const errorMsg = `Payload validation failed. Required fields missing:\n${errors.slice(0, 10).join('\n')}${errors.length > 10 ? `\n... and ${errors.length - 10} more errors` : ''}`;
    throw new Error(errorMsg);
  }
}

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

  // Merge B1:B2 - Allocated (To Date)
  sheet.mergeCells(1, 2, 2, 2);
  headerRow1.getCell(2).value = 'Allocated (To Date)';
  headerRow1.getCell(2).font = HEADER_FONT;
  headerRow1.getCell(2).fill = HEADER_FILL;
  headerRow1.getCell(2).border = BORDER_STYLE;
  headerRow1.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow2.getCell(2).border = BORDER_STYLE;

  // Merge C1:C2 - Actual (To Date)
  sheet.mergeCells(1, 3, 2, 3);
  headerRow1.getCell(3).value = 'Actual (To Date)';
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

  // Merge E1:E2 - Effort Utilized (%)
  sheet.mergeCells(1, 5, 2, 5);
  headerRow1.getCell(5).value = 'Effort Utilized (%)';
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

    // Column B: Allocated (To Date)
    const allocTotalCell = excelRow.getCell(2);
    allocTotalCell.value = rowData.allocated_effort_to_date;
    allocTotalCell.numFmt = '#,##0.00';
    allocTotalCell.border = BORDER_STYLE;
    allocTotalCell.alignment = { horizontal: 'center' };
    if (isRole) {
      allocTotalCell.font = { bold: true };
      allocTotalCell.fill = ROLE_FILL;
    }

    // Column C: Actual (To Date)
    const actualTotalCell = excelRow.getCell(3);
    actualTotalCell.value = rowData.actual_effort_to_date;
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
    grandTotalAllocated += roleRow.allocated_effort_to_date || 0;
    grandTotalActual += roleRow.actual_effort_to_date || 0;
    
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
app.post('/api/generate-excel', requireApiKey, async (req, res) => {
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

    // Validate required to-date effort fields
    validatePayloadFields(payload);

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
    } else {
      // For project tabs, use the full name from context if available
      // Format: "PRJ0100697 - Project Name" from context.narrative or sheetName
      const context = sheetPayload.meta?.context;
      if (context && context.projectNumber && context.projectName) {
        sheetName = `${context.projectNumber} - ${context.projectName}`;
      } else if (context && context.projectNumber) {
        sheetName = context.projectNumber;
      }
      // Otherwise keep original sheetName
    }

    // Calculate grand totals from role-level rows
    let grandTotalAllocated = 0;
    let grandTotalActual = 0;
    const monthlyTotals = {};

    rows.filter(r => r.level === 'role').forEach((roleRow) => {
      grandTotalAllocated += roleRow.allocated_effort_to_date || 0;
      grandTotalActual += roleRow.actual_effort_to_date || 0;

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
      allocatedToDate: row.allocated_effort_to_date || 0,
      actualToDate: row.actual_effort_to_date || 0,
      variance: (row.allocated_effort_to_date || 0) - (row.actual_effort_to_date || 0),
      effortPct: row.effortPct || 0,
      isRole: row.level === 'role'
    }));

    // Extract context block data from payload
    const context = sheetPayload.meta?.context;
    let contextBlock = null;

    if (context) {
      if (context.type === 'summary') {
        // Summary page context
        contextBlock = {
          type: 'summary',
          title: context.title || 'Summary',
          description: context.description || '',
          portfolioSpan: context.date_context?.portfolio_span || null,
          reportingPeriod: context.date_context?.reporting_period || null,
          metricDefinitions: context.metric_definitions || null,
          notes: context.notes || []
        };
      } else if (context.type === 'project') {
        // Project page context
        contextBlock = {
          type: 'project',
          title: context.title || sheetName,
          description: context.description || '',
          projectSpan: context.date_context?.project_span || null,
          reportingPeriod: context.date_context?.reporting_period || null,
          metricDefinitions: context.metric_definitions || null
        };
      }
    }

    const isSummary = sheetEntry.sheetName === 'Summary';

    sheets.push({
      sheetName,
      rows: processedRows,
      contextBlock,
      isSummary,
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

  // Collect data for stacked bar chart (projects only, for Summary page)
  // Extract just the project name (remove PRJ number and "Sanofi - " prefix)
  const extractProjectName = (sheetName) => {
    // Pattern: "PRJ0002756 - Sanofi - Azure Arc Deployment" → "Azure Arc Deployment"
    // Or: "PRJ0002756 - Project Name" → "Project Name"
    let name = sheetName;
    // Remove PRJ number prefix (e.g., "PRJ0002756 - ")
    name = name.replace(/^PRJ\d+\s*-\s*/, '');
    // Remove "Sanofi - " prefix if present
    name = name.replace(/^Sanofi\s*-\s*/i, '');
    // Truncate if still too long
    if (name.length > 25) {
      name = name.substring(0, 22) + '...';
    }
    return name;
  };

  const projectsBarData = sheets
    .filter(s => !s.isSummary)
    .map(s => ({
      label: extractProjectName(s.sheetName),
      fullName: s.sheetName,  // Keep full name for table
      allocated: s.grandTotal.allocated,
      used: s.grandTotal.actual,
      unused: Math.max(0, s.grandTotal.allocated - s.grandTotal.actual),
      utilization: s.grandTotal.effortPct
    }));

  return {
    generatedOn: payload.meta?.generatedOn || new Date().toISOString().slice(0, 19).replace('T', ' '),
    sheets,
    chartsData,
    projectsBarData
  };
}

/**
 * Build SVG gauge (semi-circular) with overflow wedge support
 * @param {number} actualHours - Actual hours consumed
 * @param {number} plannedHours - Planned/allocated hours
 * @returns {string} SVG markup string
 */
function buildGaugeSVG(actualHours, plannedHours) {
  const width = 320;
  const height = 200;
  const strokeWidth = 40;
  const cx = width / 2;
  const cy = height - 30;
  const r = Math.min(width / 2 - strokeWidth / 2 - 10, height - strokeWidth / 2 - 20);
  
  // Calculate percentage
  const pct = plannedHours > 0 ? (actualHours / plannedHours) * 100 : 0;
  const basePct = Math.max(0, Math.min(100, pct));
  const overflowPct = Math.max(0, pct - 100);
  
  // Colors
  const fillColor = '#4F7F2D';  // Green for main gauge
  const bgColor = '#E6E6E6';    // Light gray background
  const overflowColor = '#A8C98C'; // Lighter green for overflow
  
  // Helper: polar to cartesian
  function polarToCartesian(cx, cy, r, deg) {
    const rad = (deg * Math.PI) / 180;
    return {
      x: cx + r * Math.cos(rad),
      y: cy - r * Math.sin(rad)
    };
  }
  
  // Helper: describe arc path
  function describeArc(cx, cy, r, startDeg, endDeg) {
    const start = polarToCartesian(cx, cy, r, startDeg);
    const end = polarToCartesian(cx, cy, r, endDeg);
    const largeArcFlag = Math.abs(endDeg - startDeg) > 180 ? 1 : 0;
    const sweepFlag = startDeg > endDeg ? 1 : 0;
    
    return [
      'M', start.x, start.y,
      'A', r, r, 0, largeArcFlag, sweepFlag, end.x, end.y
    ].join(' ');
  }
  
  // Arc angles: 180° (left) to 0° (right) for half circle
  const theta = 180 - (basePct / 100) * 180;
  
  // Background arc (full half circle)
  const bgPath = describeArc(cx, cy, r, 180, 0);
  
  // Value arc (from left to progress point)
  const valPath = basePct > 0 ? describeArc(cx, cy, r, 180, theta) : '';
  
  // Overflow wedge (beyond 100%, cap at 50% visually)
  const overflowCapPct = 50;
  const overflowShownPct = Math.min(overflowPct, overflowCapPct);
  const overflowDeg = (overflowShownPct / 100) * 90; // Max 45 degrees for overflow
  const overflowPath = overflowShownPct > 0 ? describeArc(cx, cy, r, 0, -overflowDeg) : '';
  
  return `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
  <!-- Background arc -->
  <path d="${bgPath}" fill="none" stroke="${bgColor}" stroke-width="${strokeWidth}" stroke-linecap="butt" />
  
  <!-- Value arc -->
  ${valPath ? `<path d="${valPath}" fill="none" stroke="${fillColor}" stroke-width="${strokeWidth}" stroke-linecap="butt" />` : ''}
  
  <!-- Overflow wedge -->
  ${overflowPath ? `<path d="${overflowPath}" fill="none" stroke="${overflowColor}" stroke-width="${strokeWidth}" stroke-linecap="butt" />` : ''}
  
  <!-- Center percentage label -->
  <text x="${cx}" y="${cy - strokeWidth * 0.3}" text-anchor="middle" dominant-baseline="middle"
        font-family="Arial, sans-serif" font-size="36" font-weight="700" fill="#000">
    ${pct.toFixed(1)}%
  </text>
  
  <!-- Left tick label (0%) -->
  <text x="${strokeWidth * 0.5}" y="${height + 18}" text-anchor="start" 
        font-family="Arial, sans-serif" font-size="14" font-weight="600" fill="#355B1B">0%</text>
  
  <!-- Right tick label (100%) -->
  <text x="${width - strokeWidth * 0.5}" y="${height + 18}" text-anchor="end" 
        font-family="Arial, sans-serif" font-size="14" font-weight="600" fill="#355B1B">100%</text>
</svg>`.trim();
}

/**
 * Build SVG grouped bar chart for Hours Planned vs Consumed by Period
 * @param {Object} chartData - { labels: string[], allocated: number[], actual: number[] }
 * @returns {string} SVG markup string
 */
function buildBarChartSVG(chartData) {
  const width = 700;
  const height = 300;
  const marginTop = 60;
  const marginRight = 30;
  const marginBottom = 60;
  const marginLeft = 60;
  
  const chartWidth = width - marginLeft - marginRight;
  const chartHeight = height - marginTop - marginBottom;
  
  const labels = chartData.labels || [];
  const allocated = chartData.allocated || [];
  const actual = chartData.actual || [];
  
  if (labels.length === 0) {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
      <text x="${width/2}" y="${height/2}" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" fill="#666">No data available</text>
    </svg>`;
  }
  
  // Calculate max value for Y axis
  const allValues = [...allocated, ...actual];
  const maxVal = Math.max(...allValues, 1);
  const yMax = Math.ceil(maxVal / 100) * 100 || 100; // Round up to nearest 100
  
  // Bar dimensions
  const groupWidth = chartWidth / labels.length;
  const barWidth = groupWidth * 0.35;
  const barGap = groupWidth * 0.05;
  
  // Colors
  const allocatedColor = '#333333';
  const actualColor = '#4CAF50';
  
  // Format number for display
  const formatNum = (v) => Math.round(v).toLocaleString('en-US');
  
  // Build Y axis ticks
  const yTicks = 5;
  const yTickStep = yMax / yTicks;
  let yAxisHtml = '';
  for (let i = 0; i <= yTicks; i++) {
    const val = i * yTickStep;
    const y = marginTop + chartHeight - (val / yMax) * chartHeight;
    yAxisHtml += `
      <line x1="${marginLeft}" y1="${y}" x2="${marginLeft + chartWidth}" y2="${y}" stroke="#e0e0e0" stroke-width="1" />
      <text x="${marginLeft - 8}" y="${y + 4}" text-anchor="end" font-family="Arial, sans-serif" font-size="10" fill="#666">${formatNum(val)}</text>`;
  }
  
  // Build bars
  let barsHtml = '';
  labels.forEach((label, i) => {
    const groupX = marginLeft + i * groupWidth + groupWidth * 0.15;
    
    // Allocated bar
    const allocHeight = (allocated[i] / yMax) * chartHeight;
    const allocY = marginTop + chartHeight - allocHeight;
    barsHtml += `
      <rect x="${groupX}" y="${allocY}" width="${barWidth}" height="${allocHeight}" fill="${allocatedColor}" />
      <text x="${groupX + barWidth/2}" y="${allocY - 5}" text-anchor="middle" font-family="Arial, sans-serif" font-size="9" font-weight="bold" fill="#333">${allocated[i] > 0 ? formatNum(allocated[i]) : ''}</text>`;
    
    // Actual bar
    const actHeight = (actual[i] / yMax) * chartHeight;
    const actY = marginTop + chartHeight - actHeight;
    const actX = groupX + barWidth + barGap;
    barsHtml += `
      <rect x="${actX}" y="${actY}" width="${barWidth}" height="${actHeight}" fill="${actualColor}" />
      <text x="${actX + barWidth/2}" y="${actY - 5}" text-anchor="middle" font-family="Arial, sans-serif" font-size="9" font-weight="bold" fill="#333">${actual[i] > 0 ? formatNum(actual[i]) : ''}</text>`;
    
    // X axis label
    const labelX = groupX + barWidth + barGap/2;
    barsHtml += `
      <text x="${labelX}" y="${marginTop + chartHeight + 20}" text-anchor="middle" font-family="Arial, sans-serif" font-size="10" fill="#333">${label}</text>`;
  });
  
  return `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
  <!-- Title -->
  <text x="${width/2}" y="18" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" font-weight="bold" fill="#333">Allocated vs Actual Effort (To Date by Month)</text>
  
  <!-- Legend (under title) -->
  <rect x="${width/2 - 90}" y="26" width="14" height="14" fill="${allocatedColor}" />
  <text x="${width/2 - 72}" y="37" font-family="Arial, sans-serif" font-size="11" fill="#333">Allocated</text>
  <rect x="${width/2 + 10}" y="26" width="14" height="14" fill="${actualColor}" />
  <text x="${width/2 + 28}" y="37" font-family="Arial, sans-serif" font-size="11" fill="#333">Actual</text>
  
  <!-- Y axis -->
  <line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${marginTop + chartHeight}" stroke="#333" stroke-width="1" />
  <text x="${15}" y="${marginTop + chartHeight/2}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#333" transform="rotate(-90, 15, ${marginTop + chartHeight/2})">Hours</text>
  ${yAxisHtml}
  
  <!-- X axis -->
  <line x1="${marginLeft}" y1="${marginTop + chartHeight}" x2="${marginLeft + chartWidth}" y2="${marginTop + chartHeight}" stroke="#333" stroke-width="1" />
  <text x="${marginLeft + chartWidth/2}" y="${height - 10}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#333">Month</text>
  
  <!-- Bars -->
  ${barsHtml}
</svg>`.trim();
}

/**
 * Build SVG vertical stacked bar chart for Allocated vs Actual Hours by Project
 * @param {Array} projectsData - Array of { label: string, used: number, unused: number }
 * @returns {string} SVG markup string
 */
function buildProjectsBarChartSVG(projectsData) {
  const width = 700;
  const height = 320;
  const marginTop = 60;
  const marginRight = 30;
  const marginBottom = 80;
  const marginLeft = 60;
  
  const chartWidth = width - marginLeft - marginRight;
  const chartHeight = height - marginTop - marginBottom;
  
  if (!projectsData || projectsData.length === 0) {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
      <text x="${width/2}" y="${height/2}" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" fill="#666">No project data available</text>
    </svg>`;
  }
  
  // Calculate max value for Y axis (total of used + unused)
  const maxVal = Math.max(...projectsData.map(p => p.used + p.unused), 1);
  const yMax = Math.ceil(maxVal / 100) * 100 || 100;
  
  // Bar dimensions
  const barWidth = Math.min(60, (chartWidth / projectsData.length) * 0.7);
  const barSpacing = chartWidth / projectsData.length;
  
  // Colors
  const usedColor = '#4F7F2D';     // Green for used/actual
  const unusedColor = '#E6E6E6';   // Light gray for unused/remaining
  
  // Format number for display
  const formatNum = (v) => Math.round(v).toLocaleString('en-US');
  
  // Build Y axis ticks
  const yTicks = 5;
  const yTickStep = yMax / yTicks;
  let yAxisHtml = '';
  for (let i = 0; i <= yTicks; i++) {
    const val = i * yTickStep;
    const y = marginTop + chartHeight - (val / yMax) * chartHeight;
    yAxisHtml += `
      <line x1="${marginLeft}" y1="${y}" x2="${marginLeft + chartWidth}" y2="${y}" stroke="#e0e0e0" stroke-width="1" />
      <text x="${marginLeft - 8}" y="${y + 4}" text-anchor="end" font-family="Arial, sans-serif" font-size="10" fill="#666">${formatNum(val)}</text>`;
  }
  
  // Build stacked bars
  let barsHtml = '';
  projectsData.forEach((project, i) => {
    const barX = marginLeft + i * barSpacing + (barSpacing - barWidth) / 2;
    const total = project.used + project.unused;
    const utilPct = total > 0 ? (project.used / total) * 100 : 0;
    
    // Unused bar (bottom - full height)
    const totalHeight = (total / yMax) * chartHeight;
    const totalY = marginTop + chartHeight - totalHeight;
    barsHtml += `
      <rect x="${barX}" y="${totalY}" width="${barWidth}" height="${totalHeight}" fill="${unusedColor}" stroke="#ccc" stroke-width="1" />`;
    
    // Used bar (on top of unused, from bottom)
    const usedHeight = (project.used / yMax) * chartHeight;
    const usedY = marginTop + chartHeight - usedHeight;
    barsHtml += `
      <rect x="${barX}" y="${usedY}" width="${barWidth}" height="${usedHeight}" fill="${usedColor}" stroke="#3d6423" stroke-width="1" />`;
    
    // Value label on top of bar (show percentage instead of total)
    if (total > 0) {
      barsHtml += `
        <text x="${barX + barWidth/2}" y="${totalY - 5}" text-anchor="middle" font-family="Arial, sans-serif" font-size="9" font-weight="bold" fill="#333">${utilPct.toFixed(1)}%</text>`;
    }
    
    // X axis label (rotated for readability)
    const labelX = barX + barWidth / 2;
    const labelY = marginTop + chartHeight + 12;
    barsHtml += `
      <text x="${labelX}" y="${labelY}" text-anchor="start" font-family="Arial, sans-serif" font-size="9" fill="#333" transform="rotate(45, ${labelX}, ${labelY})">${project.label}</text>`;
  });
  
  return `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
  <!-- Title -->
  <text x="${width/2}" y="18" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" font-weight="bold" fill="#333">Allocated vs Actual Effort (To Date by Project)</text>
  
  <!-- Legend (under title) -->
  <rect x="${width/2 - 140}" y="28" width="14" height="14" fill="${usedColor}" stroke="#3d6423" stroke-width="1" />
  <text x="${width/2 - 122}" y="39" font-family="Arial, sans-serif" font-size="10" fill="#333">Used (Actual To Date)</text>
  <rect x="${width/2 + 20}" y="28" width="14" height="14" fill="${unusedColor}" stroke="#ccc" stroke-width="1" />
  <text x="${width/2 + 38}" y="39" font-family="Arial, sans-serif" font-size="10" fill="#333">Remaining (Allocated − Actual)</text>
  
  <!-- Y axis -->
  <line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${marginTop + chartHeight}" stroke="#333" stroke-width="1" />
  <text x="${15}" y="${marginTop + chartHeight/2}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#333" transform="rotate(-90, 15, ${marginTop + chartHeight/2})">Hours</text>
  ${yAxisHtml}
  
  <!-- X axis -->
  <line x1="${marginLeft}" y1="${marginTop + chartHeight}" x2="${marginLeft + chartWidth}" y2="${marginTop + chartHeight}" stroke="#333" stroke-width="1" />
  
  <!-- Bars -->
  ${barsHtml}
</svg>`.trim();
}

/**
 * Render HTML from data (replaces template engine)
 */
function renderPDFHtml(data, options = {}) {
  const { hideUserAllocatedData = false } = options;
  
  // Format number: show 2 decimals only if needed, otherwise show integer
  const formatNum = (v) => {
    const num = v || 0;
    // Check if it's a whole number
    if (num === Math.floor(num)) {
      return Math.floor(num).toLocaleString('en-US');
    }
    // Check if decimals are .00
    const formatted = num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (formatted.endsWith('.00')) {
      return Math.floor(num).toLocaleString('en-US');
    }
    return formatted;
  };
  const formatPct = (v) => formatNum(v) + '%';

  // Helper to render context block HTML
  const renderContextBlock = (contextBlock) => {
    if (!contextBlock) return '';

    let lines = [];

    if (contextBlock.type === 'summary') {
      // Summary context block - emphasize "to date" in description
      if (contextBlock.description) {
        // Ensure description reflects to-date logic
        let desc = contextBlock.description;
        if (!desc.toLowerCase().includes('to date')) {
          desc = desc.replace(/aggregates resource allocation and actual effort/i, 
            'aggregates resource allocation and actual effort <strong>to date</strong>');
        }
        lines.push(`<strong>${desc}</strong>`);
      }
      if (contextBlock.portfolioSpan) {
        lines.push(`Portfolio date span: ${contextBlock.portfolioSpan.start} → ${contextBlock.portfolioSpan.end}`);
      }
      if (contextBlock.reportingPeriod) {
        lines.push(`<span style="background: #e8f5e9; display: inline; border-radius: 3px;"><strong>Reporting period (used for all calculations):</strong> ${contextBlock.reportingPeriod.start} → ${contextBlock.reportingPeriod.end}</span>`);
      }
    } else if (contextBlock.type === 'project') {
      // Project context block
      if (contextBlock.description) {
        lines.push(contextBlock.description);
      }
      if (contextBlock.projectSpan) {
        lines.push(`Project date span: ${contextBlock.projectSpan.start} → ${contextBlock.projectSpan.end}`);
      }
      if (contextBlock.reportingPeriod) {
        lines.push(`<span style="background: #e8f5e9; display: inline; border-radius: 3px;"><strong>Reporting period (used for all calculations):</strong> ${contextBlock.reportingPeriod.start} → ${contextBlock.reportingPeriod.end}</span>`);
      }
    }

    // Add metric definitions
    if (contextBlock.metricDefinitions) {
      const defs = contextBlock.metricDefinitions;
      if (defs.allocated_effort_to_date) {
        lines.push(`<strong>Allocated Effort (To Date):</strong> ${defs.allocated_effort_to_date.replace(/^Allocated Effort \(To Date\) is the /, '').replace(/^Allocated Effort \(To Date\) is /, '')}`);
      }
      if (defs.actual_effort_to_date) {
        lines.push(`<strong>Actual Effort (To Date):</strong> ${defs.actual_effort_to_date.replace(/^Actual Effort \(To Date\) is the /, '').replace(/^Actual Effort \(To Date\) is /, '')}`);
      }
      if (defs.variance_hours) {
        lines.push(`<strong>Variance:</strong> ${defs.variance_hours.replace(/^Variance \(Hours\) = /, '')}`);
      }
      if (defs.effort_utilized_pct) {
        lines.push(`<strong>Effort Utilized (%):</strong> ${defs.effort_utilized_pct.replace(/^Effort Utilized % = /, '')}`);
      }
    }

    if (lines.length === 0) return '';

    return `
      <div class="context-block">
        ${lines.map(line => `<div class="context-line">${line}</div>`).join('')}
      </div>`;
  };

  let sheetsHtml = '';
  
  data.sheets.forEach((sheet, idx) => {
    // Build rows HTML
    let rowsHtml = '';
    sheet.rows.forEach(row => {
      const rowClass = row.isRole ? 'role-row' : 'user-row';
      const isUser = !row.isRole;
      
      // For simple PDF, hide Allocated, Variance, Effort for user rows
      const allocatedValue = (hideUserAllocatedData && isUser) ? '' : formatNum(row.allocatedToDate);
      const varianceValue = (hideUserAllocatedData && isUser) ? '' : formatNum(row.variance);
      const effortValue = (hideUserAllocatedData && isUser) ? '' : formatPct(row.effortPct);
      
      rowsHtml += `
        <tr class="${rowClass}">
          <td>${row.label}</td>
          <td>${allocatedValue}</td>
          <td>${formatNum(row.actualToDate)}</td>
          <td>${varianceValue}</td>
          <td>${effortValue}</td>
        </tr>`;
    });

    // Render context block for this sheet
    const contextBlockHtml = renderContextBlock(sheet.contextBlock);

    // Gauge chart data for this sheet
    const gaugePercent = sheet.grandTotal.effortPct;
    const gaugeAllocated = sheet.grandTotal.allocated;
    const gaugeActual = sheet.grandTotal.actual;
    
    // Build SVG gauge
    const gaugeSvg = buildGaugeSVG(gaugeActual, gaugeAllocated);
    
    // Build SVG bar chart for this sheet
    const barChartSvg = buildBarChartSVG(data.chartsData[idx]);

    // Build SVG projects bar chart for Summary page
    const projectsBarChartSvg = sheet.isSummary ? buildProjectsBarChartSVG(data.projectsBarData) : '';

    sheetsHtml += `
    <div class="sheet-section">
      <div class="sheet-title">${sheet.sheetName}</div>
      ${contextBlockHtml}
      
      <!-- SVG Gauge Chart Section -->
      <div class="gauge-section">
        <div class="gauge-title">Effort Utilized (%) (To Date)</div>
        <div class="gauge-subtitle">Based on allocated and actual effort within the reporting period</div>
        <div class="gauge-wrapper">
          <div class="gauge-svg-container">
            ${gaugeSvg}
          </div>
          <div class="kpi-block">
            <div class="kpi-item">
              <div class="kpi-label">Actual Effort<br/>(To Date)</div>
              <div class="kpi-value">${formatNum(gaugeActual)}</div>
            </div>
            <div class="kpi-item">
              <div class="kpi-label">Allocated Effort<br/>(To Date)</div>
              <div class="kpi-value">${formatNum(gaugeAllocated)}</div>
            </div>
            <div class="kpi-item">
              <div class="kpi-label">Effort Utilized<br/>(%)</div>
              <div class="kpi-value">${gaugePercent.toFixed(1)}%</div>
            </div>
          </div>
        </div>
      </div>
      
      ${sheet.isSummary ? `
      <!-- SVG Projects Bar Chart (Summary Only) -->
      <div class="projects-bar-container">
        ${projectsBarChartSvg}
      </div>
      
      <!-- Projects Summary Table -->
      <table class="projects-table">
        <thead>
          <tr>
            <th style="width: 50%">Project Name</th>
            <th style="width: 17%">Allocated (To Date)</th>
            <th style="width: 17%">Actual (To Date)</th>
            <th style="width: 16%">Effort Utilized (%)</th>
          </tr>
        </thead>
        <tbody>
          ${data.projectsBarData.map(p => `
          <tr>
            <td class="project-name-cell">${p.fullName}</td>
            <td>${formatNum(p.allocated)}</td>
            <td>${formatNum(p.used)}</td>
            <td>${p.utilization.toFixed(1)}%</td>
          </tr>`).join('')}
        </tbody>
      </table>` : ''}
      
      <!-- SVG Bar Chart -->
      <div class="bar-chart-container">
        ${barChartSvg}
      </div>
      
      <table>
        <thead>
          <tr>
            <th style="width: 35%">Role / User</th>
            <th style="width: 16%">Allocated (To Date)</th>
            <th style="width: 16%">Actual (To Date)</th>
            <th style="width: 16%">Variance</th>
            <th style="width: 17%">Effort Utilized (%)</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
          <tr class="grand-total">
            <td>GRAND TOTAL (To Date)</td>
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
  <style>
    @page { margin: 15mm; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Segoe UI', Arial, sans-serif; padding: 0; color: #333; background: #fff; }
    .header { text-align: center; margin-bottom: 0; padding: 60px 40px 40px; border-bottom: none; page-break-after: always; min-height: 100vh; display: flex; flex-direction: column; justify-content: flex-start; align-items: center; background: #fff; }
    .header h1 { font-size: 36px; color: #333; margin-bottom: 10px; font-weight: 700; letter-spacing: 1px; margin-top: 40px; }
    .header .company-name { font-size: 28px; color: #4F7F2D; font-weight: 600; margin-bottom: 20px; }
    .header .subtitle { font-size: 12px; color: #666; margin-bottom: 50px; }
    .toc { width: 100%; max-width: 500px; text-align: left; margin-top: 20px; }
    .toc-title { font-size: 16px; font-weight: 700; color: #333; margin-bottom: 15px; padding-bottom: 8px; border-bottom: 2px solid #4F7F2D; }
    .toc-item { display: flex; justify-content: space-between; align-items: baseline; padding: 8px 0; border-bottom: 1px dotted #ccc; font-size: 11px; }
    .toc-item:last-child { border-bottom: none; }
    .toc-name { color: #333; flex: 1; padding-right: 10px; }
    .toc-page { color: #666; font-weight: 600; white-space: nowrap; }
    .sheet-section { margin-bottom: 20px; }
    .sheet-section:not(:first-of-type) { page-break-before: always; }
    .sheet-title { font-size: 14px; font-weight: bold; color: #333; margin-bottom: 10px; padding: 8px; background: #f5f5f5; border-left: 4px solid #333; }
    .context-block { background: #f9f9f9; border: 1px solid #e0e0e0; border-radius: 4px; padding: 10px 12px; margin-bottom: 15px; font-size: 9pt; line-height: 1.4; color: #444; }
    .context-line { margin-bottom: 3px; }
    .context-line:last-child { margin-bottom: 0; }
    .context-line strong { color: #333; }
    .gauge-section { margin-bottom: 20px; }
    .gauge-title { text-align: center; font-size: 14px; font-weight: 700; color: #333; margin-bottom: 4px; }
    .gauge-subtitle { text-align: center; font-size: 10px; color: #666; margin-bottom: 10px; }
    .gauge-wrapper { display: flex; align-items: center; justify-content: center; gap: 30px; width: 70%; margin-left: auto; margin-right: auto; padding: 10px 30px; }
    .gauge-svg-container { flex: 1 1 auto; max-width: 320px; }
    .gauge-svg-container svg { width: 100%; height: auto; }
    .kpi-block { width: 180px; text-align: right; }
    .kpi-item { margin-bottom: 12px; }
    .kpi-item:last-child { margin-bottom: 0; }
    .kpi-label { font-size: 10px; line-height: 1.3; color: #355B1B; }
    .kpi-value { font-size: 22px; font-weight: 700; color: #333; margin-top: 2px; }
    .projects-bar-container { width: 100%; margin-bottom: 15px; }
    .projects-bar-container svg { width: 100%; height: auto; }
    .projects-table { margin-bottom: 20px; }
    .projects-table .project-name-cell { text-align: left; line-height: 1.3; word-wrap: break-word; }
    .bar-chart-container { width: 100%; margin-bottom: 20px; }
    .bar-chart-container svg { width: 100%; height: auto; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 10px; }
    th { background: #333; color: #fff; font-weight: bold; padding: 8px 6px; text-align: center; border: 1px solid #333; }
    th:first-child { text-align: left; }
    td { padding: 6px; border: 1px solid #ccc; text-align: center; }
    td:first-child { text-align: left; }
    tr.role-row { background: #e0e0e0; font-weight: bold; }
    tr.user-row td:first-child { padding-left: 15px; }
    tr.grand-total { background: #333; color: #fff; font-weight: bold; }
    tr.grand-total td { border-color: #333; }
    .footer { margin-top: 20px; padding-top: 10px; border-top: 1px solid #ccc; font-size: 9px; color: #999; text-align: center; }
  </style>
</head>
<body>
  <div class="header">
    <h1>Resource Allocation Report</h1>
    <div class="company-name">Sanofi</div>
    <div class="subtitle">Generated on ${data.generatedOn}</div>
    
    <div class="toc">
      <div class="toc-title">Table of Contents</div>
      ${data.sheets.map((sheet, idx) => `
      <div class="toc-item">
        <span class="toc-name">${sheet.sheetName}</span>
        <span class="toc-page">Page ${idx + 2}</span>
      </div>`).join('')}
    </div>
  </div>

  ${sheetsHtml}

  <div class="footer">
    Resource Management Workbench Report • Confidential • All metrics calculated using daily allocation data within the reporting period
  </div>
</body>
</html>`;
}

/**
 * Generate PDF from payload
 * @param {Object} payload - The resource allocation payload
 * @param {Object} options - Generation options
 * @param {boolean} options.hideUserAllocatedData - If true, hide Allocated, Variance, Effort for user rows
 * @returns {Buffer} PDF buffer
 */
async function generatePDF(payload, options = {}) {
  // Process payload
  const data = processPayloadForPDF(payload);

  // Render HTML directly (no template file needed now)
  const html = renderPDFHtml(data, options);

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
    margin: { top: '15mm', right: '15mm', bottom: '15mm', left: '15mm' }
  });

  await browser.close();

  return pdfBuffer;
}

// PDF Generation Endpoint (Extended - shows all data)
app.post('/api/generate-pdf-extended', requireApiKey, async (req, res) => {
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

    // Validate required to-date effort fields
    validatePayloadFields(payload);

    const pdfBuffer = await generatePDF(payload);

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const filename = `Resource_Allocation_Extended_${timestamp}.pdf`;

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

// PDF Generation Endpoint (Simple - hides user-level Allocated, Variance, Effort columns)
app.post('/api/generate-pdf', requireApiKey, async (req, res) => {
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

    // Validate required to-date effort fields
    validatePayloadFields(payload);

    // Pass option to hide user-level data
    const pdfBuffer = await generatePDF(payload, { hideUserAllocatedData: true });

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
  console.log(`🔐 API Key auth enabled (set API_KEY env var, current: ${API_KEY === 'dev-key-change-me' ? 'using default dev key' : 'custom key set'})`);
  console.log(`📋 POST /api/generate-excel - Generate Excel from JSON payload`);
  console.log(`📄 POST /api/generate-pdf - Generate PDF report (simplified: user allocation data hidden)`);
  console.log(`📄 POST /api/generate-pdf-extended - Generate PDF report (full data for all rows)`);
  console.log(`💊 GET /health - Health check (no auth required)`);
});
