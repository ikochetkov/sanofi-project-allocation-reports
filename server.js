const express = require('express');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

const app = express();

// API Key for authentication (set via environment variable)
const API_KEY = process.env.API_KEY || 'dev-key-change-me';

// Mobiz Brand Colors
const MOBIZ_COLORS = {
  primary: '#D8242A',      // Red accent - headings, borders, accents
  secondary: '#613BFE',    // Purple - H3, secondary accents
  dark: '#130E23',         // Primary text, body emphasis
  darkAlt: '#1C1631',      // Secondary dark (table headers)
  grayLight: '#F5F5F5',    // Subtle backgrounds
  grayMedium: '#666666'    // Secondary text
};

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
        font-family="Arial, sans-serif" font-size="36" font-weight="700" fill="#130E23">
    ${pct.toFixed(1)}%
  </text>

  <!-- Left tick label (0%) -->
  <text x="${strokeWidth * 0.5}" y="${height + 18}" text-anchor="start"
        font-family="Arial, sans-serif" font-size="14" font-weight="600" fill="#130E23">0%</text>

  <!-- Right tick label (100%) -->
  <text x="${width - strokeWidth * 0.5}" y="${height + 18}" text-anchor="end"
        font-family="Arial, sans-serif" font-size="14" font-weight="600" fill="#130E23">100%</text>
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
  
  // Mobiz brand colors for bar chart
  const allocatedColor = '#1C1631';  // Mobiz dark alt
  const actualColor = '#4CAF50';     // Keep green for actual (intuitive)
  
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
      <text x="${groupX + barWidth/2}" y="${allocY - 5}" text-anchor="middle" font-family="Arial, sans-serif" font-size="9" font-weight="bold" fill="#130E23">${allocated[i] > 0 ? formatNum(allocated[i]) : ''}</text>`;
    
    // Actual bar
    const actHeight = (actual[i] / yMax) * chartHeight;
    const actY = marginTop + chartHeight - actHeight;
    const actX = groupX + barWidth + barGap;
    barsHtml += `
      <rect x="${actX}" y="${actY}" width="${barWidth}" height="${actHeight}" fill="${actualColor}" />
      <text x="${actX + barWidth/2}" y="${actY - 5}" text-anchor="middle" font-family="Arial, sans-serif" font-size="9" font-weight="bold" fill="#130E23">${actual[i] > 0 ? formatNum(actual[i]) : ''}</text>`;
    
    // X axis label
    const labelX = groupX + barWidth + barGap/2;
    barsHtml += `
      <text x="${labelX}" y="${marginTop + chartHeight + 20}" text-anchor="middle" font-family="Arial, sans-serif" font-size="10" fill="#130E23">${label}</text>`;
  });
  
  return `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
  <!-- Title -->
  <text x="${width/2}" y="18" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" font-weight="bold" fill="#130E23">Allocated vs Actual Effort (To Date by Month)</text>

  <!-- Legend (under title) -->
  <rect x="${width/2 - 90}" y="26" width="14" height="14" fill="${allocatedColor}" />
  <text x="${width/2 - 72}" y="37" font-family="Arial, sans-serif" font-size="11" fill="#130E23">Allocated</text>
  <rect x="${width/2 + 10}" y="26" width="14" height="14" fill="${actualColor}" />
  <text x="${width/2 + 28}" y="37" font-family="Arial, sans-serif" font-size="11" fill="#130E23">Actual</text>

  <!-- Y axis -->
  <line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${marginTop + chartHeight}" stroke="#130E23" stroke-width="1" />
  <text x="${15}" y="${marginTop + chartHeight/2}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#130E23" transform="rotate(-90, 15, ${marginTop + chartHeight/2})">Hours</text>
  ${yAxisHtml}

  <!-- X axis -->
  <line x1="${marginLeft}" y1="${marginTop + chartHeight}" x2="${marginLeft + chartWidth}" y2="${marginTop + chartHeight}" stroke="#130E23" stroke-width="1" />
  <text x="${marginLeft + chartWidth/2}" y="${height - 10}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#130E23">Month</text>
  
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
        <text x="${barX + barWidth/2}" y="${totalY - 5}" text-anchor="middle" font-family="Arial, sans-serif" font-size="9" font-weight="bold" fill="#130E23">${utilPct.toFixed(1)}%</text>`;
    }

    // X axis label (rotated for readability)
    const labelX = barX + barWidth / 2;
    const labelY = marginTop + chartHeight + 12;
    barsHtml += `
      <text x="${labelX}" y="${labelY}" text-anchor="start" font-family="Arial, sans-serif" font-size="9" fill="#130E23" transform="rotate(45, ${labelX}, ${labelY})">${project.label}</text>`;
  });
  
  return `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
  <!-- Title -->
  <text x="${width/2}" y="18" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" font-weight="bold" fill="#130E23">Allocated vs Actual Effort (To Date by Project)</text>

  <!-- Legend (under title) -->
  <rect x="${width/2 - 140}" y="28" width="14" height="14" fill="${usedColor}" stroke="#3d6423" stroke-width="1" />
  <text x="${width/2 - 122}" y="39" font-family="Arial, sans-serif" font-size="10" fill="#130E23">Used (Actual To Date)</text>
  <rect x="${width/2 + 20}" y="28" width="14" height="14" fill="${unusedColor}" stroke="#ccc" stroke-width="1" />
  <text x="${width/2 + 38}" y="39" font-family="Arial, sans-serif" font-size="10" fill="#130E23">Remaining (Allocated − Actual)</text>

  <!-- Y axis -->
  <line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${marginTop + chartHeight}" stroke="#130E23" stroke-width="1" />
  <text x="${15}" y="${marginTop + chartHeight/2}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#130E23" transform="rotate(-90, 15, ${marginTop + chartHeight/2})">Hours</text>
  ${yAxisHtml}

  <!-- X axis -->
  <line x1="${marginLeft}" y1="${marginTop + chartHeight}" x2="${marginLeft + chartWidth}" y2="${marginTop + chartHeight}" stroke="#130E23" stroke-width="1" />

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

  // Mobiz logo (base64 encoded)
  const mobizLogo = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAACC8AAAPOCAMAAAASylBhAAAAM1BMVEVMaXHZLDPfSlA6Nzj01tfla3Bxbm/hwsPogYXlnqHYIyojHyDneHzeQkjjX2VNSkuSkJA096T4AAAACnRSTlMAxKW5EoSAKm9Qmj5P5AAAAAlwSFlzAAAuIwAALiMBeKU/dgAAIABJREFUeJzt3et246oWYGFLQUHoGCXv/7RnyE7tSqqSMgYWWgvm1/2rR59dieLLFOJyuQCAGUsmf/YPjiwh78/N1QaAwU1bnnD2D44sa9Zfe+JqA8Dg6IWx0AsAgBz0wljoBQBADnphLPQCAOB5zuf2wsqER4v8kjd/wbuzf3IAwInCNTMXtn3lD2ePu8a8P/eV+a0AMLKCXrie/bPjefQCACAHvTAWz/gCACADvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWegEAkINeGAu9AADIQS+MhV4AAOSgF8ZCLwAActALY6EXAAA56IWx0AsAgBz0wljoBQBADnphLPQCACAHvTAWd41blms4+0cHAJyIXhiL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';

  // Sanofi logo (base64 encoded SVG)
  const sanofiLogo = 'data:image/svg+xml;base64,PHN2ZyB2aWV3Qm94PSItMC4wNzI1OTQ4NjYxMTYxOTI1OSAtMC4yNDQ1NjA2MTc5ODIyOTQ1NiAyODMuMDI1MTAwMTk1NzkyODUgNzMuNDA0NjcyMDY0OTk2NDUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgd2lkdGg9IjI1MDAiIGhlaWdodD0iNjUwIj48cGF0aCBkPSJNMTAzIDQ2djE4YTEzLjU2IDEzLjU2IDAgMCAxLS4wNyAxLjQzIDMuNzcgMy43NyAwIDAgMS0yLjQ1IDMuMjUgNTEuMzIgNTEuMzIgMCAwIDEtMTIuNTcgMy41OCA1Mi4xNSA1Mi4xNSAwIDAgMS0xMC4yMi40OCAyOC4yNSAyOC4yNSAwIDAgMS05Ljg0LTIuMTggMjEuMSAyMS4xIDAgMCAxLTEyLTEyLjcxIDMyLjQ0IDMyLjQ0IDAgMCAxLS41Ni0yMS4wN2MyLjgxLTkgOS0xNC4zNiAxOC4wOS0xNi41M2EzNC44IDM0LjggMCAwIDEgMTAuMDgtLjY2IDU1LjMgNTUuMyAwIDAgMSAxNi4zNiAzLjQ5YzIuNC45MiAzLjE4IDEuOTIgMy4xOCA0LjUzVjQ2em0tMTMuMzQuMjFWMzUuODNhMi42IDIuNiAwIDAgMC0yLTIuODdjLS4zMy0uMTMtLjY4LS4yMS0xLS4zMWEyMy4zNyAyMy4zNyAwIDAgMC02LjQ5LS43OWMtNi4yMi4wNS0xMC41MyAzLjI1LTEyIDlhMjAuMzEgMjAuMzEgMCAwIDAgLjE1IDExLjA3IDExLjA1IDExLjA1IDAgMCAwIDkgOC4yNCAyMC45IDIwLjkgMCAwIDAgMTAuNDYtLjg1YzEuNDktLjQ4IDItMS4yNyAyLTIuODlxLS4xNC01LjEyLS4xNS0xMC4yNXpNMjIzLjQxIDQ2LjE5YTI4LjkyIDI4LjkyIDAgMCAxLTMuMTUgMTMuNjljLTMuODcgNy4zMi0xMCAxMS40NC0xOC4xNCAxMi42MWEyOSAyOSAwIDAgMS0xMy4wOC0xIDIzLjEgMjMuMSAwIDAgMS0xNS41OS0xNS4zNyAzMC4zNCAzMC4zNCAwIDAgMSAuMjMtMjAuNTlDMTc3LjEzIDI2LjIxIDE4NCAyMSAxOTMuOCAxOS43MmEyOCAyOCAwIDAgMSAxMy4yIDEuMzNjOC40IDMgMTMuNDMgOSAxNS41OSAxNy41OGEzMSAzMSAwIDAgMSAuODIgNy41NnptLTI1LjkzLTE0LjQ2YTExLjE1IDExLjE1IDAgMCAwLTEwLjkxIDcuNDQgMTguNDYgMTguNDYgMCAwIDAtMS4xMiA1LjU5IDE5LjEzIDE5LjEzIDAgMCAwIDEuMzIgOC43NiAxMC42OSAxMC42OSAwIDAgMCA3LjM0IDYuNDggMTYgMTYgMCAwIDAgMi41OC40NCAxMSAxMSAwIDAgMCAxMi4wNS03LjQ2IDE5Ljc4IDE5Ljc4IDAgMCAwIDAtMTMuNjcgMTAuNTYgMTAuNTYgMCAwIDAtNy43LTdjLTEuMTktLjMxLTIuMzktLjQtMy41Ni0uNTh6TTExNS40MiA0OC4xOFYyNy45MXYtMS4yYTMuODQgMy44NCAwIDAgMSAyLjY4LTMuNTggNjQuMDYgNjQuMDYgMCAwIDEgMTUuNjctMy40MiA0OS40MSA0OS40MSAwIDAgMSAxMS41My4yNyAyNS42NiAyNS42NiAwIDAgMSA3LjE4IDIuMSAxNS43IDE1LjcgMCAwIDEgOS4xOSAxMi4zNyAyNSAyNSAwIDAgMSAuMzMgNC4xNVY2OS42MWMwIDEuNDQtLjY5IDIuMTQtMi4xMSAyLjE1aC05LjA3YTEuOTEgMS45MSAwIDAgMS0yLjEyLTIuMDV2LTEuMTktMjcuNDRhMTUuODMgMTUuODMgMCAwIDAtLjMxLTMuMDcgNi44NSA2Ljg1IDAgMCAwLTQuMjMtNS4xMSAxMi41MyAxMi41MyAwIDAgMC0zLjU1LS45NSAyNi41MiAyNi41MiAwIDAgMC05LjIyLjY3IDYuNTUgNi41NSAwIDAgMC0xLjM1LjQ3IDIuMSAyLjEgMCAwIDAtMS4yNSAxLjkxdjM0LjcyYTIgMiAwIDAgMS0yLjA5IDIuMDloLTEuNjdjLTIuMTkgMC00LjM4LS4wNy02LjU2IDBzLTMuMDctLjY3LTMtM2MtLjAxLTYuODctLjA1LTEzLjc1LS4wNS0yMC42M3pNMzkuMTYgNzEuNzhoLTQuNDFjLTEuNjQgMC0yLjExLS42OS0xLjc0LTIuMjhhMTguMjMgMTguMjMgMCAwIDAgLjUzLTMuMjggNy43NCA3Ljc0IDAgMCAwLTMuNTItNyAyMC4wNiAyMC4wNiAwIDAgMC01LjY4LTIuNzFjLTMuMjktMS02LjYtMi05Ljg5LTMuMDlhMjkuNjkgMjkuNjkgMCAwIDEtNy41NC0zLjUyIDEzLjc0IDEzLjc0IDAgMCAxLTYuMy0xMS40M0MuMjQgMzAuMTMgNC43NCAyMy45IDEyLjIxIDIxQTI5LjkgMjkuOSAwIDAgMSAyNCAxOS4wOWE0My4xOSA0My4xOSAwIDAgMSAxNy42NCA0Yy43OS4zNyAxLjU3Ljc3IDIuMzMgMS4yQTIuNjIgMi42MiAwIDAgMSA0NS4yOSAyOGMtLjI4LjctLjYyIDEuMzctLjk0IDIuMDYtLjY3IDEuNDQtMS40MSAyLjg1LTIgNC4zMi0uOTMgMi4yOS0yLjgzIDIuNTEtNC41NCAxLjUzYTM0LjkzIDM0LjkzIDAgMCAwLTguMjItMy40NUEyMS4xNyAyMS4xNyAwIDAgMCAxOSAzMS45M2E5LjM4IDkuMzggMCAwIDAtMi45MyAxLjI0IDMuOSAzLjkgMCAwIDAtLjE5IDYuNjIgMTMuODIgMTMuODIgMCAwIDAgMy42OCAxLjljMy4wOCAxLjA1IDYuMjIgMS45MiA5LjM0IDIuODdhNDUuODQgNDUuODQgMCAwIDEgOC41MiAzLjM1IDIxLjY5IDIxLjY5IDAgMCAxIDQuNzQgMy4zOSAxNS45MiAxNS45MiAwIDAgMSA0Ljc4IDkuNyAyMiAyMiAwIDAgMS0uMjUgNy4yNEExNi43IDE2LjcgMCAwIDEgNDYuMiA3MGEyLjM3IDIuMzcgMCAwIDEtMi41MSAxLjc5aC00LjUzek0yNDYuNzMgMjAuNTFoMTIuNThjMS42NCAwIDIuMjguNjYgMi4zMiAyLjMzVjI5LjI4YzAgMS45My0uNTggMi41Mi0yLjQ4IDIuNTNoLTEyLjU0YzAgLjU4LS4wNSAxLjA4LS4wNSAxLjU4djM2LjM3YTEuOTMgMS45MyAwIDAgMS0yLjA2IDJoLTguNThjLTEuOTUgMC0yLjU4LS42NC0yLjU4LTIuNTdWMzQuODd2LTE0YTI4LjE1IDI4LjE1IDAgMCAxIC45MS03LjA3YzEuODktNy4xMyA2LjQ4LTExLjUgMTMuNjQtMTMuMTNhMjkgMjkgMCAwIDEgMTIuNTQgMGwuNDYuMTFjMS43Ni40NiAyLjE5IDEuMTYgMS44NSAzbC0xLjA1IDUuNXYuMTFjLS4zNyAxLjgtMSAyLjI1LTIuODQgMmEyMS40NSAyMS40NSAwIDAgMC0yLjI0LS4yOSAyNi40MSAyNi40MSAwIDAgMC00LjI3LjA5IDYuMTIgNi4xMiAwIDAgMC01LjQyIDUuMTEgMTEuNTcgMTEuNTcgMCAwIDAtLjE5IDQuMjF6TTI4Mi41MSA0Ni4xN3YyM2MwIDItLjYgMi41OS0yLjY2IDIuNmgtOC4zNGMtMS42MiAwLTIuMy0uNzEtMi4zNi0yLjMzVjIzLjY1di0xLjA4YTIgMiAwIDAgMSAyLjE1LTIuMDVoOC44MmMxLjcyIDAgMi4zNy42NiAyLjM3IDIuNHEuMDMgMTEuNjMuMDIgMjMuMjV6Ii8+PGcgZmlsbD0iIzg4MDBlZiI+PHBhdGggZD0iTTE0IDY1LjcxYy0uMDkuNjMtLjE1IDEuMjYtLjI3IDEuODhhNi4wOCA2LjA4IDAgMCAxLTUuODYgNSAxMCAxMCAwIDAgMS0yLjcyLS4yQTYuMDcgNi4wNyAwIDAgMSAwIDY2LjY1YTExIDExIDAgMCAxIC4yNS0zLjE5IDUuODUgNS44NSAwIDAgMSA0Ljg0LTQuNTMgOC42OSA4LjY5IDAgMCAxIDQuMTIuMTRjMy4wNS44MSA0LjcyIDMuMTkgNC43OSA2LjY0ek0yNzYgMTQuMDhhMTYuODEgMTYuODEgMCAwIDEtMi4xMi0uMjYgNi4xIDYuMSAwIDAgMS00LjkxLTYuMDYgMTEuNjIgMTEuNjIgMCAwIDEgLjI2LTIuOTUgNS44OCA1Ljg4IDAgMCAxIDUuMi00LjUyIDExIDExIDAgMCAxIDMuMDguMDggNi4xNSA2LjE1IDAgMCAxIDUuMzEgNi4zNSAxMC4zMiAxMC4zMiAwIDAgMS0uMzQgMi44MmMtLjU4IDIuMjktMi42NSA0LjU5LTYuNDggNC41NHoiLz48L2c+PC9zdmc+';

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Resource Allocation Report</title>
  <style>
    @page { margin: 15mm; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: Helvetica, Arial, sans-serif; padding: 0; color: #130E23; background: #fff; }
    .header { text-align: center; margin-bottom: 0; padding: 60px 40px 40px; border-bottom: none; page-break-after: always; min-height: 100vh; display: flex; flex-direction: column; justify-content: flex-start; align-items: center; background: #fff; position: relative; }
    .header .logo { width: 180px; margin-bottom: 30px; margin-top: 40px; }
    .header .red-line { width: 60%; height: 3px; background: #D8242A; margin-bottom: 40px; }
    .header h1 { font-size: 32px; color: #130E23; margin-bottom: 15px; font-weight: 700; letter-spacing: 1px; }
    .header .prepared-for { font-size: 14px; color: #666; margin-bottom: 8px; }
    .header .company-logo { width: 120px; margin-bottom: 20px; }
    .header .subtitle { font-size: 12px; color: #666; margin-bottom: 50px; }
    .header .confidential-footer { position: absolute; bottom: 30px; left: 0; right: 0; text-align: center; font-size: 9px; color: #D8242A; font-weight: 600; }
    .toc { width: 100%; max-width: 500px; text-align: left; margin-top: 20px; }
    .toc-title { font-size: 16px; font-weight: 700; color: #130E23; margin-bottom: 15px; padding-bottom: 8px; border-bottom: 2px solid #D8242A; }
    .toc-item { padding: 8px 0; padding-left: 12px; border-left: 3px solid #D8242A; border-bottom: 1px dotted #ccc; font-size: 11px; margin-bottom: 2px; }
    .toc-item:last-child { border-bottom: none; }
    .toc-name { color: #130E23; }
    .sheet-section { margin-bottom: 20px; }
    .sheet-section:not(:first-of-type) { page-break-before: always; }
    .sheet-title { font-size: 14px; font-weight: bold; color: #fff; margin-bottom: 10px; padding: 10px 12px; background: #130E23; border-left: 4px solid #D8242A; }
    .context-block { background: #F5F5F5; border: 1px solid #e0e0e0; border-radius: 4px; padding: 10px 12px; margin-bottom: 15px; font-size: 9pt; line-height: 1.4; color: #444; }
    .context-line { margin-bottom: 3px; }
    .context-line:last-child { margin-bottom: 0; }
    .context-line strong { color: #130E23; }
    .gauge-section { margin-bottom: 20px; }
    .gauge-title { text-align: center; font-size: 14px; font-weight: 700; color: #130E23; margin-bottom: 4px; }
    .gauge-subtitle { text-align: center; font-size: 10px; color: #666; margin-bottom: 10px; }
    .gauge-wrapper { display: flex; align-items: center; justify-content: center; gap: 30px; width: 70%; margin-left: auto; margin-right: auto; padding: 10px 30px; }
    .gauge-svg-container { flex: 1 1 auto; max-width: 320px; }
    .gauge-svg-container svg { width: 100%; height: auto; }
    .kpi-block { width: 180px; text-align: right; }
    .kpi-item { margin-bottom: 12px; }
    .kpi-item:last-child { margin-bottom: 0; }
    .kpi-label { font-size: 10px; line-height: 1.3; color: #D8242A; }
    .kpi-value { font-size: 22px; font-weight: 700; color: #130E23; margin-top: 2px; }
    .projects-bar-container { width: 100%; margin-bottom: 15px; }
    .projects-bar-container svg { width: 100%; height: auto; }
    .projects-table { margin-bottom: 20px; }
    .projects-table .project-name-cell { text-align: left; line-height: 1.3; word-wrap: break-word; }
    .bar-chart-container { width: 100%; margin-bottom: 20px; }
    .bar-chart-container svg { width: 100%; height: auto; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 10px; }
    th { background: #130E23; color: #fff; font-weight: bold; padding: 8px 6px; text-align: center; border: 1px solid #130E23; }
    th:first-child { text-align: left; }
    td { padding: 6px; border: 1px solid #ccc; text-align: center; }
    td:first-child { text-align: left; }
    tr.role-row { background: #F2F2F2; font-weight: bold; }
    tr.user-row td:first-child { padding-left: 15px; }
    tr.user-row:nth-child(even) { background: #F7F7F7; }
    tr.grand-total { background: #130E23; color: #fff; font-weight: bold; }
    tr.grand-total td { border-color: #130E23; }
    .footer { margin-top: 20px; padding-top: 10px; border-top: 2px solid #D8242A; font-size: 9px; text-align: center; }
    .footer .confidential { color: #D8242A; font-weight: 600; }
    .footer .info { color: #666; margin-top: 4px; }
  </style>
</head>
<body>
  <div class="header">
    <img src="${mobizLogo}" alt="Mobiz" class="logo">
    <div class="red-line"></div>
    <h1>Resource Allocation Report</h1>
    <div class="prepared-for">Prepared for</div>
    <img src="${sanofiLogo}" alt="Sanofi" class="company-logo">
    <div class="subtitle">Generated on ${data.generatedOn}</div>

    <div class="toc">
      <div class="toc-title">Table of Contents</div>
      ${data.sheets.map((sheet, idx) => `
      <div class="toc-item">
        <span class="toc-name">${sheet.sheetName}</span>
      </div>`).join('')}
    </div>

    <div class="confidential-footer">CONFIDENTIAL - Mobiz Inc.</div>
  </div>

  ${sheetsHtml}

  <div class="footer">
    <div class="confidential">CONFIDENTIAL - Mobiz Inc.</div>
    <div class="info">Resource Management Workbench Report • All metrics calculated using daily allocation data within the reporting period</div>
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
