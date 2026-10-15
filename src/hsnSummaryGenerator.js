const ExcelJS = require('exceljs');
const xlsx = require('xlsx');
const { checkSequenceGaps } = require('./auditEngine');

/**
 * Natural comparison for sorting HSN codes
 */
function naturalCompare(a, b) {
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
}

/**
 * Normalizes an HSN code applying business rules:
 * Rule: every HSN starting with 9987 followed by two digits should be grouped under the same HSN (e.g. 998719)
 */
function normalizeHsnCode(rawHsn) {
  if (!rawHsn) return 'Unspecified';
  const clean = String(rawHsn).trim();

  // If HSN starts with 9987 followed by digits (e.g. 9987, 998711, 998719, 998799)
  if (/^9987\d*/.test(clean)) {
    return '998719';
  }

  // Remove trailing .0 if Excel parsed integer as float string
  return clean.replace(/\.0+$/, '');
}

/**
 * Determines tax category based on SEZ / NIL remarks and tax amounts
 * - SEZ / NIL in remarks or 0 taxes -> "Zero GST"
 * - Intra-state or inter-state with tax -> "Taxable (SGST / CGST / IGST)"
 */
function determineTaxCategory(remarks, sgst, cgst, igst) {
  const rem = String(remarks || '').trim().toLowerCase();
  const hasSezNil = /sez|nil|zero/i.test(rem);
  const totalTax = (Number(sgst) || 0) + (Number(cgst) || 0) + (Number(igst) || 0);

  if (hasSezNil || totalTax === 0) {
    return 'Zero GST';
  }
  return 'Taxable (SGST / CGST / IGST)';
}

/**
 * Parses a cumulative sales sheet (generated from Feature 1 and possibly edited by the user)
 * and groups the data into HSN-wise GST Summary matching Image specifications.
 * 
 * @param {Buffer} buffer - Excel workbook buffer (.xlsx or .xls)
 * @param {Object} options - { periodTitleOverride }
 * @returns {Object} { periodTitle, groups, summary, rawRowCount }
 */
function parseAndGroupCumulativeSheet(buffer, options = {}) {
  const wb = xlsx.read(buffer, { type: 'buffer', cellDates: true, raw: false });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];

  if (!sheet) {
    throw new Error('No valid worksheet found in the uploaded file.');
  }

  // Extract sheet rows as 2D array
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (!rows || rows.length === 0) {
    throw new Error('The uploaded sheet is completely empty.');
  }

  // Extract period title from Row 1 or 2 if available
  let detectedPeriodTitle = '';
  for (let r = 0; r < Math.min(5, rows.length); r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const val = String(row[c] || '').trim();
      // Match period like "April-June 2025-26" or "July-September 2026" or "OUT INV | ..."
      if (/(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|\d{4}-\d{2}|\d{4})/i.test(val)) {
        if (!/^(hsn|basic|qty|sgst|cgst|igst|invoice)/i.test(val)) {
          detectedPeriodTitle = val.replace(/^OUT\s*INV\s*\|\s*/i, '').replace(/^OUT\s*INV\s*/i, '').trim();
          break;
        }
      }
    }
    if (detectedPeriodTitle) break;
  }

  const periodTitle = options.periodTitleOverride || detectedPeriodTitle || 'July–September 2026';

  // Find header row (the row containing HSN, BASIC AMOUNT, etc.)
  let headerRowIndex = -1;
  let colHsn = -1;
  let colQty = -1;
  let colBasic = -1;
  let colSgst = -1;
  let colCgst = -1;
  let colIgst = -1;
  let colTotal = -1;
  let colRemarks = -1;
  let colInvoiceNo = -1;
  let colCompany = -1;

  for (let r = 0; r < Math.min(10, rows.length); r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cellText = String(row[c] || '').trim().toLowerCase();
      if (/hsn/i.test(cellText)) colHsn = c;
      if (/qty|quantity/i.test(cellText)) colQty = c;
      if (/basic/i.test(cellText)) colBasic = c;
      if (/sgst/i.test(cellText)) colSgst = c;
      if (/cgst/i.test(cellText)) colCgst = c;
      if (/igst/i.test(cellText)) colIgst = c;
      if (/total\s*(amount|incl)/i.test(cellText)) colTotal = c;
      if (/remark|description/i.test(cellText)) colRemarks = c;
      if (/invoice\s*no/i.test(cellText)) colInvoiceNo = c;
      if (/company|buyer|customer/i.test(cellText)) colCompany = c;
    }

    if (colHsn !== -1 && (colBasic !== -1 || colQty !== -1)) {
      headerRowIndex = r;
      break;
    }
  }

  // Fallback default column indices matching Image 3 structure
  if (headerRowIndex === -1) {
    headerRowIndex = 1; // row index 1 (Excel row 2)
  }
  if (colInvoiceNo === -1) colInvoiceNo = 0;
  if (colCompany === -1) colCompany = 2;
  if (colHsn === -1) colHsn = 4;
  if (colQty === -1) colQty = 5;
  if (colBasic === -1) colBasic = 6;
  if (colSgst === -1) colSgst = 7;
  if (colCgst === -1) colCgst = 8;
  if (colIgst === -1) colIgst = 9;
  if (colTotal === -1) colTotal = 10;
  if (colRemarks === -1) colRemarks = 11;

  const groupsMap = new Map();
  let rawRowCount = 0;

  // Audit tracking arrays for Feature 2
  const missingData = [];
  const taxInconsistencies = [];
  const hsnConsolidations = [];
  const excludedSummaryRows = [];
  const seenInvoices = new Set();
  const distinctInvoices = [];

  let lastSeenInvoiceNo = '';
  let lastSeenCompany = '';

  const parseNum = (val) => {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    if (!val) return 0;
    const cleaned = String(val).replace(/[^\d.-]/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
  };

  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (row.length === 0) continue;

    const rowText = row.map(c => String(c || '').trim()).join(' ').toLowerCase();
    if (!rowText) continue;

    const excelRowNum = r + 1;
    let invNoText = String(row[colInvoiceNo] || '').trim();
    let companyText = String(row[colCompany] || '').trim();
    let rawHsnText = String(row[colHsn] || '').trim();

    // Context inheritance for blank repeated header cells
    if (invNoText) {
      lastSeenInvoiceNo = invNoText;
      if (!seenInvoices.has(invNoText) && !/^(total|grand\s*total)$/i.test(invNoText)) {
        seenInvoices.add(invNoText);
        distinctInvoices.push({ invoiceNo: invNoText });
      }
    } else {
      invNoText = lastSeenInvoiceNo;
    }

    if (companyText) {
      lastSeenCompany = companyText;
    } else {
      companyText = lastSeenCompany;
    }

    const qty = parseNum(row[colQty]);
    const basicAmount = parseNum(row[colBasic]);
    const sgst = parseNum(row[colSgst]);
    const cgst = parseNum(row[colCgst]);
    const igst = parseNum(row[colIgst]);
    const remarks = String(row[colRemarks] || '').trim();

    // 1. RULE: Original invoice TOTAL / GRAND TOTAL row excluded
    const isTotalRow =
      /^total$/i.test(companyText) ||
      /^total$/i.test(invNoText) ||
      /^grand\s*total$/i.test(companyText) ||
      /^grand\s*total$/i.test(invNoText) ||
      /^total$/i.test(rawHsnText) ||
      /^grand\s*total$/i.test(rawHsnText);

    if (isTotalRow) {
      excludedSummaryRows.push({
        excelRow: excelRowNum,
        label: rawHsnText || companyText || invNoText || 'TOTAL Row',
        basicAmount,
        total: basicAmount + sgst + cgst + igst,
        details: `Original cumulative TOTAL row detected at Excel Row #${excelRowNum} and safely excluded from HSN summary to avoid double-counting.`
      });
      continue;
    }

    // 2. Check for missing or blank HSN Code
    if (!rawHsnText) {
      if (basicAmount > 0 || qty > 0 || sgst > 0 || cgst > 0 || igst > 0) {
        missingData.push({
          excelRow: excelRowNum,
          invoiceNo: invNoText || 'N/A',
          companyName: companyText || 'N/A',
          field: 'Missing HSN Code',
          details: `Excel Row #${excelRowNum} (Invoice: "${invNoText || 'N/A'}") has Basic Amount ₹${basicAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })} / Qty ${qty}, but HSN Code is blank.`
        });
        rawHsnText = 'Unspecified';
      } else {
        continue;
      }
    }

    const cleanHsn = rawHsnText.replace(/\s+/g, '');

    // 3. Check for unusual HSN Format (neither standard 4/6/8-digit HSN nor 9987XX)
    if (cleanHsn !== 'Unspecified' && (cleanHsn.length < 4 || !/^\d+$/.test(cleanHsn))) {
      missingData.push({
        excelRow: excelRowNum,
        invoiceNo: invNoText || 'N/A',
        companyName: companyText || 'N/A',
        field: 'Unusual HSN Format',
        details: `Excel Row #${excelRowNum} (Invoice: "${invNoText || 'N/A'}"): "${rawHsnText}" does not follow standard 4, 6, or 8-digit HSN/SAC code format.`
      });
    }

    // 4. Check for missing quantity when basic amount is positive
    if (basicAmount > 0 && qty <= 0) {
      missingData.push({
        excelRow: excelRowNum,
        invoiceNo: invNoText || 'N/A',
        companyName: companyText || 'N/A',
        field: 'Missing Quantity',
        details: `Excel Row #${excelRowNum} (Invoice: "${invNoText || 'N/A'}", HSN: ${rawHsnText}) has Basic Amount ₹${basicAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}, but Quantity is 0 or blank.`
      });
    }

    // 5. Check for missing basic amount when quantity is positive
    if (qty > 0 && basicAmount <= 0) {
      missingData.push({
        excelRow: excelRowNum,
        invoiceNo: invNoText || 'N/A',
        companyName: companyText || 'N/A',
        field: 'Missing Basic Amount',
        details: `Excel Row #${excelRowNum} (Invoice: "${invNoText || 'N/A'}", HSN: ${rawHsnText}) has Quantity ${qty}, but Taxable Basic Amount is 0 or blank.`
      });
    }

    // 6. Tax Consistency Checks
    // Check 6a: Intra-state SGST and CGST mismatch
    if ((sgst > 0 || cgst > 0) && Math.abs(sgst - cgst) > 0.05) {
      taxInconsistencies.push({
        excelRow: excelRowNum,
        invoiceNo: invNoText || 'N/A',
        hsnCode: rawHsnText,
        type: 'SGST / CGST Mismatch',
        details: `Excel Row #${excelRowNum} (Invoice: "${invNoText || 'N/A'}", HSN: ${rawHsnText}): SGST (₹${sgst.toFixed(2)}) != CGST (₹${cgst.toFixed(2)}). Intra-state tax shares must be equal.`
      });
    }

    // Check 6b: Dual Tax Error (both IGST and SGST/CGST on same line)
    if (igst > 0 && (sgst > 0 || cgst > 0)) {
      taxInconsistencies.push({
        excelRow: excelRowNum,
        invoiceNo: invNoText || 'N/A',
        hsnCode: rawHsnText,
        type: 'Dual Tax Type Error',
        details: `Excel Row #${excelRowNum} (Invoice: "${invNoText || 'N/A'}", HSN: ${rawHsnText}): Both inter-state IGST (₹${igst.toFixed(2)}) and intra-state SGST/CGST (₹${(sgst + cgst).toFixed(2)}) are charged on the same line.`
      });
    }

    // Check 6c: Unclarified Zero Tax
    if (basicAmount > 0 && sgst === 0 && cgst === 0 && igst === 0 && !/sez|nil|zero|exempt|export/i.test(remarks)) {
      taxInconsistencies.push({
        excelRow: excelRowNum,
        invoiceNo: invNoText || 'N/A',
        hsnCode: rawHsnText,
        type: 'Unclarified Zero Tax',
        details: `Excel Row #${excelRowNum} (Invoice: "${invNoText || 'N/A'}", HSN: ${rawHsnText}): Tax is ₹0.00, but remarks column does not indicate SEZ, NIL, or Export status.`
      });
    }

    rawRowCount++;

    // 7. Group HSN (Exception: 9987XX -> 998719)
    const normalizedHsn = normalizeHsnCode(rawHsnText);
    const taxCategory = determineTaxCategory(remarks, sgst, cgst, igst);

    if (cleanHsn.startsWith('9987') && cleanHsn !== '998719') {
      hsnConsolidations.push({
        excelRow: excelRowNum,
        invoiceNo: invNoText || 'N/A',
        originalHsn: cleanHsn,
        unifiedHsn: '998719',
        taxCategory,
        basicAmount,
        details: `Excel Row #${excelRowNum} (Invoice: "${invNoText || 'N/A'}"): Original HSN "${cleanHsn}" consolidated under unified service HSN "998719" (${taxCategory}, ₹${basicAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}).`
      });
    }

    // Composite Group Key
    const groupKey = `${normalizedHsn}___${taxCategory}`;

    if (!groupsMap.has(groupKey)) {
      groupsMap.set(groupKey, {
        hsnCode: normalizedHsn,
        taxCategory,
        invoiceLines: 0,
        quantity: 0,
        basicAmount: 0,
        sgst: 0,
        cgst: 0,
        igst: 0,
        totalInclGst: 0,
      });
    }

    const group = groupsMap.get(groupKey);
    group.invoiceLines += 1;
    group.quantity += qty;
    group.basicAmount += basicAmount;
    group.sgst += sgst;
    group.cgst += cgst;
    group.igst += igst;
    group.totalInclGst += (basicAmount + sgst + cgst + igst);
  }

  // 8. Sequence Gap Check across invoice numbers in the cumulative sheet
  const sequenceGaps = checkSequenceGaps(distinctInvoices);

  // Convert to array and round decimals
  const groups = Array.from(groupsMap.values()).map(g => ({
    hsnCode: g.hsnCode,
    taxCategory: g.taxCategory,
    invoiceLines: g.invoiceLines,
    quantity: Math.round(g.quantity * 100) / 100,
    basicAmount: Math.round(g.basicAmount * 100) / 100,
    sgst: Math.round(g.sgst * 100) / 100,
    cgst: Math.round(g.cgst * 100) / 100,
    igst: Math.round(g.igst * 100) / 100,
    totalInclGst: Math.round((g.basicAmount + g.sgst + g.cgst + g.igst) * 100) / 100,
  }));

  // Sort groups:
  // 1. HSN Code ascending (natural compare)
  // 2. "Taxable (SGST / CGST / IGST)" comes before "Zero GST"
  groups.sort((a, b) => {
    const hsnCmp = naturalCompare(a.hsnCode, b.hsnCode);
    if (hsnCmp !== 0) return hsnCmp;
    if (a.taxCategory === 'Taxable (SGST / CGST / IGST)' && b.taxCategory !== 'Taxable (SGST / CGST / IGST)') return -1;
    if (a.taxCategory !== 'Taxable (SGST / CGST / IGST)' && b.taxCategory === 'Taxable (SGST / CGST / IGST)') return 1;
    return 0;
  });

  // Calculate Overall Totals
  const summary = groups.reduce((acc, g) => {
    acc.totalLines += g.invoiceLines;
    acc.totalQty += g.quantity;
    acc.totalBasic += g.basicAmount;
    acc.totalSgst += g.sgst;
    acc.totalCgst += g.cgst;
    acc.totalIgst += g.igst;
    acc.totalGrand += g.totalInclGst;
    return acc;
  }, {
    totalGroups: groups.length,
    totalLines: 0,
    totalQty: 0,
    totalBasic: 0,
    totalSgst: 0,
    totalCgst: 0,
    totalIgst: 0,
    totalGrand: 0,
  });

  summary.totalBasic = Math.round(summary.totalBasic * 100) / 100;
  summary.totalSgst = Math.round(summary.totalSgst * 100) / 100;
  summary.totalCgst = Math.round(summary.totalCgst * 100) / 100;
  summary.totalIgst = Math.round(summary.totalIgst * 100) / 100;
  summary.totalGrand = Math.round(summary.totalGrand * 100) / 100;

  const totalInconsistencyCount =
    missingData.length +
    taxInconsistencies.length +
    sequenceGaps.reduce((acc, g) => acc + g.totalMissing, 0);

  const auditReport = {
    hasInconsistencies: totalInconsistencyCount > 0,
    totalInconsistencyCount,
    missingDataCount: missingData.length,
    taxInconsistencyCount: taxInconsistencies.length,
    sequenceGapCount: sequenceGaps.reduce((acc, g) => acc + g.totalMissing, 0),
    consolidationCount: hsnConsolidations.length,
    excludedRowCount: excludedSummaryRows.length,
    missingData,
    taxInconsistencies,
    sequenceGaps,
    hsnConsolidations,
    excludedSummaryRows,
  };

  return {
    periodTitle,
    rawRowCount,
    groups,
    summary,
    auditReport,
  };
}

/**
 * Builds the HSN-wise GST Summary Excel workbook matching the exact style,
 * headers, colors, fonts, and formulas seen in the user reference image.
 * 
 * @param {Array} groups - Array of grouped HSN summary rows
 * @param {Object} options - { periodTitle }
 * @returns {Promise<Buffer>}
 */
async function generateHsnSummaryExcel(groups, options = {}) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'HSN GST Summary Generator';
  workbook.lastModifiedBy = 'HSN GST Summary Generator';
  workbook.created = new Date();
  workbook.modified = new Date();

  const worksheet = workbook.addWorksheet('HSN_Wise_GST_Summary', {
    views: [{ showGridLines: true }]
  });

  const periodTitle = options.periodTitle || 'July–September 2026';

  // Row 1: Sheet Title Header
  worksheet.getCell('A1').value = `HSN-wise GST Summary | ${periodTitle}`;
  worksheet.getCell('A1').font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
  worksheet.getRow(1).height = 20;

  // Row 2: Subtitle / Exclusion Note matching image
  worksheet.getCell('A2').value = 'Original invoice TOTAL row excluded; zero GST basic shown separately.';
  worksheet.getCell('A2').font = { name: 'Calibri', size: 10, italic: false, color: { argb: 'FF333333' } };
  worksheet.getRow(2).height = 18;

  // Row 3: Blank spacing row
  worksheet.getRow(3).height = 12;

  // Row 4: Column Headers (Deep Navy Blue background with White Bold text)
  const headers = [
    'HSN Code',
    'Tax Category',
    'Invoice Lines',
    'Quantity',
    'Basic Amount',
    'SGST',
    'CGST',
    'IGST',
    'Total incl. GST',
  ];

  const headerRow = worksheet.getRow(4);
  headers.forEach((header, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = header;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF162E4A' } // Deep Navy Blue as seen in image
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: ['Invoice Lines', 'Quantity', 'Basic Amount', 'SGST', 'CGST', 'IGST', 'Total incl. GST'].includes(header) ? 'right' : 'left'
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF0F2033' } },
      left: { style: 'thin', color: { argb: 'FF0F2033' } },
      bottom: { style: 'medium', color: { argb: 'FF0A1624' } },
      right: { style: 'thin', color: { argb: 'FF0F2033' } }
    };
  });
  headerRow.height = 24;

  let currentRow = 5;

  groups.forEach((g) => {
    const row = worksheet.getRow(currentRow);

    // Col A: HSN Code
    row.getCell(1).value = g.hsnCode || '';
    row.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };

    // Col B: Tax Category
    row.getCell(2).value = g.taxCategory || '';
    row.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };

    // Col C: Invoice Lines
    row.getCell(3).value = g.invoiceLines || 0;
    row.getCell(3).numFmt = '#,##0';
    row.getCell(3).alignment = { vertical: 'middle', horizontal: 'right' };

    // Col D: Quantity
    row.getCell(4).value = g.quantity || 0;
    row.getCell(4).numFmt = '#,##0';
    row.getCell(4).alignment = { vertical: 'middle', horizontal: 'right' };

    // Col E: Basic Amount
    row.getCell(5).value = g.basicAmount || 0;
    row.getCell(5).numFmt = '#,##0.00';
    row.getCell(5).alignment = { vertical: 'middle', horizontal: 'right' };

    // Col F: SGST
    row.getCell(6).value = g.sgst || 0;
    row.getCell(6).numFmt = '#,##0.00';
    row.getCell(6).alignment = { vertical: 'middle', horizontal: 'right' };

    // Col G: CGST
    row.getCell(7).value = g.cgst || 0;
    row.getCell(7).numFmt = '#,##0.00';
    row.getCell(7).alignment = { vertical: 'middle', horizontal: 'right' };

    // Col H: IGST
    row.getCell(8).value = g.igst || 0;
    row.getCell(8).numFmt = '#,##0.00';
    row.getCell(8).alignment = { vertical: 'middle', horizontal: 'right' };

    // Col I: Total incl. GST
    row.getCell(9).value = g.totalInclGst || 0;
    row.getCell(9).numFmt = '#,##0.00';
    row.getCell(9).alignment = { vertical: 'middle', horizontal: 'right' };

    // Borders & Fonts
    for (let c = 1; c <= 9; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Calibri', size: 10 };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
        bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
        left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
        right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
      };
    }

    row.height = 20;
    currentRow++;
  });

  // GRAND TOTAL Row (Matching light blue fill and formulas in image)
  const totalRow = worksheet.getRow(currentRow);
  const lastDataRow = currentRow - 1;

  totalRow.getCell(1).value = 'GRAND TOTAL';
  totalRow.getCell(2).value = '';

  if (lastDataRow >= 5) {
    totalRow.getCell(3).value = { formula: `SUM(C5:C${lastDataRow})` };
    totalRow.getCell(3).numFmt = '#,##0';

    totalRow.getCell(4).value = { formula: `SUM(D5:D${lastDataRow})` };
    totalRow.getCell(4).numFmt = '#,##0';

    totalRow.getCell(5).value = { formula: `SUM(E5:E${lastDataRow})` };
    totalRow.getCell(5).numFmt = '#,##0.00';

    totalRow.getCell(6).value = { formula: `SUM(F5:F${lastDataRow})` };
    totalRow.getCell(6).numFmt = '#,##0.00';

    totalRow.getCell(7).value = { formula: `SUM(G5:G${lastDataRow})` };
    totalRow.getCell(7).numFmt = '#,##0.00';

    totalRow.getCell(8).value = { formula: `SUM(H5:H${lastDataRow})` };
    totalRow.getCell(8).numFmt = '#,##0.00';

    totalRow.getCell(9).value = { formula: `SUM(I5:I${lastDataRow})` };
    totalRow.getCell(9).numFmt = '#,##0.00';
  }

  for (let c = 1; c <= 9; c++) {
    const cell = totalRow.getCell(c);
    cell.font = { name: 'Calibri', size: 10, bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFDDEBF7' } // Light Blue fill as seen in row 23 of image
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'double', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FFB8CCE4' } },
      right: { style: 'thin', color: { argb: 'FFB8CCE4' } }
    };
  }

  totalRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
  totalRow.getCell(3).alignment = { vertical: 'middle', horizontal: 'right' };
  totalRow.getCell(4).alignment = { vertical: 'middle', horizontal: 'right' };
  totalRow.getCell(5).alignment = { vertical: 'middle', horizontal: 'right' };
  totalRow.getCell(6).alignment = { vertical: 'middle', horizontal: 'right' };
  totalRow.getCell(7).alignment = { vertical: 'middle', horizontal: 'right' };
  totalRow.getCell(8).alignment = { vertical: 'middle', horizontal: 'right' };
  totalRow.getCell(9).alignment = { vertical: 'middle', horizontal: 'right' };
  totalRow.height = 22;

  // Set Column Widths for clean alignment
  worksheet.columns = [
    { width: 16 }, // HSN Code
    { width: 32 }, // Tax Category
    { width: 16 }, // Invoice Lines
    { width: 14 }, // Quantity
    { width: 18 }, // Basic Amount
    { width: 15 }, // SGST
    { width: 15 }, // CGST
    { width: 16 }, // IGST
    { width: 20 }, // Total incl. GST
  ];

  return await workbook.xlsx.writeBuffer();
}

/**
 * Builds CSV export for HSN Summary
 */
function generateHsnSummaryCsv(groups, options = {}) {
  const lines = [];
  const period = options.periodTitle || 'July–September 2026';

  lines.push(`"HSN-wise GST Summary | ${period.replace(/"/g, '""')}",,,,,,,,`);
  lines.push('"Original invoice TOTAL row excluded; zero GST basic shown separately.",,,,,,,,');
  lines.push('');
  lines.push('HSN Code,Tax Category,Invoice Lines,Quantity,Basic Amount,SGST,CGST,IGST,Total incl. GST');

  let totalLines = 0;
  let totalQty = 0;
  let totalBasic = 0;
  let totalSgst = 0;
  let totalCgst = 0;
  let totalIgst = 0;
  let totalGrand = 0;

  groups.forEach(g => {
    totalLines += g.invoiceLines;
    totalQty += g.quantity;
    totalBasic += g.basicAmount;
    totalSgst += g.sgst;
    totalCgst += g.cgst;
    totalIgst += g.igst;
    totalGrand += g.totalInclGst;

    lines.push([
      `"${(g.hsnCode || '').replace(/"/g, '""')}"`,
      `"${(g.taxCategory || '').replace(/"/g, '""')}"`,
      g.invoiceLines,
      g.quantity,
      g.basicAmount.toFixed(2),
      g.sgst.toFixed(2),
      g.cgst.toFixed(2),
      g.igst.toFixed(2),
      g.totalInclGst.toFixed(2),
    ].join(','));
  });

  lines.push([
    'GRAND TOTAL',
    '',
    totalLines,
    totalQty,
    totalBasic.toFixed(2),
    totalSgst.toFixed(2),
    totalCgst.toFixed(2),
    totalIgst.toFixed(2),
    totalGrand.toFixed(2),
  ].join(','));

  return lines.join('\r\n');
}

module.exports = {
  normalizeHsnCode,
  determineTaxCategory,
  parseAndGroupCumulativeSheet,
  generateHsnSummaryExcel,
  generateHsnSummaryCsv,
};
