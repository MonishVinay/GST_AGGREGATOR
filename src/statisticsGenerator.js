const ExcelJS = require('exceljs');

/**
 * Builds the comprehensive Statistics & Analytics Excel workbook
 * 
 * Contains:
 * 1. Company-wise Summary (Client's explicit requirement)
 * 2. HSN-wise Summary (GSTR-1 Table 12 compliance)
 * 3. Tax Category Breakdown (Inter-state, Intra-state, SEZ)
 * 4. Monthly / Timeline Breakdown
 * 5. Executive KPIs & Highlights
 * 
 * @param {Array} invoices - Parsed invoice objects
 * @param {Object} options - Customization options
 * @returns {Promise<Buffer>}
 */
async function generateStatisticsExcel(invoices, options = {}) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'GST Invoice Aggregator Analytics';
  workbook.created = new Date();

  const periodTitle = options.periodTitle || 'April-June 2025-26';

  // Overall aggregates
  let grandBasic = 0;
  let grandSgst = 0;
  let grandCgst = 0;
  let grandIgst = 0;
  let grandQty = 0;
  let grandItems = 0;

  invoices.forEach(inv => {
    (inv.items || []).forEach(it => {
      grandBasic += (it.basicAmount || 0);
      grandSgst += (it.sgst || 0);
      grandCgst += (it.cgst || 0);
      grandIgst += (it.igst || 0);
      grandQty += (it.qty || 0);
      grandItems++;
    });
  });
  const grandTax = grandSgst + grandCgst + grandIgst;
  const grandTotal = grandBasic + grandTax;

  // -------------------------------------------------------------
  // SHEET 1: Executive KPI Dashboard
  // -------------------------------------------------------------
  const kpiSheet = workbook.addWorksheet('Executive Overview', {
    views: [{ showGridLines: true }]
  });

  kpiSheet.columns = [
    { width: 4 },
    { width: 34 },
    { width: 24 },
    { width: 30 }
  ];

  // Header Title
  kpiSheet.mergeCells('B2:D2');
  const titleCell = kpiSheet.getCell('B2');
  titleCell.value = `GST SALES & INVOICE ANALYTICS SUMMARY (${periodTitle})`;
  titleCell.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  kpiSheet.getRow(2).height = 36;

  // Key Metrics
  const kpiData = [
    ['Total Invoices Processed', invoices.length, 'Cumulative invoices parsed'],
    ['Total Line Items', grandItems, 'Individual product/service lines'],
    ['Total Quantity Supplied', grandQty, 'Units across all items'],
    ['Total Taxable / Basic Value', grandBasic, 'Net turnover before taxes'],
    ['Total Central GST (CGST)', grandSgst, 'Intra-state central revenue'],
    ['Total State GST (SGST)', grandCgst, 'Intra-state state revenue'],
    ['Total Integrated GST (IGST)', grandIgst, 'Inter-state integrated revenue'],
    ['Total Tax Collected', grandTax, 'Cumulative GST liability'],
    ['Grand Total Invoiced Value', grandTotal, 'Gross value including all taxes']
  ];

  kpiSheet.getCell('B4').value = 'KEY PERFORMANCE METRIC';
  kpiSheet.getCell('C4').value = 'VALUE';
  kpiSheet.getCell('D4').value = 'DESCRIPTION';
  ['B4', 'C4', 'D4'].forEach(c => {
    const cell = kpiSheet.getCell(c);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
    cell.alignment = { vertical: 'middle', horizontal: c === 'C4' ? 'right' : 'left' };
  });
  kpiSheet.getRow(4).height = 24;

  let kpiRowIdx = 5;
  kpiData.forEach(([label, val, desc]) => {
    const row = kpiSheet.getRow(kpiRowIdx);
    row.getCell(2).value = label;
    row.getCell(3).value = val;
    row.getCell(4).value = desc;

    row.getCell(2).font = { bold: true };
    if (typeof val === 'number') {
      if (label.includes('Total Taxable') || label.includes('GST') || label.includes('Value')) {
        row.getCell(3).numFmt = '₹ #,##0.00';
      } else {
        row.getCell(3).numFmt = '#,##0';
      }
    }
    row.getCell(3).alignment = { horizontal: 'right' };
    row.getCell(4).font = { color: { argb: 'FF64748B' }, italic: true };

    for (let col = 2; col <= 4; col++) {
      row.getCell(col).border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
    }
    row.height = 22;
    kpiRowIdx++;
  });

  // -------------------------------------------------------------
  // SHEET 2: Company-Wise Summary (Client Requirement)
  // -------------------------------------------------------------
  const companySheet = workbook.addWorksheet('Company Wise Stats', {
    views: [{ showGridLines: true }]
  });

  // Aggregate by Company
  const companyMap = new Map();
  invoices.forEach(inv => {
    const key = (inv.companyName || 'Unknown Company').trim();
    if (!companyMap.has(key)) {
      companyMap.set(key, {
        companyName: key,
        gstNo: inv.gstNo || '',
        stateCode: inv.gstNo ? inv.gstNo.substring(0, 2) : '',
        invoices: new Set(),
        itemCount: 0,
        qty: 0,
        basicAmount: 0,
        sgst: 0,
        cgst: 0,
        igst: 0,
        isSez: false,
      });
    }
    const cData = companyMap.get(key);
    cData.invoices.add(inv.invoiceNo);
    if (!cData.gstNo && inv.gstNo) {
      cData.gstNo = inv.gstNo;
      cData.stateCode = inv.gstNo.substring(0, 2);
    }
    if (inv.isSez) cData.isSez = true;

    (inv.items || []).forEach(it => {
      cData.itemCount++;
      cData.qty += (it.qty || 0);
      cData.basicAmount += (it.basicAmount || 0);
      cData.sgst += (it.sgst || 0);
      cData.cgst += (it.cgst || 0);
      cData.igst += (it.igst || 0);
    });
  });

  const companyList = Array.from(companyMap.values());
  // Sort descending by highest sales
  companyList.sort((a, b) => b.basicAmount - a.basicAmount);

  // Headers
  const compHeaders = [
    'S.No',
    'Company Name',
    'GSTIN',
    'State Code',
    'Supply Type',
    'Invoices Count',
    'Total Qty',
    'Taxable Value (Basic INR)',
    'SGST INR',
    'CGST INR',
    'IGST INR',
    'Total Tax INR',
    'Grand Total INR',
    '% Share of Sales'
  ];

  companySheet.columns = [
    { width: 8 },  // S.No
    { width: 40 }, // Company Name
    { width: 20 }, // GSTIN
    { width: 12 }, // State Code
    { width: 14 }, // Supply Type
    { width: 14 }, // Invoices Count
    { width: 12 }, // Total Qty
    { width: 24 }, // Basic
    { width: 16 }, // SGST
    { width: 16 }, // CGST
    { width: 16 }, // IGST
    { width: 18 }, // Total Tax
    { width: 22 }, // Grand Total
    { width: 16 }, // % Share
  ];

  // Header styling
  const compHeaderRow = companySheet.getRow(1);
  compHeaders.forEach((h, i) => {
    const cell = compHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.alignment = { vertical: 'middle', horizontal: i >= 5 ? 'right' : 'left' };
    cell.border = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };
  });
  compHeaderRow.height = 28;

  let compRowIdx = 2;
  companyList.forEach((c, idx) => {
    const row = companySheet.getRow(compRowIdx);
    const supplyType = c.isSez ? 'SEZ (Nil)' : (c.igst > 0 ? 'Inter-State' : 'Intra-State');
    const taxSum = c.sgst + c.cgst + c.igst;
    const grossVal = c.basicAmount + taxSum;
    const share = grandBasic > 0 ? (c.basicAmount / grandBasic) : 0;

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = c.companyName;
    row.getCell(3).value = c.gstNo;
    row.getCell(4).value = c.stateCode;
    row.getCell(5).value = supplyType;
    row.getCell(6).value = c.invoices.size;
    row.getCell(7).value = c.qty;
    row.getCell(8).value = c.basicAmount;
    row.getCell(9).value = c.sgst;
    row.getCell(10).value = c.cgst;
    row.getCell(11).value = c.igst;
    row.getCell(12).value = taxSum;
    row.getCell(13).value = grossVal;
    row.getCell(14).value = share;

    // Formats
    row.getCell(6).numFmt = '#,##0';
    row.getCell(7).numFmt = '#,##0';
    row.getCell(8).numFmt = '#,##0.00';
    row.getCell(9).numFmt = '#,##0.00';
    row.getCell(10).numFmt = '#,##0.00';
    row.getCell(11).numFmt = '#,##0.00';
    row.getCell(12).numFmt = '#,##0.00';
    row.getCell(13).numFmt = '#,##0.00';
    row.getCell(14).numFmt = '0.00%';

    for (let col = 1; col <= 14; col++) {
      const cell = row.getCell(col);
      cell.font = { name: 'Calibri', size: 10 };
      cell.border = {
        top: { style: 'hair', color: { argb: 'FFD1D5DB' } },
        bottom: { style: 'hair', color: { argb: 'FFD1D5DB' } },
        left: { style: 'hair', color: { argb: 'FFD1D5DB' } },
        right: { style: 'hair', color: { argb: 'FFD1D5DB' } }
      };
      if (col >= 6) cell.alignment = { horizontal: 'right' };
    }

    compRowIdx++;
  });

  // Totals Row
  const compTotalRow = companySheet.getRow(compRowIdx);
  compTotalRow.getCell(2).value = 'TOTAL';
  compTotalRow.getCell(6).value = { formula: `SUM(F2:F${compRowIdx - 1})` };
  compTotalRow.getCell(7).value = { formula: `SUM(G2:G${compRowIdx - 1})` };
  compTotalRow.getCell(8).value = { formula: `SUM(H2:H${compRowIdx - 1})` };
  compTotalRow.getCell(9).value = { formula: `SUM(I2:I${compRowIdx - 1})` };
  compTotalRow.getCell(10).value = { formula: `SUM(J2:J${compRowIdx - 1})` };
  compTotalRow.getCell(11).value = { formula: `SUM(K2:K${compRowIdx - 1})` };
  compTotalRow.getCell(12).value = { formula: `SUM(L2:L${compRowIdx - 1})` };
  compTotalRow.getCell(13).value = { formula: `SUM(M2:M${compRowIdx - 1})` };
  compTotalRow.getCell(14).value = 1.0;

  compTotalRow.getCell(6).numFmt = '#,##0';
  compTotalRow.getCell(7).numFmt = '#,##0';
  for (let c = 8; c <= 13; c++) compTotalRow.getCell(c).numFmt = '#,##0.00';
  compTotalRow.getCell(14).numFmt = '0.00%';

  for (let col = 1; col <= 14; col++) {
    const cell = compTotalRow.getCell(col);
    cell.font = { name: 'Calibri', size: 10, bold: true };
    cell.border = {
      top: { style: 'thin' },
      bottom: { style: 'double' }
    };
    if (col >= 6) cell.alignment = { horizontal: 'right' };
  }

  // -------------------------------------------------------------
  // SHEET 3: HSN Code Wise Summary (GSTR-1 HSN compliance)
  // -------------------------------------------------------------
  const hsnSheet = workbook.addWorksheet('HSN Code Summary', {
    views: [{ showGridLines: true }]
  });

  const hsnMap = new Map();
  invoices.forEach(inv => {
    (inv.items || []).forEach(it => {
      const code = String(it.hsnCode || 'Unspecified').trim();
      if (!hsnMap.has(code)) {
        hsnMap.set(code, {
          hsnCode: code,
          invoices: new Set(),
          itemCount: 0,
          qty: 0,
          basicAmount: 0,
          sgst: 0,
          cgst: 0,
          igst: 0,
          descriptions: new Set(),
        });
      }
      const hData = hsnMap.get(code);
      hData.invoices.add(inv.invoiceNo);
      hData.itemCount++;
      hData.qty += (it.qty || 0);
      hData.basicAmount += (it.basicAmount || 0);
      hData.sgst += (it.sgst || 0);
      hData.cgst += (it.cgst || 0);
      hData.igst += (it.igst || 0);
      if (it.description) hData.descriptions.add(it.description);
    });
  });

  const hsnList = Array.from(hsnMap.values());
  hsnList.sort((a, b) => b.basicAmount - a.basicAmount);

  const hsnHeaders = [
    'S.No',
    'HSN / SAC Code',
    'Sample Description',
    'Invoices Count',
    'Total Qty',
    'Taxable Value (Basic INR)',
    'SGST INR',
    'CGST INR',
    'IGST INR',
    'Total Tax INR',
    'Grand Total INR',
    '% Share of Turnover'
  ];

  hsnSheet.columns = [
    { width: 8 },  // S.No
    { width: 18 }, // HSN Code
    { width: 36 }, // Sample Description
    { width: 14 }, // Invoices Count
    { width: 12 }, // Total Qty
    { width: 24 }, // Basic
    { width: 16 }, // SGST
    { width: 16 }, // CGST
    { width: 16 }, // IGST
    { width: 18 }, // Total Tax
    { width: 22 }, // Grand Total
    { width: 18 }, // % Share
  ];

  const hsnHeaderRow = hsnSheet.getRow(1);
  hsnHeaders.forEach((h, i) => {
    const cell = hsnHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D9488' } };
    cell.alignment = { vertical: 'middle', horizontal: i >= 3 ? 'right' : 'left' };
    cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });
  hsnHeaderRow.height = 28;

  let hsnRowIdx = 2;
  hsnList.forEach((h, idx) => {
    const row = hsnSheet.getRow(hsnRowIdx);
    const taxSum = h.sgst + h.cgst + h.igst;
    const grossVal = h.basicAmount + taxSum;
    const share = grandBasic > 0 ? (h.basicAmount / grandBasic) : 0;
    const descSample = Array.from(h.descriptions).slice(0, 2).join('; ');

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = h.hsnCode;
    row.getCell(3).value = descSample;
    row.getCell(4).value = h.invoices.size;
    row.getCell(5).value = h.qty;
    row.getCell(6).value = h.basicAmount;
    row.getCell(7).value = h.sgst;
    row.getCell(8).value = h.cgst;
    row.getCell(9).value = h.igst;
    row.getCell(10).value = taxSum;
    row.getCell(11).value = grossVal;
    row.getCell(12).value = share;

    row.getCell(4).numFmt = '#,##0';
    row.getCell(5).numFmt = '#,##0';
    for (let c = 6; c <= 11; c++) row.getCell(c).numFmt = '#,##0.00';
    row.getCell(12).numFmt = '0.00%';

    for (let col = 1; col <= 12; col++) {
      const cell = row.getCell(col);
      cell.font = { name: 'Calibri', size: 10 };
      cell.border = {
        top: { style: 'hair', color: { argb: 'FFD1D5DB' } },
        bottom: { style: 'hair', color: { argb: 'FFD1D5DB' } },
        left: { style: 'hair', color: { argb: 'FFD1D5DB' } },
        right: { style: 'hair', color: { argb: 'FFD1D5DB' } }
      };
      if (col >= 4) cell.alignment = { horizontal: 'right' };
    }

    hsnRowIdx++;
  });

  // HSN Totals Row
  const hsnTotalRow = hsnSheet.getRow(hsnRowIdx);
  hsnTotalRow.getCell(2).value = 'TOTAL';
  hsnTotalRow.getCell(4).value = { formula: `SUM(D2:D${hsnRowIdx - 1})` };
  hsnTotalRow.getCell(5).value = { formula: `SUM(E2:E${hsnRowIdx - 1})` };
  hsnTotalRow.getCell(6).value = { formula: `SUM(F2:F${hsnRowIdx - 1})` };
  hsnTotalRow.getCell(7).value = { formula: `SUM(G2:G${hsnRowIdx - 1})` };
  hsnTotalRow.getCell(8).value = { formula: `SUM(H2:H${hsnRowIdx - 1})` };
  hsnTotalRow.getCell(9).value = { formula: `SUM(I2:I${hsnRowIdx - 1})` };
  hsnTotalRow.getCell(10).value = { formula: `SUM(J2:J${hsnRowIdx - 1})` };
  hsnTotalRow.getCell(11).value = { formula: `SUM(K2:K${hsnRowIdx - 1})` };
  hsnTotalRow.getCell(12).value = 1.0;

  hsnTotalRow.getCell(4).numFmt = '#,##0';
  hsnTotalRow.getCell(5).numFmt = '#,##0';
  for (let c = 6; c <= 11; c++) hsnTotalRow.getCell(c).numFmt = '#,##0.00';
  hsnTotalRow.getCell(12).numFmt = '0.00%';

  for (let col = 1; col <= 12; col++) {
    const cell = hsnTotalRow.getCell(col);
    cell.font = { name: 'Calibri', size: 10, bold: true };
    cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
    if (col >= 4) cell.alignment = { horizontal: 'right' };
  }

  // -------------------------------------------------------------
  // SHEET 4: Tax Type Breakdown
  // -------------------------------------------------------------
  const taxSheet = workbook.addWorksheet('Tax Breakdown', {
    views: [{ showGridLines: true }]
  });

  taxSheet.columns = [
    { width: 28 },
    { width: 14 },
    { width: 24 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 18 },
    { width: 22 },
  ];

  const taxHeaders = [
    'Supply Category',
    'Invoices Count',
    'Taxable Value (Basic INR)',
    'SGST INR',
    'CGST INR',
    'IGST INR',
    'Total Tax INR',
    'Grand Total INR'
  ];

  const taxHeaderRow = taxSheet.getRow(1);
  taxHeaders.forEach((h, i) => {
    const cell = taxHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4338CA' } };
    cell.alignment = { vertical: 'middle', horizontal: i >= 1 ? 'right' : 'left' };
    cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });
  taxHeaderRow.height = 28;

  // Classify invoices
  const categories = {
    'Inter-State (IGST 18%)': { count: 0, basic: 0, sgst: 0, cgst: 0, igst: 0 },
    'Intra-State (CGST 9% + SGST 9%)': { count: 0, basic: 0, sgst: 0, cgst: 0, igst: 0 },
    'Zero-Rated / SEZ (GST NIL)': { count: 0, basic: 0, sgst: 0, cgst: 0, igst: 0 },
  };

  invoices.forEach(inv => {
    let catKey = 'Inter-State (IGST 18%)';
    if (inv.isSez || (inv.items && inv.items.some(it => /sez|nil/i.test(it.remarks)))) {
      catKey = 'Zero-Rated / SEZ (GST NIL)';
    } else if (inv.totalCgst > 0 || inv.totalSgst > 0 || inv.taxType === 'CGST_SGST') {
      catKey = 'Intra-State (CGST 9% + SGST 9%)';
    }

    const cat = categories[catKey];
    cat.count++;
    cat.basic += inv.totalBasicAmount;
    cat.sgst += inv.totalSgst;
    cat.cgst += inv.totalCgst;
    cat.igst += inv.totalIgst;
  });

  let taxRowIdx = 2;
  Object.entries(categories).forEach(([name, data]) => {
    const row = taxSheet.getRow(taxRowIdx);
    const taxTot = data.sgst + data.cgst + data.igst;
    row.getCell(1).value = name;
    row.getCell(2).value = data.count;
    row.getCell(3).value = data.basic;
    row.getCell(4).value = data.sgst;
    row.getCell(5).value = data.cgst;
    row.getCell(6).value = data.igst;
    row.getCell(7).value = taxTot;
    row.getCell(8).value = data.basic + taxTot;

    row.getCell(2).numFmt = '#,##0';
    for (let c = 3; c <= 8; c++) row.getCell(c).numFmt = '#,##0.00';

    for (let col = 1; col <= 8; col++) {
      const cell = row.getCell(col);
      cell.font = { name: 'Calibri', size: 10 };
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
      if (col >= 2) cell.alignment = { horizontal: 'right' };
    }
    taxRowIdx++;
  });

  // -------------------------------------------------------------
  // SHEET 5: Audit & Inconsistencies Report
  // -------------------------------------------------------------
  const { runFullAudit } = require('./auditEngine');
  const audit = runFullAudit(invoices, options);

  const auditSheet = workbook.addWorksheet('Audit & Inconsistencies', {
    views: [{ showGridLines: true }]
  });

  auditSheet.columns = [
    { width: 20 },
    { width: 32 },
    { width: 25 },
    { width: 55 },
  ];

  // Title
  auditSheet.mergeCells('A1:D1');
  const auditTitle = auditSheet.getCell('A1');
  auditTitle.value = `INVOICE DATA AUDIT & INCONSISTENCIES REPORT (${periodTitle})`;
  auditTitle.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  auditTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF991B1B' } };
  auditTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  auditSheet.getRow(1).height = 30;

  // Header
  const auditHeaders = ['Category', 'Invoice No / Series', 'Filename / Company', 'Details & Inconsistency Description'];
  const aHeaderRow = auditSheet.getRow(3);
  auditHeaders.forEach((h, i) => {
    const cell = aHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
    cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });
  aHeaderRow.height = 24;

  let aRowIdx = 4;

  // 1. Sequence Gaps
  if (audit.sequenceGaps.length > 0) {
    audit.sequenceGaps.forEach(g => {
      const row = auditSheet.getRow(aRowIdx);
      row.getCell(1).value = 'Missing in Sequence';
      row.getCell(2).value = g.series;
      row.getCell(3).value = g.rangeAnalyzed;
      row.getCell(4).value = `Missing ${g.totalMissing} invoices: ${g.missingInvoices.join(', ')}`;
      row.getCell(1).font = { bold: true, color: { argb: 'FFB45309' } };
      aRowIdx++;
    });
  }

  // 2. Missing Data
  if (audit.missingData.length > 0) {
    audit.missingData.forEach(m => {
      const row = auditSheet.getRow(aRowIdx);
      row.getCell(1).value = 'Missing Required Data';
      row.getCell(2).value = m.invoiceNo;
      row.getCell(3).value = `${m.filename} (${m.companyName})`;
      row.getCell(4).value = m.missingFields.join('; ');
      row.getCell(1).font = { bold: true, color: { argb: 'FFDC2626' } };
      aRowIdx++;
    });
  }

  // 3. Duplicates
  if (audit.duplicates.length > 0) {
    audit.duplicates.forEach(d => {
      const row = auditSheet.getRow(aRowIdx);
      row.getCell(1).value = 'Duplicate Invoice No';
      row.getCell(2).value = d.invoiceNo;
      row.getCell(3).value = d.files.join(', ');
      row.getCell(4).value = `Appears ${d.occurrences} times across files`;
      row.getCell(1).font = { bold: true, color: { argb: 'FF7C3AED' } };
      aRowIdx++;
    });
  }

  // 4. Clubbing Log
  if (audit.clubbingLogs.length > 0) {
    audit.clubbingLogs.forEach(c => {
      const row = auditSheet.getRow(aRowIdx);
      row.getCell(1).value = 'HSN Clubbed Invoice';
      row.getCell(2).value = c.invoiceNo;
      row.getCell(3).value = c.companyName;
      row.getCell(4).value = c.details;
      row.getCell(1).font = { bold: true, color: { argb: 'FF2563EB' } };
      aRowIdx++;
    });
  }

  if (aRowIdx === 4) {
    const row = auditSheet.getRow(4);
    row.getCell(1).value = 'Clean';
    row.getCell(4).value = 'No sequence gaps or missing data inconsistencies detected.';
    row.getCell(1).font = { bold: true, color: { argb: 'FF16A34A' } };
    aRowIdx++;
  }

  for (let r = 4; r < aRowIdx; r++) {
    const row = auditSheet.getRow(r);
    for (let c = 1; c <= 4; c++) {
      row.getCell(c).border = { top: { style: 'thin', color: { argb: 'FFE5E7EB' } }, bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } }, left: { style: 'thin', color: { argb: 'FFE5E7EB' } }, right: { style: 'thin', color: { argb: 'FFE5E7EB' } } };
    }
  }

  return await workbook.xlsx.writeBuffer();
}

module.exports = {
  generateStatisticsExcel,
};
