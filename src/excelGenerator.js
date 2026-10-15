const ExcelJS = require('exceljs');

/**
 * Clubs items sharing the same HSN code within the SAME invoice into a single row
 */
function clubInvoiceItemsByHsn(invoices) {
  return invoices.map(inv => {
    if (!inv.items || inv.items.length <= 1) return inv;

    const hsnMap = new Map();
    for (const item of inv.items) {
      const hsnKey = String(item.hsnCode || '').trim();
      if (!hsnMap.has(hsnKey)) {
        hsnMap.set(hsnKey, {
          hsnCode: item.hsnCode,
          qty: 0,
          basicAmount: 0,
          sgst: 0,
          cgst: 0,
          igst: 0,
          descriptions: [],
          remarks: item.remarks || '',
        });
      }
      const grouped = hsnMap.get(hsnKey);
      grouped.qty += (item.qty || 0);
      grouped.basicAmount += (item.basicAmount || 0);
      grouped.sgst += (item.sgst || 0);
      grouped.cgst += (item.cgst || 0);
      grouped.igst += (item.igst || 0);
      if (item.description && !grouped.descriptions.includes(item.description)) {
        grouped.descriptions.push(item.description);
      }
      if (!grouped.remarks && item.remarks) {
        grouped.remarks = item.remarks;
      }
    }

    const clubbedItems = Array.from(hsnMap.values()).map(g => ({
      hsnCode: g.hsnCode,
      qty: g.qty,
      basicAmount: Math.round(g.basicAmount * 100) / 100,
      sgst: Math.round(g.sgst * 100) / 100,
      cgst: Math.round(g.cgst * 100) / 100,
      igst: Math.round(g.igst * 100) / 100,
      description: g.descriptions.join('; '),
      remarks: g.remarks,
    }));

    return {
      ...inv,
      items: clubbedItems,
    };
  });
}

/**
 * Collapses each invoice into a single row using the HSN code of the item with the highest amount,
 * summing all quantities, basic amounts, and taxes on that invoice.
 */
function collapseInvoiceToHighestHsn(invoices) {
  return invoices.map(inv => {
    if (!inv.items || inv.items.length === 0) return inv;
    if (inv.items.length === 1) return inv;

    // Find the item with the highest basicAmount
    let highestItem = inv.items[0];
    for (let i = 1; i < inv.items.length; i++) {
      if ((inv.items[i].basicAmount || 0) > (highestItem.basicAmount || 0)) {
        highestItem = inv.items[i];
      }
    }

    const totalQty = inv.items.reduce((s, it) => s + (it.qty || 0), 0);
    const totalBasic = inv.items.reduce((s, it) => s + (it.basicAmount || 0), 0);
    const totalSgst = inv.items.reduce((s, it) => s + (it.sgst || 0), 0);
    const totalCgst = inv.items.reduce((s, it) => s + (it.cgst || 0), 0);
    const totalIgst = inv.items.reduce((s, it) => s + (it.igst || 0), 0);

    const remarksList = inv.items.map(it => it.remarks).filter(Boolean);
    const finalRemark = remarksList.length > 0 ? remarksList[0] : '';
    const descriptions = inv.items.map(it => it.description).filter(Boolean);

    const singleItem = {
      hsnCode: highestItem.hsnCode || '',
      qty: totalQty,
      basicAmount: Math.round(totalBasic * 100) / 100,
      sgst: Math.round(totalSgst * 100) / 100,
      cgst: Math.round(totalCgst * 100) / 100,
      igst: Math.round(totalIgst * 100) / 100,
      description: descriptions.join('; '),
      remarks: finalRemark,
    };

    return {
      ...inv,
      items: [singleItem],
    };
  });
}

/**
 * Builds the cumulative Excel workbook identical to Image 3
 * 
 * @param {Array} invoices - Array of parsed invoice objects
 * @param {Object} options - Customization options (periodTitle, repeatInvoiceInfo, clubByHsn, highestHsnOnly, etc.)
 * @returns {Promise<Buffer>}
 */
async function generateCumulativeExcel(invoices, options = {}) {
  let activeInvoices = invoices;
  if (options.highestHsnOnly === true) {
    activeInvoices = collapseInvoiceToHighestHsn(invoices);
  } else if (options.clubByHsn === true) {
    activeInvoices = clubInvoiceItemsByHsn(invoices);
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Cumulative Invoice Aggregator';
  workbook.lastModifiedBy = 'Cumulative Invoice Aggregator';
  workbook.created = new Date();
  workbook.modified = new Date();

  const worksheet = workbook.addWorksheet('Sheet1', {
    views: [{ showGridLines: true }]
  });

  // Determine period title (e.g., "April-June 2025-26")
  let periodTitle = options.periodTitle || 'April-June 2025-26';
  const repeatInvoiceInfo = options.repeatInvoiceInfo === true;

  // Row 1: Title Header
  worksheet.getCell('A1').value = 'OUT INV';
  worksheet.getCell('B1').value = periodTitle;

  worksheet.getCell('A1').font = { name: 'Calibri', size: 11, bold: true };
  worksheet.getCell('B1').font = { name: 'Calibri', size: 11, bold: true };

  // Row 2: Table Column Headers
  const headers = [
    'INVOICE NO',
    'INV DATE',
    'Company Name',
    'GST NO',
    'HSN Code',
    'Qty',
    'BASIC AMOUNT',
    'SGST',
    'CGST',
    'IGST',
    'TOTAL AMOUNT',
    'description / remarks'
  ];

  const headerRow = worksheet.getRow(2);
  headers.forEach((header, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = header;
    cell.font = { name: 'Calibri', size: 10, bold: true };
    cell.alignment = {
      vertical: 'middle',
      horizontal: ['Qty', 'BASIC AMOUNT', 'SGST', 'CGST', 'IGST', 'TOTAL AMOUNT'].includes(header) ? 'right' : 'left'
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFD4D4D4' } },
      left: { style: 'thin', color: { argb: 'FFD4D4D4' } },
      bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
      right: { style: 'thin', color: { argb: 'FFD4D4D4' } }
    };
  });
  headerRow.height = 24;

  let currentRowIndex = 3;

  for (const inv of activeInvoices) {
    if (!inv.items || inv.items.length === 0) continue;

    inv.items.forEach((item, itemIdx) => {
      const row = worksheet.getRow(currentRowIndex);
      const isFirstItem = itemIdx === 0;

      // Columns A-D: Only on first item unless repeatInvoiceInfo is enabled
      if (isFirstItem || repeatInvoiceInfo) {
        row.getCell(1).value = inv.invoiceNo || '';
        row.getCell(2).value = inv.invoiceDate || '';
        row.getCell(3).value = inv.companyName || '';
        row.getCell(4).value = inv.gstNo || '';
      } else {
        row.getCell(1).value = '';
        row.getCell(2).value = '';
        row.getCell(3).value = '';
        row.getCell(4).value = '';
      }

      // Column E: HSN Code (string format)
      row.getCell(5).value = item.hsnCode ? String(item.hsnCode) : '';

      // Column F: Qty (number)
      row.getCell(6).value = item.qty || 0;
      row.getCell(6).numFmt = '#,##0';

      // Column G: BASIC AMOUNT (number)
      const basic = item.basicAmount || 0;
      row.getCell(7).value = basic;
      row.getCell(7).numFmt = '#,##0.00';

      // Column H: SGST (number)
      const sgst = item.sgst || 0;
      row.getCell(8).value = sgst;
      row.getCell(8).numFmt = '#,##0.00';

      // Column I: CGST (number)
      const cgst = item.cgst || 0;
      row.getCell(9).value = cgst;
      row.getCell(9).numFmt = '#,##0.00';

      // Column J: IGST (number)
      const igst = item.igst || 0;
      row.getCell(10).value = igst;
      row.getCell(10).numFmt = '#,##0.00';

      // Column K: TOTAL AMOUNT (number)
      const lineTotal = Math.round((basic + sgst + cgst + igst) * 100) / 100;
      row.getCell(11).value = lineTotal;
      row.getCell(11).numFmt = '#,##0.00';

      // Column L: description / remarks
      row.getCell(12).value = item.remarks || '';

      // Alignments & Fonts
      for (let c = 1; c <= 12; c++) {
        const cell = row.getCell(c);
        cell.font = { name: 'Calibri', size: 10 };
        cell.border = {
          top: { style: 'hair', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'hair', color: { argb: 'FFE0E0E0' } },
          left: { style: 'hair', color: { argb: 'FFE0E0E0' } },
          right: { style: 'hair', color: { argb: 'FFE0E0E0' } }
        };
      }

      // Numbers right aligned
      row.getCell(6).alignment = { horizontal: 'right' };
      row.getCell(7).alignment = { horizontal: 'right' };
      row.getCell(8).alignment = { horizontal: 'right' };
      row.getCell(9).alignment = { horizontal: 'right' };
      row.getCell(10).alignment = { horizontal: 'right' };
      row.getCell(11).alignment = { horizontal: 'right' };

      currentRowIndex++;
    });
  }

  // Summary Row (Totals)
  const summaryRow = worksheet.getRow(currentRowIndex);
  summaryRow.getCell(3).value = 'Total';
  summaryRow.getCell(3).font = { name: 'Calibri', size: 10, bold: true };

  // Formulas for totals
  const lastDataRow = currentRowIndex - 1;
  if (lastDataRow >= 3) {
    summaryRow.getCell(6).value = { formula: `SUM(F3:F${lastDataRow})` };
    summaryRow.getCell(6).numFmt = '#,##0';

    summaryRow.getCell(7).value = { formula: `SUM(G3:G${lastDataRow})` };
    summaryRow.getCell(7).numFmt = '#,##0.00';

    summaryRow.getCell(8).value = { formula: `SUM(H3:H${lastDataRow})` };
    summaryRow.getCell(8).numFmt = '#,##0.00';

    summaryRow.getCell(9).value = { formula: `SUM(I3:I${lastDataRow})` };
    summaryRow.getCell(9).numFmt = '#,##0.00';

    summaryRow.getCell(10).value = { formula: `SUM(J3:J${lastDataRow})` };
    summaryRow.getCell(10).numFmt = '#,##0.00';

    summaryRow.getCell(11).value = { formula: `SUM(K3:K${lastDataRow})` };
    summaryRow.getCell(11).numFmt = '#,##0.00';
  }

  for (let c = 1; c <= 12; c++) {
    const cell = summaryRow.getCell(c);
    cell.font = { name: 'Calibri', size: 10, bold: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'double', color: { argb: 'FF000000' } }
    };
  }
  summaryRow.getCell(6).alignment = { horizontal: 'right' };
  summaryRow.getCell(7).alignment = { horizontal: 'right' };
  summaryRow.getCell(8).alignment = { horizontal: 'right' };
  summaryRow.getCell(9).alignment = { horizontal: 'right' };
  summaryRow.getCell(10).alignment = { horizontal: 'right' };
  summaryRow.getCell(11).alignment = { horizontal: 'right' };

  // Set Column Widths for comfortable readability matching Image 3
  worksheet.columns = [
    { width: 16 }, // INVOICE NO
    { width: 13 }, // INV DATE
    { width: 38 }, // Company Name
    { width: 22 }, // GST NO
    { width: 14 }, // HSN Code
    { width: 8 },  // Qty
    { width: 16 }, // BASIC AMOUNT
    { width: 13 }, // SGST
    { width: 13 }, // CGST
    { width: 14 }, // IGST
    { width: 16 }, // TOTAL AMOUNT
    { width: 25 }, // description / remarks
  ];

  return await workbook.xlsx.writeBuffer();
}

/**
 * Builds CSV string
 */
function generateCumulativeCsv(invoices, options = {}) {
  let activeInvoices = invoices;
  if (options.highestHsnOnly === true) {
    activeInvoices = collapseInvoiceToHighestHsn(invoices);
  } else if (options.clubByHsn === true) {
    activeInvoices = clubInvoiceItemsByHsn(invoices);
  }

  const repeatInvoiceInfo = options.repeatInvoiceInfo === true;
  const lines = [];

  // Header row 1
  lines.push(`OUT INV,${options.periodTitle || 'April-June 2025-26'},,,,,,,,,,`);

  // Header row 2
  lines.push('INVOICE NO,INV DATE,Company Name,GST NO,HSN Code,Qty,BASIC AMOUNT,SGST,CGST,IGST,TOTAL AMOUNT,description / remarks');

  for (const inv of activeInvoices) {
    if (!inv.items || inv.items.length === 0) continue;

    inv.items.forEach((item, itemIdx) => {
      const isFirstItem = itemIdx === 0;
      const invNo = (isFirstItem || repeatInvoiceInfo) ? `"${(inv.invoiceNo || '').replace(/"/g, '""')}"` : '';
      const invDate = (isFirstItem || repeatInvoiceInfo) ? `"${(inv.invoiceDate || '').replace(/"/g, '""')}"` : '';
      const compName = (isFirstItem || repeatInvoiceInfo) ? `"${(inv.companyName || '').replace(/"/g, '""')}"` : '';
      const gstNo = (isFirstItem || repeatInvoiceInfo) ? `"${(inv.gstNo || '').replace(/"/g, '""')}"` : '';
      const hsn = `"${(item.hsnCode || '').replace(/"/g, '""')}"`;
      const qty = item.qty || 0;
      const basicNum = item.basicAmount || 0;
      const sgstNum = item.sgst || 0;
      const cgstNum = item.cgst || 0;
      const igstNum = item.igst || 0;
      const totalNum = basicNum + sgstNum + cgstNum + igstNum;

      const basic = basicNum.toFixed(2);
      const sgst = sgstNum.toFixed(2);
      const cgst = cgstNum.toFixed(2);
      const igst = igstNum.toFixed(2);
      const total = totalNum.toFixed(2);
      const remarks = `"${(item.remarks || '').replace(/"/g, '""')}"`;

      lines.push(`${invNo},${invDate},${compName},${gstNo},${hsn},${qty},${basic},${sgst},${cgst},${igst},${total},${remarks}`);
    });
  }

  return lines.join('\r\n');
}

module.exports = {
  clubInvoiceItemsByHsn,
  collapseInvoiceToHighestHsn,
  generateCumulativeExcel,
  generateCumulativeCsv,
};
