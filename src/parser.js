const xlsx = require('xlsx');

/**
 * Normalizes dates to D/M/YYYY format matching Image 3 (e.g. 15/4/2025)
 */
function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return `${val.getDate()}/${val.getMonth() + 1}/${val.getFullYear()}`;
  }

  // If it's an Excel serial date number
  if (typeof val === 'number') {
    const dateObj = xlsx.SSF.parse_date_code(val);
    if (dateObj) {
      return `${dateObj.d}/${dateObj.m}/${dateObj.y}`;
    }
  }

  const str = String(val).trim();
  // Match D/M/YYYY or M/D/YYYY or YYYY-M-D
  const isoMatch = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (isoMatch) {
    return `${parseInt(isoMatch[3], 10)}/${parseInt(isoMatch[2], 10)}/${isoMatch[1]}`;
  }

  const slashMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (slashMatch) {
    const p1 = parseInt(slashMatch[1], 10);
    const p2 = parseInt(slashMatch[2], 10);
    const year = slashMatch[3];
    // If p1 > 12, it must be day/month/year
    if (p1 > 12) {
      return `${p1}/${p2}/${year}`;
    }
    // If p2 > 12, then p1 is month, p2 is day (e.g. 4/15/2025 -> 15/4/2025)
    if (p2 > 12) {
      return `${p2}/${p1}/${year}`;
    }
    // Default to day/month/year or month/day/year based on context
    return `${p1}/${p2}/${year}`;
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return `${parsed.getDate()}/${parsed.getMonth() + 1}/${parsed.getFullYear()}`;
  }

  return str;
}

/**
 * Natural sort comparator for invoice numbers (e.g. TS-01, TS-02, ..., TS-10, TS-100)
 */
function naturalCompare(a, b) {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

/**
 * GSTIN Regex for Indian GST Numbers
 */
const GSTIN_REGEX = /\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}[Z]{1}[0-9A-Z]{1})\b/i;

/**
 * Clean numeric values from strings (e.g. "256,500.00" -> 256500)
 */
function parseNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const cleaned = String(val).replace(/,/g, '').trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

/**
 * Parse an Excel invoice workbook (buffer or file path)
 */
function parseInvoiceWorkbook(buffer, originalFilename = '') {
  const workbook = xlsx.read(buffer, {
    type: 'buffer',
    cellDates: true,
    cellNF: true,
    cellText: true,
  });

  // Search sheets for the invoice sheet (prefer 'ORIGINAL', 'Tax Invoice', or sheet containing "TAX INVOICE")
  let targetSheetName = workbook.SheetNames[0];
  for (const name of workbook.SheetNames) {
    if (/original|tax.*inv|invoice/i.test(name)) {
      targetSheetName = name;
      break;
    }
  }

  const sheet = workbook.Sheets[targetSheetName];
  if (!sheet) {
    throw new Error(`No sheets found in ${originalFilename}`);
  }

  // Convert sheet to 2D array of raw values and formulas
  const range = xlsx.utils.decode_range(sheet['!ref'] || 'A1:Z100');
  const grid = [];
  for (let r = 0; r <= range.e.r; r++) {
    const row = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cellAddress = xlsx.utils.encode_cell({ r, c });
      const cell = sheet[cellAddress];
      let val = '';
      if (cell) {
        val = cell.w !== undefined ? cell.w : (cell.v !== undefined ? cell.v : '');
      }
      row.push(val);
    }
    grid.push(row);
  }

  // Extract Invoice Metadata
  let invoiceNo = '';
  let invoiceDate = '';
  let consigneeName = '';
  let buyerName = '';
  let consigneeGstin = '';
  let buyerGstin = '';
  let sellerGstin = '';
  let isSez = false;
  let taxRatePercent = 18; // Standard default 18%
  let hasIgst = false;
  let hasCgstSgst = false;
  let taxNilOrZero = false;

  // 1. Scan for Invoice No and Dated
  for (let r = 0; r < Math.min(grid.length, 30); r++) {
    for (let c = 0; c < grid[r].length; c++) {
      const cell = String(grid[r][c] || '').trim();

      // Look for Invoice No label
      if (/^invoice\s*no\.?$/i.test(cell) || /inv\s*no\.?/i.test(cell)) {
        // Value might be below (r+1) or to the right (c+1)
        if (grid[r + 1] && grid[r + 1][c]) {
          const cand = String(grid[r + 1][c]).trim();
          if (cand && !/^dated/i.test(cand) && !invoiceNo) invoiceNo = cand;
        }
        if (!invoiceNo && grid[r][c + 1]) {
          const cand = String(grid[r][c + 1]).trim();
          if (cand && !invoiceNo) invoiceNo = cand;
        }
      } else if (/invoice\s*no\.?\s*[:=]\s*(.+)/i.test(cell)) {
        const match = cell.match(/invoice\s*no\.?\s*[:=]\s*(.+)/i);
        if (match && !invoiceNo) invoiceNo = match[1].trim();
      }

      // Look for Dated label
      if (/^dated$/i.test(cell) || /^inv\s*date$/i.test(cell) || /^invoice\s*date$/i.test(cell)) {
        if (grid[r + 1] && grid[r + 1][c]) {
          const cand = grid[r + 1][c];
          if (cand && !invoiceDate) invoiceDate = formatDate(cand);
        }
        if (!invoiceDate && grid[r][c + 1]) {
          const cand = grid[r][c + 1];
          if (cand && !invoiceDate) invoiceDate = formatDate(cand);
        }
      } else if (/dated?\s*[:=]\s*(.+)/i.test(cell)) {
        const match = cell.match(/dated?\s*[:=]\s*(.+)/i);
        if (match && !invoiceDate) invoiceDate = formatDate(match[1]);
      }

      // Look for Consignee (only in left columns c <= 4)
      if (c <= 4 && (/^consignee$/i.test(cell) || /^consignee\s*[:=]/i.test(cell))) {
        for (let ro = 1; ro <= 4; ro++) {
          if (grid[r + ro] && grid[r + ro][c]) {
            const cand = String(grid[r + ro][c]).trim();
            if (cand && !/address|gstin|state|buyer/i.test(cand) && !consigneeName) {
              consigneeName = cand;
              break;
            }
          }
        }
      }

      // Look for Buyer (exclude 'buyer order' / 'buyer's order' / 'order no', only in left columns c <= 4)
      if (c <= 4 && /buyer(?!\s*['’]s?\s*order)/i.test(cell) && !/order/i.test(cell)) {
        for (let ro = 1; ro <= 4; ro++) {
          if (grid[r + ro] && grid[r + ro][c]) {
            const cand = String(grid[r + ro][c]).trim();
            if (cand && !/address|gstin|state|other\s+than|order/i.test(cand) && !buyerName) {
              buyerName = cand;
              break;
            }
          }
        }
      }

      // Look for GSTIN in Consignee / Buyer area
      const gstinMatch = cell.match(GSTIN_REGEX);
      if (gstinMatch) {
        const gstinFound = gstinMatch[1].toUpperCase();
        if (r < 25) {
          // It's in the top section (buyer/consignee)
          if (!buyerGstin) buyerGstin = gstinFound;
          else if (!consigneeGstin) consigneeGstin = gstinFound;
        } else if (r > 40) {
          // Bottom section is seller GSTIN
          sellerGstin = gstinFound;
        }
      }

      // Check for SEZ mentions in the header
      if (/sez|special\s+economic\s+zone|lut\b|zero\s*rated/i.test(cell)) {
        isSez = true;
      }
    }
  }

  // Fallback for Invoice No if not found: regex match TS-XX/XX-XX or similar pattern
  if (!invoiceNo) {
    for (let r = 0; r < Math.min(grid.length, 30); r++) {
      for (let c = 0; c < grid[r].length; c++) {
        const cell = String(grid[r][c] || '').trim();
        const m = cell.match(/\b([A-Z]{1,5}-\d{1,5}\/\d{2}-\d{2})\b/i);
        if (m) {
          invoiceNo = m[1];
          break;
        }
      }
      if (invoiceNo) break;
    }
  }

  // Fallback for Invoice No from filename (e.g. TS-01.xlsx or TS-01_25-26.xlsx)
  if (!invoiceNo && originalFilename) {
    const fnMatch = originalFilename.match(/\b(TS-?\d+[\w-]*)\b/i);
    if (fnMatch) invoiceNo = fnMatch[1];
    else invoiceNo = originalFilename.replace(/\.xlsx?$/i, '');
  }

  // Fallback for Seller GSTIN
  if (!sellerGstin) {
    for (let r = grid.length - 1; r >= Math.max(0, grid.length - 25); r--) {
      for (let c = 0; c < grid[r].length; c++) {
        const cell = String(grid[r][c] || '').trim();
        const m = cell.match(GSTIN_REGEX);
        if (m) {
          sellerGstin = m[1].toUpperCase();
          break;
        }
      }
      if (sellerGstin) break;
    }
  }

  // Bottom section: Check Tax Rows (Sub-Total, IGST, CGST, SGST, NIL)
  for (let r = 25; r < grid.length; r++) {
    const rowStr = grid[r].map(v => String(v || '').trim()).join(' ');

    if (/igst/i.test(rowStr)) {
      hasIgst = true;
      const rateM = rowStr.match(/igst\s*@?\s*(\d+(\.\d+)?)%/i);
      if (rateM) taxRatePercent = parseFloat(rateM[1]);
    }
    if (/cgst/i.test(rowStr) || /sgst/i.test(rowStr)) {
      hasCgstSgst = true;
      const rateM = rowStr.match(/(?:cgst|sgst)\s*@?\s*(\d+(\.\d+)?)%/i);
      if (rateM) taxRatePercent = parseFloat(rateM[1]) * 2; // e.g. CGST 9% means total 18%
    }
    if (/nil|zero|exempt|sez/i.test(rowStr)) {
      if (/gst\s*nil|nil\s*tax|tax\s*nil|sez/i.test(rowStr)) {
        taxNilOrZero = true;
        isSez = true;
      }
    }
  }

  // Determine Company Name
  const companyName = buyerName || consigneeName || 'Unknown Company';
  const finalGstin = buyerGstin || consigneeGstin || '';

  // Determine Seller State Code and Buyer State Code
  const sellerStateCode = sellerGstin ? sellerGstin.substring(0, 2) : '36'; // Default 36 (Telangana)
  const buyerStateCode = finalGstin ? finalGstin.substring(0, 2) : '';

  // Inter-state vs Intra-state:
  // If explicitly has IGST in invoice or buyerStateCode != sellerStateCode, it's IGST
  // If explicitly has CGST/SGST or buyerStateCode === sellerStateCode, it's CGST + SGST
  let taxType = 'IGST';
  if (hasCgstSgst || (buyerStateCode && buyerStateCode === sellerStateCode)) {
    taxType = 'CGST_SGST';
  } else if (hasIgst || (buyerStateCode && buyerStateCode !== sellerStateCode)) {
    taxType = 'IGST';
  }

  // 2. Identify Table Header Row (S No, Description, HSN/SAC, Qty, Rate, Amount)
  let headerRowIndex = -1;
  let colSNo = 0;
  let colDesc = 2;
  let colHsn = 6;
  let colQty = 7;
  let colRate = 8;
  let colAmount = 9;

  for (let r = 15; r < Math.min(grid.length, 35); r++) {
    const row = grid[r];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim().toLowerCase();
      if (/^s\s*\.?\s*no/i.test(cell) || /^sl\s*\.?\s*no/i.test(cell) || /^item\s*no/i.test(cell)) {
        colSNo = c;
        headerRowIndex = r;
      }
      if (/description/i.test(cell) || /item\s*name/i.test(cell) || /particulars/i.test(cell)) {
        colDesc = c;
      }
      if (/hsn/i.test(cell) || /sac/i.test(cell)) {
        colHsn = c;
      }
      if (/^qty/i.test(cell) || /quantity/i.test(cell)) {
        colQty = c;
      }
      if (/rate/i.test(cell)) {
        colRate = c;
      }
      if (/amount/i.test(cell)) {
        colAmount = c;
      }
    }
    if (headerRowIndex !== -1) break;
  }

  if (headerRowIndex === -1) {
    // Fallback: look for row 22 (0-indexed 21)
    headerRowIndex = 21;
    colSNo = 0;
    colDesc = 2;
    colHsn = 6;
    colQty = 7;
    colRate = 8;
    colAmount = 9;
  }

  // 3. Extract Line Items
  const items = [];
  let currentItem = null;

  for (let r = headerRowIndex + 1; r < grid.length; r++) {
    const row = grid[r];
    const rowStr = row.map(v => String(v || '').trim()).join(' ');

    // Stop if we reach Sub-Total, Total, Amount in Words, or Bank Details
    if (/sub-?total|total|amount\s+chargeable|indian\s+rupees|bank\s+details|authorised\s+signatory/i.test(rowStr)) {
      break;
    }

    const sNoCell = String(row[colSNo] || '').trim();
    const hsnCell = String(row[colHsn] || '').trim();
    const qtyCell = parseNumber(row[colQty]);
    const rateCell = parseNumber(row[colRate]);
    const amountCell = parseNumber(row[colAmount]);
    const descCell = String(row[colDesc] || '').trim();

    // Check if this row is a new item (has S No like 1, 2, 3... or has valid HSN and amount)
    const isNewItem = (/^\d+$/.test(sNoCell) && parseInt(sNoCell, 10) > 0) || (hsnCell && amountCell > 0);

    if (isNewItem) {
      if (currentItem) {
        items.push(currentItem);
      }
      currentItem = {
        sNo: sNoCell ? parseInt(sNoCell, 10) : items.length + 1,
        description: descCell,
        hsnCode: hsnCell || '',
        qty: qtyCell,
        rate: rateCell,
        basicAmount: amountCell,
      };
    } else if (currentItem) {
      // Continuation line (e.g. description multi-line, or HSN / amount on subsequent row)
      if (descCell) {
        currentItem.description += ' ' + descCell;
      }
      if (!currentItem.hsnCode && hsnCell) {
        currentItem.hsnCode = hsnCell;
      }
      if (!currentItem.qty && qtyCell > 0) {
        currentItem.qty = qtyCell;
      }
      if (!currentItem.rate && rateCell > 0) {
        currentItem.rate = rateCell;
      }
      if (!currentItem.basicAmount && amountCell > 0) {
        currentItem.basicAmount = amountCell;
      }
    }
  }

  if (currentItem) {
    items.push(currentItem);
  }

  // Filter out any empty items
  const validItems = items.filter(it => it.basicAmount > 0 || it.hsnCode || it.qty > 0);

  // If no items found, create a fallback item with total if available
  if (validItems.length === 0) {
    // Look for subtotal or amount
    let fallbackAmount = 0;
    for (let r = 20; r < grid.length; r++) {
      for (let c = 0; c < grid[r].length; c++) {
        if (/sub-?total/i.test(String(grid[r][c] || ''))) {
          for (let cc = c + 1; cc < grid[r].length; cc++) {
            const num = parseNumber(grid[r][cc]);
            if (num > 0) {
              fallbackAmount = num;
              break;
            }
          }
        }
      }
      if (fallbackAmount > 0) break;
    }
    validItems.push({
      sNo: 1,
      description: 'Invoice Item',
      hsnCode: '',
      qty: 1,
      rate: fallbackAmount,
      basicAmount: fallbackAmount,
    });
  }

  // Calculate tax per item
  const processedItems = validItems.map(item => {
    let sgst = 0;
    let cgst = 0;
    let igst = 0;
    let remarks = '';

    if (isSez || taxNilOrZero) {
      remarks = 'GST NIL sez';
      sgst = 0;
      cgst = 0;
      igst = 0;
    } else if (taxType === 'CGST_SGST') {
      const halfRate = (taxRatePercent / 2) / 100;
      sgst = Math.round(item.basicAmount * halfRate * 100) / 100;
      cgst = Math.round(item.basicAmount * halfRate * 100) / 100;
      igst = 0;
    } else {
      // IGST
      const fullRate = taxRatePercent / 100;
      igst = Math.round(item.basicAmount * fullRate * 100) / 100;
      sgst = 0;
      cgst = 0;
    }

    return {
      hsnCode: item.hsnCode,
      qty: item.qty,
      basicAmount: item.basicAmount,
      sgst,
      cgst,
      igst,
      description: item.description,
      remarks,
    };
  });

  return {
    filename: originalFilename,
    invoiceNo: invoiceNo || 'UNKNOWN',
    invoiceDate: invoiceDate || '',
    companyName: companyName,
    gstNo: finalGstin,
    isSez,
    taxType,
    taxRatePercent,
    items: processedItems,
    totalBasicAmount: processedItems.reduce((acc, it) => acc + it.basicAmount, 0),
    totalSgst: processedItems.reduce((acc, it) => acc + it.sgst, 0),
    totalCgst: processedItems.reduce((acc, it) => acc + it.cgst, 0),
    totalIgst: processedItems.reduce((acc, it) => acc + it.igst, 0),
  };
}

module.exports = {
  parseInvoiceWorkbook,
  formatDate,
  naturalCompare,
};
