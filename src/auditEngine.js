/**
 * Audit & Inconsistency Engine for GST Invoices
 * 
 * Performs 4 key checks:
 * 1. Missing Data (Missing Invoice No, Date, Company, GSTIN, HSN, Qty, Amount)
 * 2. Sequence Gaps (Detects missing numbers in invoice series, e.g. TS-04, TS-06 to TS-07)
 * 3. Duplicate Invoices (Duplicate invoice numbers across multiple files)
 * 4. Clubbing Log (Details exactly which invoices and HSN codes were merged)
 */

const GSTIN_REGEX = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/i;

/**
 * Checks for missing fields in invoices and line items
 */
function checkMissingData(invoices) {
  const issues = [];

  invoices.forEach(inv => {
    const invLabel = inv.invoiceNo && inv.invoiceNo !== 'UNKNOWN' ? inv.invoiceNo : (inv.filename || 'Unknown File');
    const missingFields = [];

    if (!inv.invoiceNo || inv.invoiceNo === 'UNKNOWN') {
      missingFields.push('Invoice Number');
    }
    if (!inv.invoiceDate) {
      missingFields.push('Invoice Date');
    }
    if (!inv.companyName || /unknown/i.test(inv.companyName)) {
      missingFields.push('Company / Buyer Name');
    }
    if (!inv.gstNo) {
      missingFields.push('Buyer GSTIN');
    } else {
      const cleanGst = inv.gstNo.trim().toUpperCase();
      if (!GSTIN_REGEX.test(cleanGst)) {
        missingFields.push(`Invalid GSTIN format ("${inv.gstNo}" - must be 15-character alphanumeric)`);
      } else {
        const stateCode = parseInt(cleanGst.substring(0, 2), 10);
        if (isNaN(stateCode) || stateCode < 1 || (stateCode > 38 && stateCode !== 97 && stateCode !== 99)) {
          missingFields.push(`Invalid GSTIN State Code ("${cleanGst.substring(0, 2)}" in "${inv.gstNo}")`);
        }
      }
    }

    // Check line items
    if (!inv.items || inv.items.length === 0) {
      missingFields.push('No line items found');
    } else {
      inv.items.forEach((it, idx) => {
        const itemIssues = [];
        const rawHsn = String(it.hsnCode || '').trim();
        const cleanHsn = rawHsn.replace(/\s+/g, '');

        // 1. HSN Code format validation
        if (!rawHsn) {
          itemIssues.push('HSN Code is blank');
        } else if (!/^\d+$/.test(cleanHsn)) {
          itemIssues.push(`Non-numeric HSN Code ("${rawHsn}")`);
        } else if (cleanHsn.length < 4) {
          itemIssues.push(`Unusual HSN Format ("${rawHsn}" has only ${cleanHsn.length} digits, minimum 4 required)`);
        }

        // 2. Quantity & Basic Amount validation
        const qty = it.qty || 0;
        const basic = it.basicAmount || 0;
        if (qty < 0) {
          itemIssues.push(`Negative Quantity (${qty})`);
        } else if (basic > 0 && qty === 0) {
          itemIssues.push('Quantity is 0 or blank');
        }

        if (basic < 0) {
          itemIssues.push(`Negative Basic Amount (₹${basic})`);
        } else if (qty > 0 && basic === 0) {
          itemIssues.push('Basic Amount is 0 or blank');
        }

        // 3. Tax validation
        const sgst = it.sgst || 0;
        const cgst = it.cgst || 0;
        const igst = it.igst || 0;
        const remarks = it.remarks || '';

        // Intra-state SGST != CGST mismatch
        if ((sgst > 0 || cgst > 0) && Math.abs(sgst - cgst) > 0.05) {
          itemIssues.push(`SGST (₹${sgst.toFixed(2)}) != CGST (₹${cgst.toFixed(2)}) intra-state mismatch`);
        }

        // Dual tax conflict (both IGST and SGST/CGST on same item)
        if (igst > 0 && (sgst > 0 || cgst > 0)) {
          itemIssues.push(`Dual Tax Conflict (IGST ₹${igst.toFixed(2)} + SGST/CGST ₹${(sgst + cgst).toFixed(2)})`);
        }

        // Unclarified zero tax
        if (basic > 0 && sgst === 0 && cgst === 0 && igst === 0 && !/sez|nil|zero|exempt|export/i.test(remarks)) {
          itemIssues.push('Zero Tax without SEZ / NIL / Export indicator in remarks');
        }

        if (itemIssues.length > 0) {
          missingFields.push(`Item #${idx + 1}: ${itemIssues.join('; ')}`);
        }
      });
    }

    if (missingFields.length > 0) {
      issues.push({
        invoiceNo: inv.invoiceNo || 'UNKNOWN',
        filename: inv.filename || '',
        companyName: inv.companyName || '',
        missingFields,
      });
    }
  });

  return issues;
}

/**
 * Detects gaps in invoice numbering sequences (e.g. TS-01, TS-02, TS-05 -> missing TS-03, TS-04)
 */
function checkSequenceGaps(invoices) {
  const seriesMap = new Map();

  invoices.forEach(inv => {
    const rawNo = (inv.invoiceNo || '').trim();
    if (!rawNo || rawNo === 'UNKNOWN') return;

    // Match series patterns like "TS-01/25-26", "INV-100", "01/24-25", "TI-2025-001"
    const match = rawNo.match(/^(.*?)(0*)(\d+)(.*)$/);
    if (!match) return;

    const prefix = match[1];
    const padZeros = match[2];
    const num = parseInt(match[3], 10);
    const suffix = match[4];
    const padLength = (padZeros + match[3]).length;

    const seriesKey = `${prefix}___${suffix}`;
    if (!seriesMap.has(seriesKey)) {
      seriesMap.set(seriesKey, {
        prefix,
        suffix,
        padLength,
        numbers: [],
        rawMap: new Map(),
      });
    }

    const s = seriesMap.get(seriesKey);
    s.numbers.push(num);
    s.rawMap.set(num, rawNo);
  });

  const gapReports = [];

  seriesMap.forEach((s) => {
    if (s.numbers.length < 2) return;

    const sortedNums = Array.from(new Set(s.numbers)).sort((a, b) => a - b);
    const min = sortedNums[0];
    const max = sortedNums[sortedNums.length - 1];

    const presentSet = new Set(sortedNums);
    const missingNums = [];

    for (let n = min; n <= max; n++) {
      if (!presentSet.has(n)) {
        missingNums.push(n);
      }
    }

    if (missingNums.length > 0) {
      // Group consecutive missing numbers into readable ranges
      const ranges = [];
      let rangeStart = missingNums[0];
      let rangeEnd = missingNums[0];

      for (let i = 1; i < missingNums.length; i++) {
        if (missingNums[i] === rangeEnd + 1) {
          rangeEnd = missingNums[i];
        } else {
          ranges.push({ start: rangeStart, end: rangeEnd });
          rangeStart = missingNums[i];
          rangeEnd = missingNums[i];
        }
      }
      ranges.push({ start: rangeStart, end: rangeEnd });

      const formatNum = (n) => {
        const padded = String(n).padStart(s.padLength, '0');
        return `${s.prefix}${padded}${s.suffix}`;
      };

      const missingDescriptions = ranges.map(r => {
        if (r.start === r.end) {
          return formatNum(r.start);
        } else if (r.end === r.start + 1) {
          return `${formatNum(r.start)}, ${formatNum(r.end)}`;
        } else {
          const count = r.end - r.start + 1;
          return `${formatNum(r.start)} to ${formatNum(r.end)} (${count} invoices)`;
        }
      });

      gapReports.push({
        series: `${s.prefix}[NUM]${s.suffix}`,
        rangeAnalyzed: `${formatNum(min)} to ${formatNum(max)}`,
        totalPresent: sortedNums.length,
        totalMissing: missingNums.length,
        missingInvoices: missingDescriptions,
      });
    }
  });

  return gapReports;
}

/**
 * Checks for duplicate invoice numbers across uploaded files
 */
function checkDuplicateInvoices(invoices) {
  const counts = new Map();

  invoices.forEach(inv => {
    const no = (inv.invoiceNo || '').trim();
    if (!no || no === 'UNKNOWN') return;
    if (!counts.has(no)) {
      counts.set(no, []);
    }
    counts.get(no).push(inv.filename || 'File');
  });

  const duplicates = [];
  counts.forEach((files, no) => {
    if (files.length > 1) {
      duplicates.push({
        invoiceNo: no,
        occurrences: files.length,
        files,
      });
    }
  });

  return duplicates;
}

/**
 * Analyzes and details exactly which invoices had items clubbed
 */
function analyzeClubbing(invoices, options = {}) {
  const clubbedLogs = [];
  const highestHsnOnly = options.highestHsnOnly === true;
  const clubByHsn = options.clubByHsn === true;

  invoices.forEach(inv => {
    if (!inv.items || inv.items.length <= 1) return;

    const originalItemsCount = inv.items.length;
    const hsnBreakdown = new Map();

    inv.items.forEach(it => {
      const code = String(it.hsnCode || 'Unspecified').trim();
      if (!hsnBreakdown.has(code)) {
        hsnBreakdown.set(code, {
          hsnCode: code,
          count: 0,
          totalQty: 0,
          totalBasic: 0,
        });
      }
      const b = hsnBreakdown.get(code);
      b.count++;
      b.totalQty += (it.qty || 0);
      b.totalBasic += (it.basicAmount || 0);
    });

    const distinctHsns = Array.from(hsnBreakdown.values());
    const hasMultipleSameHsn = distinctHsns.some(h => h.count > 1);
    const hasMultipleDifferentHsn = distinctHsns.length > 1;

    if (highestHsnOnly) {
      // Find highest amount item
      let highestItem = inv.items[0];
      for (let i = 1; i < inv.items.length; i++) {
        if ((inv.items[i].basicAmount || 0) > (highestItem.basicAmount || 0)) {
          highestItem = inv.items[i];
        }
      }

      const topHsn = highestItem.hsnCode || 'Unspecified';
      const totalQty = inv.items.reduce((s, it) => s + (it.qty || 0), 0);
      const totalBasic = inv.items.reduce((s, it) => s + (it.basicAmount || 0), 0);

      const hsnDetails = distinctHsns.map(h => 
        `HSN ${h.hsnCode}: ₹${h.totalBasic.toLocaleString('en-IN', { minimumFractionDigits: 2 })} (${h.count} item${h.count > 1 ? 's' : ''}, Qty ${h.totalQty})`
      ).join(', ');

      clubbedLogs.push({
        invoiceNo: inv.invoiceNo,
        companyName: inv.companyName,
        action: '1 Row (Top HSN Collapse)',
        originalItemCount: originalItemsCount,
        resultItemCount: 1,
        selectedHsn: topHsn,
        totalQty,
        totalBasic,
        details: `Had ${originalItemsCount} items across ${distinctHsns.length} HSN code(s) [${hsnDetails}]. Collapsed into Top HSN ${topHsn} with total basic ₹${totalBasic.toLocaleString('en-IN', { minimumFractionDigits: 2 })}.`,
      });
    } else if (clubByHsn) {
      if (hasMultipleSameHsn) {
        const mergedHsns = distinctHsns.filter(h => h.count > 1);
        const details = mergedHsns.map(h => 
          `${h.count} items with HSN ${h.hsnCode} merged into 1 row (Total Qty: ${h.totalQty}, Total Basic: ₹${h.totalBasic.toLocaleString('en-IN', { minimumFractionDigits: 2 })})`
        ).join('; ');

        clubbedLogs.push({
          invoiceNo: inv.invoiceNo,
          companyName: inv.companyName,
          action: 'Same HSN Clubbing',
          originalItemCount: originalItemsCount,
          resultItemCount: distinctHsns.length,
          mergedHsnCodes: mergedHsns.map(h => h.hsnCode),
          details,
        });
      }
    } else {
      // Informational: noting invoices that have multiple items / multiple HSNs
      if (hasMultipleDifferentHsn || hasMultipleSameHsn) {
        const details = distinctHsns.map(h => 
          `HSN ${h.hsnCode} (${h.count} item${h.count > 1 ? 's' : ''}, Qty ${h.totalQty}, ₹${h.totalBasic.toLocaleString('en-IN', { minimumFractionDigits: 2 })})`
        ).join(', ');

        clubbedLogs.push({
          invoiceNo: inv.invoiceNo,
          companyName: inv.companyName,
          action: 'Itemized (Multiple Lines)',
          originalItemCount: originalItemsCount,
          resultItemCount: originalItemsCount,
          details: `Invoice contains ${originalItemsCount} items across ${distinctHsns.length} HSN code(s): ${details}. Currently shown as separate rows.`,
        });
      }
    }
  });

  return clubbedLogs;
}

/**
 * Runs full audit checks and returns structured consistency report
 */
function runFullAudit(invoices, options = {}) {
  const missingData = checkMissingData(invoices);
  const sequenceGaps = checkSequenceGaps(invoices);
  const duplicates = checkDuplicateInvoices(invoices);
  const clubbingLogs = analyzeClubbing(invoices, options);

  const totalInconsistencyCount = 
    missingData.length + 
    sequenceGaps.reduce((acc, g) => acc + g.totalMissing, 0) + 
    duplicates.length;

  return {
    hasInconsistencies: totalInconsistencyCount > 0,
    totalInconsistencyCount,
    missingDataCount: missingData.length,
    missingSequenceCount: sequenceGaps.reduce((acc, g) => acc + g.totalMissing, 0),
    duplicateCount: duplicates.length,
    clubbedInvoicesCount: clubbingLogs.length,
    missingData,
    sequenceGaps,
    duplicates,
    clubbingLogs,
  };
}

module.exports = {
  runFullAudit,
  checkMissingData,
  checkSequenceGaps,
  checkDuplicateInvoices,
  analyzeClubbing,
};
