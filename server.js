const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const { parseInvoiceWorkbook, naturalCompare } = require('./src/parser');
const { generateCumulativeExcel, generateCumulativeCsv } = require('./src/excelGenerator');
const { generateStatisticsExcel } = require('./src/statisticsGenerator');
const { runFullAudit } = require('./src/auditEngine');
const {
  parseAndGroupCumulativeSheet,
  generateHsnSummaryExcel,
  generateHsnSummaryCsv,
} = require('./src/hsnSummaryGenerator');

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS and body parsers
app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// Serve static frontend files
app.use(express.static(path.join(__dirname, 'public')));

// Configure Multer for in-memory batch processing of up to 250 files
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 30 * 1024 * 1024, // 30 MB per file
    files: 250, // Up to 250 files at once
  },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xlsx' || ext === '.xls') {
      cb(null, true);
    } else {
      cb(new Error(`Unsupported file type: ${file.originalname}. Only .xlsx and .xls are supported.`));
    }
  },
});

/**
 * Healthcheck endpoint for Render / monitoring
 */
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'ok', uptime: process.uptime() });
});

/**
 * Upload & Process Excel Invoices (supports 1 to 200+ files)
 */
app.post('/api/upload', upload.array('invoices', 250), async (req, res) => {
  try {
    const files = req.files;
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'No Excel files were uploaded.' });
    }

    const sortOrder = req.body.sortOrder || 'invoiceNo'; // 'invoiceNo', 'date', or 'file'
    const results = [];
    const errors = [];

    for (const file of files) {
      try {
        const parsed = parseInvoiceWorkbook(file.buffer, file.originalname);
        results.push(parsed);
      } catch (err) {
        console.error(`Error parsing file ${file.originalname}:`, err.message);
        errors.push({ filename: file.originalname, error: err.message });
      }
    }

    // Sort invoices
    if (sortOrder === 'invoiceNo') {
      results.sort((a, b) => naturalCompare(a.invoiceNo || '', b.invoiceNo || ''));
    } else if (sortOrder === 'date') {
      results.sort((a, b) => (a.invoiceDate || '').localeCompare(b.invoiceDate || ''));
    }

    // Calculate overall aggregates
    let totalInvoices = results.length;
    let totalItems = 0;
    let totalBasicAmount = 0;
    let totalSgst = 0;
    let totalCgst = 0;
    let totalIgst = 0;

    results.forEach(inv => {
      totalItems += inv.items.length;
      totalBasicAmount += inv.totalBasicAmount;
      totalSgst += inv.totalSgst;
      totalCgst += inv.totalCgst;
      totalIgst += inv.totalIgst;
    });

    const totalGrand = totalBasicAmount + totalSgst + totalCgst + totalIgst;

    const highestHsnOnly = req.body.highestHsnOnly === 'true' || req.body.highestHsnOnly === true;
    const clubByHsn = req.body.clubByHsn === 'true' || req.body.clubByHsn === true;
    const auditReport = runFullAudit(results, { highestHsnOnly, clubByHsn });

    res.json({
      success: true,
      processedCount: results.length,
      errorCount: errors.length,
      errors,
      auditReport,
      summary: {
        totalInvoices,
        totalItems,
        totalBasicAmount: Math.round(totalBasicAmount * 100) / 100,
        totalSgst: Math.round(totalSgst * 100) / 100,
        totalCgst: Math.round(totalCgst * 100) / 100,
        totalIgst: Math.round(totalIgst * 100) / 100,
        totalGrand: Math.round(totalGrand * 100) / 100,
      },
      invoices: results,
    });
  } catch (err) {
    console.error('Upload processing error:', err);
    res.status(500).json({ error: err.message || 'Internal server error while processing invoices.' });
  }
});

/**
 * Generate and download cumulative Excel (.xlsx) file matching Image 3
 */
app.post('/api/export-excel', async (req, res) => {
  try {
    const { invoices, periodTitle, repeatInvoiceInfo, clubByHsn, highestHsnOnly } = req.body;
    if (!invoices || !Array.isArray(invoices) || invoices.length === 0) {
      return res.status(400).json({ error: 'No invoice data provided for export.' });
    }

    const buffer = await generateCumulativeExcel(invoices, {
      periodTitle: periodTitle || 'April-June 2025-26',
      repeatInvoiceInfo: repeatInvoiceInfo === true,
      clubByHsn: clubByHsn === true,
      highestHsnOnly: highestHsnOnly === true,
    });

    const filename = `OUT_INV_${(periodTitle || 'Cumulative').replace(/[^a-zA-Z0-9_-]/g, '_')}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('Excel generation error:', err);
    res.status(500).json({ error: 'Failed to generate Excel file.' });
  }
});

/**
 * Generate and download cumulative CSV file matching Image 3
 */
app.post('/api/export-csv', (req, res) => {
  try {
    const { invoices, periodTitle, repeatInvoiceInfo, clubByHsn, highestHsnOnly } = req.body;
    if (!invoices || !Array.isArray(invoices) || invoices.length === 0) {
      return res.status(400).json({ error: 'No invoice data provided for export.' });
    }

    const csvContent = generateCumulativeCsv(invoices, {
      periodTitle: periodTitle || 'April-June 2025-26',
      repeatInvoiceInfo: repeatInvoiceInfo === true,
      clubByHsn: clubByHsn === true,
      highestHsnOnly: highestHsnOnly === true,
    });

    const filename = `OUT_INV_${(periodTitle || 'Cumulative').replace(/[^a-zA-Z0-9_-]/g, '_')}.csv`;

    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(csvContent);
  } catch (err) {
    console.error('CSV generation error:', err);
    res.status(500).json({ error: 'Failed to generate CSV file.' });
  }
});

/**
 * Generate and download comprehensive Statistics & Analytics Excel workbook
 */
app.post('/api/export-statistics', async (req, res) => {
  try {
    const { invoices, periodTitle } = req.body;
    if (!invoices || !Array.isArray(invoices) || invoices.length === 0) {
      return res.status(400).json({ error: 'No invoice data provided for statistics.' });
    }

    const buffer = await generateStatisticsExcel(invoices, {
      periodTitle: periodTitle || 'April-June 2025-26',
    });

    const filename = `GST_Statistics_Report_${(periodTitle || 'Summary').replace(/[^a-zA-Z0-9_-]/g, '_')}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('Statistics generation error:', err);
    res.status(500).json({ error: 'Failed to generate statistics Excel.' });
  }
});

/**
 * Return sample invoices for instant 1-click test drive
 */
app.get('/api/sample-data', (req, res) => {
  try {
    const sampleDir = path.join(__dirname, 'sample_invoices');
    if (!fs.existsSync(sampleDir)) {
      return res.status(404).json({ error: 'Sample files not found.' });
    }

    const files = fs.readdirSync(sampleDir).filter(f => f.startsWith('Invoice_') && f.endsWith('.xlsx'));
    const results = [];

    for (const f of files) {
      const buf = fs.readFileSync(path.join(sampleDir, f));
      const parsed = parseInvoiceWorkbook(buf, f);
      results.push(parsed);
    }

    results.sort((a, b) => naturalCompare(a.invoiceNo || '', b.invoiceNo || ''));

    let totalInvoices = results.length;
    let totalItems = 0;
    let totalBasicAmount = 0;
    let totalSgst = 0;
    let totalCgst = 0;
    let totalIgst = 0;

    results.forEach(inv => {
      totalItems += inv.items.length;
      totalBasicAmount += inv.totalBasicAmount;
      totalSgst += inv.totalSgst;
      totalCgst += inv.totalCgst;
      totalIgst += inv.totalIgst;
    });

    const auditReport = runFullAudit(results, {});

    res.json({
      success: true,
      processedCount: results.length,
      auditReport,
      summary: {
        totalInvoices,
        totalItems,
        totalBasicAmount: Math.round(totalBasicAmount * 100) / 100,
        totalSgst: Math.round(totalSgst * 100) / 100,
        totalCgst: Math.round(totalCgst * 100) / 100,
        totalIgst: Math.round(totalIgst * 100) / 100,
        totalGrand: Math.round((totalBasicAmount + totalSgst + totalCgst + totalIgst) * 100) / 100,
      },
      invoices: results,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/**
 * Re-calculate Audit Report on the fly (when toggles change)
 */
app.post('/api/audit', (req, res) => {
  try {
    const { invoices, highestHsnOnly, clubByHsn } = req.body;
    if (!invoices || !Array.isArray(invoices)) {
      return res.status(400).json({ error: 'Invoices array required.' });
    }
    const auditReport = runFullAudit(invoices, {
      highestHsnOnly: highestHsnOnly === true,
      clubByHsn: clubByHsn === true,
    });
    res.json({ success: true, auditReport });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/**
 * Upload and process modified cumulative sales sheet for HSN-wise GST Summary
 */
app.post('/api/hsn-summary/upload', upload.single('cumulativeSheet'), (req, res) => {
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'Please upload a valid cumulative Excel file (.xlsx or .xls).' });
    }

    const periodTitleOverride = req.body.periodTitle ? req.body.periodTitle.trim() : '';
    const result = parseAndGroupCumulativeSheet(file.buffer, { periodTitleOverride });

    res.json({
      success: true,
      filename: file.originalname,
      periodTitle: result.periodTitle,
      rawRowCount: result.rawRowCount,
      groups: result.groups,
      summary: result.summary,
      auditReport: result.auditReport,
    });
  } catch (err) {
    console.error('HSN Summary parsing error:', err);
    res.status(500).json({ error: err.message || 'Failed to process cumulative sheet for HSN summary.' });
  }
});

/**
 * Export HSN-wise GST Summary as Excel (.xlsx) matching user reference image
 */
app.post('/api/hsn-summary/export-excel', async (req, res) => {
  try {
    const { groups, periodTitle } = req.body;
    if (!groups || !Array.isArray(groups) || groups.length === 0) {
      return res.status(400).json({ error: 'No HSN summary groups data provided for export.' });
    }

    const buffer = await generateHsnSummaryExcel(groups, {
      periodTitle: periodTitle || 'July–September 2026',
    });

    const safeTitle = (periodTitle || 'Summary').replace(/[^a-zA-Z0-9_-]/g, '_');
    const filename = `HSN_Wise_GST_Summary_${safeTitle}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('HSN Summary Excel export error:', err);
    res.status(500).json({ error: 'Failed to generate HSN Summary Excel workbook.' });
  }
});

/**
 * Export HSN-wise GST Summary as CSV
 */
app.post('/api/hsn-summary/export-csv', (req, res) => {
  try {
    const { groups, periodTitle } = req.body;
    if (!groups || !Array.isArray(groups) || groups.length === 0) {
      return res.status(400).json({ error: 'No HSN summary groups data provided for export.' });
    }

    const csvContent = generateHsnSummaryCsv(groups, {
      periodTitle: periodTitle || 'July–September 2026',
    });

    const safeTitle = (periodTitle || 'Summary').replace(/[^a-zA-Z0-9_-]/g, '_');
    const filename = `HSN_Wise_GST_Summary_${safeTitle}.csv`;

    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(csvContent);
  } catch (err) {
    console.error('HSN Summary CSV export error:', err);
    res.status(500).json({ error: 'Failed to generate HSN Summary CSV file.' });
  }
});

/**
 * Load sample cumulative sheet to test HSN Summary instantly
 */
app.get('/api/hsn-summary/sample', async (req, res) => {
  try {
    const sampleDir = path.join(__dirname, 'sample_invoices');
    if (!fs.existsSync(sampleDir)) {
      return res.status(404).json({ error: 'Sample files not found.' });
    }

    const files = fs.readdirSync(sampleDir).filter(f => f.startsWith('Invoice_') && f.endsWith('.xlsx'));
    const results = [];

    for (const f of files) {
      const buf = fs.readFileSync(path.join(sampleDir, f));
      const parsed = parseInvoiceWorkbook(buf, f);
      results.push(parsed);
    }

    results.sort((a, b) => naturalCompare(a.invoiceNo || '', b.invoiceNo || ''));

    // Generate cumulative Excel buffer in memory
    const cumBuffer = await generateCumulativeExcel(results, { periodTitle: 'July–September 2026' });

    // Parse into HSN summary
    const hsnResult = parseAndGroupCumulativeSheet(cumBuffer, { periodTitleOverride: 'July–September 2026' });

    res.json({
      success: true,
      periodTitle: hsnResult.periodTitle,
      rawRowCount: hsnResult.rawRowCount,
      groups: hsnResult.groups,
      summary: hsnResult.summary,
      auditReport: hsnResult.auditReport,
    });
  } catch (err) {
    console.error('Sample HSN Summary error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  const url = `http://localhost:${PORT}`;
  console.log(`=======================================================`);
  console.log(` TAX INVOICE AGGREGATOR SERVER RUNNING`);
  console.log(` Local URL: ${url}`);
  console.log(` Network:   http://0.0.0.0:${PORT}`);
  console.log(` Ready to aggregate 50 to 150+ Excel invoices!`);
  console.log(`=======================================================`);

  // Open default web browser automatically
  try {
    const { exec } = require('child_process');
    if (process.platform === 'win32') {
      exec(`start ${url}`);
    } else if (process.platform === 'darwin') {
      exec(`open ${url}`);
    } else {
      exec(`xdg-open ${url}`);
    }
  } catch (e) {
    // Non-critical, user can open manually
  }
});
