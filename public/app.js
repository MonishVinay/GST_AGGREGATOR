// State
let selectedFiles = [];
let processedInvoices = [];
let currentSummary = null;
let currentAuditReport = null;

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileSelectedBadge = document.getElementById('fileSelectedBadge');
const fileCountNumber = document.getElementById('fileCountNumber');
const periodTitleInput = document.getElementById('periodTitleInput');
const sortOrderSelect = document.getElementById('sortOrderSelect');
const repeatInvoiceInfoToggle = document.getElementById('repeatInvoiceInfoToggle');
const clubByHsnToggle = document.getElementById('clubByHsnToggle');
const btnQuickClubToggle = document.getElementById('btnQuickClubToggle');
const quickClubText = document.getElementById('quickClubText');
const highestHsnToggle = document.getElementById('highestHsnToggle');
const btnQuickHighestToggle = document.getElementById('btnQuickHighestToggle');
const quickHighestText = document.getElementById('quickHighestText');
const btnProcess = document.getElementById('btnProcess');
const processSpinner = document.getElementById('processSpinner');
const btnLoadSample = document.getElementById('btnLoadSample');
const progressContainer = document.getElementById('progressContainer');
const progressBar = document.getElementById('progressBar');
const progressStatus = document.getElementById('progressStatus');
const progressPercent = document.getElementById('progressPercent');

const resultsSection = document.getElementById('resultsSection');
const auditCard = document.getElementById('auditCard');
const auditStatusIcon = document.getElementById('auditStatusIcon');
const auditTitleText = document.getElementById('auditTitleText');
const auditSubtitleText = document.getElementById('auditSubtitleText');
const auditMainBadge = document.getElementById('auditMainBadge');
const auditBadgesArea = document.getElementById('auditBadgesArea');
const badgeCountSeq = document.getElementById('badgeCountSeq');
const badgeCountMissing = document.getElementById('badgeCountMissing');
const badgeCountClubbed = document.getElementById('badgeCountClubbed');
const badgeCountDup = document.getElementById('badgeCountDup');
const auditSeqList = document.getElementById('auditSeqList');
const auditMissingList = document.getElementById('auditMissingList');
const auditClubbedList = document.getElementById('auditClubbedList');
const auditDupList = document.getElementById('auditDupList');

const metricTotalInvoices = document.getElementById('metricTotalInvoices');
const metricTotalItems = document.getElementById('metricTotalItems');
const metricBasicAmount = document.getElementById('metricBasicAmount');
const metricSgst = document.getElementById('metricSgst');
const metricCgst = document.getElementById('metricCgst');
const metricIgst = document.getElementById('metricIgst');
const metricGrandTotal = document.getElementById('metricGrandTotal');

const tableHeaderPeriod = document.getElementById('tableHeaderPeriod');
const reportPeriodTag = document.getElementById('reportPeriodTag');
const tableBody = document.getElementById('tableBody');
const tableFoot = document.getElementById('tableFoot');
const searchInput = document.getElementById('searchInput');
const rowCountInfo = document.getElementById('rowCountInfo');

const btnExportExcel = document.getElementById('btnExportExcel');
const btnExportStats = document.getElementById('btnExportStats');
const btnExportCsv = document.getElementById('btnExportCsv');
const btnCopyTable = document.getElementById('btnCopyTable');

const deployModal = document.getElementById('deployModal');
const btnOpenDeployModal = document.getElementById('btnOpenDeployModal');
const btnCloseDeployModal = document.getElementById('btnCloseDeployModal');
const tabButtons = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');

// Number formatter for Indian Rupee style (e.g. 2,56,500.00)
const formatCurrency = (num) => {
  if (num === null || num === undefined || isNaN(num)) return '₹ 0.00';
  return '₹ ' + Number(num).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
};

const formatNumber = (num, decimals = 2) => {
  if (num === null || num === undefined || isNaN(num) || num === 0) return '0.00';
  return Number(num).toLocaleString('en-IN', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
};

// GSTIN & HSN Validation Helpers
const GSTIN_REGEX = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/i;

function validateHsnCode(rawHsn) {
  const hsn = String(rawHsn || '').trim();
  if (!hsn) {
    return { valid: false, error: 'Blank HSN', isError: true };
  }
  const clean = hsn.replace(/\s+/g, '');
  if (!/^\d+$/.test(clean)) {
    return { valid: false, error: 'Non-numeric HSN', isError: true, clean };
  }
  if (clean.length < 4) {
    return { valid: false, error: `Short HSN (${clean.length} digits)`, isWarning: true, clean };
  }
  return { valid: true, clean };
}

function validateGstinFormat(rawGstin) {
  const g = String(rawGstin || '').trim().toUpperCase();
  if (!g) {
    return { valid: false, error: 'Missing GSTIN', isError: true };
  }
  if (!GSTIN_REGEX.test(g)) {
    return { valid: false, error: 'Invalid 15-char GSTIN', isWarning: true };
  }
  const stateCode = parseInt(g.substring(0, 2), 10);
  if (isNaN(stateCode) || stateCode < 1 || (stateCode > 38 && stateCode !== 97 && stateCode !== 99)) {
    return { valid: false, error: `Invalid State Code (${g.substring(0, 2)})`, isWarning: true };
  }
  return { valid: true };
}

// Drag and Drop listeners
dropZone.addEventListener('click', () => fileInput.click());

['dragenter', 'dragover'].forEach(eventName => {
  dropZone.addEventListener(eventName, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.add('drag-over');
  });
});

['dragleave', 'drop'].forEach(eventName => {
  dropZone.addEventListener(eventName, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove('drag-over');
  });
});

dropZone.addEventListener('drop', (e) => {
  const dt = e.dataTransfer;
  if (dt.files && dt.files.length > 0) {
    handleFilesSelected(dt.files);
  }
});

fileInput.addEventListener('change', (e) => {
  if (e.target.files && e.target.files.length > 0) {
    handleFilesSelected(e.target.files);
  }
});

function handleFilesSelected(files) {
  const validFiles = Array.from(files).filter(f => {
    const ext = f.name.toLowerCase();
    return ext.endsWith('.xlsx') || ext.endsWith('.xls');
  });

  if (validFiles.length === 0) {
    alert('Please select valid Excel files (.xlsx or .xls).');
    return;
  }

  selectedFiles = validFiles;
  fileCountNumber.textContent = selectedFiles.length;
  fileSelectedBadge.classList.remove('hidden');
  btnProcess.disabled = false;
  btnProcess.querySelector('.btn-text').textContent = `Process ${selectedFiles.length} Invoices`;
}

// Upload & Process invoices with progress bar
btnProcess.addEventListener('click', () => {
  if (selectedFiles.length === 0) return;

  const formData = new FormData();
  selectedFiles.forEach(file => {
    formData.append('invoices', file);
  });
  formData.append('sortOrder', sortOrderSelect.value);

  btnProcess.disabled = true;
  processSpinner.classList.remove('hidden');
  progressContainer.classList.remove('hidden');
  progressBar.style.width = '0%';
  progressPercent.textContent = '0%';
  progressStatus.textContent = `Uploading ${selectedFiles.length} invoice files...`;

  const xhr = new XMLHttpRequest();
  xhr.open('POST', '/api/upload', true);

  xhr.upload.onprogress = (e) => {
    if (e.lengthComputable) {
      const percent = Math.round((e.loaded / e.total) * 75);
      progressBar.style.width = `${percent}%`;
      progressPercent.textContent = `${percent}%`;
      if (percent >= 75) {
        progressStatus.textContent = 'Processing and extracting tax invoice lines...';
      }
    }
  };

  xhr.onload = () => {
    btnProcess.disabled = false;
    processSpinner.classList.add('hidden');

    if (xhr.status === 200) {
      progressBar.style.width = '100%';
      progressPercent.textContent = '100%';
      progressStatus.textContent = 'Completed successfully!';

      setTimeout(() => {
        progressContainer.classList.add('hidden');
      }, 1200);

      const res = JSON.parse(xhr.responseText);
      processedInvoices = res.invoices;
      currentSummary = res.summary;
      currentAuditReport = res.auditReport;
      renderResults();
    } else {
      progressBar.style.width = '0%';
      progressContainer.classList.add('hidden');
      try {
        const err = JSON.parse(xhr.responseText);
        alert(`Error: ${err.error || 'Failed to process files.'}`);
      } catch (e) {
        alert('Server error occurred during processing.');
      }
    }
  };

  xhr.onerror = () => {
    btnProcess.disabled = false;
    processSpinner.classList.add('hidden');
    progressContainer.classList.add('hidden');
    alert('Network error while uploading. Please check connection and try again.');
  };

  xhr.send(formData);
});

// Load Sample Demo Data
btnLoadSample.addEventListener('click', async () => {
  btnLoadSample.disabled = true;
  btnLoadSample.textContent = 'Loading samples...';

  try {
    const res = await fetch('/api/sample-data');
    const data = await res.json();
    if (data.success) {
      processedInvoices = data.invoices;
      currentSummary = data.summary;
      currentAuditReport = data.auditReport;
      renderResults();
    } else {
      alert('Could not load sample data.');
    }
  } catch (err) {
    alert('Error loading sample data: ' + err.message);
  } finally {
    btnLoadSample.disabled = false;
    btnLoadSample.textContent = 'Try Demo Invoices';
  }
});

// Clubs items sharing the same HSN code within the SAME invoice
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

// Collapses each invoice into a single row using the HSN code of the highest amount item
function collapseInvoiceToHighestHsn(invoices) {
  return invoices.map(inv => {
    if (!inv.items || inv.items.length === 0) return inv;
    if (inv.items.length === 1) return inv;

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

function getDisplayInvoices() {
  if (highestHsnToggle.checked) {
    return collapseInvoiceToHighestHsn(processedInvoices);
  }
  if (clubByHsnToggle.checked) {
    return clubInvoiceItemsByHsn(processedInvoices);
  }
  return processedInvoices;
}

// Audit Tab switching (card-scoped so Module 1 and Module 2 operate independently)
document.addEventListener('click', (e) => {
  const btn = e.target.closest('.audit-tab-btn');
  if (!btn) return;
  const card = btn.closest('.audit-card');
  if (!card) return;

  const btns = card.querySelectorAll('.audit-tab-btn');
  const panels = card.querySelectorAll('.audit-panel');
  btns.forEach(b => b.classList.remove('active'));
  panels.forEach(p => p.classList.remove('active'));

  btn.classList.add('active');
  const targetId = btn.getAttribute('data-atab');
  const targetPanel = card.querySelector('#' + targetId);
  if (targetPanel) targetPanel.classList.add('active');
});

/**
 * Fetch and update live audit report (e.g. when clubbing options change)
 */
async function refreshAuditReport() {
  if (!processedInvoices || processedInvoices.length === 0) return;
  try {
    const res = await fetch('/api/audit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        invoices: processedInvoices,
        highestHsnOnly: highestHsnToggle.checked,
        clubByHsn: clubByHsnToggle.checked,
      }),
    });
    const data = await res.json();
    if (data.success && data.auditReport) {
      currentAuditReport = data.auditReport;
      renderAuditReport(currentAuditReport);
    }
  } catch (err) {
    console.error('Failed to refresh audit report:', err);
  }
}

/**
 * Renders the Audit & Inconsistencies card with sequence gaps, missing data, clubbed invoices log, and duplicates
 */
function renderAuditReport(audit) {
  if (!audit) {
    auditCard.classList.add('hidden');
    return;
  }
  auditCard.classList.remove('hidden');

  // Reset card state styling
  auditCard.classList.remove('audit-clean', 'audit-error');

  const seqCount = audit.missingSequenceCount || 0;
  const missingCount = audit.missingDataCount || 0;
  const dupCount = audit.duplicateCount || 0;
  const clubCount = audit.clubbedInvoicesCount || 0;
  const totalIssues = (audit.totalInconsistencyCount !== undefined)
    ? audit.totalInconsistencyCount
    : (seqCount + missingCount + dupCount);

  // Update tab badges
  badgeCountSeq.textContent = seqCount;
  badgeCountMissing.textContent = missingCount;
  badgeCountClubbed.textContent = clubCount;
  badgeCountDup.textContent = dupCount;

  // Header status & Icon
  if (totalIssues === 0) {
    auditCard.classList.add('audit-clean');
    auditStatusIcon.textContent = '✅';
    auditTitleText.textContent = 'Data Consistency & Integrity Check Passed';
    auditSubtitleText.textContent = 'All invoice sequences, company fields, GSTINs, and HSN codes verified without critical errors.';
    auditMainBadge.className = 'badge pill-green';
    auditMainBadge.textContent = '100% Consistent';
  } else {
    auditCard.classList.add('audit-error');
    auditStatusIcon.textContent = '⚠️';
    auditTitleText.textContent = `Found ${totalIssues} Inconsistency Warning${totalIssues > 1 ? 's' : ''}`;
    auditSubtitleText.textContent = 'Review missing sequence numbers, blank fields, or duplicate invoices below:';
    auditMainBadge.className = 'badge pill-amber';
    auditMainBadge.textContent = `${totalIssues} Discrepanc${totalIssues > 1 ? 'ies' : 'y'}`;
  }

  // Header quick pills
  let pillsHtml = '';
  if (seqCount > 0) pillsHtml += `<span class="audit-pill pill-amber">⚠️ ${seqCount} Missing Sequence Number${seqCount > 1 ? 's' : ''}</span>`;
  if (missingCount > 0) pillsHtml += `<span class="audit-pill pill-red">❌ ${missingCount} Invoice${missingCount > 1 ? 's' : ''} Missing Fields</span>`;
  if (dupCount > 0) pillsHtml += `<span class="audit-pill pill-red">🔁 ${dupCount} Duplicate Invoice${dupCount > 1 ? 's' : ''}</span>`;
  if (clubCount > 0) pillsHtml += `<span class="audit-pill pill-blue">ℹ️ ${clubCount} Invoice${clubCount > 1 ? 's' : ''} with Clubbed Items</span>`;
  auditBadgesArea.innerHTML = pillsHtml;

  // 1. Missing in Sequence List
  auditSeqList.innerHTML = '';
  if (!audit.sequenceGaps || audit.sequenceGaps.length === 0) {
    auditSeqList.innerHTML = `
      <div class="audit-empty-state">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>
        <span>No missing sequence numbers. Invoice sequence is complete and continuous!</span>
      </div>
    `;
  } else {
    audit.sequenceGaps.forEach(gap => {
      const itemEl = document.createElement('div');
      itemEl.className = 'audit-item';
      const missingList = gap.missingInvoices && gap.missingInvoices.length > 0 
        ? gap.missingInvoices.join('; ')
        : (gap.missingNumbers ? gap.missingNumbers.join(', ') : 'None');
      itemEl.innerHTML = `
        <div class="audit-item-header">
          <span>Series: <code>${escapeHtml(gap.series)}</code></span>
          <span class="audit-pill pill-amber">${gap.totalMissing} Missing</span>
        </div>
        <div class="audit-item-meta">
          Range analyzed: ${escapeHtml(gap.rangeAnalyzed || 'N/A')} (${gap.totalPresent || 0} present)
        </div>
        <div class="audit-item-desc">
          <strong>Missing invoices:</strong> ${escapeHtml(missingList)}
        </div>
      `;
      auditSeqList.appendChild(itemEl);
    });
  }

  // 2. Missing Required Data List
  auditMissingList.innerHTML = '';
  if (!audit.missingData || audit.missingData.length === 0) {
    auditMissingList.innerHTML = `
      <div class="audit-empty-state">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>
        <span>No missing data. All invoices have valid numbers, dates, buyers, GSTINs, HSNs & amounts.</span>
      </div>
    `;
  } else {
    audit.missingData.forEach(m => {
      const itemEl = document.createElement('div');
      itemEl.className = 'audit-item';
      itemEl.innerHTML = `
        <div class="audit-item-header">
          <span>Invoice: <strong>${escapeHtml(m.invoiceNo)}</strong></span>
          <span class="audit-pill pill-red">${m.missingFields.length} Issue${m.missingFields.length > 1 ? 's' : ''}</span>
        </div>
        <div class="audit-item-meta">
          Company: ${escapeHtml(m.companyName || 'N/A')} | File: <code>${escapeHtml(m.filename || 'N/A')}</code>
        </div>
        <div class="audit-item-desc">
          <strong>Issues:</strong> ${escapeHtml(m.missingFields.join('; '))}
        </div>
      `;
      auditMissingList.appendChild(itemEl);
    });
  }

  // 3. Clubbed Invoices Log List
  auditClubbedList.innerHTML = '';
  if (!audit.clubbingLogs || audit.clubbingLogs.length === 0) {
    auditClubbedList.innerHTML = `
      <div class="audit-empty-state">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        <span>No invoices currently clubbed. Each item is displayed as its own line.</span>
      </div>
    `;
  } else {
    audit.clubbingLogs.forEach(c => {
      const itemEl = document.createElement('div');
      itemEl.className = 'audit-item';
      itemEl.innerHTML = `
        <div class="audit-item-header">
          <span>Invoice: <strong>${escapeHtml(c.invoiceNo)}</strong></span>
          <span class="audit-pill pill-blue">${escapeHtml(c.action)}</span>
        </div>
        <div class="audit-item-meta">
          Company: ${escapeHtml(c.companyName || 'N/A')} | Original items: ${c.originalItemCount} &rarr; Shown as: ${c.resultItemCount} row(s)
        </div>
        <div class="audit-item-desc">
          ${escapeHtml(c.details || '')}
        </div>
      `;
      auditClubbedList.appendChild(itemEl);
    });
  }

  // 4. Duplicate Invoices List
  auditDupList.innerHTML = '';
  if (!audit.duplicates || audit.duplicates.length === 0) {
    auditDupList.innerHTML = `
      <div class="audit-empty-state">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>
        <span>No duplicate invoices detected. Every invoice number is unique.</span>
      </div>
    `;
  } else {
    audit.duplicates.forEach(d => {
      const itemEl = document.createElement('div');
      itemEl.className = 'audit-item';
      itemEl.innerHTML = `
        <div class="audit-item-header">
          <span>Duplicate Invoice: <strong>${escapeHtml(d.invoiceNo)}</strong></span>
          <span class="audit-pill pill-red">Found ${d.count} Times</span>
        </div>
        <div class="audit-item-meta">
          Found in files: ${escapeHtml((d.files || []).join(', '))}
        </div>
        <div class="audit-item-desc">
          Ensure duplicate files are not inflating cumulative sales figures.
        </div>
      `;
      auditDupList.appendChild(itemEl);
    });
  }
}

// Re-render when toggling settings
repeatInvoiceInfoToggle.addEventListener('change', () => {
  if (processedInvoices.length > 0) {
    renderTable();
  }
});

clubByHsnToggle.addEventListener('change', () => {
  if (clubByHsnToggle.checked && highestHsnToggle.checked) {
    highestHsnToggle.checked = false;
    updateHighestToggles(false);
  }
  updateClubToggles(clubByHsnToggle.checked);
  if (processedInvoices.length > 0) {
    renderTable();
    refreshAuditReport();
  }
});

btnQuickClubToggle.addEventListener('click', () => {
  const newState = !clubByHsnToggle.checked;
  if (newState && highestHsnToggle.checked) {
    highestHsnToggle.checked = false;
    updateHighestToggles(false);
  }
  clubByHsnToggle.checked = newState;
  updateClubToggles(newState);
  if (processedInvoices.length > 0) {
    renderTable();
    refreshAuditReport();
  }
});

highestHsnToggle.addEventListener('change', () => {
  if (highestHsnToggle.checked && clubByHsnToggle.checked) {
    clubByHsnToggle.checked = false;
    updateClubToggles(false);
  }
  updateHighestToggles(highestHsnToggle.checked);
  if (processedInvoices.length > 0) {
    renderTable();
    refreshAuditReport();
  }
});

btnQuickHighestToggle.addEventListener('click', () => {
  const newState = !highestHsnToggle.checked;
  if (newState && clubByHsnToggle.checked) {
    clubByHsnToggle.checked = false;
    updateClubToggles(false);
  }
  highestHsnToggle.checked = newState;
  updateHighestToggles(newState);
  if (processedInvoices.length > 0) {
    renderTable();
    refreshAuditReport();
  }
});

function updateClubToggles(active) {
  if (active) {
    btnQuickClubToggle.classList.add('active');
    quickClubText.textContent = 'Club same HSN: ON';
  } else {
    btnQuickClubToggle.classList.remove('active');
    quickClubText.textContent = 'Club same HSN: OFF';
  }
}

function updateHighestToggles(active) {
  if (active) {
    btnQuickHighestToggle.classList.add('active');
    quickHighestText.textContent = '1 Row (Top HSN): ON';
  } else {
    btnQuickHighestToggle.classList.remove('active');
    quickHighestText.textContent = '1 Row (Top HSN): OFF';
  }
}

periodTitleInput.addEventListener('input', () => {
  const title = periodTitleInput.value.trim() || 'April-June 2025-26';
  tableHeaderPeriod.textContent = title;
  reportPeriodTag.textContent = `OUT INV | ${title}`;
});

// Render the metrics and table
function renderResults() {
  resultsSection.classList.remove('hidden');

  const title = periodTitleInput.value.trim() || 'April-June 2025-26';
  tableHeaderPeriod.textContent = title;
  reportPeriodTag.textContent = `OUT INV | ${title}`;

  // Update Metrics
  if (currentSummary) {
    metricTotalInvoices.textContent = currentSummary.totalInvoices;
    metricBasicAmount.textContent = formatCurrency(currentSummary.totalBasicAmount);
    metricSgst.textContent = formatCurrency(currentSummary.totalSgst);
    metricCgst.textContent = formatCurrency(currentSummary.totalCgst);
    metricIgst.textContent = formatCurrency(currentSummary.totalIgst);
    metricGrandTotal.textContent = formatCurrency(currentSummary.totalGrand);
  }

  // Render Audit Report Card
  renderAuditReport(currentAuditReport);

  renderTable();

  // Scroll smoothly to results
  resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Render Table Rows matching Image 3 layout
function renderTable() {
  tableBody.innerHTML = '';
  tableFoot.innerHTML = '';

  const displayInvoices = getDisplayInvoices();
  const totalItemCount = displayInvoices.reduce((acc, inv) => acc + (inv.items ? inv.items.length : 0), 0);
  let modeBadge = '';
  if (highestHsnToggle.checked) modeBadge = ' (1 row / top HSN)';
  else if (clubByHsnToggle.checked) modeBadge = ' (clubbed HSN)';
  metricTotalItems.textContent = `${totalItemCount} line items${modeBadge}`;

  const repeatInfo = repeatInvoiceInfoToggle.checked;
  const searchTerm = searchInput.value.trim().toLowerCase();

  let visibleRowCount = 0;
  let totalQty = 0;
  let totalBasic = 0;
  let totalSgst = 0;
  let totalCgst = 0;
  let totalIgst = 0;

  displayInvoices.forEach(inv => {
    if (!inv.items || inv.items.length === 0) return;

    inv.items.forEach((item, itemIdx) => {
      const isFirstItem = itemIdx === 0;

      // Search match test across all fields
      const rowSearchString = [
        inv.invoiceNo,
        inv.invoiceDate,
        inv.companyName,
        inv.gstNo,
        item.hsnCode,
        item.description,
        item.remarks,
        String(item.basicAmount),
      ].join(' ').toLowerCase();

      if (searchTerm && !rowSearchString.includes(searchTerm)) {
        return;
      }

      visibleRowCount++;
      totalQty += (item.qty || 0);
      totalBasic += (item.basicAmount || 0);
      totalSgst += (item.sgst || 0);
      totalCgst += (item.cgst || 0);
      totalIgst += (item.igst || 0);

      const tr = document.createElement('tr');
      if (isFirstItem) {
        tr.classList.add('new-invoice');
      } else {
        tr.classList.add('subsequent-item');
      }

      // Display invoice fields (blank on subsequent rows if repeatInfo is false)
      const showHeaderInfo = isFirstItem || repeatInfo;
      const invNoText = showHeaderInfo ? (inv.invoiceNo || '') : '';
      const invDateText = showHeaderInfo ? (inv.invoiceDate || '') : '';
      const compNameText = showHeaderInfo ? (inv.companyName || '') : '';
      const gstNoText = showHeaderInfo ? (inv.gstNo || '') : '';

      // 1. INVOICE NO Cell
      let invNoHtml = escapeHtml(invNoText);
      let invNoClass = '';
      if (showHeaderInfo && (!inv.invoiceNo || inv.invoiceNo === 'UNKNOWN')) {
        invNoClass = 'cell-has-error';
        invNoHtml = `<span class="val-warn-pill pill-red" title="Invoice number is missing or invalid">⚠️ Missing Inv No</span>`;
      }

      // 2. INV DATE Cell
      let invDateHtml = escapeHtml(invDateText);
      let invDateClass = '';
      if (showHeaderInfo && !inv.invoiceDate) {
        invDateClass = 'cell-has-error';
        invDateHtml = `<span class="val-warn-pill pill-red" title="Invoice date is missing">⚠️ Missing Date</span>`;
      }

      // 3. Company Name Cell
      let compNameHtml = escapeHtml(compNameText);
      let compNameClass = '';
      if (showHeaderInfo && (!inv.companyName || /unknown/i.test(inv.companyName))) {
        compNameClass = 'cell-has-error';
        compNameHtml = `<span class="val-warn-pill pill-red" title="Company name is missing">⚠️ Missing Name</span>`;
      }

      // 4. GST NO Cell
      let gstNoHtml = escapeHtml(gstNoText);
      let gstNoClass = '';
      if (showHeaderInfo) {
        const gstVal = validateGstinFormat(inv.gstNo);
        if (!gstVal.valid) {
          gstNoClass = gstVal.isError ? 'cell-has-error' : 'cell-has-warning';
          if (gstVal.isError && !inv.gstNo) {
            gstNoHtml = `<span class="val-warn-pill pill-red" title="Buyer GSTIN is missing">⚠️ Missing GSTIN</span>`;
          } else {
            gstNoHtml = `${escapeHtml(gstNoText)} <span class="val-warn-pill ${gstVal.isError ? 'pill-red' : 'pill-amber'}" title="${escapeHtml(gstVal.error)}">⚠️ ${escapeHtml(gstVal.error)}</span>`;
          }
        }
      }

      // 5. HSN Code Cell
      let hsnHtml = escapeHtml(item.hsnCode || '');
      let hsnClass = '';
      const hsnVal = validateHsnCode(item.hsnCode);
      if (!hsnVal.valid) {
        hsnClass = hsnVal.isError ? 'cell-has-error' : 'cell-has-warning';
        if (hsnVal.isError && !item.hsnCode) {
          hsnHtml = `<span class="val-warn-pill pill-red" title="HSN code is missing">⚠️ Blank HSN</span>`;
        } else {
          hsnHtml = `${escapeHtml(item.hsnCode || '')} <span class="val-warn-pill ${hsnVal.isError ? 'pill-red' : 'pill-amber'}" title="${escapeHtml(hsnVal.error)}">⚠️ ${escapeHtml(hsnVal.error)}</span>`;
        }
      }

      const basicNum = item.basicAmount || 0;
      const sgstNum = item.sgst || 0;
      const cgstNum = item.cgst || 0;
      const igstNum = item.igst || 0;
      const qtyNum = item.qty || 0;
      const lineTotal = basicNum + sgstNum + cgstNum + igstNum;

      // 6. QTY Cell
      let qtyHtml = `${qtyNum}`;
      let qtyClass = 'num-col';
      if (qtyNum < 0) {
        qtyClass += ' cell-has-error';
        qtyHtml = `${qtyNum} <span class="val-warn-pill pill-red" title="Negative Quantity">⚠️ Neg</span>`;
      } else if (basicNum > 0 && qtyNum === 0) {
        qtyClass += ' cell-has-warning';
        qtyHtml = `0 <span class="val-warn-pill pill-amber" title="Quantity is 0 while basic amount is positive">⚠️ Qty 0</span>`;
      }

      // 7. BASIC AMOUNT Cell
      let basicHtml = formatNumber(basicNum);
      let basicClass = 'num-col';
      if (basicNum < 0) {
        basicClass += ' cell-has-error';
        basicHtml = `${basicHtml} <span class="val-warn-pill pill-red" title="Negative Basic Amount">⚠️ Neg</span>`;
      } else if (qtyNum > 0 && basicNum === 0) {
        basicClass += ' cell-has-warning';
        basicHtml = `0.00 <span class="val-warn-pill pill-amber" title="Basic amount is 0 while quantity is positive">⚠️ ₹0.00</span>`;
      }

      // 8. SGST & CGST Cells (Intra-state check)
      let sgstHtml = formatNumber(sgstNum);
      let sgstClass = 'num-col';
      let cgstHtml = formatNumber(cgstNum);
      let cgstClass = 'num-col';

      const isMismatch = (sgstNum > 0 || cgstNum > 0) && Math.abs(sgstNum - cgstNum) > 0.05;
      if (isMismatch) {
        sgstClass += ' cell-has-warning';
        cgstClass += ' cell-has-warning';
        sgstHtml = `${sgstHtml} <span class="val-warn-pill pill-amber" title="Intra-state Tax Mismatch: SGST ₹${sgstNum.toFixed(2)} does not match CGST ₹${cgstNum.toFixed(2)}">⚠️ Mismatch</span>`;
        cgstHtml = `${cgstHtml} <span class="val-warn-pill pill-amber" title="Intra-state Tax Mismatch: CGST ₹${cgstNum.toFixed(2)} does not match SGST ₹${sgstNum.toFixed(2)}">⚠️ Mismatch</span>`;
      }

      // 9. IGST Cell (Dual-tax check)
      let igstHtml = formatNumber(igstNum);
      let igstClass = 'num-col';
      const isDualTax = igstNum > 0 && (sgstNum > 0 || cgstNum > 0);
      if (isDualTax) {
        igstClass += ' cell-has-error';
        igstHtml = `${igstHtml} <span class="val-warn-pill pill-red" title="Dual Tax Conflict: Both Inter-state IGST and Intra-state SGST/CGST charged">⚠️ Dual Tax</span>`;
      }

      // 10. REMARKS Cell
      let remarksHtml = '';
      let remarksClass = '';
      if (item.remarks && /sez|nil/i.test(item.remarks)) {
        remarksHtml = `<span class="badge-sez">${escapeHtml(item.remarks)}</span>`;
      } else if (item.remarks) {
        remarksHtml = escapeHtml(item.remarks);
      }

      const isUnclarifiedZero = basicNum > 0 && sgstNum === 0 && cgstNum === 0 && igstNum === 0 && !/sez|nil|zero|exempt|export/i.test(item.remarks || '');
      if (isUnclarifiedZero) {
        remarksClass = 'cell-has-warning';
        remarksHtml += ` <span class="val-warn-pill pill-amber" title="Zero Tax: Remarks do not specify SEZ / NIL / Export status">⚠️ Unclarified 0% Tax</span>`;
      }

      tr.innerHTML = `
        <td class="${invNoClass}" data-copy-val="${escapeHtml(invNoText)}">${invNoHtml}</td>
        <td class="${invDateClass}" data-copy-val="${escapeHtml(invDateText)}">${invDateHtml}</td>
        <td class="${compNameClass}" data-copy-val="${escapeHtml(compNameText)}">${compNameHtml}</td>
        <td class="${gstNoClass}" data-copy-val="${escapeHtml(gstNoText)}">${gstNoHtml}</td>
        <td class="${hsnClass}" data-copy-val="${escapeHtml(item.hsnCode || '')}">${hsnHtml}</td>
        <td class="${qtyClass}" data-copy-val="${qtyNum}">${qtyHtml}</td>
        <td class="${basicClass}" data-copy-val="${formatNumber(basicNum)}">${basicHtml}</td>
        <td class="${sgstClass}" data-copy-val="${formatNumber(sgstNum)}">${sgstHtml}</td>
        <td class="${cgstClass}" data-copy-val="${formatNumber(cgstNum)}">${cgstHtml}</td>
        <td class="${igstClass}" data-copy-val="${formatNumber(igstNum)}">${igstHtml}</td>
        <td class="num-col" data-copy-val="${formatNumber(lineTotal)}">${formatNumber(lineTotal)}</td>
        <td class="${remarksClass}" data-copy-val="${escapeHtml(item.remarks || '')}">${remarksHtml}</td>
      `;

      tableBody.appendChild(tr);
    });
  });

  // Render Table Foot (Totals)
  const totalGrand = totalBasic + totalSgst + totalCgst + totalIgst;
  const footTr = document.createElement('tr');
  footTr.innerHTML = `
    <td></td>
    <td></td>
    <td data-copy-val="Total"><strong>Total</strong></td>
    <td></td>
    <td></td>
    <td class="num-col" data-copy-val="${totalQty}"><strong>${totalQty}</strong></td>
    <td class="num-col" data-copy-val="${formatNumber(totalBasic)}"><strong>${formatNumber(totalBasic)}</strong></td>
    <td class="num-col" data-copy-val="${formatNumber(totalSgst)}"><strong>${formatNumber(totalSgst)}</strong></td>
    <td class="num-col" data-copy-val="${formatNumber(totalCgst)}"><strong>${formatNumber(totalCgst)}</strong></td>
    <td class="num-col" data-copy-val="${formatNumber(totalIgst)}"><strong>${formatNumber(totalIgst)}</strong></td>
    <td class="num-col" data-copy-val="${formatNumber(totalGrand)}"><strong>${formatNumber(totalGrand)}</strong></td>
    <td></td>
  `;
  tableFoot.appendChild(footTr);

  rowCountInfo.textContent = `Showing ${visibleRowCount} invoice item rows across ${processedInvoices.length} invoices`;
}

// Search filtering
searchInput.addEventListener('input', () => {
  renderTable();
});

// Export to Excel (.xlsx)
btnExportExcel.addEventListener('click', async () => {
  if (processedInvoices.length === 0) return;

  btnExportExcel.disabled = true;
  btnExportExcel.textContent = 'Generating Excel...';

  try {
    const res = await fetch('/api/export-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        invoices: processedInvoices,
        periodTitle: periodTitleInput.value.trim() || 'April-June 2025-26',
        repeatInvoiceInfo: repeatInvoiceInfoToggle.checked,
        clubByHsn: clubByHsnToggle.checked,
        highestHsnOnly: highestHsnToggle.checked,
      }),
    });

    if (!res.ok) throw new Error('Failed to generate Excel file.');

    const blob = await res.blob();
    const downloadUrl = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = downloadUrl;
    const title = (periodTitleInput.value.trim() || 'Cumulative').replace(/[^a-zA-Z0-9_-]/g, '_');
    a.download = `OUT_INV_${title}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(downloadUrl);
  } catch (err) {
    alert('Error exporting Excel: ' + err.message);
  } finally {
    btnExportExcel.disabled = false;
    btnExportExcel.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      Download Excel (.xlsx)
    `;
  }
});

// Export Comprehensive Statistics Report (.xlsx)
btnExportStats.addEventListener('click', async () => {
  if (processedInvoices.length === 0) return;

  btnExportStats.disabled = true;
  btnExportStats.textContent = 'Generating Stats...';

  try {
    const res = await fetch('/api/export-statistics', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        invoices: processedInvoices,
        periodTitle: periodTitleInput.value.trim() || 'April-June 2025-26',
      }),
    });

    if (!res.ok) throw new Error('Failed to generate Statistics report.');

    const blob = await res.blob();
    const downloadUrl = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = downloadUrl;
    const title = (periodTitleInput.value.trim() || 'Summary').replace(/[^a-zA-Z0-9_-]/g, '_');
    a.download = `GST_Statistics_Report_${title}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(downloadUrl);
  } catch (err) {
    alert('Error exporting Statistics: ' + err.message);
  } finally {
    btnExportStats.disabled = false;
    btnExportStats.innerHTML = `
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <line x1="18" y1="20" x2="18" y2="10"></line>
        <line x1="12" y1="20" x2="12" y2="4"></line>
        <line x1="6" y1="20" x2="6" y2="14"></line>
      </svg>
      Download Stats Report (.xlsx)
    `;
  }
});

// Export to CSV
btnExportCsv.addEventListener('click', async () => {
  if (processedInvoices.length === 0) return;

  btnExportCsv.disabled = true;
  btnExportCsv.textContent = 'Generating CSV...';

  try {
    const res = await fetch('/api/export-csv', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        invoices: processedInvoices,
        periodTitle: periodTitleInput.value.trim() || 'April-June 2025-26',
        repeatInvoiceInfo: repeatInvoiceInfoToggle.checked,
        clubByHsn: clubByHsnToggle.checked,
        highestHsnOnly: highestHsnToggle.checked,
      }),
    });

    if (!res.ok) throw new Error('Failed to generate CSV file.');

    const blob = await res.blob();
    const downloadUrl = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = downloadUrl;
    const title = (periodTitleInput.value.trim() || 'Cumulative').replace(/[^a-zA-Z0-9_-]/g, '_');
    a.download = `OUT_INV_${title}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(downloadUrl);
  } catch (err) {
    alert('Error exporting CSV: ' + err.message);
  } finally {
    btnExportCsv.disabled = false;
    btnExportCsv.textContent = 'Download CSV';
  }
});

// Copy Data (TSV format for direct paste into Excel or Google Sheets)
btnCopyTable.addEventListener('click', () => {
  if (processedInvoices.length === 0) return;

  const repeatInfo = repeatInvoiceInfoToggle.checked;
  const displayInvoices = getDisplayInvoices();
  const lines = [];

  // Super Header
  lines.push(`OUT INV\t${periodTitleInput.value.trim() || 'April-June 2025-26'}`);
  // Columns
  lines.push('INVOICE NO\tINV DATE\tCompany Name\tGST NO\tHSN Code\tQty\tBASIC AMOUNT\tSGST\tCGST\tIGST\tTOTAL AMOUNT\tdescription / remarks');

  displayInvoices.forEach(inv => {
    if (!inv.items) return;
    inv.items.forEach((item, idx) => {
      const showHeader = idx === 0 || repeatInfo;
      const basicNum = item.basicAmount || 0;
      const sgstNum = item.sgst || 0;
      const cgstNum = item.cgst || 0;
      const igstNum = item.igst || 0;
      const lineTotal = basicNum + sgstNum + cgstNum + igstNum;

      const row = [
        showHeader ? (inv.invoiceNo || '') : '',
        showHeader ? (inv.invoiceDate || '') : '',
        showHeader ? (inv.companyName || '') : '',
        showHeader ? (inv.gstNo || '') : '',
        item.hsnCode || '',
        item.qty || 0,
        item.basicAmount || 0,
        item.sgst || 0,
        item.cgst || 0,
        item.igst || 0,
        lineTotal.toFixed(2),
        item.remarks || '',
      ];
      lines.push(row.join('\t'));
    });
  });

  const text = lines.join('\n');
  navigator.clipboard.writeText(text).then(() => {
    const originalText = btnCopyTable.textContent;
    btnCopyTable.textContent = 'Copied to Clipboard!';
    setTimeout(() => {
      btnCopyTable.textContent = originalText;
    }, 2000);
  }).catch(() => {
    alert('Unable to copy to clipboard.');
  });
});

// Deploy Modal Logic
btnOpenDeployModal.addEventListener('click', () => {
  deployModal.classList.remove('hidden');
});

btnCloseDeployModal.addEventListener('click', () => {
  deployModal.classList.add('hidden');
});

deployModal.addEventListener('click', (e) => {
  if (e.target === deployModal) {
    deployModal.classList.add('hidden');
  }
});

tabButtons.forEach(btn => {
  btn.addEventListener('click', () => {
    tabButtons.forEach(b => b.classList.remove('active'));
    tabContents.forEach(c => c.classList.remove('active'));

    btn.classList.add('active');
    const targetTab = btn.getAttribute('data-tab');
    document.getElementById(targetTab).classList.add('active');
  });
});

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// =========================================================
// FUNCTIONAL WORKFLOW NAVIGATION TABS
// =========================================================
const funcTabBtns = document.querySelectorAll('.func-tab-btn');
const viewPanels = document.querySelectorAll('.view-panel');

funcTabBtns.forEach(btn => {
  btn.addEventListener('click', () => {
    const targetId = btn.getAttribute('data-tab-target');
    if (!targetId || btn.classList.contains('disabled-func-tab')) return;

    funcTabBtns.forEach(b => b.classList.remove('active'));
    viewPanels.forEach(p => p.classList.add('hidden'));

    btn.classList.add('active');
    const targetPanel = document.getElementById(targetId);
    if (targetPanel) {
      targetPanel.classList.remove('hidden');
    }
  });
});

// =========================================================
// MODULE 2: HSN-WISE GST SUMMARY GENERATOR
// =========================================================
let selectedHsnFile = null;
let currentHsnData = null;

const hsnDropZone = document.getElementById('hsnDropZone');
const hsnFileInput = document.getElementById('hsnFileInput');
const hsnFileSelectedBadge = document.getElementById('hsnFileSelectedBadge');
const hsnFileName = document.getElementById('hsnFileName');
const hsnPeriodInput = document.getElementById('hsnPeriodInput');
const btnProcessHsn = document.getElementById('btnProcessHsn');
const hsnProcessSpinner = document.getElementById('hsnProcessSpinner');
const btnLoadSampleHsn = document.getElementById('btnLoadSampleHsn');

const hsnResultsSection = document.getElementById('hsnResultsSection');
const hsnMetricGroups = document.getElementById('hsnMetricGroups');
const hsnMetricLines = document.getElementById('hsnMetricLines');
const hsnMetricQty = document.getElementById('hsnMetricQty');
const hsnMetricBasic = document.getElementById('hsnMetricBasic');
const hsnMetricSgst = document.getElementById('hsnMetricSgst');
const hsnMetricCgst = document.getElementById('hsnMetricCgst');
const hsnMetricIgst = document.getElementById('hsnMetricIgst');
const hsnMetricGrand = document.getElementById('hsnMetricGrand');

const hsnTableTitle = document.getElementById('hsnTableTitle');
const hsnTableBody = document.getElementById('hsnTableBody');
const hsnTableFoot = document.getElementById('hsnTableFoot');
const hsnRowCountInfo = document.getElementById('hsnRowCountInfo');

const btnExportHsnExcel = document.getElementById('btnExportHsnExcel');
const btnExportHsnCsv = document.getElementById('btnExportHsnCsv');
const btnCopyHsnTable = document.getElementById('btnCopyHsnTable');

// HSN Audit & Inconsistencies DOM Elements
const hsnAuditCard = document.getElementById('hsnAuditCard');
const hsnAuditStatusIcon = document.getElementById('hsnAuditStatusIcon');
const hsnAuditTitleText = document.getElementById('hsnAuditTitleText');
const hsnAuditSubtitleText = document.getElementById('hsnAuditSubtitleText');
const hsnAuditMainBadge = document.getElementById('hsnAuditMainBadge');
const hsnAuditBadgesArea = document.getElementById('hsnAuditBadgesArea');
const badgeHsnMissing = document.getElementById('badgeHsnMissing');
const badgeHsnTax = document.getElementById('badgeHsnTax');
const badgeHsnSeq = document.getElementById('badgeHsnSeq');
const badgeHsnConsol = document.getElementById('badgeHsnConsol');
const hsnAuditMissingList = document.getElementById('hsnAuditMissingList');
const hsnAuditTaxList = document.getElementById('hsnAuditTaxList');
const hsnAuditSeqList = document.getElementById('hsnAuditSeqList');
const hsnAuditConsolList = document.getElementById('hsnAuditConsolList');

// HSN Drop Zone listeners
if (hsnDropZone) {
  hsnDropZone.addEventListener('click', () => hsnFileInput.click());

  ['dragenter', 'dragover'].forEach(eventName => {
    hsnDropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      hsnDropZone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    hsnDropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      hsnDropZone.classList.remove('drag-over');
    });
  });

  hsnDropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt.files && dt.files.length > 0) {
      handleHsnFileSelected(dt.files[0]);
    }
  });

  hsnFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleHsnFileSelected(e.target.files[0]);
    }
  });
}

function handleHsnFileSelected(file) {
  const ext = file.name.toLowerCase();
  if (!ext.endsWith('.xlsx') && !ext.endsWith('.xls')) {
    alert('Please select a valid Excel file (.xlsx or .xls).');
    return;
  }

  selectedHsnFile = file;
  hsnFileName.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
  hsnFileSelectedBadge.classList.remove('hidden');
  btnProcessHsn.disabled = false;
  btnProcessHsn.querySelector('.btn-text').textContent = 'Generate HSN Summary';
}

// Process HSN Summary
if (btnProcessHsn) {
  btnProcessHsn.addEventListener('click', async () => {
    if (!selectedHsnFile) return;

    btnProcessHsn.disabled = true;
    hsnProcessSpinner.classList.remove('hidden');
    btnProcessHsn.querySelector('.btn-text').textContent = 'Processing HSN Groups...';

    const formData = new FormData();
    formData.append('cumulativeSheet', selectedHsnFile);
    formData.append('periodTitle', hsnPeriodInput.value.trim());

    try {
      const res = await fetch('/api/hsn-summary/upload', {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to process HSN summary.');
      }

      currentHsnData = data;
      if (data.periodTitle && !hsnPeriodInput.value.trim()) {
        hsnPeriodInput.value = data.periodTitle;
      }
      renderHsnResults(data);
    } catch (err) {
      alert('Error: ' + err.message);
    } finally {
      btnProcessHsn.disabled = false;
      hsnProcessSpinner.classList.add('hidden');
      btnProcessHsn.querySelector('.btn-text').textContent = 'Generate HSN Summary';
    }
  });
}

// Load Demo Cumulative Sheet
if (btnLoadSampleHsn) {
  btnLoadSampleHsn.addEventListener('click', async () => {
    btnLoadSampleHsn.disabled = true;
    btnLoadSampleHsn.textContent = 'Loading sample...';

    try {
      const res = await fetch('/api/hsn-summary/sample');
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to load demo HSN summary.');
      }

      currentHsnData = data;
      if (data.periodTitle) {
        hsnPeriodInput.value = data.periodTitle;
      }
      renderHsnResults(data);
    } catch (err) {
      alert('Error loading sample: ' + err.message);
    } finally {
      btnLoadSampleHsn.disabled = false;
      btnLoadSampleHsn.textContent = 'Try Demo Cumulative Sheet';
    }
  });
}

// Render Audit & Inconsistencies Report for Feature 2 HSN Summary
function renderHsnAuditReport(audit) {
  if (!hsnAuditCard) return;
  if (!audit) {
    hsnAuditCard.classList.add('hidden');
    return;
  }
  hsnAuditCard.classList.remove('hidden');
  hsnAuditCard.classList.remove('audit-clean', 'audit-error');

  const missingCount = audit.missingDataCount || 0;
  const taxCount = audit.taxInconsistencyCount || 0;
  const seqCount = audit.sequenceGapCount || 0;
  const consolCount = audit.consolidationCount || 0;
  const excludedCount = audit.excludedRowCount || 0;
  const totalIssues = audit.totalInconsistencyCount || 0;

  // Update tab badge numbers
  if (badgeHsnMissing) badgeHsnMissing.textContent = missingCount;
  if (badgeHsnTax) badgeHsnTax.textContent = taxCount;
  if (badgeHsnSeq) badgeHsnSeq.textContent = seqCount;
  if (badgeHsnConsol) badgeHsnConsol.textContent = consolCount;

  // Header status & Icon
  if (totalIssues === 0) {
    hsnAuditCard.classList.add('audit-clean');
    if (hsnAuditStatusIcon) hsnAuditStatusIcon.textContent = '✅';
    if (hsnAuditTitleText) hsnAuditTitleText.textContent = 'Data Consistency & Integrity Check Passed';
    if (hsnAuditSubtitleText) {
      hsnAuditSubtitleText.textContent = 'All HSN codes, taxable values, SGST/CGST rates, and invoice sequences verified without critical issues.';
    }
    if (hsnAuditMainBadge) {
      hsnAuditMainBadge.className = 'badge pill-green';
      hsnAuditMainBadge.textContent = '100% Consistent';
    }
  } else {
    hsnAuditCard.classList.add('audit-error');
    if (hsnAuditStatusIcon) hsnAuditStatusIcon.textContent = '⚠️';
    if (hsnAuditTitleText) {
      hsnAuditTitleText.textContent = `Found ${totalIssues} Data Discrepanc${totalIssues > 1 ? 'ies' : 'y'} / Warning${totalIssues > 1 ? 's' : ''}`;
    }
    if (hsnAuditSubtitleText) {
      hsnAuditSubtitleText.textContent = 'Review missing HSNs, blank fields, tax mismatches, or invoice sequence gaps below:';
    }
    if (hsnAuditMainBadge) {
      hsnAuditMainBadge.className = 'badge pill-amber';
      hsnAuditMainBadge.textContent = `${totalIssues} Issue${totalIssues > 1 ? 's' : ''}`;
    }
  }

  // Header quick pills
  let pillsHtml = '';
  if (missingCount > 0) pillsHtml += `<span class="audit-pill pill-red">❌ ${missingCount} Missing / Format Issue${missingCount > 1 ? 's' : ''}</span>`;
  if (taxCount > 0) pillsHtml += `<span class="audit-pill pill-amber">⚠️ ${taxCount} Tax Discrepanc${taxCount > 1 ? 'ies' : 'y'}</span>`;
  if (seqCount > 0) pillsHtml += `<span class="audit-pill pill-purple">🔍 ${seqCount} Missing in Sequence</span>`;
  if (consolCount > 0) pillsHtml += `<span class="audit-pill pill-blue">ℹ️ ${consolCount} Service Line${consolCount > 1 ? 's' : ''} Grouped under 998719</span>`;
  if (excludedCount > 0) pillsHtml += `<span class="audit-pill pill-green">🛡️ ${excludedCount} Total Row${excludedCount > 1 ? 's' : ''} Safely Excluded</span>`;
  if (hsnAuditBadgesArea) hsnAuditBadgesArea.innerHTML = pillsHtml;

  // 1. Missing Required Data List
  if (hsnAuditMissingList) {
    hsnAuditMissingList.innerHTML = '';
    if (!audit.missingData || audit.missingData.length === 0) {
      hsnAuditMissingList.innerHTML = `
        <div class="audit-empty-state">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>
          <span>No missing data. All rows contain valid HSN codes, quantities, and taxable basic amounts.</span>
        </div>
      `;
    } else {
      audit.missingData.forEach(m => {
        const itemEl = document.createElement('div');
        itemEl.className = 'audit-item';
        itemEl.innerHTML = `
          <div class="audit-item-header">
            <span>Excel Row: <strong>#${m.excelRow || 'N/A'}</strong> (Invoice: <code>${escapeHtml(m.invoiceNo || 'N/A')}</code>)</span>
            <span class="audit-pill pill-red">${escapeHtml(m.field || 'Missing Data')}</span>
          </div>
          <div class="audit-item-meta">
            Company: ${escapeHtml(m.companyName || 'N/A')}
          </div>
          <div class="audit-item-desc">
            ${escapeHtml(m.details || '')}
          </div>
        `;
        hsnAuditMissingList.appendChild(itemEl);
      });
    }
  }

  // 2. Tax Discrepancies List
  if (hsnAuditTaxList) {
    hsnAuditTaxList.innerHTML = '';
    if (!audit.taxInconsistencies || audit.taxInconsistencies.length === 0) {
      hsnAuditTaxList.innerHTML = `
        <div class="audit-empty-state">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>
          <span>No tax discrepancies. All SGST = CGST shares match, zero GST remarks are clear, and no dual tax conflicts exist.</span>
        </div>
      `;
    } else {
      audit.taxInconsistencies.forEach(t => {
        const itemEl = document.createElement('div');
        itemEl.className = 'audit-item';
        itemEl.innerHTML = `
          <div class="audit-item-header">
            <span>Excel Row: <strong>#${t.excelRow || 'N/A'}</strong> (Invoice: <code>${escapeHtml(t.invoiceNo || 'N/A')}</code>)</span>
            <span class="audit-pill pill-amber">${escapeHtml(t.type || 'Tax Discrepancy')}</span>
          </div>
          <div class="audit-item-meta">
            HSN Code: <code>${escapeHtml(t.hsnCode || 'N/A')}</code>
          </div>
          <div class="audit-item-desc">
            ${escapeHtml(t.details || '')}
          </div>
        `;
        hsnAuditTaxList.appendChild(itemEl);
      });
    }
  }

  // 3. Missing in Sequence List
  if (hsnAuditSeqList) {
    hsnAuditSeqList.innerHTML = '';
    if (!audit.sequenceGaps || audit.sequenceGaps.length === 0) {
      hsnAuditSeqList.innerHTML = `
        <div class="audit-empty-state">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>
          <span>No missing sequence numbers. Invoice sequence is complete and unbroken!</span>
        </div>
      `;
    } else {
      audit.sequenceGaps.forEach(gap => {
        const itemEl = document.createElement('div');
        itemEl.className = 'audit-item';
        const missingList = gap.missingInvoices && gap.missingInvoices.length > 0 
          ? gap.missingInvoices.join('; ')
          : (gap.missingNumbers ? gap.missingNumbers.join(', ') : 'None');
        itemEl.innerHTML = `
          <div class="audit-item-header">
            <span>Series: <code>${escapeHtml(gap.series)}</code></span>
            <span class="audit-pill pill-purple">${gap.totalMissing} Missing</span>
          </div>
          <div class="audit-item-meta">
            Range analyzed: ${escapeHtml(gap.rangeAnalyzed || 'N/A')} (${gap.totalPresent || 0} invoices present)
          </div>
          <div class="audit-item-desc">
            <strong>Missing invoices in sequence:</strong> ${escapeHtml(missingList)}
          </div>
        `;
        hsnAuditSeqList.appendChild(itemEl);
      });
    }
  }

  // 4. 9987 Consolidations Log List
  if (hsnAuditConsolList) {
    hsnAuditConsolList.innerHTML = '';
    const hasConsol = audit.hsnConsolidations && audit.hsnConsolidations.length > 0;
    const hasExclusions = audit.excludedSummaryRows && audit.excludedSummaryRows.length > 0;

    if (!hasConsol && !hasExclusions) {
      hsnAuditConsolList.innerHTML = `
        <div class="audit-empty-state">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          <span>No 9987XX service codes or summary rows were modified. All HSN codes mapped directly.</span>
        </div>
      `;
    } else {
      if (hasConsol) {
        audit.hsnConsolidations.forEach(c => {
          const itemEl = document.createElement('div');
          itemEl.className = 'audit-item';
          itemEl.innerHTML = `
            <div class="audit-item-header">
              <span>Original HSN: <code>${escapeHtml(c.originalHsn)}</code> &rarr; Unified: <code>${escapeHtml(c.unifiedHsn)}</code></span>
              <span class="audit-pill pill-blue">${escapeHtml(c.taxCategory || 'Consolidated')}</span>
            </div>
            <div class="audit-item-meta">
              Excel Row: #${c.excelRow || 'N/A'} | Invoice: <code>${escapeHtml(c.invoiceNo || 'N/A')}</code>
            </div>
            <div class="audit-item-desc">
              ${escapeHtml(c.details || '')}
            </div>
          `;
          hsnAuditConsolList.appendChild(itemEl);
        });
      }
      if (hasExclusions) {
        audit.excludedSummaryRows.forEach(ex => {
          const itemEl = document.createElement('div');
          itemEl.className = 'audit-item';
          itemEl.innerHTML = `
            <div class="audit-item-header">
              <span>Excluded Row: <strong>${escapeHtml(ex.label || 'TOTAL Row')}</strong> (Excel Row #${ex.excelRow})</span>
              <span class="audit-pill pill-green">Safely Excluded</span>
            </div>
            <div class="audit-item-meta">
              Detected Total: ₹${(ex.total || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 })}
            </div>
            <div class="audit-item-desc">
              ${escapeHtml(ex.details || '')}
            </div>
          `;
          hsnAuditConsolList.appendChild(itemEl);
        });
      }
    }
  }
}

function renderHsnResults(data) {
  if (!data || !data.groups) return;

  const { groups, summary, periodTitle, rawRowCount, auditReport } = data;
  const currentTitle = hsnPeriodInput.value.trim() || periodTitle || 'July–September 2026';

  hsnResultsSection.classList.remove('hidden');
  hsnTableTitle.textContent = `HSN-wise GST Summary | ${currentTitle}`;

  // Render HSN Audit & Inconsistencies Card
  renderHsnAuditReport(auditReport);

  // Update Metrics Cards
  hsnMetricGroups.textContent = summary.totalGroups || groups.length;
  hsnMetricLines.textContent = `${rawRowCount || summary.totalLines} raw invoice lines`;
  hsnMetricQty.textContent = (summary.totalQty || 0).toLocaleString('en-IN');
  hsnMetricBasic.textContent = formatCurrency(summary.totalBasic);
  hsnMetricSgst.textContent = formatCurrency(summary.totalSgst);
  hsnMetricCgst.textContent = formatCurrency(summary.totalCgst);
  hsnMetricIgst.textContent = formatCurrency(summary.totalIgst);
  hsnMetricGrand.textContent = formatCurrency(summary.totalGrand);

  // Render Table Body
  hsnTableBody.innerHTML = '';
  groups.forEach(g => {
    const tr = document.createElement('tr');
    const isZeroGst = g.taxCategory === 'Zero GST';
    const catBadgeClass = isZeroGst ? 'badge-zero-gst' : 'badge-taxable-gst';

    const hsnVal = validateHsnCode(g.hsnCode);
    let hsnClass = '';
    let hsnHtml = `<strong>${escapeHtml(g.hsnCode)}</strong>`;
    if (!hsnVal.valid) {
      hsnClass = hsnVal.isError ? 'cell-has-error' : 'cell-has-warning';
      hsnHtml += ` <span class="val-warn-pill ${hsnVal.isError ? 'pill-red' : 'pill-amber'}" title="${escapeHtml(hsnVal.error)}">⚠️ ${escapeHtml(hsnVal.error)}</span>`;
    }

    let qtyClass = 'num-col';
    let qtyHtml = (g.quantity || 0).toLocaleString('en-IN');
    if (g.basicAmount > 0 && (g.quantity || 0) <= 0) {
      qtyClass += ' cell-has-warning';
      qtyHtml += ` <span class="val-warn-pill pill-amber" title="Quantity is 0 while basic amount is positive">⚠️ Qty 0</span>`;
    }

    let sgstClass = 'num-col';
    let sgstHtml = formatNumber(g.sgst);
    let cgstClass = 'num-col';
    let cgstHtml = formatNumber(g.cgst);
    if ((g.sgst > 0 || g.cgst > 0) && Math.abs(g.sgst - g.cgst) > 0.05) {
      sgstClass += ' cell-has-warning';
      cgstClass += ' cell-has-warning';
      sgstHtml += ` <span class="val-warn-pill pill-amber" title="SGST != CGST mismatch">⚠️ Mismatch</span>`;
      cgstHtml += ` <span class="val-warn-pill pill-amber" title="SGST != CGST mismatch">⚠️ Mismatch</span>`;
    }

    tr.innerHTML = `
      <td class="${hsnClass}" data-copy-val="${escapeHtml(g.hsnCode)}">${hsnHtml}</td>
      <td data-copy-val="${escapeHtml(g.taxCategory)}"><span class="badge-tax-cat ${catBadgeClass}">${escapeHtml(g.taxCategory)}</span></td>
      <td class="num-col" data-copy-val="${g.invoiceLines}">${g.invoiceLines}</td>
      <td class="${qtyClass}" data-copy-val="${(g.quantity || 0).toLocaleString('en-IN')}">${qtyHtml}</td>
      <td class="num-col" data-copy-val="${formatNumber(g.basicAmount)}">${formatNumber(g.basicAmount)}</td>
      <td class="${sgstClass}" data-copy-val="${formatNumber(g.sgst)}">${sgstHtml}</td>
      <td class="${cgstClass}" data-copy-val="${formatNumber(g.cgst)}">${cgstHtml}</td>
      <td class="num-col" data-copy-val="${formatNumber(g.igst)}">${formatNumber(g.igst)}</td>
      <td class="num-col font-bold" data-copy-val="${formatNumber(g.totalInclGst)}">${formatNumber(g.totalInclGst)}</td>
    `;
    hsnTableBody.appendChild(tr);
  });

  // Render Table Foot (Matching reference image Row 23 light blue)
  hsnTableFoot.innerHTML = `
    <tr>
      <td data-copy-val="GRAND TOTAL"><strong>GRAND TOTAL</strong></td>
      <td></td>
      <td class="num-col" data-copy-val="${summary.totalLines}"><strong>${summary.totalLines}</strong></td>
      <td class="num-col" data-copy-val="${(summary.totalQty || 0).toLocaleString('en-IN')}"><strong>${(summary.totalQty || 0).toLocaleString('en-IN')}</strong></td>
      <td class="num-col" data-copy-val="${formatNumber(summary.totalBasic)}"><strong>${formatNumber(summary.totalBasic)}</strong></td>
      <td class="num-col" data-copy-val="${formatNumber(summary.totalSgst)}"><strong>${formatNumber(summary.totalSgst)}</strong></td>
      <td class="num-col" data-copy-val="${formatNumber(summary.totalCgst)}"><strong>${formatNumber(summary.totalCgst)}</strong></td>
      <td class="num-col" data-copy-val="${formatNumber(summary.totalIgst)}"><strong>${formatNumber(summary.totalIgst)}</strong></td>
      <td class="num-col" data-copy-val="${formatNumber(summary.totalGrand)}"><strong>${formatNumber(summary.totalGrand)}</strong></td>
    </tr>
  `;

  hsnRowCountInfo.textContent = `Showing ${groups.length} HSN summary groups (aggregated from ${rawRowCount || summary.totalLines} cumulative sales rows)`;

  // Scroll smoothly to results
  hsnResultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Export HSN Summary Excel (.xlsx)
if (btnExportHsnExcel) {
  btnExportHsnExcel.addEventListener('click', async () => {
    if (!currentHsnData || !currentHsnData.groups || currentHsnData.groups.length === 0) return;

    btnExportHsnExcel.disabled = true;
    btnExportHsnExcel.textContent = 'Generating Excel...';

    try {
      const res = await fetch('/api/hsn-summary/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          groups: currentHsnData.groups,
          periodTitle: hsnPeriodInput.value.trim() || currentHsnData.periodTitle || 'July–September 2026',
        }),
      });

      if (!res.ok) throw new Error('Failed to generate HSN Excel.');

      const blob = await res.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = downloadUrl;
      const safeTitle = (hsnPeriodInput.value.trim() || 'Summary').replace(/[^a-zA-Z0-9_-]/g, '_');
      a.download = `HSN_Wise_GST_Summary_${safeTitle}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(downloadUrl);
    } catch (err) {
      alert('Error exporting HSN Summary Excel: ' + err.message);
    } finally {
      btnExportHsnExcel.disabled = false;
      btnExportHsnExcel.innerHTML = `
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
          <polyline points="7 10 12 15 17 10"></polyline>
          <line x1="12" y1="15" x2="12" y2="3"></line>
        </svg>
        Download HSN Summary (.xlsx)
      `;
    }
  });
}

// Export HSN Summary CSV
if (btnExportHsnCsv) {
  btnExportHsnCsv.addEventListener('click', async () => {
    if (!currentHsnData || !currentHsnData.groups || currentHsnData.groups.length === 0) return;

    btnExportHsnCsv.disabled = true;
    btnExportHsnCsv.textContent = 'Generating CSV...';

    try {
      const res = await fetch('/api/hsn-summary/export-csv', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          groups: currentHsnData.groups,
          periodTitle: hsnPeriodInput.value.trim() || currentHsnData.periodTitle || 'July–September 2026',
        }),
      });

      if (!res.ok) throw new Error('Failed to generate HSN CSV.');

      const blob = await res.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = downloadUrl;
      const safeTitle = (hsnPeriodInput.value.trim() || 'Summary').replace(/[^a-zA-Z0-9_-]/g, '_');
      a.download = `HSN_Wise_GST_Summary_${safeTitle}.csv`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(downloadUrl);
    } catch (err) {
      alert('Error exporting HSN Summary CSV: ' + err.message);
    } finally {
      btnExportHsnCsv.disabled = false;
      btnExportHsnCsv.textContent = 'Download CSV';
    }
  });
}

// Copy HSN Summary Table as TSV for direct paste into Excel
if (btnCopyHsnTable) {
  btnCopyHsnTable.addEventListener('click', () => {
    if (!currentHsnData || !currentHsnData.groups || currentHsnData.groups.length === 0) return;

    const period = hsnPeriodInput.value.trim() || currentHsnData.periodTitle || 'July–September 2026';
    const lines = [];

    lines.push(`HSN-wise GST Summary | ${period}`);
    lines.push('Original invoice TOTAL row excluded; zero GST basic shown separately.');
    lines.push('');
    lines.push('HSN Code\tTax Category\tInvoice Lines\tQuantity\tBasic Amount\tSGST\tCGST\tIGST\tTotal incl. GST');

    currentHsnData.groups.forEach(g => {
      lines.push([
        g.hsnCode || '',
        g.taxCategory || '',
        g.invoiceLines || 0,
        g.quantity || 0,
        g.basicAmount || 0,
        g.sgst || 0,
        g.cgst || 0,
        g.igst || 0,
        g.totalInclGst || 0,
      ].join('\t'));
    });

    const s = currentHsnData.summary;
    lines.push([
      'GRAND TOTAL',
      '',
      s.totalLines || 0,
      s.totalQty || 0,
      s.totalBasic || 0,
      s.totalSgst || 0,
      s.totalCgst || 0,
      s.totalIgst || 0,
      s.totalGrand || 0,
    ].join('\t'));

    const text = lines.join('\n');
    navigator.clipboard.writeText(text).then(() => {
      const origText = btnCopyHsnTable.textContent;
      btnCopyHsnTable.textContent = 'Copied to Clipboard!';
      setTimeout(() => {
        btnCopyHsnTable.textContent = origText;
      }, 2000);
    }).catch(() => {
      alert('Unable to copy to clipboard.');
    });
  });
}

// 1-Click Jump from Consolidate Invoices to HSN Summary
const btnJumpToHsnSummary = document.getElementById('btnJumpToHsnSummary');
if (btnJumpToHsnSummary) {
  btnJumpToHsnSummary.addEventListener('click', async () => {
    if (!processedInvoices || processedInvoices.length === 0) {
      alert('Please process some invoices first.');
      return;
    }

    btnJumpToHsnSummary.disabled = true;
    const origText = btnJumpToHsnSummary.innerHTML;
    btnJumpToHsnSummary.textContent = 'Preparing HSN Summary...';

    try {
      // 1. Generate cumulative Excel buffer from current processed invoices
      const res = await fetch('/api/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          invoices: processedInvoices,
          periodTitle: periodTitleInput.value.trim() || 'April-June 2025-26',
          repeatInvoiceInfo: repeatInvoiceInfoToggle.checked,
          clubByHsn: clubByHsnToggle.checked,
          highestHsnOnly: highestHsnToggle.checked,
        }),
      });

      if (!res.ok) throw new Error('Failed to generate cumulative data.');
      const blob = await res.blob();

      // 2. Upload to HSN summary endpoint
      const formData = new FormData();
      formData.append('cumulativeSheet', blob, 'cumulative_export.xlsx');
      formData.append('periodTitle', periodTitleInput.value.trim() || 'July–September 2026');

      const hsnRes = await fetch('/api/hsn-summary/upload', {
        method: 'POST',
        body: formData,
      });

      const data = await hsnRes.json();
      if (!hsnRes.ok || !data.success) {
        throw new Error(data.error || 'Failed to generate HSN summary.');
      }

      currentHsnData = data;
      hsnPeriodInput.value = periodTitleInput.value.trim() || data.periodTitle || 'July–September 2026';

      // 3. Switch to HSN Summary Tab seamlessly on the same page
      const hsnTabBtn = document.querySelector('[data-tab-target="viewHsnSummary"]');
      if (hsnTabBtn) hsnTabBtn.click();

      // 4. Render HSN results
      renderHsnResults(data);
    } catch (err) {
      alert('Error transferring data to HSN summary: ' + err.message);
    } finally {
      btnJumpToHsnSummary.disabled = false;
      btnJumpToHsnSummary.innerHTML = origText;
    }
  });
}

// =========================================================
// CLICK-TO-COPY CELL VALUE (FEATURE 1 & FEATURE 2)
// =========================================================
const copyToastContainer = document.getElementById('copyToastContainer');

function showCopyToast(message, val = '', type = 'success') {
  if (!copyToastContainer) return;

  const toast = document.createElement('div');
  toast.className = `copy-toast ${type}`;

  const iconSvg = (type === 'info') 
    ? `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>`
    : `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#10b981" stroke-width="2.2"><path d="M20 6L9 17l-5-5"/></svg>`;

  let contentHtml = `<span class="copy-toast-icon">${iconSvg}</span> <span>${escapeHtml(message)}</span>`;
  if (val) {
    contentHtml += ` <span class="copy-toast-val">${escapeHtml(val)}</span>`;
  }

  toast.innerHTML = contentHtml;
  copyToastContainer.appendChild(toast);

  // Trigger smooth enter animation
  requestAnimationFrame(() => {
    toast.classList.add('show');
  });

  // Automatically remove toast
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => {
      if (toast.parentNode) toast.parentNode.removeChild(toast);
    }, 250);
  }, 2200);
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      // Fallback below
    }
  }

  // Universal fallback for insecure contexts / local network IPs / older browsers
  try {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.top = '-9999px';
    textArea.style.left = '-9999px';
    textArea.setAttribute('readonly', '');
    document.body.appendChild(textArea);
    textArea.select();
    const successful = document.execCommand('copy');
    document.body.removeChild(textArea);
    return successful;
  } catch (err) {
    console.error('Copy fallback failed:', err);
    return false;
  }
}

// Global delegated event listener for clicks inside table cells
document.addEventListener('click', (e) => {
  // Target any td in Feature 1 table or Feature 2 table
  const td = e.target.closest('#dataTable tbody td, #dataTable tfoot td, #hsnDataTable tbody td, #hsnDataTable tfoot td');
  if (!td) return;

  // If user selected text manually with mouse drag, don't trigger cell copy
  const selection = window.getSelection();
  if (selection && selection.toString().length > 1 && td.contains(selection.anchorNode) && td.contains(selection.focusNode)) {
    return;
  }

  const copyVal = td.getAttribute('data-copy-val');
  const rawText = td.innerText !== undefined ? td.innerText : td.textContent;
  const value = copyVal !== null ? copyVal : (rawText ? rawText.trim() : '');

  if (!value) {
    showCopyToast('Cell is empty', '', 'info');
    return;
  }

  copyTextToClipboard(value);

  // Flash highlight animation on the clicked cell
  td.classList.remove('cell-copied-flash');
  void td.offsetWidth; // force DOM reflow
  td.classList.add('cell-copied-flash');
  setTimeout(() => {
    td.classList.remove('cell-copied-flash');
  }, 500);

  showCopyToast('Copied to clipboard:', value, 'success');
});

// Dynamic tooltip on hover
document.addEventListener('mouseover', (e) => {
  const td = e.target.closest('#dataTable tbody td, #dataTable tfoot td, #hsnDataTable tbody td, #hsnDataTable tfoot td');
  if (td && !td.getAttribute('data-has-copy-tip')) {
    const copyVal = td.getAttribute('data-copy-val');
    const txt = (copyVal !== null ? copyVal : (td.innerText || td.textContent || '')).trim();
    if (txt) {
      td.title = `Click to copy: "${txt}"`;
      td.setAttribute('data-has-copy-tip', 'true');
    }
  }
});
