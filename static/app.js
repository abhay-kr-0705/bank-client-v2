/**
 * BankTech Valuation OCR & Dynamic Real-Time Excel Spreadsheet Engine
 * Frontend Interactive Controller
 */

document.addEventListener('DOMContentLoaded', () => {
  // Session State - Isolated per browser session or URL query parameter
  const urlParams = new URLSearchParams(window.location.search);
  let currentSessionId = urlParams.get('session_id') || sessionStorage.getItem('banktech_session_id');
  if (!currentSessionId) {
    currentSessionId = 'sess_' + Date.now().toString(36) + '_' + Math.random().toString(36).substring(2, 7);
    sessionStorage.setItem('banktech_session_id', currentSessionId);
  }
  let currentReportData = null;
  let currentGridData = null;
  let activeFiles = [];
  let selectedCellCoord = "A1";
  let currentActiveSheet = 0;

  const defaultFloorNames = [
    "Basement", "Stilt Floor", "Ground Floor", "First Floor",
    "Second Floor", "Third Floor", "Fouth Floor", "Fifth Floor",
    "Six Floor", "Seven Floor"
  ];

  // DOM Elements - Navigation & Actions
  const btnThemeToggle = document.getElementById('btnThemeToggle');
  const themeIcon = document.getElementById('themeIcon');
  const themeText = document.getElementById('themeText');
  const btnSettingsModal = document.getElementById('btnSettingsModal');
  const btnDownloadReport = document.getElementById('btnDownloadReport');
  const btnGenerateExcel = document.getElementById('btnGenerateExcel');
  const btnSyncFormToGrid = document.getElementById('btnSyncFormToGrid');
  
  // Template Management
  const btnUploadTemplate = document.getElementById('btnUploadTemplate');
  const templateFileInput = document.getElementById('templateFileInput');
  const btnPreviewCleanTemplate = document.getElementById('btnPreviewCleanTemplate');
  const btnResetTemplate = document.getElementById('btnResetTemplate');
  const templateNameText = document.getElementById('templateNameText');
  const templatePreviewModal = document.getElementById('templatePreviewModal');
  const btnCloseTemplatePreview = document.getElementById('btnCloseTemplatePreview');
  const btnConfirmTemplatePreview = document.getElementById('btnConfirmTemplatePreview');
  const templateModalViewport = document.getElementById('templateModalViewport');

  // Ingestion Elements
  const btnLoadSunitaCase = document.getElementById('btnLoadSunitaCase');
  const btnRunFolderPath = document.getElementById('btnRunFolderPath');
  const folderPathInput = document.getElementById('folderPathInput');
  const fileDropzone = document.getElementById('fileDropzone');
  const fileInput = document.getElementById('fileInput');
  const multiFileInput = document.getElementById('multiFileInput');
  const btnSelectFiles = document.getElementById('btnSelectFiles');
  const btnSelectFolder = document.getElementById('btnSelectFolder');
  
  const pipelineTracker = document.getElementById('pipelineTracker');
  const pipelineProgressBar = document.getElementById('pipelineProgressBar');
  const filesListContainer = document.getElementById('filesListContainer');
  const filesCountBadge = document.getElementById('filesCountBadge');
  const fileFilterInput = document.getElementById('fileFilterInput');
  const btnProcessAllFiles = document.getElementById('btnProcessAllFiles');
  const casePillBadge = document.getElementById('casePillBadge');

  // Spreadsheet Grid Elements
  const spreadsheetViewport = document.getElementById('spreadsheetViewport');
  const activeCellCoord = document.getElementById('activeCellCoord');
  const formulaBarInput = document.getElementById('formulaBarInput');
  const sheetTabsContainer = document.getElementById('sheetTabsContainer');
  const btnSheet1 = document.getElementById('btnSheet1');
  const btnSheet2 = document.getElementById('btnSheet2');

  // Document Preview Elements
  const docPreviewWrapper = document.getElementById('docPreviewWrapper');
  const previewFileName = document.getElementById('previewFileName');
  const previewBody = document.getElementById('previewBody');
  const btnClosePreview = document.getElementById('btnClosePreview');

  // Settings Modal
  const settingsModal = document.getElementById('settingsModal');
  const btnCloseSettings = document.getElementById('btnCloseSettings');
  const btnCancelSettings = document.getElementById('btnCancelSettings');
  const btnSaveSettings = document.getElementById('btnSaveSettings');
  const modalApiKey = document.getElementById('modalApiKey');
  const modalModelSelect = document.getElementById('modalModelSelect');
  const engineStatusBadge = document.getElementById('engineStatusBadge');
  const engineStatusText = document.getElementById('engineStatusText');

  // Floor Table Body
  const floorsTableBody = document.getElementById('floorsTableBody');

  // Initialize Application
  initTheme();
  initFloorTable();
  initTabs();
  initEventListeners();
  loadInitialGridState();
  checkBackendHealth();

  // 0. Theme Toggle
  function initTheme() {
    const savedTheme = localStorage.getItem('banktech_theme') || 'dark';
    applyTheme(savedTheme);
  }

  function applyTheme(theme) {
    if (theme === 'light') {
      document.body.classList.add('light-theme');
      document.body.classList.remove('dark-theme');
      if (themeIcon) themeIcon.className = 'fa-solid fa-moon';
      if (themeText) themeText.textContent = 'Dark Mode';
    } else {
      document.body.classList.add('dark-theme');
      document.body.classList.remove('light-theme');
      if (themeIcon) themeIcon.className = 'fa-solid fa-sun';
      if (themeText) themeText.textContent = 'Light Mode';
    }
    localStorage.setItem('banktech_theme', theme);
  }

  function toggleTheme() {
    const currentIsLight = document.body.classList.contains('light-theme');
    applyTheme(currentIsLight ? 'dark' : 'light');
    showToast(`Switched to ${currentIsLight ? 'Dark' : 'Light'} Mode`, 'info');
  }

  // 1. Initial Setup & Health Check
  async function checkBackendHealth() {
    try {
      const res = await fetch('/api/health');
      const data = await res.json();
      engineStatusText.textContent = `Offline RapidOCR & Local Semantic Engine Active`;
      engineStatusBadge.style.borderColor = 'rgba(16, 185, 129, 0.4)';
      engineStatusBadge.style.color = '#6ee7b7';
    } catch (e) {
      console.warn("Backend connection pending:", e);
    }
  }

  // 2. Tab Navigation
  function initTabs() {
    const tabs = document.querySelectorAll('.tab-btn');
    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        tabs.forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
        
        tab.classList.add('active');
        const targetId = tab.getAttribute('data-tab');
        const targetPane = document.getElementById(targetId);
        if (targetPane) targetPane.classList.add('active');

        // If switched to grid tab, refresh view
        if (targetId === 'tab-grid' && currentGridData) {
          renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
        }
      });
    });
  }

  // 3. Load Initial Grid State
  async function loadInitialGridState(sheet = 0) {
    currentActiveSheet = sheet;
    try {
      const res = await fetch(`/api/session/live-grid?session_id=${currentSessionId}&sheet=${sheet}`);
      const data = await res.json();
      if (data.success && data.grid) {
        currentGridData = data.grid;
        renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
        updateSheetTabs(data.grid.all_sheets || ['Sheet1', 'Sheet2'], sheet);
      }
    } catch (e) {
      console.error("Initial grid load error:", e);
    }
  }

  // Update Sheet Switcher Tabs
  function updateSheetTabs(sheetNames, activeIdx) {
    if (!sheetTabsContainer) return;
    sheetTabsContainer.innerHTML = '';
    sheetNames.forEach((name, idx) => {
      const btn = document.createElement('button');
      btn.className = `sheet-tab ${idx === activeIdx || name === activeIdx ? 'active' : ''}`;
      btn.innerHTML = `<i class="fa-regular fa-file-lines"></i> ${name}`;
      btn.addEventListener('click', () => {
        loadInitialGridState(idx);
      });
      sheetTabsContainer.appendChild(btn);
    });
  }

  // 4. Interactive Spreadsheet Grid Renderer
  function renderSpreadsheetGrid(gridData, container, isEditable = true) {
    if (!gridData || !gridData.rows || gridData.rows.length === 0) {
      container.innerHTML = `<div class="p-4 text-center text-muted">No grid data available.</div>`;
      return;
    }

    const table = document.createElement('table');
    table.className = 'excel-grid-table';
    table.tabIndex = 0; // Make table focusable for keyboard navigation

    // 1. Column Header Row (Corner + A, B, C, D...)
    const thead = document.createElement('thead');
    const colHeaderRow = document.createElement('tr');
    
    // Corner empty cell
    const cornerTh = document.createElement('th');
    cornerTh.className = 'corner-header';
    cornerTh.textContent = '#';
    colHeaderRow.appendChild(cornerTh);

    const firstRowCells = gridData.rows[0].cells;
    firstRowCells.forEach(cell => {
      const th = document.createElement('th');
      th.className = 'col-header';
      th.textContent = cell.col_letter;
      colHeaderRow.appendChild(th);
    });
    thead.appendChild(colHeaderRow);
    table.appendChild(thead);

    // 2. Data Rows
    const tbody = document.createElement('tbody');
    gridData.rows.forEach(rowObj => {
      const tr = document.createElement('tr');

      // Row Number Header
      const rowTh = document.createElement('th');
      rowTh.className = 'row-header';
      rowTh.textContent = rowObj.row;
      tr.appendChild(rowTh);

      // Data Cells
      rowObj.cells.forEach(cellObj => {
        const td = document.createElement('td');
        td.id = `cell_${cellObj.coord}`;
        td.setAttribute('data-coord', cellObj.coord);
        td.setAttribute('data-row', cellObj.row);
        td.setAttribute('data-col', cellObj.col);
        td.setAttribute('data-val', cellObj.value || '');

        let cellClass = 'grid-data-cell';
        if (cellObj.is_header) cellClass += ' grid-header-cell';
        if (cellObj.is_formula) cellClass += ' grid-formula-cell';
        if (cellObj.coord === selectedCellCoord) cellClass += ' cell-selected';
        
        // Review Highlighting (Confidence Decision)
        if (cellObj.needs_review) {
          cellClass += ' cell-needs-review';
          td.style.backgroundColor = '#fff3cd';
          td.style.color = '#92400e';
          td.style.border = '1px solid #f59e0b';
          const conf = Math.round((cellObj.confidence || 0.5) * 100);
          td.title = `[REVIEW NEEDED - ${conf}% Confidence]\n${cellObj.review_reason || 'Please verify this field value'}`;
        } else if (cellObj.fill_color) {
          td.style.backgroundColor = `#${cellObj.fill_color.slice(-6)}`;
        }
        td.className = cellClass;

        td.textContent = cellObj.value || '';

        // Cell Click Handler
        td.addEventListener('click', () => {
          selectCell(cellObj.coord, cellObj.value || '', td, container);
        });

        // Double click for direct in-cell editing
        if (isEditable) {
          td.addEventListener('dblclick', () => {
            makeCellEditable(td, cellObj.coord);
          });
        }

        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    container.innerHTML = '';
    container.appendChild(table);

    // Update bottom status info
    const statsElem = document.getElementById('gridStatsText');
    if (statsElem) {
      statsElem.textContent = `${gridData.max_row} Rows • ${gridData.max_col} Columns • Formula Engine Active`;
    }

    // Keyboard Navigation for Excel Grid
    initGridKeyboardNavigation(table, gridData, container, isEditable);
  }

  // Keyboard navigation (Arrow keys, Enter, Tab, F2)
  function initGridKeyboardNavigation(table, gridData, container, isEditable) {
    table.addEventListener('keydown', (e) => {
      if (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA') return;

      const currentTd = document.getElementById(`cell_${selectedCellCoord}`);
      if (!currentTd) return;

      const currRow = parseInt(currentTd.getAttribute('data-row') || '1');
      const currCol = parseInt(currentTd.getAttribute('data-col') || '1');
      let targetRow = currRow;
      let targetCol = currCol;

      if (e.key === 'ArrowUp') {
        targetRow = Math.max(1, currRow - 1);
        e.preventDefault();
      } else if (e.key === 'ArrowDown') {
        targetRow = Math.min(gridData.max_row, currRow + 1);
        e.preventDefault();
      } else if (e.key === 'ArrowLeft') {
        targetCol = Math.max(1, currCol - 1);
        e.preventDefault();
      } else if (e.key === 'ArrowRight' || e.key === 'Tab') {
        targetCol = Math.min(gridData.max_col, currCol + 1);
        e.preventDefault();
      } else if (e.key === 'Enter') {
        targetRow = Math.min(gridData.max_row, currRow + 1);
        e.preventDefault();
      } else if (e.key === 'F2' && isEditable) {
        makeCellEditable(currentTd, selectedCellCoord);
        e.preventDefault();
        return;
      }

      if (targetRow !== currRow || targetCol !== currCol) {
        const colLetter = getColLetter(targetCol);
        const nextCoord = `${colLetter}${targetRow}`;
        const nextTd = document.getElementById(`cell_${nextCoord}`);
        if (nextTd) {
          selectCell(nextCoord, nextTd.textContent || '', nextTd, container);
          nextTd.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        }
      }
    });
  }

  function getColLetter(colIdx) {
    let temp, letter = '';
    while (colIdx > 0) {
      temp = (colIdx - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      colIdx = (colIdx - temp - 1) / 26;
    }
    return letter;
  }

  function selectCell(coord, val, tdElem, container) {
    selectedCellCoord = coord;
    activeCellCoord.textContent = coord;
    formulaBarInput.value = val;

    container.querySelectorAll('.excel-grid-table td').forEach(td => td.classList.remove('cell-selected'));
    if (tdElem) tdElem.classList.add('cell-selected');
  }

  function makeCellEditable(td, coord) {
    const currentVal = td.textContent;
    td.innerHTML = `<input type="text" class="grid-inline-input" value="${currentVal.replace(/"/g, '&quot;')}">`;
    const input = td.querySelector('input');
    input.focus();
    input.select();

    const commitChange = async () => {
      const newVal = input.value;
      td.textContent = newVal;
      formulaBarInput.value = newVal;
      await syncCellUpdate(coord, newVal);
    };

    input.addEventListener('blur', commitChange);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        input.blur();
      } else if (e.key === 'Escape') {
        td.textContent = currentVal;
      }
    });
  }

  // Formula Bar Realtime Input
  formulaBarInput.addEventListener('keydown', async (e) => {
    if (e.key === 'Enter' && selectedCellCoord) {
      const newVal = formulaBarInput.value;
      const td = document.getElementById(`cell_${selectedCellCoord}`);
      if (td) td.textContent = newVal;
      await syncCellUpdate(selectedCellCoord, newVal);
      showToast(`Cell ${selectedCellCoord} updated`, 'info');
    }
  });

  async function syncCellUpdate(coord, val) {
    try {
      await fetch('/api/session/update-cell', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session_id: currentSessionId, coordinate: coord, value: val })
      });
    } catch (e) {
      console.error("Cell update error:", e);
    }
  }

  // Grid Modification Actions: Add Row, Add Col, Clear Data
  const btnAddGridRow = document.getElementById('btnAddGridRow');
  const btnAddGridCol = document.getElementById('btnAddGridCol');
  const btnClearGridData = document.getElementById('btnClearGridData');

  if (btnAddGridRow) {
    btnAddGridRow.addEventListener('click', async () => {
      showToast('Adding new row to Excel grid...', 'info');
      try {
        const res = await fetch(`/api/session/add-row?session_id=${currentSessionId}`, { method: 'POST' });
        const data = await res.json();
        if (data.success && data.grid) {
          currentGridData = data.grid;
          renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
          showToast(`Row ${currentGridData.max_row} added!`, 'success');
        }
      } catch (e) {
        showToast('Failed to add row', 'error');
      }
    });
  }

  if (btnAddGridCol) {
    btnAddGridCol.addEventListener('click', async () => {
      showToast('Adding new column to Excel grid...', 'info');
      try {
        const res = await fetch(`/api/session/add-column?session_id=${currentSessionId}`, { method: 'POST' });
        const data = await res.json();
        if (data.success && data.grid) {
          currentGridData = data.grid;
          renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
          showToast(`Column added! Total: ${currentGridData.max_col}`, 'success');
        }
      } catch (e) {
        showToast('Failed to add column', 'error');
      }
    });
  }

  if (btnClearGridData) {
    btnClearGridData.addEventListener('click', async () => {
      if (!confirm('Are you sure you want to clear all data cells? (Headers, labels, and formulas will be preserved)')) return;
      try {
        const res = await fetch(`/api/session/clear-data?session_id=${currentSessionId}`, { method: 'POST' });
        const data = await res.json();
        if (data.success && data.grid) {
          currentGridData = data.grid;
          currentReportData = data.report_data;
          renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
          populateFormWithData(currentReportData);
          showToast('Data cells cleared. Template structure & formulas preserved.', 'info');
        }
      } catch (e) {
        showToast('Failed to clear data', 'error');
      }
    });
  }

  // Collapsible Ingestion Deck
  const btnToggleIngestDeck = document.getElementById('btnToggleIngestDeck');
  const toggleIngestIcon = document.getElementById('toggleIngestIcon');
  const ingestCard = document.querySelector('.ingest-card');

  if (btnToggleIngestDeck && ingestCard) {
    btnToggleIngestDeck.addEventListener('click', () => {
      ingestCard.classList.toggle('collapsed');
      const isCollapsed = ingestCard.classList.contains('collapsed');
      if (toggleIngestIcon) {
        toggleIngestIcon.className = isCollapsed ? 'fa-solid fa-chevron-down' : 'fa-solid fa-chevron-up';
      }
    });
  }

  // Expand Grid View Toggle
  const btnToggleExpandGrid = document.getElementById('btnToggleExpandGrid');
  const expandGridIcon = document.getElementById('expandGridIcon');
  const expandGridText = document.getElementById('expandGridText');
  const workspaceArea = document.getElementById('workspaceArea');

  if (btnToggleExpandGrid && workspaceArea) {
    btnToggleExpandGrid.addEventListener('click', () => {
      workspaceArea.classList.toggle('grid-expanded');
      const isExpanded = workspaceArea.classList.contains('grid-expanded');
      if (expandGridIcon) expandGridIcon.className = isExpanded ? 'fa-solid fa-compress' : 'fa-solid fa-expand';
      if (expandGridText) expandGridText.textContent = isExpanded ? 'Restore View' : 'Expand View';
      showToast(isExpanded ? 'Spreadsheet expanded to full workspace' : 'Workspace layout restored', 'info');
    });
  }

  // 5. Template Management (Upload, Preview Blank, Reset)
  btnUploadTemplate.addEventListener('click', () => templateFileInput.click());

  templateFileInput.addEventListener('change', async (e) => {
    if (!e.target.files || e.target.files.length === 0) return;
    const file = e.target.files[0];
    showToast(`Sanitizing custom template: ${file.name}...`, 'info');

    const formData = new FormData();
    formData.append('file', file);
    formData.append('session_id', currentSessionId);

    try {
      const res = await fetch('/api/template/upload', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (data.success) {
        templateNameText.textContent = `Template: ${data.template_name}`;
        currentGridData = data.grid;
        renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
        showToast(`Template ${file.name} loaded & sanitized! All headers & formulas preserved.`, 'success');
      }
    } catch (err) {
      showToast(`Template upload failed: ${err.message}`, 'error');
    }
  });

  btnResetTemplate.addEventListener('click', async () => {
    try {
      const res = await fetch(`/api/template/reset?session_id=${currentSessionId}`, { method: 'POST' });
      const data = await res.json();
      templateNameText.textContent = `Template: Base Template (India Shelter)`;
      currentGridData = data.grid;
      renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
      showToast('Reset to default base template.', 'info');
    } catch (e) {
      showToast('Reset failed', 'error');
    }
  });

  btnPreviewCleanTemplate.addEventListener('click', async () => {
    showToast('Loading clean blank blueprint...', 'info');
    try {
      const res = await fetch(`/api/template/clean-preview?session_id=${currentSessionId}`);
      const grid = await res.json();
      templatePreviewModal.style.display = 'flex';
      renderSpreadsheetGrid(grid, templateModalViewport, false);
    } catch (e) {
      showToast('Failed to load template preview', 'error');
    }
  });

  btnCloseTemplatePreview.addEventListener('click', () => templatePreviewModal.style.display = 'none');
  btnConfirmTemplatePreview.addEventListener('click', () => templatePreviewModal.style.display = 'none');

  // Safe JSON parser that catches HTML 502/504/524 proxy errors gracefully
  async function safeJson(res) {
    const text = await res.text();
    try {
      return JSON.parse(text);
    } catch (e) {
      if (text.includes("524") || text.includes("timeout")) {
        return { detail: "Cloudflare proxy timeout (524). Background task is actively continuing on host server." };
      }
      if (text.includes("502") || text.includes("Bad Gateway")) {
        return { detail: "Bad Gateway (502). Host server connection was interrupted." };
      }
      return { detail: `Server error (${res.status} ${res.statusText || 'non-JSON response'})` };
    }
  }

  let isSyncingCompletedState = false;
  async function syncCompletedSessionState(message = "") {
    if (isSyncingCompletedState) return;
    isSyncingCompletedState = true;
    try {
      const res = await fetch(`/api/session/live-grid?session_id=${currentSessionId}`);
      const data = await safeJson(res);
      if (data.success && data.grid) {
        currentReportData = data.report_data;
        currentGridData = data.grid;
        populateFormWithData(currentReportData);
        renderFilesList(data.files || []);
        renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
        btnDownloadReport.disabled = false;
        finishPipelineTracker(true, message || 'Extraction complete! Grid and formulas populated.');
        showToast('Extraction complete! Live grid & formulas populated.', 'success');
      }
    } catch (e) {
      console.warn("Failed to sync completed session:", e);
      finishPipelineTracker(true, message);
    } finally {
      isSyncingCompletedState = false;
    }
  }

  // 6. Incremental File & Multi-format Ingestion
  async function handleFileUploads(files) {
    if (!files || files.length === 0) return;

    showToast(`Uploading ${files.length} document(s) & queuing OCR...`, 'info');
    startPipelineTracker(`Uploading & extracting ${files.length} document(s)...`);

    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
      formData.append('files', files[i]);
    }
    formData.append('session_id', currentSessionId);

    try {
      const res = await fetch('/api/session/upload-files', {
        method: 'POST',
        body: formData
      });

      const result = await safeJson(res);
      if (!res.ok) {
        throw new Error(result.detail || 'Upload failed');
      }

      if (result.status === 'processing') {
        showToast(`Files uploaded! Running RapidOCR in background...`, 'info');
      } else if (result.grid) {
        syncCompletedSessionState(`Successfully processed ${files.length} document(s)!`);
      }
    } catch (e) {
      finishPipelineTracker(false, e.message);
      showToast(`Upload Error: ${e.message}`, 'error');
      console.error(e);
    }
  }

  // Direct Folder Ingestion
  async function processFolderPath(folderPath) {
    if (!folderPath || !folderPath.trim()) {
      showToast('Please enter a valid folder path', 'error');
      return;
    }

    showToast(`Scanning & Extracting: ${folderPath}`, 'info');
    startPipelineTracker(`Scanning host folder: ${folderPath}...`);

    try {
      const res = await fetch('/api/process-folder', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder_path: folderPath.trim(), session_id: currentSessionId })
      });

      const result = await safeJson(res);
      if (!res.ok) {
        throw new Error(result.detail || 'Extraction failed');
      }

      if (result.status === 'processing') {
        showToast(`Folder queued! RapidOCR pipeline running in background...`, 'info');
      } else if (result.grid) {
        syncCompletedSessionState(`Folder processed! All files mapped.`);
      }
    } catch (e) {
      finishPipelineTracker(false, e.message);
      showToast(`Error: ${e.message}`, 'error');
      console.error(e);
    }
  }

  // Single-File Force Extraction Endpoint Call
  async function processSingleFile(filePath, fileName) {
    showToast(`Force extracting from: ${fileName}...`, 'info');
    startPipelineTracker(`Force analyzing single file: ${fileName}...`);
    try {
      const res = await fetch('/api/session/process-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_path: filePath, session_id: currentSessionId })
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'File extraction failed');
      }

      const result = await res.json();
      finishPipelineTracker(true, `Extracted ${fileName} successfully!`);
      currentReportData = result.report_data;
      currentGridData = result.grid;
      
      populateFormWithData(currentReportData);
      renderFilesList(result.files || []);
      renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
      
      btnDownloadReport.disabled = false;
      showToast(`Extracted data from ${fileName} & updated live grid!`, 'success');
    } catch (e) {
      finishPipelineTracker(false, e.message);
      showToast(`File Extraction Error: ${e.message}`, 'error');
    }
  }

  // Remove File from Session
  async function removeSessionFile(filePath, fileName) {
    if (!confirm(`Remove ${fileName} from active session?`)) return;
    try {
      const res = await fetch('/api/session/remove-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file_path: filePath, session_id: currentSessionId })
      });
      const data = await res.json();
      if (data.success) {
        renderFilesList(data.files || []);
        showToast(`${fileName} removed from session.`, 'info');
      }
    } catch (e) {
      showToast(`Failed to remove file`, 'error');
    }
  }

  // 7. Real-Time Pipeline Progress Tracker (Live Backend Polling & Accurate Timer)
  let pipelinePollInterval = null;
  let pipelineTimerInterval = null;
  let pipelineStartTime = 0;

  function startPipelineTracker(customMsg = "Initializing OCR and document processing...") {
    if (pipelinePollInterval) clearInterval(pipelinePollInterval);
    if (pipelineTimerInterval) clearInterval(pipelineTimerInterval);

    const s1 = document.getElementById('step1');
    const s2 = document.getElementById('step2');
    const s3 = document.getElementById('step3');
    const s4 = document.getElementById('step4');
    const statusText = document.getElementById('pipelineStatusText');
    const timerBadge = document.getElementById('pipelineTimerBadge');

    pipelineTracker.style.display = 'block';
    pipelineProgressBar.className = 'progress-bar';
    pipelineProgressBar.style.width = '10%';

    s1.className = 'step active';
    s2.className = 'step';
    s3.className = 'step';
    s4.className = 'step';

    if (statusText) statusText.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin text-accent"></i> ${customMsg}`;
    pipelineStartTime = Date.now();

    // Elapsed timer updater
    pipelineTimerInterval = setInterval(() => {
      const elapsed = Math.floor((Date.now() - pipelineStartTime) / 1000);
      const mins = Math.floor(elapsed / 60);
      const secs = elapsed % 60;
      const timeStr = `${mins > 0 ? mins + 'm ' : ''}${secs}s`;
      if (timerBadge) {
        timerBadge.innerHTML = `<i class="fa-solid fa-stopwatch text-accent"></i> Elapsed: ${timeStr} • Est: ~30-50s`;
      }
    }, 1000);

    // Live polling from backend
    pipelinePollInterval = setInterval(async () => {
      try {
        const res = await fetch(`/api/session/progress?session_id=${currentSessionId}`);
        if (!res.ok) return;
        const p = await safeJson(res);

        if (p && p.percent !== undefined) {
          pipelineProgressBar.style.width = Math.max(10, Math.min(p.percent, 96)) + '%';
          
          if (statusText && p.message) {
            statusText.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin text-accent"></i> ${p.message}`;
          }

          // Update step states based on real step (1 to 4)
          const currentStep = p.step || 1;
          s1.className = currentStep > 1 ? 'step completed' : (currentStep === 1 ? 'step active' : 'step');
          s2.className = currentStep > 2 ? 'step completed' : (currentStep === 2 ? 'step active' : 'step');
          s3.className = currentStep > 3 ? 'step completed' : (currentStep === 3 ? 'step active' : 'step');
          s4.className = currentStep === 4 ? 'step active' : 'step';

          // Automatic completion and error handling
          if (p.state === 'completed') {
            await syncCompletedSessionState(p.message);
          } else if (p.state === 'error') {
            finishPipelineTracker(false, p.message);
          }
        }
      } catch (err) {
        // Silently retry on transient network errors
      }
    }, 600);
  }

  function finishPipelineTracker(success = true, message = "") {
    if (pipelinePollInterval) {
      clearInterval(pipelinePollInterval);
      pipelinePollInterval = null;
    }
    if (pipelineTimerInterval) {
      clearInterval(pipelineTimerInterval);
      pipelineTimerInterval = null;
    }

    const s1 = document.getElementById('step1');
    const s2 = document.getElementById('step2');
    const s3 = document.getElementById('step3');
    const s4 = document.getElementById('step4');
    const statusText = document.getElementById('pipelineStatusText');
    const timerBadge = document.getElementById('pipelineTimerBadge');

    const totalSeconds = Math.floor((Date.now() - pipelineStartTime) / 1000);

    if (success) {
      pipelineProgressBar.style.width = '100%';
      s1.className = 'step completed';
      s2.className = 'step completed';
      s3.className = 'step completed';
      s4.className = 'step completed';
      
      if (statusText) {
        statusText.innerHTML = `<i class="fa-solid fa-circle-check text-success"></i> ${message || 'Extraction complete! Grid and formulas populated.'}`;
      }
      if (timerBadge) {
        timerBadge.innerHTML = `<i class="fa-solid fa-check text-success"></i> Total Time: ${totalSeconds}s`;
      }
    } else {
      pipelineProgressBar.classList.add('error');
      if (statusText) {
        statusText.innerHTML = `<i class="fa-solid fa-circle-xmark text-danger"></i> ${message || 'Extraction interrupted or failed.'}`;
      }
      if (timerBadge) {
        timerBadge.innerHTML = `<i class="fa-solid fa-triangle-exclamation text-danger"></i> Stopped at ${totalSeconds}s`;
      }
    }
  }

  // 8. Render Files List in Left Pane with Granular Action Controls
  function renderFilesList(files) {
    activeFiles = files;
    filesCountBadge.textContent = files.length;

    if (files.length === 0) {
      filesListContainer.innerHTML = `
        <div class="empty-state-card">
          <i class="fa-regular fa-folder-open"></i>
          <p>No documents found in session.</p>
          <span class="text-xs text-muted">Upload files, folders, or ZIP archives above</span>
        </div>`;
      return;
    }

    filesListContainer.innerHTML = '';
    const filterQuery = (fileFilterInput?.value || '').toLowerCase().trim();

    files.forEach(file => {
      if (filterQuery && !file.filename.toLowerCase().includes(filterQuery)) return;

      const div = document.createElement('div');
      div.className = 'file-item';
      
      let icon = 'fa-solid fa-file';
      const ext = (file.type || '').replace('.', '').toLowerCase();
      if (['jpg', 'jpeg', 'png', 'webp', 'site_photo', 'bmp', 'tiff'].includes(ext)) icon = 'fa-solid fa-file-image text-warning';
      else if (['pdf', 'property_pdf'].includes(ext)) icon = 'fa-solid fa-file-pdf text-danger';
      else if (['docx', 'doc', 'field_notes_docx'].includes(ext)) icon = 'fa-solid fa-file-word text-accent';
      else if (['xlsx', 'csv', 'excel', 'target_excel_template'].includes(ext)) icon = 'fa-solid fa-file-excel text-success';
      else if (['zip'].includes(ext)) icon = 'fa-solid fa-file-zipper text-warning';

      const status = file.status || 'Ready';
      let statusBadgeClass = 'badge-ready';
      if (status === 'Processed') statusBadgeClass = 'badge-processed';
      if (status === 'Updated') statusBadgeClass = 'badge-updated';

      div.innerHTML = `
        <div class="file-info-group">
          <i class="${icon} file-icon"></i>
          <div class="file-details">
            <div class="file-name truncate" title="${file.filename}">${file.filename}</div>
            <div class="file-subtext">
              <span>${file.size_display || ''}</span> &bull; 
              <span class="file-status-badge ${statusBadgeClass}">${status}</span>
            </div>
          </div>
        </div>
        <div class="file-actions-btn-group">
          <button class="btn-action-sm btn-extract" title="Extract / Re-extract this document only">
            <i class="fa-solid fa-bolt"></i> Extract
          </button>
          <button class="btn-action-sm btn-preview" title="Preview document">
            <i class="fa-solid fa-eye"></i>
          </button>
          <button class="btn-action-sm btn-delete" title="Remove from session">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      `;

      // Extract this file button
      const extractBtn = div.querySelector('.btn-extract');
      extractBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        processSingleFile(file.path, file.filename);
      });

      // Preview button
      const previewBtn = div.querySelector('.btn-preview');
      previewBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        previewFile(file);
      });

      // Delete button
      const deleteBtn = div.querySelector('.btn-delete');
      deleteBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        removeSessionFile(file.path, file.filename);
      });

      div.addEventListener('click', () => previewFile(file));
      filesListContainer.appendChild(div);
    });
  }

  // Filter input event
  if (fileFilterInput) {
    fileFilterInput.addEventListener('input', () => {
      renderFilesList(activeFiles);
    });
  }

  // Process all files button
  if (btnProcessAllFiles) {
    btnProcessAllFiles.addEventListener('click', () => {
      const path = folderPathInput.value || "d:\\ABHAY VICKY\\Banking client project\\SUNITA DEVI BACHHAN KUMAR";
      processFolderPath(path);
    });
  }

  // 9. Document Previewer
  function previewFile(file) {
    previewFileName.textContent = file.filename;
    docPreviewWrapper.style.display = 'block';
    previewBody.innerHTML = '';

    const ext = file.filename.split('.').pop().toLowerCase();
    const fileUrl = `/api/view-file?filepath=${encodeURIComponent(file.path)}`;

    if (['jpg', 'jpeg', 'png', 'webp', 'bmp'].includes(ext)) {
      previewBody.innerHTML = `<img src="${fileUrl}" alt="${file.filename}">`;
    } else if (ext === 'pdf') {
      previewBody.innerHTML = `<iframe src="${fileUrl}#toolbar=0"></iframe>`;
    } else {
      previewBody.innerHTML = `
        <div class="text-center p-3 text-muted text-sm">
          <i class="fa-solid fa-file-lines fa-2x mb-2 text-accent"></i>
          <p>${file.filename}</p>
          <a href="${fileUrl}" target="_blank" class="btn btn-sm btn-outline mt-2">Open Raw File</a>
        </div>
      `;
    }
  }

  btnClosePreview.addEventListener('click', () => {
    docPreviewWrapper.style.display = 'none';
  });

  // 10. Populate Form Fields from Extracted JSON
  function populateFormWithData(data) {
    if (!data) return;

    casePillBadge.innerHTML = `<i class="fa-solid fa-id-card"></i> Case: ${data.header?.applicant_name || data.case_name || 'Active Case'}`;

    // Header
    const h = data.header || {};
    document.getElementById('inp_report_title').value = h.report_title || "India Shelter Report";
    document.getElementById('inp_application_id').value = h.application_id || "";
    document.getElementById('inp_applicant_name').value = h.applicant_name || "";
    document.getElementById('inp_property_type').value = h.property_type || "";
    document.getElementById('inp_completion_percent').value = (h.completion_percent !== undefined && h.completion_percent > 0) ? h.completion_percent : "";
    document.getElementById('inp_structure_type').value = h.structure_type || "";
    document.getElementById('inp_age_of_property').value = h.age_of_property || "";
    document.getElementById('inp_dwelling_units_owned').value = h.dwelling_units_owned || "";
    document.getElementById('inp_geo_tag').value = h.geo_tag || "";

    // Address
    const a = data.address || {};
    document.getElementById('inp_street_name').value = a.street_name || "";
    document.getElementById('inp_nearest_landmark').value = a.nearest_landmark || "";
    document.getElementById('inp_village_name').value = a.village_name || "";
    document.getElementById('inp_city').value = a.city || "";
    document.getElementById('inp_plot_house_khasra').value = a.plot_house_khasra || "";
    document.getElementById('inp_floor_number').value = a.floor_number || "";
    document.getElementById('inp_colony_name').value = a.colony_name || "";
    document.getElementById('inp_address_site').value = a.address_site || "";
    document.getElementById('inp_address_docs').value = a.address_docs || "";
    document.getElementById('inp_pincode').value = a.pincode || "";
    document.getElementById('inp_district').value = a.district || "";

    // Boundaries
    const b = data.boundaries || {};
    document.getElementById('inp_site_east').value = b.site_east || "";
    document.getElementById('inp_site_west').value = b.site_west || "";
    document.getElementById('inp_site_north').value = b.site_north || "";
    document.getElementById('inp_site_south').value = b.site_south || "";
    document.getElementById('inp_deed_east').value = b.deed_east || "";
    document.getElementById('inp_deed_west').value = b.deed_west || "";
    document.getElementById('inp_deed_north').value = b.deed_north || "";
    document.getElementById('inp_deed_south').value = b.deed_south || "";
    document.getElementById('inp_boundary_matching').value = b.boundary_matching ? b.boundary_matching.trim() : "";
    document.getElementById('inp_mismatch_remarks').value = b.mismatch_remarks || "";
    document.getElementById('inp_occupancy_status').value = b.occupancy_status ? b.occupancy_status.trim() : "";

    // Land
    const l = data.land_measurements || {};
    document.getElementById('inp_land_length').value = l.land_length || "";
    document.getElementById('inp_land_breadth').value = l.land_breadth || "";
    document.getElementById('inp_land_area_site_sqft').value = l.land_area_site_sqft || "";
    document.getElementById('inp_adopted_land_area_sqft').value = l.adopted_land_area_sqft || "";
    document.getElementById('inp_per_unit_land_rate').value = l.per_unit_land_rate || "";

    // Floors Table
    const f = data.construction_floors || {};
    initFloorTable(f.floors);
    document.getElementById('inp_built_up_rate').value = f.built_up_rate || "";
    document.getElementById('inp_total_property_value').value = f.total_property_value || "";
    recalculateFloorTotals();

    // Legal Checks
    const leg = data.legal_checks || {};
    document.getElementById('inp_documents_name').value = leg.documents_name || "";
    document.getElementById('inp_person_met').value = leg.person_met || "";
    document.getElementById('inp_relation_with_owner').value = leg.relation_with_owner || "";
    document.getElementById('inp_property_situated_at').value = leg.property_situated_at ? leg.property_situated_at.trim() : "";
    document.getElementById('inp_is_sanction_plan_compliant').value = leg.is_sanction_plan_compliant || "";
    document.getElementById('inp_pathway_clear').value = leg.pathway_clear || "";
    document.getElementById('inp_approach_by_public_road').value = leg.approach_by_public_road ? leg.approach_by_public_road.trim() : "";
    document.getElementById('inp_width_of_public_road').value = leg.width_of_public_road || "";
    document.getElementById('inp_current_uses').value = leg.current_uses || "";
    document.getElementById('inp_opinion_about_report').value = leg.opinion_about_report ? leg.opinion_about_report.trim() : "";
    document.getElementById('inp_occupancy_250m').value = leg.occupancy_250m || "";
    document.getElementById('inp_tentative_rent').value = leg.tentative_rent || "";
    document.getElementById('inp_development_250m').value = leg.development_250m || "";
    document.getElementById('inp_property_limit').value = leg.property_limit || "";
    document.getElementById('inp_adm').value = leg.adm || "";

    // Reference
    const r = data.reference || {};
    document.getElementById('inp_reference_name').value = r.reference_name || "";
    document.getElementById('inp_reference_mobile').value = r.reference_mobile || "";
    document.getElementById('inp_feedback').value = r.feedback || "";

    // Remarks
    document.getElementById('inp_remarks').value = data.remarks || "";
  }

  // Sync Form to Grid Action
  if (btnSyncFormToGrid) {
    btnSyncFormToGrid.addEventListener('click', async () => {
      showToast('Syncing form values to live grid & server session...', 'info');
      
      const payload = {
        report_title: document.getElementById('inp_report_title').value,
        header: {
          report_title: document.getElementById('inp_report_title').value,
          application_id: document.getElementById('inp_application_id').value,
          applicant_name: document.getElementById('inp_applicant_name').value,
          property_type: document.getElementById('inp_property_type').value,
          completion_percent: parseFloat(document.getElementById('inp_completion_percent').value || 0.0),
          structure_type: document.getElementById('inp_structure_type').value,
          age_of_property: document.getElementById('inp_age_of_property').value,
          dwelling_units_owned: parseInt(document.getElementById('inp_dwelling_units_owned').value || 0),
          geo_tag: document.getElementById('inp_geo_tag').value
        },
        address: {
          street_name: document.getElementById('inp_street_name').value,
          nearest_landmark: document.getElementById('inp_nearest_landmark').value,
          village_name: document.getElementById('inp_village_name').value,
          city: document.getElementById('inp_city').value,
          plot_house_khasra: document.getElementById('inp_plot_house_khasra').value,
          floor_number: document.getElementById('inp_floor_number').value,
          colony_name: document.getElementById('inp_colony_name').value,
          address_site: document.getElementById('inp_address_site').value,
          address_docs: document.getElementById('inp_address_docs').value,
          pincode: document.getElementById('inp_pincode').value,
          district: document.getElementById('inp_district').value
        },
        boundaries: {
          site_east: document.getElementById('inp_site_east').value,
          site_west: document.getElementById('inp_site_west').value,
          site_north: document.getElementById('inp_site_north').value,
          site_south: document.getElementById('inp_site_south').value,
          deed_east: document.getElementById('inp_deed_east').value,
          deed_west: document.getElementById('inp_deed_west').value,
          deed_north: document.getElementById('inp_deed_north').value,
          deed_south: document.getElementById('inp_deed_south').value,
          boundary_matching: document.getElementById('inp_boundary_matching').value,
          mismatch_remarks: document.getElementById('inp_mismatch_remarks').value,
          occupancy_status: document.getElementById('inp_occupancy_status').value
        },
        land_measurements: {
          land_length: parseFloat(document.getElementById('inp_land_length').value || 0.0),
          land_breadth: parseFloat(document.getElementById('inp_land_breadth').value || 0.0),
          land_area_site_sqft: document.getElementById('inp_land_area_site_sqft').value,
          adopted_land_area_sqft: parseFloat(document.getElementById('inp_adopted_land_area_sqft').value || 0.0),
          per_unit_land_rate: parseFloat(document.getElementById('inp_per_unit_land_rate').value || 0.0),
          total_land_value: "=B46*B47"
        },
        legal_checks: {
          documents_name: document.getElementById('inp_documents_name').value,
          person_met: document.getElementById('inp_person_met').value,
          relation_with_owner: document.getElementById('inp_relation_with_owner').value,
          property_situated_at: document.getElementById('inp_property_situated_at').value,
          is_sanction_plan_compliant: document.getElementById('inp_is_sanction_plan_compliant').value,
          pathway_clear: document.getElementById('inp_pathway_clear').value,
          approach_by_public_road: document.getElementById('inp_approach_by_public_road').value,
          width_of_public_road: document.getElementById('inp_width_of_public_road').value,
          current_uses: document.getElementById('inp_current_uses').value,
          opinion_about_report: document.getElementById('inp_opinion_about_report').value,
          occupancy_250m: document.getElementById('inp_occupancy_250m').value,
          tentative_rent: document.getElementById('inp_tentative_rent').value,
          development_250m: document.getElementById('inp_development_250m').value,
          property_limit: document.getElementById('inp_property_limit').value,
          adm: document.getElementById('inp_adm').value
        },
        reference: {
          reference_name: document.getElementById('inp_reference_name').value,
          reference_mobile: document.getElementById('inp_reference_mobile').value,
          feedback: document.getElementById('inp_feedback').value
        },
        remarks: document.getElementById('inp_remarks').value
      };

      try {
        const res = await fetch('/api/session/update-report', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            session_id: currentSessionId,
            report_data: payload
          })
        });
        const data = await res.json();
        if (data.success) {
          currentReportData = data.report_data;
          currentGridData = data.grid;
          renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
          showToast('Form successfully saved to session & live grid updated!', 'success');
        }
      } catch (err) {
        showToast(`Sync failed: ${err.message}`, 'error');
      }
    });
  }

  // 11. Floors Table Init
  function initFloorTable(floorsData = null) {
    floorsTableBody.innerHTML = '';
    defaultFloorNames.forEach((name, idx) => {
      const floor = (floorsData && floorsData[idx]) ? floorsData[idx] : { name, actual_area: 0, permissible_area: 0, adopted_area: 0 };
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="font-bold">${floor.name}</td>
        <td><input type="number" class="floor-actual" data-idx="${idx}" value="${floor.actual_area}"></td>
        <td><input type="number" class="floor-permissible" data-idx="${idx}" value="${floor.permissible_area}"></td>
        <td><input type="number" class="floor-adopted" data-idx="${idx}" value="${floor.adopted_area}"></td>
      `;
      floorsTableBody.appendChild(tr);
    });

    document.querySelectorAll('.floor-actual, .floor-adopted').forEach(input => {
      input.addEventListener('input', recalculateFloorTotals);
    });
  }

  function recalculateFloorTotals() {
    let sumActual = 0;
    let sumAdopted = 0;
    document.querySelectorAll('.floor-actual').forEach(inp => sumActual += parseFloat(inp.value || 0));
    document.querySelectorAll('.floor-adopted').forEach(inp => sumAdopted += parseFloat(inp.value || 0));
    
    document.getElementById('totalActualArea').textContent = `=SUM(B52:B61) (${sumActual})`;
    document.getElementById('totalAdoptedArea').textContent = `=SUM(D52:D61) (${sumAdopted})`;
  }

  // 12. Final Excel Generation & Download with Real-Time HUD Tracker
  let excelGenTimerInterval = null;
  let excelGenPollInterval = null;

  async function generateAndDownloadExcel() {
    const buttons = [btnDownloadReport, btnGenerateExcel].filter(Boolean);
    
    // Check if already running
    if (buttons.some(b => b.classList.contains('btn-generating'))) return;

    // Cache original button content
    const originalButtonContents = buttons.map(b => b.innerHTML);
    buttons.forEach(b => {
      b.disabled = true;
      b.classList.add('btn-generating');
      b.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Generating... <span class="btn-gen-time">0.0s</span>`;
    });

    // Show Excel HUD
    const hud = document.getElementById('excelGenHud');
    const hudProgressBar = document.getElementById('excelGenProgressBar');
    const hudStatusMsg = document.getElementById('excelGenStatusMsg');
    const hudPercent = document.getElementById('excelGenPercent');
    const hudTimer = document.getElementById('excelGenTimer');
    const hudTitle = document.getElementById('excelGenTitle');
    const genSteps = [
      document.getElementById('genStep1'),
      document.getElementById('genStep2'),
      document.getElementById('genStep3'),
      document.getElementById('genStep4')
    ];
    const genLines = [
      document.getElementById('genLine1'),
      document.getElementById('genLine2'),
      document.getElementById('genLine3')
    ];

    if (hud) {
      hud.style.display = 'block';
      hud.style.opacity = '1';
      hud.style.transform = 'translateY(0)';
      if (hudProgressBar) hudProgressBar.style.width = '15%';
      if (hudPercent) hudPercent.textContent = '15%';
      if (hudTitle) hudTitle.textContent = 'Generating Valuation Report';
      if (hudStatusMsg) {
        hudStatusMsg.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin text-accent"></i> Loading template blueprint & styles...`;
      }
      if (hudTimer) hudTimer.innerHTML = `<i class="fa-solid fa-stopwatch"></i> 0.0s`;
      genSteps.forEach((s, idx) => {
        if (s) s.className = 'gen-step' + (idx === 0 ? ' active' : '');
      });
      genLines.forEach(l => { if (l) l.className = 'gen-step-line'; });
    }

    const startTime = performance.now();
    let currentPercent = 15;

    // Stopwatch ticker (updates every 100ms)
    excelGenTimerInterval = setInterval(() => {
      const elapsed = ((performance.now() - startTime) / 1000).toFixed(1);
      if (hudTimer) hudTimer.innerHTML = `<i class="fa-solid fa-stopwatch"></i> ${elapsed}s`;
      document.querySelectorAll('.btn-gen-time').forEach(el => el.textContent = `${elapsed}s`);
    }, 100);

    // Live progress poller from backend session state (updates every 250ms)
    excelGenPollInterval = setInterval(async () => {
      try {
        const res = await fetch(`/api/session/progress?session_id=${currentSessionId}`);
        if (!res.ok) return;
        const p = await res.json();
        if (p.state === 'generating_excel' || p.state === 'completed') {
          const pct = Math.max(currentPercent, Math.min(p.percent || 15, 96));
          currentPercent = pct;
          if (hudProgressBar) hudProgressBar.style.width = pct + '%';
          if (hudPercent) hudPercent.textContent = pct + '%';
          if (hudStatusMsg && p.message) {
            hudStatusMsg.innerHTML = `<i class="fa-solid fa-gear fa-spin text-accent"></i> ${p.message}`;
          }
          // Update Stepper badges
          const stepNum = p.step || 1;
          genSteps.forEach((s, idx) => {
            if (s) {
              if (idx < stepNum - 1) s.className = 'gen-step done';
              else if (idx === stepNum - 1) s.className = 'gen-step active';
              else s.className = 'gen-step';
            }
          });
          genLines.forEach((l, idx) => {
            if (l) l.className = 'gen-step-line' + (idx < stepNum - 1 ? ' active' : '');
          });
        }
      } catch (err) {
        // Silently tolerate transient polling error
      }
    }, 250);

    try {
      const res = await fetch(`/api/session/download-excel?session_id=${currentSessionId}`, {
        method: 'POST'
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Download failed');
      }

      // Read filename from Content-Disposition header if available
      let filename = 'Valuation_Report.xlsx';
      const disposition = res.headers.get('Content-Disposition');
      if (disposition && disposition.indexOf('filename=') !== -1) {
        const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
        if (matches != null && matches[1]) {
          filename = matches[1].replace(/['"]/g, '');
        }
      } else {
        const applicantName = (currentReportData?.header?.applicant_name || 'Valuation_Report').replace(/[^a-zA-Z0-9_-]/g, '_');
        const appId = (currentReportData?.header?.application_id || '').replace(/[^a-zA-Z0-9_-]/g, '_');
        filename = `${applicantName}_${appId}.xlsx`.replace(/__+/g, '_');
      }

      const blob = await res.blob();
      const totalElapsed = ((performance.now() - startTime) / 1000).toFixed(1);

      // Trigger browser download
      const blobUrl = window.URL.createObjectURL(blob);
      const downloadLink = document.createElement('a');
      downloadLink.style.display = 'none';
      downloadLink.href = blobUrl;
      downloadLink.setAttribute('download', filename);
      document.body.appendChild(downloadLink);
      downloadLink.click();

      setTimeout(() => {
        window.URL.revokeObjectURL(blobUrl);
        downloadLink.remove();
      }, 500);

      // Clean up intervals
      clearInterval(excelGenTimerInterval);
      clearInterval(excelGenPollInterval);

      // Complete HUD state
      if (hudProgressBar) hudProgressBar.style.width = '100%';
      if (hudPercent) hudPercent.textContent = '100%';
      if (hudTitle) hudTitle.textContent = 'Report Downloaded!';
      if (hudStatusMsg) {
        hudStatusMsg.innerHTML = `<i class="fa-solid fa-circle-check text-success"></i> ${filename} ready in ${totalElapsed}s!`;
      }
      genSteps.forEach(s => { if (s) s.className = 'gen-step done'; });
      genLines.forEach(l => { if (l) l.className = 'gen-step-line active'; });

      // Buttons show success state
      buttons.forEach(b => {
        b.classList.remove('btn-generating');
        b.innerHTML = `<i class="fa-solid fa-circle-check text-success"></i> Downloaded (${totalElapsed}s)!`;
      });

      showToast(`Excel file "${filename}" compiled and downloaded in ${totalElapsed}s!`, 'success');

      // Auto-hide HUD after 3 seconds and restore buttons
      setTimeout(() => {
        if (hud) {
          hud.style.opacity = '0';
          hud.style.transform = 'translateY(12px)';
          setTimeout(() => { hud.style.display = 'none'; }, 300);
        }
        buttons.forEach((b, idx) => {
          b.disabled = false;
          b.innerHTML = originalButtonContents[idx];
        });
      }, 3000);

    } catch (e) {
      clearInterval(excelGenTimerInterval);
      clearInterval(excelGenPollInterval);

      if (hud) {
        if (hudStatusMsg) hudStatusMsg.innerHTML = `<i class="fa-solid fa-triangle-exclamation text-error"></i> Error: ${e.message}`;
        setTimeout(() => { hud.style.display = 'none'; }, 3500);
      }

      buttons.forEach((b, idx) => {
        b.disabled = false;
        b.classList.remove('btn-generating');
        b.innerHTML = originalButtonContents[idx];
      });

      showToast(`Generation Error: ${e.message}`, 'error');
      console.error(e);
    }
  }

  // 13. Event Listeners
  function initEventListeners() {
    if (btnThemeToggle) btnThemeToggle.addEventListener('click', toggleTheme);

    btnLoadSunitaCase.addEventListener('click', () => {
      const path = folderPathInput.value || "d:\\ABHAY VICKY\\Banking client project\\SUNITA DEVI BACHHAN KUMAR";
      processFolderPath(path);
    });

    btnRunFolderPath.addEventListener('click', () => {
      processFolderPath(folderPathInput.value);
    });

    fileDropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      fileDropzone.classList.add('dragover');
    });

    fileDropzone.addEventListener('dragleave', () => {
      fileDropzone.classList.remove('dragover');
    });

    fileDropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      fileDropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleFileUploads(e.dataTransfer.files);
      }
    });

    btnSelectFiles.addEventListener('click', (e) => {
      e.stopPropagation();
      multiFileInput.click();
    });

    btnSelectFolder.addEventListener('click', (e) => {
      e.stopPropagation();
      fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleFileUploads(e.target.files);
      }
    });

    multiFileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleFileUploads(e.target.files);
      }
    });

    btnGenerateExcel.addEventListener('click', generateAndDownloadExcel);
    btnDownloadReport.addEventListener('click', generateAndDownloadExcel);

    btnSettingsModal.addEventListener('click', () => settingsModal.style.display = 'flex');
    btnCloseSettings.addEventListener('click', () => settingsModal.style.display = 'none');
    btnCancelSettings.addEventListener('click', () => settingsModal.style.display = 'none');

    btnSaveSettings.addEventListener('click', async () => {
      const apiKey = modalApiKey.value.trim();
      const model = modalModelSelect.value;
      
      try {
        const res = await fetch('/api/settings', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ gemini_api_key: apiKey, model_name: model })
        });
        const result = await res.json();
        showToast(result.message, 'success');
        settingsModal.style.display = 'none';
        checkBackendHealth();
      } catch (e) {
        showToast('Failed to save settings', 'error');
      }
    });
  }

  // 14. Toast Notification Utility
  function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    let icon = 'fa-solid fa-circle-info';
    if (type === 'success') icon = 'fa-solid fa-circle-check';
    else if (type === 'error') icon = 'fa-solid fa-circle-exclamation';

    toast.innerHTML = `<i class="${icon}"></i> <span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }
});
