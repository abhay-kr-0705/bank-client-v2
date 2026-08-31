/**
 * BankTech Valuation OCR & Dynamic Excel Spreadsheet Engine
 * Frontend Interactive Controller
 */

document.addEventListener('DOMContentLoaded', () => {
  // Session State
  let currentSessionId = "default_session";
  let currentReportData = null;
  let currentGridData = null;
  let activeFiles = [];
  let selectedCellCoord = "A1";

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
  const btnExportJson = document.getElementById('btnExportJson');
  
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
  const casePillBadge = document.getElementById('casePillBadge');

  // Spreadsheet Grid Elements
  const spreadsheetViewport = document.getElementById('spreadsheetViewport');
  const activeCellCoord = document.getElementById('activeCellCoord');
  const formulaBarInput = document.getElementById('formulaBarInput');

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

  // Initialize
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
      if (data.has_api_key) {
        engineStatusText.textContent = `Multimodal Gemini (${data.model_name}) Active`;
        engineStatusBadge.style.borderColor = 'rgba(16, 185, 129, 0.4)';
        engineStatusBadge.style.color = '#6ee7b7';
      } else {
        engineStatusText.textContent = `Local Dynamic Hybrid OCR Active`;
      }
    } catch (e) {
      console.warn("Backend not connected yet:", e);
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
  async function loadInitialGridState() {
    try {
      const res = await fetch(`/api/session/live-grid?session_id=${currentSessionId}`);
      const data = await res.json();
      if (data.success && data.grid) {
        currentGridData = data.grid;
        renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
      }
    } catch (e) {
      console.error("Initial grid load error:", e);
    }
  }

  // 4. Interactive Spreadsheet Grid Renderer
  function renderSpreadsheetGrid(gridData, container, isEditable = true) {
    if (!gridData || !gridData.rows || gridData.rows.length === 0) {
      container.innerHTML = `<div class="p-4 text-center text-muted">No grid data available.</div>`;
      return;
    }

    const table = document.createElement('table');
    table.className = 'excel-grid-table';

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

        let cellClass = 'grid-data-cell';
        if (cellObj.is_header) cellClass += ' grid-header-cell';
        if (cellObj.is_formula) cellClass += ' grid-formula-cell';
        if (cellObj.coord === selectedCellCoord) cellClass += ' cell-selected';
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
  }

  // Grid Modification Actions: Add Row, Add Col, Clear Data
  const btnAddGridRow = document.getElementById('btnAddGridRow');
  const btnAddGridCol = document.getElementById('btnAddGridCol');
  const btnClearGridData = document.getElementById('btnClearGridData');
  const btnSheet1 = document.getElementById('btnSheet1');
  const btnSheet2 = document.getElementById('btnSheet2');

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

  if (btnSheet1 && btnSheet2) {
    btnSheet1.addEventListener('click', () => {
      btnSheet1.classList.add('active');
      btnSheet2.classList.remove('active');
      if (currentGridData) renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
    });

    btnSheet2.addEventListener('click', () => {
      btnSheet2.classList.add('active');
      btnSheet1.classList.remove('active');
      showToast('Sheet2 displays predefined bank validation dropdown lists.', 'info');
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

  function selectCell(coord, val, tdElem, container) {
    selectedCellCoord = coord;
    activeCellCoord.textContent = coord;
    formulaBarInput.value = val;

    container.querySelectorAll('.excel-grid-table td').forEach(td => td.classList.remove('cell-selected'));
    if (tdElem) tdElem.classList.add('cell-selected');
  }

  function makeCellEditable(td, coord) {
    const currentVal = td.textContent;
    td.innerHTML = `<input type="text" class="form-control text-xs p-1" value="${currentVal}">`;
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

  // 6. Incremental File & Multi-format Ingestion
  async function handleFileUploads(files) {
    if (!files || files.length === 0) return;

    showToast(`Ingesting ${files.length} document(s) & extracting data...`, 'info');
    runPipelineAnimation();

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

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Upload failed');
      }

      const result = await res.json();
      currentReportData = result.report_data;
      currentGridData = result.grid;
      
      populateFormWithData(currentReportData);
      renderFilesList(result.files || []);
      renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);

      btnDownloadReport.disabled = false;
      showToast('Documents ingested & mapped to template cells!', 'success');
    } catch (e) {
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
    runPipelineAnimation();

    try {
      const res = await fetch('/api/process-folder', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder_path: folderPath.trim(), session_id: currentSessionId })
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Extraction failed');
      }

      const result = await res.json();
      currentReportData = result.data;
      currentGridData = result.grid;
      
      populateFormWithData(currentReportData);
      renderFilesList(currentReportData.raw_files_summary || []);
      renderSpreadsheetGrid(currentGridData, spreadsheetViewport, true);
      
      btnDownloadReport.disabled = false;
      showToast('Case extracted and live grid populated!', 'success');
    } catch (e) {
      showToast(`Error: ${e.message}`, 'error');
      console.error(e);
    }
  }

  // 7. Pipeline Tracker Animation
  function runPipelineAnimation() {
    pipelineTracker.style.display = 'block';
    pipelineProgressBar.style.width = '15%';
    
    const s1 = document.getElementById('step1');
    const s2 = document.getElementById('step2');
    const s3 = document.getElementById('step3');
    const s4 = document.getElementById('step4');

    s1.className = 'step active';
    s2.className = 'step';
    s3.className = 'step';
    s4.className = 'step';

    setTimeout(() => {
      pipelineProgressBar.style.width = '45%';
      s1.className = 'step completed';
      s2.className = 'step active';
    }, 400);

    setTimeout(() => {
      pipelineProgressBar.style.width = '75%';
      s2.className = 'step completed';
      s3.className = 'step active';
    }, 900);

    setTimeout(() => {
      pipelineProgressBar.style.width = '100%';
      s3.className = 'step completed';
      s4.className = 'step completed';
    }, 1400);
  }

  // 8. Render Files List in Left Pane
  function renderFilesList(files) {
    activeFiles = files;
    filesCountBadge.textContent = files.length;

    if (files.length === 0) {
      filesListContainer.innerHTML = `
        <div class="empty-state-card">
          <i class="fa-regular fa-folder-open"></i>
          <p>No documents found in session.</p>
        </div>`;
      return;
    }

    filesListContainer.innerHTML = '';
    files.forEach(file => {
      const div = document.createElement('div');
      div.className = 'file-item';
      
      let icon = 'fa-solid fa-file';
      const ext = (file.type || '').replace('.', '');
      if (['jpg', 'jpeg', 'png', 'webp', 'site_photo'].includes(ext)) icon = 'fa-solid fa-file-image text-warning';
      else if (['pdf', 'property_pdf'].includes(ext)) icon = 'fa-solid fa-file-pdf text-danger';
      else if (['docx', 'doc', 'field_notes_docx'].includes(ext)) icon = 'fa-solid fa-file-word text-accent';
      else if (['xlsx', 'csv', 'excel'].includes(ext)) icon = 'fa-solid fa-file-excel text-success';
      else if (['zip'].includes(ext)) icon = 'fa-solid fa-file-zipper text-warning';

      div.innerHTML = `
        <div class="file-info">
          <i class="${icon} file-icon"></i>
          <div>
            <div class="file-name truncate" title="${file.filename}">${file.filename}</div>
            <div class="file-size">${file.size_display || ''} &bull; ${file.type || 'document'}</div>
          </div>
        </div>
        <button class="btn-icon" title="Preview document"><i class="fa-regular fa-eye"></i></button>
      `;

      div.addEventListener('click', () => previewFile(file));
      filesListContainer.appendChild(div);
    });
  }

  // 9. Document Previewer
  function previewFile(file) {
    previewFileName.textContent = file.filename;
    docPreviewWrapper.style.display = 'block';
    previewBody.innerHTML = '';

    const ext = file.filename.split('.').pop().toLowerCase();
    const fileUrl = `/api/view-file?filepath=${encodeURIComponent(file.path)}`;

    if (['jpg', 'jpeg', 'png', 'webp'].includes(ext)) {
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
    document.getElementById('inp_property_type').value = h.property_type || "Row House";
    document.getElementById('inp_completion_percent').value = h.completion_percent !== undefined ? h.completion_percent : 1.0;
    document.getElementById('inp_structure_type').value = h.structure_type || "RCC";
    document.getElementById('inp_age_of_property').value = h.age_of_property || "08 Years";
    document.getElementById('inp_dwelling_units_owned').value = h.dwelling_units_owned || 1;
    document.getElementById('inp_geo_tag').value = h.geo_tag || "";

    // Address
    const a = data.address || {};
    document.getElementById('inp_street_name').value = a.street_name || "";
    document.getElementById('inp_nearest_landmark').value = a.nearest_landmark || "";
    document.getElementById('inp_village_name').value = a.village_name || "";
    document.getElementById('inp_city').value = a.city || "";
    document.getElementById('inp_plot_house_khasra').value = a.plot_house_khasra || "";
    document.getElementById('inp_floor_number').value = a.floor_number || "Entire Property";
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
    document.getElementById('inp_boundary_matching').value = b.boundary_matching ? b.boundary_matching.trim() : "Yes";
    document.getElementById('inp_mismatch_remarks').value = b.mismatch_remarks || "NA";
    document.getElementById('inp_occupancy_status').value = b.occupancy_status ? b.occupancy_status.trim() : "Seller";

    // Land
    const l = data.land_measurements || {};
    document.getElementById('inp_land_length').value = l.land_length || 38;
    document.getElementById('inp_land_breadth').value = l.land_breadth || 15;
    document.getElementById('inp_land_area_site_sqft').value = l.land_area_site_sqft || "569.7 Sqft";
    document.getElementById('inp_adopted_land_area_sqft').value = l.adopted_land_area_sqft || 569.7;
    document.getElementById('inp_per_unit_land_rate').value = l.per_unit_land_rate || 0;

    // Solar & Roof
    const s = data.solar_roof_vicinity || {};
    document.getElementById('inp_solar_install_location').value = s.solar_install_location ? s.solar_install_location.trim() : "Ground";
    document.getElementById('inp_roof_length_sqft').value = s.roof_length_sqft || 15;
    document.getElementById('inp_roof_breadth_sqft').value = s.roof_breadth_sqft || 38;
    document.getElementById('inp_shadow_free_roof_sqft').value = s.shadow_free_roof_sqft || 0;
    document.getElementById('inp_parapet_wall_height').value = s.parapet_wall_height || 0;
    document.getElementById('inp_cracks_in_roof').value = s.cracks_in_roof || 0;
    document.getElementById('inp_is_outreach').value = s.is_outreach || "No";
    document.getElementById('inp_population_1km').value = s.population_1km || "Above 5000";
    document.getElementById('inp_primary_schools_1km').value = s.primary_schools_1km || 1;
    document.getElementById('inp_secondary_schools_1km').value = s.secondary_schools_1km || 1;

    // Floors Table
    const f = data.construction_floors || {};
    initFloorTable(f.floors);
    document.getElementById('inp_built_up_rate').value = f.built_up_rate || 0;
    document.getElementById('inp_total_property_value').value = f.total_property_value || 0;
    recalculateFloorTotals();

    // Accommodation
    const acc = data.accommodation || {};
    document.getElementById('inp_no_of_floors').value = acc.no_of_floors || 5;
    document.getElementById('inp_toilet_available').value = acc.toilet_available ? acc.toilet_available.trim() : "Yes";
    document.getElementById('inp_no_of_lifts').value = acc.no_of_lifts || 0;
    document.getElementById('inp_apartments_per_floor').value = acc.apartments_per_floor || 1;
    document.getElementById('inp_electricity_meter_installed').value = acc.electricity_meter_installed ? acc.electricity_meter_installed.trim() : "Yes";
    document.getElementById('inp_electricity_meter_number').value = acc.electricity_meter_number || "NA";

    // Legal Checks
    const leg = data.legal_checks || {};
    document.getElementById('inp_documents_name').value = leg.documents_name || "Other";
    document.getElementById('inp_person_met').value = leg.person_met || "Mr. Gauarv";
    document.getElementById('inp_relation_with_owner').value = leg.relation_with_owner || "Applicant's Son";
    document.getElementById('inp_property_situated_at').value = leg.property_situated_at ? leg.property_situated_at.trim() : "MC";
    document.getElementById('inp_is_sanction_plan_compliant').value = leg.is_sanction_plan_compliant || "No";
    document.getElementById('inp_pathway_clear').value = leg.pathway_clear || "Yes";
    document.getElementById('inp_is_disaster_prone').value = leg.is_disaster_prone || "No";
    document.getElementById('inp_approach_by_public_road').value = leg.approach_by_public_road ? leg.approach_by_public_road.trim() : "Yes";
    document.getElementById('inp_width_of_public_road').value = leg.width_of_public_road || "23 Ft Wide";
    document.getElementById('inp_utilities_in_vicinity').value = leg.utilities_in_vicinity || "Yes";
    document.getElementById('inp_approved_land_master_plan').value = leg.approved_land_master_plan || "Residential";
    document.getElementById('inp_current_uses').value = leg.current_uses || "Residential";
    document.getElementById('inp_opinion_about_report').value = leg.opinion_about_report ? leg.opinion_about_report.trim() : "Negative";
    document.getElementById('inp_occupancy_250m').value = leg.occupancy_250m || "80%-90%";
    document.getElementById('inp_development_250m').value = leg.development_250m || "80%-90%";
    document.getElementById('inp_property_limit').value = leg.property_limit || "Within MC Limit";
    document.getElementById('inp_adm').value = leg.adm || "Average";

    // Reference
    const r = data.reference || {};
    document.getElementById('inp_reference_name').value = r.reference_name || "Local Enquiry";
    document.getElementById('inp_reference_mobile').value = r.reference_mobile || "9540637533";
    document.getElementById('inp_feedback').value = r.feedback || "1 L to 1.10 L per Sqyds";

    // Remarks
    document.getElementById('inp_remarks').value = data.remarks || "";
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

  // 12. Final Excel Generation & Download
  async function generateAndDownloadExcel() {
    showToast('Compiling Excel Report with Formula Engine...', 'info');

    try {
      const res = await fetch(`/api/session/download-excel?session_id=${currentSessionId}`, {
        method: 'POST'
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Download failed');
      }

      const blob = await res.blob();
      const applicantName = (currentReportData?.header?.applicant_name || 'Valuation_Report').replace(/[^a-zA-Z0-9_-]/g, '_');
      const appId = (currentReportData?.header?.application_id || '').replace(/[^a-zA-Z0-9_-]/g, '_');
      const filename = `${applicantName}_${appId}.xlsx`.replace(/__+/g, '_');

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

      showToast(`Excel file ${filename} downloaded successfully!`, 'success');
    } catch (e) {
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
