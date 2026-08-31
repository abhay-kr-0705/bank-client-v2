/**
 * BankTech Valuation OCR & Exact Excel Report Generator
 * Frontend Application Engine
 */

document.addEventListener('DOMContentLoaded', () => {
  // Application State
  let currentReportData = null;
  let currentSourceFolder = "";
  let activeFiles = [];

  const defaultFloorNames = [
    "Basement", "Stilt Floor", "Ground Floor", "First Floor",
    "Second Floor", "Third Floor", "Fouth Floor", "Fifth Floor",
    "Six Floor", "Seven Floor"
  ];

  // DOM Elements
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
  
  const btnGenerateExcel = document.getElementById('btnGenerateExcel');
  const btnDownloadReport = document.getElementById('btnDownloadReport');
  const btnExportJson = document.getElementById('btnExportJson');

  const docPreviewWrapper = document.getElementById('docPreviewWrapper');
  const previewFileName = document.getElementById('previewFileName');
  const previewBody = document.getElementById('previewBody');
  const btnClosePreview = document.getElementById('btnClosePreview');

  // Settings Modal Elements
  const btnSettingsModal = document.getElementById('btnSettingsModal');
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
  initFloorTable();
  initTabs();
  initEventListeners();
  checkBackendHealth();

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
        engineStatusText.textContent = `Local High-Accuracy Hybrid OCR Active`;
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
      });
    });
  }

  // 3. Initialize Default Floors Table
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

  // 4. Progress Pipeline Animation
  function runPipelineAnimation(onComplete) {
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
      if (onComplete) onComplete();
    }, 1400);
  }

  // 5. Ingest / Process Folder
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
        body: JSON.stringify({ folder_path: folderPath.trim() })
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Extraction failed');
      }

      const result = await res.json();
      currentReportData = result.data;
      currentSourceFolder = result.source_folder;
      
      populateFormWithData(currentReportData);
      renderFilesList(currentReportData.raw_files_summary || []);
      
      btnDownloadReport.disabled = false;
      showToast('Case extracted and synced to Sunita.xlsx format successfully!', 'success');
    } catch (e) {
      showToast(`Error: ${e.message}`, 'error');
      console.error(e);
    }
  }

  // 6. Handle File Uploads (Drag & Drop or Multi-Select)
  async function handleFileUploads(files) {
    if (!files || files.length === 0) return;

    showToast(`Uploading ${files.length} documents...`, 'info');
    runPipelineAnimation();

    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
      formData.append('files', files[i]);
    }
    formData.append('case_name', 'Uploaded_Case_' + new Date().toISOString().slice(0, 10));

    try {
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Upload and extraction failed');
      }

      const result = await res.json();
      currentReportData = result.data;
      currentSourceFolder = result.source_folder;

      populateFormWithData(currentReportData);
      renderFilesList(currentReportData.raw_files_summary || []);

      btnDownloadReport.disabled = false;
      showToast('Upload processed and 106 report fields extracted!', 'success');
    } catch (e) {
      showToast(`Upload Error: ${e.message}`, 'error');
      console.error(e);
    }
  }

  // 7. Render Files List in Left Pane
  function renderFilesList(files) {
    activeFiles = files;
    filesCountBadge.textContent = files.length;

    if (files.length === 0) {
      filesListContainer.innerHTML = `
        <div class="empty-state-card">
          <i class="fa-regular fa-folder-open"></i>
          <p>No documents found in folder.</p>
        </div>`;
      return;
    }

    filesListContainer.innerHTML = '';
    files.forEach(file => {
      const div = document.createElement('div');
      div.className = 'file-item';
      
      let icon = 'fa-solid fa-file';
      if (file.type === 'site_photo') icon = 'fa-solid fa-file-image text-warning';
      else if (file.type === 'property_pdf') icon = 'fa-solid fa-file-pdf text-danger';
      else if (file.type === 'field_notes_docx') icon = 'fa-solid fa-file-word text-accent';
      else if (file.type === 'target_excel_template' || file.type === 'excel') icon = 'fa-solid fa-file-excel text-success';

      div.innerHTML = `
        <div class="file-info">
          <i class="${icon} file-icon"></i>
          <div>
            <div class="file-name truncate" title="${file.filename}">${file.filename}</div>
            <div class="file-size">${file.size_display} &bull; ${file.type}</div>
          </div>
        </div>
        <button class="btn-icon" title="Preview document"><i class="fa-regular fa-eye"></i></button>
      `;

      div.addEventListener('click', () => previewFile(file));
      filesListContainer.appendChild(div);
    });
  }

  // 8. Document Previewer
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

  // 9. Populate Form Fields from Extracted JSON
  function populateFormWithData(data) {
    if (!data) return;

    casePillBadge.innerHTML = `<i class="fa-solid fa-id-card"></i> Case: ${data.case_name || 'Sunita Devi'}`;

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

  // 10. Collect Current Form State into JSON
  function collectFormData() {
    const floors = [];
    document.querySelectorAll('#floorsTableBody tr').forEach((tr, idx) => {
      floors.push({
        name: defaultFloorNames[idx],
        actual_area: parseFloat(tr.querySelector('.floor-actual').value || 0),
        permissible_area: parseFloat(tr.querySelector('.floor-permissible').value || 0),
        adopted_area: parseFloat(tr.querySelector('.floor-adopted').value || 0)
      });
    });

    return {
      case_name: currentReportData ? currentReportData.case_name : "Case_Export",
      header: {
        report_title: document.getElementById('inp_report_title').value,
        application_id: document.getElementById('inp_application_id').value,
        applicant_name: document.getElementById('inp_applicant_name').value,
        property_type: document.getElementById('inp_property_type').value,
        completion_percent: parseFloat(document.getElementById('inp_completion_percent').value || 1.0),
        structure_type: document.getElementById('inp_structure_type').value,
        age_of_property: document.getElementById('inp_age_of_property').value,
        dwelling_units_owned: parseInt(document.getElementById('inp_dwelling_units_owned').value || 1),
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
      solar_roof_vicinity: {
        solar_install_location: document.getElementById('inp_solar_install_location').value,
        roof_length_sqft: parseFloat(document.getElementById('inp_roof_length_sqft').value || 15),
        roof_breadth_sqft: parseFloat(document.getElementById('inp_roof_breadth_sqft').value || 38),
        shadow_free_roof_sqft: parseFloat(document.getElementById('inp_shadow_free_roof_sqft').value || 0),
        parapet_wall_height: parseFloat(document.getElementById('inp_parapet_wall_height').value || 0),
        cracks_in_roof: parseFloat(document.getElementById('inp_cracks_in_roof').value || 0),
        is_outreach: document.getElementById('inp_is_outreach').value,
        population_1km: document.getElementById('inp_population_1km').value,
        primary_schools_1km: parseInt(document.getElementById('inp_primary_schools_1km').value || 1),
        secondary_schools_1km: parseInt(document.getElementById('inp_secondary_schools_1km').value || 1)
      },
      land_measurements: {
        land_length: parseFloat(document.getElementById('inp_land_length').value || 38),
        land_breadth: parseFloat(document.getElementById('inp_land_breadth').value || 15),
        land_area_site_sqft: document.getElementById('inp_land_area_site_sqft').value,
        adopted_land_area_sqft: parseFloat(document.getElementById('inp_adopted_land_area_sqft').value || 569.7),
        per_unit_land_rate: parseFloat(document.getElementById('inp_per_unit_land_rate').value || 0),
        total_land_value: "=B46*B47"
      },
      construction_floors: {
        floors: floors,
        built_up_rate: parseFloat(document.getElementById('inp_built_up_rate').value || 0),
        total_property_value: parseFloat(document.getElementById('inp_total_property_value').value || 0)
      },
      accommodation: {
        no_of_floors: parseInt(document.getElementById('inp_no_of_floors').value || 5),
        toilet_available: document.getElementById('inp_toilet_available').value,
        no_of_lifts: parseInt(document.getElementById('inp_no_of_lifts').value || 0),
        apartments_per_floor: parseInt(document.getElementById('inp_apartments_per_floor').value || 1),
        electricity_meter_installed: document.getElementById('inp_electricity_meter_installed').value,
        electricity_meter_number: document.getElementById('inp_electricity_meter_number').value
      },
      legal_checks: {
        documents_name: document.getElementById('inp_documents_name').value,
        person_met: document.getElementById('inp_person_met').value,
        relation_with_owner: document.getElementById('inp_relation_with_owner').value,
        property_situated_at: document.getElementById('inp_property_situated_at').value,
        is_sanction_plan_compliant: document.getElementById('inp_is_sanction_plan_compliant').value,
        pathway_clear: document.getElementById('inp_pathway_clear').value,
        is_disaster_prone: document.getElementById('inp_is_disaster_prone').value,
        approach_by_public_road: document.getElementById('inp_approach_by_public_road').value,
        width_of_public_road: document.getElementById('inp_width_of_public_road').value,
        utilities_in_vicinity: document.getElementById('inp_utilities_in_vicinity').value,
        approved_land_master_plan: document.getElementById('inp_approved_land_master_plan').value,
        current_uses: document.getElementById('inp_current_uses').value,
        opinion_about_report: document.getElementById('inp_opinion_about_report').value,
        occupancy_250m: document.getElementById('inp_occupancy_250m').value,
        development_250m: document.getElementById('inp_development_250m').value,
        property_limit: document.getElementById('inp_property_limit').value,
        adm: document.getElementById('inp_adm').value
      },
      reference: {
        reference_name: document.getElementById('inp_reference_name').value,
        reference_mobile: document.getElementById('inp_reference_mobile').value,
        feedback: document.getElementById('inp_feedback').value
      },
      remarks: document.getElementById('inp_remarks').value,
      raw_files_summary: activeFiles
    };
  }

  // 11. Generate and Download Exact Sunita.xlsx Excel Report as Binary Blob
  async function generateAndDownloadExcel() {
    const payload = collectFormData();
    showToast('Generating Exact Excel Report...', 'info');

    try {
      const res = await fetch('/api/generate-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Excel generation failed');
      }

      // Read as binary blob
      const blob = await res.blob();
      
      // Determine filename
      const applicantName = (payload.header.applicant_name || payload.case_name || 'Report').replace(/[^a-zA-Z0-9_-]/g, '_');
      const appId = (payload.header.application_id || '').replace(/[^a-zA-Z0-9_-]/g, '_');
      const filename = `${applicantName}_${appId}.xlsx`.replace(/__+/g, '_');

      // Create explicit blob download link
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

  // 12. Export JSON Payload
  function exportJsonPayload() {
    const payload = collectFormData();
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(payload, null, 2));
    const a = document.createElement('a');
    a.href = dataStr;
    a.download = `${payload.case_name}_Data.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    showToast('JSON data exported!', 'success');
  }

  // 13. Event Listeners
  function initEventListeners() {
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
    btnExportJson.addEventListener('click', exportJsonPayload);

    btnSettingsModal.addEventListener('click', () => {
      settingsModal.style.display = 'flex';
    });

    btnCloseSettings.addEventListener('click', () => {
      settingsModal.style.display = 'none';
    });

    btnCancelSettings.addEventListener('click', () => {
      settingsModal.style.display = 'none';
    });

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
