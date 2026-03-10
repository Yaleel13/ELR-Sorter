/* =========================================================
   ELR Sorter – app.js
   Elaria Visual Asset Organizer
   Plain vanilla JavaScript, no build step required.
   ========================================================= */

'use strict';

// ── State ────────────────────────────────────────────────
const state = {
  workbookRaw: null,         // original XLSX workbook object
  workbookFileName: '',      // original workbook filename
  registry: [],              // parsed ELR registry rows (objects)
  images: [],                // uploaded image objects
  assignments: {},           // { elrId: imageName }
  selectedElrId: null,       // currently selected ELR row
  searchQuery: '',           // current search string
};

// ELR IDs from 001–077
const VALID_ELR_IDS = Array.from({ length: 77 }, (_, i) =>
  'ELR ' + String(i + 1).padStart(3, '0')
);

// ── DOM references ───────────────────────────────────────
const $ = id => document.getElementById(id);
const workbookInput    = $('workbook-input');
const imagesInput      = $('images-input');
const autoAssignBtn    = $('auto-assign-btn');
const clearAssignBtn   = $('clear-assignments-btn');
const exportMappingBtn = $('export-mapping-btn');
const exportRenameBtn  = $('export-rename-btn');
const exportWbBtn      = $('export-workbook-btn');
const statusMsg        = $('status-message');
const searchInput      = $('search-input');
const registryTbody    = $('registry-tbody');
const galleryGrid      = $('gallery-grid');
const dropZone         = $('drop-zone');
const selectedInfo     = $('selected-info');
const selectedInfoGrid = $('selected-info-grid');
const deselectBtn      = $('deselect-btn');
const statTotal        = $('stat-total');
const statImages       = $('stat-images');
const statMapped       = $('stat-mapped');
const statUnmapped     = $('stat-unmapped');
const registryCount    = $('registry-count');
const galleryCount     = $('gallery-count');

// ── Status helpers ───────────────────────────────────────
function setStatus(msg, type = '') {
  statusMsg.textContent = msg;
  statusMsg.className = type;
}

// ── ELR Detection ───────────────────────────────────────
/**
 * Given a filename string, attempt to extract an ELR ID.
 * Supports: ELR_001, ELR-001, ELR 001, elr001, ELR001, etc.
 * Returns normalised "ELR 001" or null.
 */
function detectElrFromFilename(filename) {
  if (!filename) return null;
  const upper = filename.toUpperCase();
  // Match ELR followed by optional separator and 1–3 digits
  const match = upper.match(/ELR[\s_\-]?(\d{1,3})/);
  if (!match) return null;
  const num = parseInt(match[1], 10);
  if (num < 1 || num > 77) return null;
  return 'ELR ' + String(num).padStart(3, '0');
}

/**
 * Normalise a raw value from the workbook to a valid ELR ID or null.
 */
function normaliseElrId(raw) {
  if (raw == null || raw === '') return null;
  const str = String(raw).trim().toUpperCase();
  // Already in correct format
  if (/^ELR \d{3}$/.test(str)) {
    const num = parseInt(str.slice(4), 10);
    return (num >= 1 && num <= 77) ? str : null;
  }
  return detectElrFromFilename(str);
}

// ── Workbook Loading ─────────────────────────────────────
workbookInput.addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  loadWorkbook(file);
  // Reset so same file can be re-loaded
  workbookInput.value = '';
});

function loadWorkbook(file) {
  setStatus('Reading workbook…', 'info');
  const reader = new FileReader();
  reader.onload = evt => {
    try {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: 'array' });

      const SHEET_NAME = 'ELR Master Registry';
      if (!wb.SheetNames.includes(SHEET_NAME)) {
        setStatus(
          `Error: Sheet "${SHEET_NAME}" not found. Available sheets: ${wb.SheetNames.join(', ')}`,
          'error'
        );
        return;
      }

      state.workbookRaw = wb;
      state.workbookFileName = file.name;

      const sheet = wb.Sheets[SHEET_NAME];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      const parsed = parseRegistryRows(rows);
      if (parsed.length === 0) {
        setStatus('Error: No valid ELR rows (ELR 001–077) found in the sheet.', 'error');
        return;
      }

      state.registry = parsed;
      state.assignments = {};
      state.selectedElrId = null;

      renderRegistry();
      updateStats();
      updateButtonStates();
      setStatus(`Workbook loaded: ${file.name} — ${parsed.length} ELR rows found.`, 'success');
    } catch (err) {
      setStatus('Error reading workbook: ' + err.message, 'error');
      console.error(err);
    }
  };
  reader.onerror = () => setStatus('Error: Could not read file.', 'error');
  reader.readAsArrayBuffer(file);
}

// ── Row Parsing ──────────────────────────────────────────
const COL_MAP = {
  elrId:          ['ELR ID', 'ELR_ID', 'ELRID'],
  seq:            ['Seq', 'SEQ', 'Sequence'],
  assetCategory:  ['Asset Category'],
  collection:     ['Collection'],
  appSection:     ['App Section'],
  screenPlacement:['Screen / Placement', 'Screen/Placement', 'Screen Placement'],
  assetTitle:     ['Asset Title'],
  usageObjective: ['Usage Objective'],
  orientation:    ['Orientation'],
  recommendedSize:['Recommended Size'],
  visualDirection:['Visual Direction'],
  artNotes:       ['Prompt / Art Notes', 'Art Notes', 'Prompt/Art Notes'],
  filename:       ['Filename', 'File Name', 'FileName'],
  format:         ['Format'],
  supabaseBucket: ['Supabase Bucket'],
  storageFolder:  ['Storage Folder'],
  supabaseTable:  ['Supabase Table'],
  dbField:        ['DB Field'],
  linkedContentId:['Linked Content ID'],
  buildStatus:    ['Build Status'],
  uploadStatus:   ['Upload Status'],
  priority:       ['Priority'],
  finalUrl:       ['Final URL'],
  qaReview:       ['QA Review'],
  notes:          ['Notes'],
};

function getCell(row, keys) {
  for (const k of keys) {
    if (row[k] != null && row[k] !== '') return String(row[k]).trim();
  }
  return '';
}

function parseRegistryRows(rows) {
  const parsed = [];
  for (const row of rows) {
    const elrId = normaliseElrId(getCell(row, COL_MAP.elrId));
    if (!elrId) continue;

    parsed.push({
      elrId,
      seq:             getCell(row, COL_MAP.seq),
      assetCategory:   getCell(row, COL_MAP.assetCategory),
      collection:      getCell(row, COL_MAP.collection),
      appSection:      getCell(row, COL_MAP.appSection),
      screenPlacement: getCell(row, COL_MAP.screenPlacement),
      assetTitle:      getCell(row, COL_MAP.assetTitle),
      usageObjective:  getCell(row, COL_MAP.usageObjective),
      orientation:     getCell(row, COL_MAP.orientation),
      recommendedSize: getCell(row, COL_MAP.recommendedSize),
      visualDirection: getCell(row, COL_MAP.visualDirection),
      artNotes:        getCell(row, COL_MAP.artNotes),
      filename:        getCell(row, COL_MAP.filename),
      format:          getCell(row, COL_MAP.format),
      supabaseBucket:  getCell(row, COL_MAP.supabaseBucket),
      storageFolder:   getCell(row, COL_MAP.storageFolder),
      supabaseTable:   getCell(row, COL_MAP.supabaseTable),
      dbField:         getCell(row, COL_MAP.dbField),
      linkedContentId: getCell(row, COL_MAP.linkedContentId),
      buildStatus:     getCell(row, COL_MAP.buildStatus),
      uploadStatus:    getCell(row, COL_MAP.uploadStatus),
      priority:        getCell(row, COL_MAP.priority),
      finalUrl:        getCell(row, COL_MAP.finalUrl),
      qaReview:        getCell(row, COL_MAP.qaReview),
      notes:           getCell(row, COL_MAP.notes),
    });
  }

  // Sort by Seq ascending; fall back to ELR number
  parsed.sort((a, b) => {
    const seqA = parseInt(a.seq, 10);
    const seqB = parseInt(b.seq, 10);
    if (!isNaN(seqA) && !isNaN(seqB)) return seqA - seqB;
    if (!isNaN(seqA)) return -1;
    if (!isNaN(seqB)) return 1;
    const numA = parseInt(a.elrId.slice(4), 10);
    const numB = parseInt(b.elrId.slice(4), 10);
    return numA - numB;
  });

  return parsed;
}

// ── Image Loading ────────────────────────────────────────
imagesInput.addEventListener('change', e => {
  loadImageFiles(Array.from(e.target.files));
  imagesInput.value = '';
});

// Drag-and-drop
dropZone.addEventListener('dragenter', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', e => { if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over'); });
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const files = Array.from(e.dataTransfer.files).filter(f => isImageFile(f));
  if (files.length) loadImageFiles(files);
  else setStatus('No supported image files dropped.', 'warn');
});
dropZone.addEventListener('click', () => imagesInput.click());

function isImageFile(file) {
  return /^image\/(png|jpeg|webp|gif)$/.test(file.type) ||
    /\.(png|jpe?g|webp|gif)$/i.test(file.name);
}

function loadImageFiles(files) {
  if (!files.length) return;

  const supported = files.filter(isImageFile);
  const skipped = files.length - supported.length;
  let added = 0;
  let duplicates = 0;

  const promises = supported.map(file => new Promise(resolve => {
    // Deduplicate by name + size
    const exists = state.images.some(
      img => img.name === file.name && img.size === file.size
    );
    if (exists) { duplicates++; resolve(); return; }

    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const w = img.naturalWidth;
      const h = img.naturalHeight;
      let orient = 'Unknown';
      if (w > 0 && h > 0) {
        if (w === h) orient = 'Square';
        else if (w > h) orient = 'Landscape';
        else orient = 'Portrait';
      }
      const ext = file.name.split('.').pop().toLowerCase();
      state.images.push({
        name: file.name,
        size: file.size,
        ext,
        width: w,
        height: h,
        orientation: orient,
        previewUrl: url,
        detectedElrId: detectElrFromFilename(file.name),
      });
      added++;
      resolve();
    };
    img.onerror = () => {
      // Still add without dimensions
      const ext = file.name.split('.').pop().toLowerCase();
      state.images.push({
        name: file.name,
        size: file.size,
        ext,
        width: 0,
        height: 0,
        orientation: 'Unknown',
        previewUrl: url,
        detectedElrId: detectElrFromFilename(file.name),
      });
      added++;
      resolve();
    };
    img.src = url;
  }));

  Promise.all(promises).then(() => {
    renderGallery();
    updateStats();
    updateButtonStates();
    let msg = `${added} image(s) added.`;
    if (duplicates) msg += ` ${duplicates} duplicate(s) skipped.`;
    if (skipped) msg += ` ${skipped} unsupported file(s) ignored.`;
    setStatus(msg, added > 0 ? 'success' : 'warn');
  });
}

// ── Selection State ──────────────────────────────────────
function selectRow(elrId) {
  state.selectedElrId = elrId;
  renderRegistry();
  renderSelectedInfo();
  renderGallery(); // re-render to update button states
}

function deselectRow() {
  state.selectedElrId = null;
  renderRegistry();
  selectedInfo.hidden = true;
  renderGallery();
}

deselectBtn.addEventListener('click', deselectRow);

// ── Assignment State ─────────────────────────────────────
function assignImage(elrId, imageName) {
  // Remove this image from any other ELR row
  for (const [id, imgName] of Object.entries(state.assignments)) {
    if (imgName === imageName && id !== elrId) {
      delete state.assignments[id];
    }
  }
  state.assignments[elrId] = imageName;
  renderRegistry();
  renderGallery();
  renderSelectedInfo();
  updateStats();
  updateButtonStates();
}

function clearAssignment(elrId) {
  delete state.assignments[elrId];
  renderRegistry();
  renderGallery();
  renderSelectedInfo();
  updateStats();
  updateButtonStates();
}

clearAssignBtn.addEventListener('click', () => {
  if (!Object.keys(state.assignments).length) return;
  state.assignments = {};
  renderRegistry();
  renderGallery();
  renderSelectedInfo();
  updateStats();
  updateButtonStates();
  setStatus('All assignments cleared.', 'info');
});

// ── Auto-Assign ──────────────────────────────────────────
autoAssignBtn.addEventListener('click', autoAssign);

function autoAssign() {
  if (!state.registry.length) { setStatus('Load a workbook first.', 'error'); return; }
  if (!state.images.length)   { setStatus('Load images first.', 'error'); return; }

  let matched = 0;

  for (const img of state.images) {
    if (!img.detectedElrId) continue;
    const elrId = img.detectedElrId;
    // Check the ELR row exists in registry
    const rowExists = state.registry.some(r => r.elrId === elrId);
    if (!rowExists) continue;
    // Only assign if row not already assigned
    if (state.assignments[elrId]) continue;
    // Check this image is not already assigned to another row
    const alreadyUsed = Object.values(state.assignments).includes(img.name);
    if (alreadyUsed) continue;
    state.assignments[elrId] = img.name;
    matched++;
  }

  renderRegistry();
  renderGallery();
  renderSelectedInfo();
  updateStats();
  updateButtonStates();

  if (matched > 0) {
    setStatus(`Auto-assign complete: ${matched} image(s) matched automatically.`, 'success');
  } else {
    setStatus('Auto-assign complete: no new matches found based on filenames.', 'info');
  }
}

// ── Render Registry ──────────────────────────────────────
function renderRegistry() {
  if (!state.registry.length) {
    registryTbody.innerHTML = '<tr class="empty-row"><td colspan="6">Load a workbook to see registry rows.</td></tr>';
    registryCount.textContent = '';
    return;
  }

  const q = state.searchQuery.toLowerCase();
  const filtered = q ? state.registry.filter(row => matchesSearch(row, q)) : state.registry;
  registryCount.textContent = `(${filtered.length} of ${state.registry.length})`;

  if (!filtered.length) {
    registryTbody.innerHTML = '<tr class="empty-row"><td colspan="6">No results for current search.</td></tr>';
    return;
  }

  registryTbody.innerHTML = filtered.map(row => {
    const assigned = state.assignments[row.elrId] || '';
    const isSelected = row.elrId === state.selectedElrId;
    const statusHtml = assigned
      ? `<span class="status-badge status-mapped">Mapped</span>`
      : `<span class="status-badge status-unmapped">Unmapped</span>`;
    const assignedCell = assigned
      ? `<span class="cell-truncate" title="${escHtml(assigned)}">${escHtml(assigned)}</span>`
      : `<span style="color:var(--text-muted)">—</span>`;

    return `<tr class="${isSelected ? 'row-selected' : ''}" data-elr="${escHtml(row.elrId)}">
      <td>${escHtml(row.elrId)}</td>
      <td class="cell-truncate" title="${escHtml(row.assetTitle)}">${escHtml(row.assetTitle) || '—'}</td>
      <td class="cell-truncate" title="${escHtml(row.appSection)}">${escHtml(row.appSection) || '—'}</td>
      <td class="cell-truncate" title="${escHtml(row.storageFolder)}">${escHtml(row.storageFolder) || '—'}</td>
      <td class="cell-truncate">${assignedCell}</td>
      <td>${statusHtml}</td>
    </tr>`;
  }).join('');

  // Attach click handlers
  registryTbody.querySelectorAll('tr[data-elr]').forEach(tr => {
    tr.addEventListener('click', () => {
      const elrId = tr.dataset.elr;
      if (elrId === state.selectedElrId) {
        deselectRow();
      } else {
        selectRow(elrId);
      }
    });
  });
}

function matchesSearch(row, q) {
  const assigned = state.assignments[row.elrId] || '';
  return (
    row.elrId.toLowerCase().includes(q)          ||
    row.assetTitle.toLowerCase().includes(q)      ||
    row.collection.toLowerCase().includes(q)      ||
    row.appSection.toLowerCase().includes(q)      ||
    row.storageFolder.toLowerCase().includes(q)   ||
    row.filename.toLowerCase().includes(q)        ||
    assigned.toLowerCase().includes(q)
  );
}

// ── Render Selected Info ─────────────────────────────────
function renderSelectedInfo() {
  if (!state.selectedElrId) {
    selectedInfo.hidden = true;
    return;
  }
  const row = state.registry.find(r => r.elrId === state.selectedElrId);
  if (!row) { selectedInfo.hidden = true; return; }

  const assigned = state.assignments[row.elrId] || '';
  const fields = [
    { label: 'ELR ID',            value: row.elrId },
    { label: 'Asset Title',       value: row.assetTitle },
    { label: 'Collection',        value: row.collection },
    { label: 'App Section',       value: row.appSection },
    { label: 'Storage Folder',    value: row.storageFolder },
    { label: 'Workbook Filename', value: row.filename },
    { label: 'Orientation',       value: row.orientation },
    { label: 'Assigned Image',    value: assigned, fullWidth: true },
  ];

  selectedInfoGrid.innerHTML = fields.map(f =>
    `<div class="info-field${f.fullWidth ? ' full-width' : ''}">
      <span class="info-label">${escHtml(f.label)}</span>
      <span class="info-value${f.value ? '' : ' empty'}">${escHtml(f.value) || 'Not set'}</span>
    </div>`
  ).join('');

  selectedInfo.hidden = false;
}

// ── Render Gallery ───────────────────────────────────────
function renderGallery() {
  if (!state.images.length) {
    galleryGrid.innerHTML = '';
    galleryCount.textContent = '';
    return;
  }

  galleryCount.textContent = `(${state.images.length})`;

  galleryGrid.innerHTML = state.images.map(img => {
    const assigned = getImageAssignment(img.name);
    const isAssigned = !!assigned;
    const dims = (img.width && img.height) ? `${img.width}×${img.height}` : 'unknown';

    const detectedTag = img.detectedElrId
      ? `<span class="card-tag tag-elr">🔍 ${escHtml(img.detectedElrId)}</span>`
      : '';
    const assignedTag = isAssigned
      ? `<span class="card-tag tag-assigned">✓ ${escHtml(assigned)}</span>`
      : '';

    const canAssign = !!state.selectedElrId;

    return `<div class="image-card ${isAssigned ? 'card-assigned' : ''}" data-img="${escHtml(img.name)}">
      <img class="card-thumb" src="${escHtml(img.previewUrl)}" alt="${escHtml(img.name)}" loading="lazy" />
      <div class="card-body">
        <div class="card-filename" title="${escHtml(img.name)}">${escHtml(img.name)}</div>
        <div class="card-meta">
          <span class="card-tag tag-dims">${escHtml(dims)}</span>
          <span class="card-tag tag-orient">${escHtml(img.orientation)}</span>
          ${detectedTag}
          ${assignedTag}
        </div>
        <div class="card-actions">
          <button
            class="btn btn-sm btn-assign"
            data-action="assign"
            data-img="${escHtml(img.name)}"
            ${canAssign ? '' : 'disabled'}
            title="${canAssign ? 'Assign to selected ELR row: ' + escHtml(state.selectedElrId) : 'Select an ELR row first'}"
          >${canAssign ? 'Assign to ' + escHtml(state.selectedElrId) : 'Select ELR row first'}</button>
          <button
            class="btn btn-sm btn-clear-assign"
            data-action="clear"
            data-img="${escHtml(img.name)}"
            ${isAssigned ? '' : 'disabled'}
            title="${isAssigned ? 'Remove assignment for ' + escHtml(assigned) : 'Not assigned'}"
          >Clear assignment</button>
        </div>
      </div>
    </div>`;
  }).join('');

  // Attach handlers
  galleryGrid.querySelectorAll('button[data-action]').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const imgName = btn.dataset.img;
      if (btn.dataset.action === 'assign') {
        if (!state.selectedElrId) { setStatus('Select an ELR row first.', 'warn'); return; }
        assignImage(state.selectedElrId, imgName);
        setStatus(`Assigned "${imgName}" → ${state.selectedElrId}`, 'success');
      } else if (btn.dataset.action === 'clear') {
        const elrId = getImageAssignment(imgName);
        if (elrId) { clearAssignment(elrId); setStatus(`Cleared assignment for ${elrId}.`, 'info'); }
      }
    });
  });
}

/** Returns the ELR ID this image is currently assigned to, or null */
function getImageAssignment(imageName) {
  for (const [id, name] of Object.entries(state.assignments)) {
    if (name === imageName) return id;
  }
  return null;
}

// ── Stats ────────────────────────────────────────────────
function updateStats() {
  const total    = state.registry.length;
  const images   = state.images.length;
  const mapped   = Object.keys(state.assignments).length;
  const unmapped = total - mapped;

  statTotal.textContent   = total;
  statImages.textContent  = images;
  statMapped.textContent  = mapped;
  statUnmapped.textContent = Math.max(0, unmapped);
}

// ── Button States ────────────────────────────────────────
function updateButtonStates() {
  const hasRegistry = state.registry.length > 0;
  const hasImages   = state.images.length > 0;
  const hasAny      = Object.keys(state.assignments).length > 0;

  autoAssignBtn.disabled    = !(hasRegistry && hasImages);
  clearAssignBtn.disabled   = !hasAny;
  exportMappingBtn.disabled = !hasRegistry;
  exportRenameBtn.disabled  = !hasRegistry;
  exportWbBtn.disabled      = !(hasRegistry && state.workbookRaw);
}

// ── Search ───────────────────────────────────────────────
searchInput.addEventListener('input', () => {
  state.searchQuery = searchInput.value;
  renderRegistry();
});

// ── Exports ──────────────────────────────────────────────
exportMappingBtn.addEventListener('click', () => {
  if (!state.registry.length) { setStatus('Load a workbook first.', 'error'); return; }
  exportMappingCsv();
});

exportRenameBtn.addEventListener('click', () => {
  if (!state.registry.length) { setStatus('Load a workbook first.', 'error'); return; }
  exportRenameCsv();
});

exportWbBtn.addEventListener('click', () => {
  if (!state.workbookRaw) { setStatus('Load a workbook first.', 'error'); return; }
  exportPatchedWorkbook();
});

/**
 * Export Mapping CSV
 * Fields: elr_id, seq, asset_title, collection, app_section,
 *         storage_folder, workbook_filename, source_image_name, assigned
 */
function exportMappingCsv() {
  const headers = [
    'elr_id','seq','asset_title','collection','app_section',
    'storage_folder','workbook_filename','source_image_name','assigned'
  ];
  const rows = state.registry.map(r => {
    const imgName = state.assignments[r.elrId] || '';
    return [
      r.elrId,
      r.seq,
      r.assetTitle,
      r.collection,
      r.appSection,
      r.storageFolder,
      r.filename,
      imgName,
      imgName ? 'yes' : 'no',
    ];
  });
  downloadCsv([headers, ...rows], 'elr_mapping.csv');
  setStatus('Mapping CSV exported.', 'success');
}

/**
 * Export Rename CSV
 * Fields: elr_id, source_image_name, rename_to, target_folder, final_path
 */
function exportRenameCsv() {
  const headers = [
    'elr_id','source_image_name','rename_to','target_folder','final_path'
  ];
  const rows = state.registry.map(r => {
    const imgName = state.assignments[r.elrId] || '';
    if (!imgName) {
      return [r.elrId, '', '', r.storageFolder, ''];
    }
    const img = state.images.find(i => i.name === imgName);
    const ext = img ? img.ext : imgName.split('.').pop().toLowerCase();

    let renameTo = r.filename || '';
    if (!renameTo) {
      // Fallback: ELR_001.png style
      const num = r.elrId.replace('ELR ', 'ELR_');
      renameTo = `${num}.${ext}`;
    }

    const folder = r.storageFolder || '';
    const finalPath = folder ? `/${folder}/${renameTo}` : `/${renameTo}`;

    return [r.elrId, imgName, renameTo, folder, finalPath];
  });
  downloadCsv([headers, ...rows], 'elr_rename.csv');
  setStatus('Rename CSV exported.', 'success');
}

/**
 * Export Patched Workbook
 * Clones workbook, updates rows that have assignments.
 */
function exportPatchedWorkbook() {
  try {
    const wb = state.workbookRaw;
    const SHEET_NAME = 'ELR Master Registry';
    const ws = wb.Sheets[SHEET_NAME];

    // Get the range
    const range = XLSX.utils.decode_range(ws['!ref']);
    // Find header row (row 0)
    const headers = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
      if (cell && cell.v != null) headers[String(cell.v).trim()] = c;
    }

    // Column indices we'll update
    const colFilename     = findColumnIndex(headers, COL_MAP.filename);
    const colUploadStatus = findColumnIndex(headers, COL_MAP.uploadStatus);
    const colFinalUrl     = findColumnIndex(headers, COL_MAP.finalUrl);
    const colNotes        = findColumnIndex(headers, COL_MAP.notes);

    // Clone workbook data (SheetJS works on the same object; we rebuild)
    const cloneWb = XLSX.utils.book_new();
    // Copy all sheets
    for (const sheetName of wb.SheetNames) {
      const cloneWs = Object.assign({}, wb.Sheets[sheetName]);
      XLSX.utils.book_append_sheet(cloneWb, cloneWs, sheetName);
    }
    const targetWs = cloneWb.Sheets[SHEET_NAME];

    // Apply patches
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const elrCell = targetWs[XLSX.utils.encode_cell({ r, c: findColumnIndex(headers, COL_MAP.elrId) })];
      if (!elrCell) continue;
      const elrId = normaliseElrId(elrCell.v);
      if (!elrId) continue;
      const imgName = state.assignments[elrId];
      if (!imgName) continue;

      const registryRow = state.registry.find(row => row.elrId === elrId);
      if (!registryRow) continue;

      const img = state.images.find(i => i.name === imgName);
      const ext = img ? img.ext : imgName.split('.').pop().toLowerCase();
      let renameTo = registryRow.filename || '';
      if (!renameTo) {
        const num = elrId.replace('ELR ', 'ELR_');
        renameTo = `${num}.${ext}`;
      }
      const folder = registryRow.storageFolder || '';
      const finalPath = folder ? `/${folder}/${renameTo}` : `/${renameTo}`;
      const noteAppend = `mapped by ELR Sorter from source image ${imgName}`;

      if (colFilename !== -1) setCellValue(targetWs, r, colFilename, renameTo);
      if (colUploadStatus !== -1) setCellValue(targetWs, r, colUploadStatus, 'Mapped');
      if (colFinalUrl !== -1) setCellValue(targetWs, r, colFinalUrl, finalPath);
      if (colNotes !== -1) {
        const existing = getCellValue(targetWs, r, colNotes);
        const newNote = existing ? `${existing}; ${noteAppend}` : noteAppend;
        setCellValue(targetWs, r, colNotes, newNote);
      }
    }

    // Determine output filename
    const baseName = state.workbookFileName.replace(/\.[^.]+$/, '');
    const outName = `${baseName}_patched.xlsx`;
    XLSX.writeFile(cloneWb, outName);
    setStatus(`Patched workbook exported: ${outName}`, 'success');
  } catch (err) {
    setStatus('Error exporting workbook: ' + err.message, 'error');
    console.error(err);
  }
}

function findColumnIndex(headers, keys) {
  for (const k of keys) {
    if (headers[k] !== undefined) return headers[k];
  }
  return -1;
}

function setCellValue(ws, r, c, value) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (!ws[addr]) ws[addr] = {};
  ws[addr].v = value;
  ws[addr].t = 's';
}

function getCellValue(ws, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (!ws[addr]) return '';
  return ws[addr].v != null ? String(ws[addr].v) : '';
}

// ── CSV Helpers ──────────────────────────────────────────
function downloadCsv(rows, filename) {
  const content = rows.map(row =>
    row.map(cell => {
      const str = String(cell == null ? '' : cell);
      // Escape cells that contain commas, quotes, or newlines
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',')
  ).join('\r\n');

  const blob = new Blob(['\uFEFF' + content], { type: 'text/csv;charset=utf-8;' });
  triggerDownload(blob, filename);
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ── Utility ──────────────────────────────────────────────
function escHtml(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ── Init ─────────────────────────────────────────────────
updateStats();
updateButtonStates();
