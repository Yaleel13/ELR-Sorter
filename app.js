/* =========================================================
   ELR Sorter – app.js
   Elaria Visual Asset Organizer
   Plain vanilla JavaScript, no build step required.
   ========================================================= */

'use strict';

// ── State ────────────────────────────────────────────────
const state = {
  workbookRaw: null,           // original XLSX workbook object
  workbookFileName: '',        // original workbook filename
  registry: [],                // parsed ELR registry rows (objects)
  images: [],                  // uploaded image objects  { name, size, ext, width, height, orientation, previewUrl, detectedElrId, file }
  assignments: {},             // { elrId: imageName }
  assignmentMeta: {},          // { elrId: { source, score, confidence, reason } }
  selectedElrId: null,         // currently selected ELR row
  searchQuery: '',             // current search string
  transformersApiPromise: null,// lazy-loaded Transformers.js API
  visualModelPromise: null,    // lazy-loaded CLIP pipeline promise
  isAutoAssigning: false,      // guard to avoid concurrent auto-assign runs
  isExporting: false,          // guard to avoid concurrent export runs
  visualThreshold: 0.24,       // confidence floor for visual matching
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
const exportAssetsZipBtn  = $('export-assets-zip-btn');
const exportPackageZipBtn = $('export-package-zip-btn');
const statusMsg        = $('status-message');
const searchInput      = $('search-input');
const visualThresholdInput = $('visual-threshold-input');
const visualThresholdValue = $('visual-threshold-value');
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

function readVisualThreshold() {
  const raw = parseFloat(visualThresholdInput.value);
  if (Number.isNaN(raw)) return state.visualThreshold;
  return Math.max(0.15, Math.min(0.55, raw));
}

function updateVisualThresholdLabel() {
  visualThresholdValue.textContent = state.visualThreshold.toFixed(2);
}

// ── Safe Filename / Folder Helpers ───────────────────────
function sanitizeFilename(name) {
  if (!name) return '';
  return name
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^\.+/, '_');
}

function sanitizeFolder(folder) {
  if (!folder) return 'unfiled';
  return folder
    .replace(/\\/g, '/')
    .replace(/[<>:"|?*\x00-\x1f]/g, '_')
    .replace(/\/+/g, '/')
    .replace(/^\/+|\/+$/g, '')
    .trim() || 'unfiled';
}

function getTargetFilename(row, img) {
  const ext = img ? img.ext : 'png';
  let name = row.filename || '';
  if (name) {
    if (!/\.\w+$/.test(name)) {
      name = name + '.' + ext;
    }
  } else {
    const num = row.elrId.replace('ELR ', 'ELR_');
    name = num + '.' + ext;
  }
  return sanitizeFilename(name);
}

function getTargetFolder(row) {
  return sanitizeFolder(row.storageFolder);
}

function getTargetPath(row, img) {
  return getTargetFolder(row) + '/' + getTargetFilename(row, img);
}

// ── Assignment Source Tracking ────────────────────────────
function setAssignmentMeta(elrId, source, score, reason) {
  var confidence = 'unknown';
  if (source === 'manual') {
    confidence = 'approved manually';
  } else if (source === 'filename-detected') {
    confidence = 'filename match';
  } else if (source === 'visual-analysis' && typeof score === 'number') {
    if (score >= 0.4) confidence = 'high';
    else if (score >= 0.3) confidence = 'medium';
    else confidence = 'low';
  }
  state.assignmentMeta[elrId] = {
    source: source,
    score: typeof score === 'number' ? score : null,
    confidence: confidence,
    reason: reason || source,
  };
}

function getAssignmentMeta(elrId) {
  return state.assignmentMeta[elrId] || {
    source: 'unknown',
    score: null,
    confidence: 'unknown',
    reason: 'unknown',
  };
}

function getAssignmentSourceLabel(meta) {
  switch (meta.source) {
    case 'filename-detected': return 'Mapped by ELR Sorter filename detection';
    case 'visual-analysis':   return 'Matched by ELR Sorter visual analysis';
    case 'manual':            return 'Matched manually after visual analysis';
    default:                  return 'Mapped by ELR Sorter';
  }
}

// ── ELR Detection ───────────────────────────────────────
function detectElrFromFilename(filename) {
  if (!filename) return null;
  var upper = filename.toUpperCase();
  var match = upper.match(/ELR[\s_\-]?(\d{1,3})/);
  if (!match) return null;
  var num = parseInt(match[1], 10);
  if (num < 1 || num > 77) return null;
  return 'ELR ' + String(num).padStart(3, '0');
}

function normaliseElrId(raw) {
  if (raw == null || raw === '') return null;
  var str = String(raw).trim().toUpperCase();
  if (/^ELR \d{3}$/.test(str)) {
    var num = parseInt(str.slice(4), 10);
    return (num >= 1 && num <= 77) ? str : null;
  }
  return detectElrFromFilename(str);
}

// ── Workbook Loading ─────────────────────────────────────
workbookInput.addEventListener('change', function(e) {
  var file = e.target.files[0];
  if (!file) return;
  loadWorkbook(file);
  workbookInput.value = '';
});

function loadWorkbook(file) {
  setStatus('Reading workbook…', 'info');
  var reader = new FileReader();
  reader.onload = function(evt) {
    try {
      var data = new Uint8Array(evt.target.result);
      var wb = XLSX.read(data, { type: 'array' });

      var SHEET_NAME = 'ELR Master Registry';
      if (!wb.SheetNames.includes(SHEET_NAME)) {
        setStatus(
          'Error: Sheet "' + SHEET_NAME + '" not found. Available sheets: ' + wb.SheetNames.join(', '),
          'error'
        );
        return;
      }

      state.workbookRaw = wb;
      state.workbookFileName = file.name;

      var sheet = wb.Sheets[SHEET_NAME];
      var rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

      var parsed = parseRegistryRows(rows);
      if (parsed.length === 0) {
        setStatus('Error: No valid ELR rows (ELR 001–077) found in the sheet.', 'error');
        return;
      }

      state.registry = parsed;
      state.assignments = {};
      state.assignmentMeta = {};
      state.selectedElrId = null;

      renderRegistry();
      updateStats();
      updateButtonStates();
      setStatus('Workbook loaded: ' + file.name + ' — ' + parsed.length + ' ELR rows found.', 'success');
    } catch (err) {
      setStatus('Error reading workbook: ' + err.message, 'error');
      console.error(err);
    }
  };
  reader.onerror = function() { setStatus('Error: Could not read file.', 'error'); };
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
  for (var i = 0; i < keys.length; i++) {
    if (row[keys[i]] != null && row[keys[i]] !== '') return String(row[keys[i]]).trim();
  }
  return '';
}

function parseRegistryRows(rows) {
  var parsed = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var elrId = normaliseElrId(getCell(row, COL_MAP.elrId));
    if (!elrId) continue;

    parsed.push({
      elrId:           elrId,
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

  parsed.sort(function(a, b) {
    var seqA = parseInt(a.seq, 10);
    var seqB = parseInt(b.seq, 10);
    if (!isNaN(seqA) && !isNaN(seqB)) return seqA - seqB;
    if (!isNaN(seqA)) return -1;
    if (!isNaN(seqB)) return 1;
    var numA = parseInt(a.elrId.slice(4), 10);
    var numB = parseInt(b.elrId.slice(4), 10);
    return numA - numB;
  });

  return parsed;
}

// ── Image Loading ────────────────────────────────────────
imagesInput.addEventListener('change', function(e) {
  loadImageFiles(Array.from(e.target.files));
  imagesInput.value = '';
});

dropZone.addEventListener('dragenter', function(e) { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragover',  function(e) { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', function(e) { if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over'); });
dropZone.addEventListener('drop', function(e) {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  var files = Array.from(e.dataTransfer.files).filter(function(f) { return isImageFile(f); });
  if (files.length) loadImageFiles(files);
  else setStatus('No supported image files dropped.', 'warn');
});
dropZone.addEventListener('click', function() { imagesInput.click(); });

function isImageFile(file) {
  return /^image\/(png|jpeg|webp|gif)$/.test(file.type) ||
    /\.(png|jpe?g|webp|gif)$/i.test(file.name);
}

function loadImageFiles(files) {
  if (!files.length) return;

  var supported = files.filter(isImageFile);
  var skipped = files.length - supported.length;
  var added = 0;
  var duplicates = 0;

  var promises = supported.map(function(file) {
    return new Promise(function(resolve) {
      var exists = state.images.some(function(img) {
        return img.name === file.name && img.size === file.size;
      });
      if (exists) { duplicates++; resolve(); return; }

      var url = URL.createObjectURL(file);
      var imgEl = new Image();
      var addImage = function(w, h) {
        var orient = 'Unknown';
        if (w > 0 && h > 0) {
          if (w === h) orient = 'Square';
          else if (w > h) orient = 'Landscape';
          else orient = 'Portrait';
        }
        var ext = file.name.split('.').pop().toLowerCase();
        state.images.push({
          name: file.name,
          size: file.size,
          ext: ext,
          width: w,
          height: h,
          orientation: orient,
          previewUrl: url,
          detectedElrId: detectElrFromFilename(file.name),
          file: file,  // keep original File object for ZIP export
        });
        added++;
        resolve();
      };
      imgEl.onload = function() { addImage(imgEl.naturalWidth, imgEl.naturalHeight); };
      imgEl.onerror = function() { addImage(0, 0); };
      imgEl.src = url;
    });
  });

  Promise.all(promises).then(function() {
    renderGallery();
    updateStats();
    updateButtonStates();
    var msg = added + ' image(s) added.';
    if (duplicates) msg += ' ' + duplicates + ' duplicate(s) skipped.';
    if (skipped) msg += ' ' + skipped + ' unsupported file(s) ignored.';
    setStatus(msg, added > 0 ? 'success' : 'warn');
  });
}

// ── Selection State ──────────────────────────────────────
function selectRow(elrId) {
  state.selectedElrId = elrId;
  renderRegistry();
  renderSelectedInfo();
  renderGallery();
}

function deselectRow() {
  state.selectedElrId = null;
  renderRegistry();
  selectedInfo.hidden = true;
  renderGallery();
}

deselectBtn.addEventListener('click', deselectRow);

// ── Assignment State ─────────────────────────────────────
function assignImage(elrId, imageName, source, score, reason) {
  // Remove this image from any other ELR row
  for (var id in state.assignments) {
    if (state.assignments[id] === imageName && id !== elrId) {
      delete state.assignments[id];
      delete state.assignmentMeta[id];
    }
  }
  state.assignments[elrId] = imageName;
  setAssignmentMeta(elrId, source || 'manual', score, reason || source || 'manual');
  renderRegistry();
  renderGallery();
  renderSelectedInfo();
  updateStats();
  updateButtonStates();
}

function clearAssignment(elrId) {
  delete state.assignments[elrId];
  delete state.assignmentMeta[elrId];
  renderRegistry();
  renderGallery();
  renderSelectedInfo();
  updateStats();
  updateButtonStates();
}

clearAssignBtn.addEventListener('click', function() {
  if (!Object.keys(state.assignments).length) return;
  state.assignments = {};
  state.assignmentMeta = {};
  renderRegistry();
  renderGallery();
  renderSelectedInfo();
  updateStats();
  updateButtonStates();
  setStatus('All assignments cleared.', 'info');
});

// ── Auto-Assign ──────────────────────────────────────────
autoAssignBtn.addEventListener('click', autoAssign);

async function autoAssign() {
  if (!state.registry.length) { setStatus('Load a workbook first.', 'error'); return; }
  if (!state.images.length)   { setStatus('Load images first.', 'error'); return; }
  if (state.isAutoAssigning)  { setStatus('Auto-assign already running…', 'info'); return; }

  state.isAutoAssigning = true;
  updateButtonStates();

  try {
    state.visualThreshold = readVisualThreshold();
    updateVisualThresholdLabel();

    var matchedByFilename = 0;
    var matchedByVisual = 0;

    setStatus('Auto-assign step 1/2: filename match pass…', 'info');
    matchedByFilename = autoAssignByFilename();

    var remainingRows = state.registry.filter(function(r) { return !state.assignments[r.elrId]; });
    var remainingImages = state.images.filter(function(img) { return !getImageAssignment(img.name); });

    if (remainingRows.length && remainingImages.length) {
      setStatus('Auto-assign step 2/2: loading visual matching model…', 'info');
      matchedByVisual = await autoAssignByVisualSimilarity(remainingRows, remainingImages);
    }

    renderRegistry();
    renderGallery();
    renderSelectedInfo();
    updateStats();

    var total = matchedByFilename + matchedByVisual;
    if (total > 0) {
      setStatus(
        'Auto-assign complete: ' + total + ' mapped (' + matchedByVisual + ' visual, ' + matchedByFilename + ' filename, threshold ' + state.visualThreshold.toFixed(2) + ').',
        'success'
      );
    } else {
      setStatus('Auto-assign complete: no confident visual or filename matches found.', 'info');
    }
  } catch (err) {
    console.error(err);
    setStatus('Auto-assign failed: ' + err.message, 'error');
  } finally {
    state.isAutoAssigning = false;
    updateButtonStates();
  }
}

function autoAssignByFilename() {
  var matched = 0;
  for (var i = 0; i < state.images.length; i++) {
    var img = state.images[i];
    if (!img.detectedElrId) continue;
    var elrId = img.detectedElrId;
    var rowExists = state.registry.some(function(r) { return r.elrId === elrId; });
    if (!rowExists) continue;
    if (state.assignments[elrId]) continue;
    var alreadyUsed = Object.values(state.assignments).includes(img.name);
    if (alreadyUsed) continue;
    state.assignments[elrId] = img.name;
    setAssignmentMeta(elrId, 'filename-detected', null, 'ELR ID detected in filename');
    matched++;
  }
  return matched;
}

async function autoAssignByVisualSimilarity(candidateRows, candidateImages) {
  var classifier = await getVisualClassifier();

  var rowsWithPrompts = candidateRows.map(function(row) {
    return { row: row, prompt: buildVisualPrompt(row) };
  });

  var proposals = [];
  var completed = 0;

  for (var i = 0; i < candidateImages.length; i++) {
    var img = candidateImages[i];
    completed++;
    setStatus(
      'Visual match in progress: scoring image ' + completed + '/' + candidateImages.length + '…',
      'info'
    );

    var labels = rowsWithPrompts.map(function(x) { return x.prompt; });
    var raw = await classifier(img.previewUrl, labels, {
      topk: Math.min(5, labels.length),
    });

    var predictions = Array.isArray(raw) ? raw : [];
    for (var j = 0; j < predictions.length; j++) {
      var pred = predictions[j];
      if (!pred || typeof pred.label !== 'string' || typeof pred.score !== 'number') continue;
      var match = rowsWithPrompts.find(function(x) { return x.prompt === pred.label; });
      if (!match) continue;
      proposals.push({
        elrId: match.row.elrId,
        imageName: img.name,
        score: pred.score,
      });
    }
  }

  proposals.sort(function(a, b) { return b.score - a.score; });
  var usedRows = new Set();
  var usedImages = new Set();
  var matched = 0;

  for (var k = 0; k < proposals.length; k++) {
    var p = proposals[k];
    if (p.score < state.visualThreshold) continue;
    if (state.assignments[p.elrId]) continue;
    if (getImageAssignment(p.imageName)) continue;
    if (usedRows.has(p.elrId) || usedImages.has(p.imageName)) continue;

    state.assignments[p.elrId] = p.imageName;
    setAssignmentMeta(p.elrId, 'visual-analysis', p.score, 'CLIP visual similarity match');
    usedRows.add(p.elrId);
    usedImages.add(p.imageName);
    matched++;
  }

  return matched;
}

function buildVisualPrompt(row) {
  var tokens = [
    row.assetTitle,
    row.appSection,
    row.collection,
    row.usageObjective,
    row.visualDirection,
    row.artNotes,
  ].filter(Boolean);

  if (!tokens.length) {
    return 'fantasy game interface artwork';
  }

  var compact = tokens.join(', ').replace(/\s+/g, ' ').trim();
  return compact.slice(0, 220);
}

async function getVisualClassifier() {
  if (state.visualModelPromise) return state.visualModelPromise;

  var transformersApi = await getTransformersApi();
  state.visualModelPromise = transformersApi.pipeline(
    'zero-shot-image-classification',
    'Xenova/clip-vit-base-patch32'
  );

  return state.visualModelPromise;
}

async function getTransformersApi() {
  if (window.transformers && typeof window.transformers.pipeline === 'function') {
    return window.transformers;
  }
  if (state.transformersApiPromise) return state.transformersApiPromise;

  state.transformersApiPromise = (async function() {
    var urls = [
      'https://cdn.jsdelivr.net/npm/@xenova/transformers@2.17.2/+esm',
      'https://cdn.jsdelivr.net/npm/@xenova/transformers@2.17.2',
      'https://unpkg.com/@xenova/transformers@2.17.2',
    ];

    for (var i = 0; i < urls.length; i++) {
      try {
        var mod = await import(urls[i]);
        if (mod && typeof mod.pipeline === 'function') return mod;
        if (mod && mod.default && typeof mod.default.pipeline === 'function') return mod.default;
      } catch (err) {
        console.warn('Transformers import failed for', urls[i], err);
      }
    }

    throw new Error('Could not load visual model library from CDN. Check internet access and reload.');
  })();

  return state.transformersApiPromise;
}

// ── Render Registry ──────────────────────────────────────
function renderRegistry() {
  if (!state.registry.length) {
    registryTbody.innerHTML = '<tr class="empty-row"><td colspan="6">Load a workbook to see registry rows.</td></tr>';
    registryCount.textContent = '';
    return;
  }

  var q = state.searchQuery.toLowerCase();
  var filtered = q ? state.registry.filter(function(row) { return matchesSearch(row, q); }) : state.registry;
  registryCount.textContent = '(' + filtered.length + ' of ' + state.registry.length + ')';

  if (!filtered.length) {
    registryTbody.innerHTML = '<tr class="empty-row"><td colspan="6">No results for current search.</td></tr>';
    return;
  }

  registryTbody.innerHTML = filtered.map(function(row) {
    var assigned = state.assignments[row.elrId] || '';
    var isSelected = row.elrId === state.selectedElrId;
    var statusHtml = assigned
      ? '<span class="status-badge status-mapped">Mapped</span>'
      : '<span class="status-badge status-unmapped">Unmapped</span>';
    var assignedCell = assigned
      ? '<span class="cell-truncate" title="' + escHtml(assigned) + '">' + escHtml(assigned) + '</span>'
      : '<span style="color:var(--text-muted)">—</span>';

    return '<tr class="' + (isSelected ? 'row-selected' : '') + '" data-elr="' + escHtml(row.elrId) + '">' +
      '<td>' + escHtml(row.elrId) + '</td>' +
      '<td class="cell-truncate" title="' + escHtml(row.assetTitle) + '">' + (escHtml(row.assetTitle) || '—') + '</td>' +
      '<td class="cell-truncate" title="' + escHtml(row.appSection) + '">' + (escHtml(row.appSection) || '—') + '</td>' +
      '<td class="cell-truncate" title="' + escHtml(row.storageFolder) + '">' + (escHtml(row.storageFolder) || '—') + '</td>' +
      '<td class="cell-truncate">' + assignedCell + '</td>' +
      '<td>' + statusHtml + '</td>' +
    '</tr>';
  }).join('');

  registryTbody.querySelectorAll('tr[data-elr]').forEach(function(tr) {
    tr.addEventListener('click', function() {
      var elrId = tr.dataset.elr;
      if (elrId === state.selectedElrId) deselectRow();
      else selectRow(elrId);
    });
  });
}

function matchesSearch(row, q) {
  var assigned = state.assignments[row.elrId] || '';
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
  var row = state.registry.find(function(r) { return r.elrId === state.selectedElrId; });
  if (!row) { selectedInfo.hidden = true; return; }

  var assigned = state.assignments[row.elrId] || '';
  var meta = getAssignmentMeta(row.elrId);
  var fields = [
    { label: 'ELR ID',            value: row.elrId },
    { label: 'Asset Title',       value: row.assetTitle },
    { label: 'Collection',        value: row.collection },
    { label: 'App Section',       value: row.appSection },
    { label: 'Storage Folder',    value: row.storageFolder },
    { label: 'Workbook Filename', value: row.filename },
    { label: 'Orientation',       value: row.orientation },
    { label: 'Assigned Image',    value: assigned, fullWidth: true },
  ];
  if (assigned) {
    fields.push({ label: 'Match Source', value: meta.source });
    fields.push({ label: 'Confidence', value: meta.confidence });
    if (meta.score != null) {
      fields.push({ label: 'Match Score', value: meta.score.toFixed(4) });
    }
  }

  selectedInfoGrid.innerHTML = fields.map(function(f) {
    return '<div class="info-field' + (f.fullWidth ? ' full-width' : '') + '">' +
      '<span class="info-label">' + escHtml(f.label) + '</span>' +
      '<span class="info-value' + (f.value ? '' : ' empty') + '">' + (escHtml(f.value) || 'Not set') + '</span>' +
    '</div>';
  }).join('');

  selectedInfo.hidden = false;
}

// ── Render Gallery ───────────────────────────────────────
function renderGallery() {
  if (!state.images.length) {
    galleryGrid.innerHTML = '';
    galleryCount.textContent = '';
    return;
  }

  galleryCount.textContent = '(' + state.images.length + ')';

  galleryGrid.innerHTML = state.images.map(function(img) {
    var assigned = getImageAssignment(img.name);
    var isAssigned = !!assigned;
    var dims = (img.width && img.height) ? img.width + '×' + img.height : 'unknown';

    var detectedTag = img.detectedElrId
      ? '<span class="card-tag tag-elr">🔍 ' + escHtml(img.detectedElrId) + '</span>'
      : '';
    var assignedTag = isAssigned
      ? '<span class="card-tag tag-assigned">✓ ' + escHtml(assigned) + '</span>'
      : '';

    var canAssign = !!state.selectedElrId;

    return '<div class="image-card ' + (isAssigned ? 'card-assigned' : '') + '" data-img="' + escHtml(img.name) + '">' +
      '<img class="card-thumb" src="' + escHtml(img.previewUrl) + '" alt="' + escHtml(img.name) + '" loading="lazy" />' +
      '<div class="card-body">' +
        '<div class="card-filename" title="' + escHtml(img.name) + '">' + escHtml(img.name) + '</div>' +
        '<div class="card-meta">' +
          '<span class="card-tag tag-dims">' + escHtml(dims) + '</span>' +
          '<span class="card-tag tag-orient">' + escHtml(img.orientation) + '</span>' +
          detectedTag +
          assignedTag +
        '</div>' +
        '<div class="card-actions">' +
          '<button class="btn btn-sm btn-assign" data-action="assign" data-img="' + escHtml(img.name) + '"' +
            (canAssign ? '' : ' disabled') +
            ' title="' + (canAssign ? 'Assign to selected ELR row: ' + escHtml(state.selectedElrId) : 'Select an ELR row first') + '">' +
            (canAssign ? 'Assign to ' + escHtml(state.selectedElrId) : 'Select ELR row first') +
          '</button>' +
          '<button class="btn btn-sm btn-clear-assign" data-action="clear" data-img="' + escHtml(img.name) + '"' +
            (isAssigned ? '' : ' disabled') +
            ' title="' + (isAssigned ? 'Remove assignment for ' + escHtml(assigned) : 'Not assigned') + '">' +
            'Clear assignment' +
          '</button>' +
        '</div>' +
      '</div>' +
    '</div>';
  }).join('');

  galleryGrid.querySelectorAll('button[data-action]').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      var imgName = btn.dataset.img;
      if (btn.dataset.action === 'assign') {
        if (!state.selectedElrId) { setStatus('Select an ELR row first.', 'warn'); return; }
        assignImage(state.selectedElrId, imgName, 'manual', null, 'assigned manually by user');
        setStatus('Assigned "' + imgName + '" → ' + state.selectedElrId, 'success');
      } else if (btn.dataset.action === 'clear') {
        var elrId = getImageAssignment(imgName);
        if (elrId) { clearAssignment(elrId); setStatus('Cleared assignment for ' + elrId + '.', 'info'); }
      }
    });
  });
}

function getImageAssignment(imageName) {
  for (var id in state.assignments) {
    if (state.assignments[id] === imageName) return id;
  }
  return null;
}

// ── Stats ────────────────────────────────────────────────
function updateStats() {
  var total    = state.registry.length;
  var images   = state.images.length;
  var mapped   = Object.keys(state.assignments).length;
  var unmapped = total - mapped;

  statTotal.textContent   = total;
  statImages.textContent  = images;
  statMapped.textContent  = mapped;
  statUnmapped.textContent = Math.max(0, unmapped);
}

// ── Button States ────────────────────────────────────────
function updateButtonStates() {
  var hasRegistry = state.registry.length > 0;
  var hasImages   = state.images.length > 0;
  var hasAny      = Object.keys(state.assignments).length > 0;
  var busy        = state.isAutoAssigning || state.isExporting;

  autoAssignBtn.disabled       = busy || !(hasRegistry && hasImages);
  clearAssignBtn.disabled      = !hasAny || busy;
  exportMappingBtn.disabled    = !hasRegistry || busy;
  exportRenameBtn.disabled     = !hasRegistry || busy;
  exportWbBtn.disabled         = !(hasRegistry && state.workbookRaw) || busy;
  exportAssetsZipBtn.disabled  = !hasAny || busy;
  exportPackageZipBtn.disabled = !(hasRegistry && state.workbookRaw) || busy;
}

// ── Search ───────────────────────────────────────────────
searchInput.addEventListener('input', function() {
  state.searchQuery = searchInput.value;
  renderRegistry();
});

visualThresholdInput.addEventListener('input', function() {
  state.visualThreshold = readVisualThreshold();
  updateVisualThresholdLabel();
});

// ── Data Builders ────────────────────────────────────────

function buildMappingRows() {
  return state.registry.map(function(r) {
    var imgName = state.assignments[r.elrId] || '';
    var meta = getAssignmentMeta(r.elrId);
    return {
      elr_id: r.elrId,
      seq: r.seq,
      asset_title: r.assetTitle,
      collection: r.collection,
      app_section: r.appSection,
      storage_folder: r.storageFolder,
      workbook_filename: r.filename,
      source_image_name: imgName,
      assigned: imgName ? 'yes' : 'no',
      match_score: imgName ? (meta.score != null ? meta.score.toFixed(4) : 'manual') : '',
      confidence: imgName ? meta.confidence : '',
      assignment_source: imgName ? meta.source : '',
      match_reason: imgName ? meta.reason : '',
    };
  });
}

function buildRenameRows(exportedInZip) {
  return state.registry.map(function(r) {
    var imgName = state.assignments[r.elrId] || '';
    var meta = getAssignmentMeta(r.elrId);
    if (!imgName) {
      return {
        elr_id: r.elrId,
        seq: r.seq,
        asset_title: r.assetTitle,
        source_image_name: '',
        source_extension: '',
        rename_to: '',
        target_folder: r.storageFolder,
        final_path: '',
        confidence: '',
        assignment_source: '',
        match_reason: '',
        exported_in_zip: 'no',
      };
    }
    var img = state.images.find(function(i) { return i.name === imgName; });
    var ext = img ? img.ext : imgName.split('.').pop().toLowerCase();
    var renameTo = getTargetFilename(r, img);
    var folder = getTargetFolder(r);
    var finalPath = folder + '/' + renameTo;

    return {
      elr_id: r.elrId,
      seq: r.seq,
      asset_title: r.assetTitle,
      source_image_name: imgName,
      source_extension: ext,
      rename_to: renameTo,
      target_folder: folder,
      final_path: finalPath,
      confidence: meta.confidence,
      assignment_source: meta.source,
      match_reason: meta.reason,
      exported_in_zip: exportedInZip ? 'yes' : 'no',
    };
  });
}

function buildPatchedWorkbookBlob() {
  var wb = state.workbookRaw;
  var SHEET_NAME = 'ELR Master Registry';
  var ws = wb.Sheets[SHEET_NAME];

  var range = XLSX.utils.decode_range(ws['!ref']);
  var headers = {};
  for (var c = range.s.c; c <= range.e.c; c++) {
    var cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c: c })];
    if (cell && cell.v != null) headers[String(cell.v).trim()] = c;
  }

  var colFilename     = findColumnIndex(headers, COL_MAP.filename);
  var colUploadStatus = findColumnIndex(headers, COL_MAP.uploadStatus);
  var colFinalUrl     = findColumnIndex(headers, COL_MAP.finalUrl);
  var colNotes        = findColumnIndex(headers, COL_MAP.notes);

  var cloneWb = XLSX.read(
    XLSX.write(wb, { bookType: 'xlsx', type: 'array' }),
    { type: 'array' }
  );
  var targetWs = cloneWb.Sheets[SHEET_NAME];

  for (var r = range.s.r + 1; r <= range.e.r; r++) {
    var elrCell = targetWs[XLSX.utils.encode_cell({ r: r, c: findColumnIndex(headers, COL_MAP.elrId) })];
    if (!elrCell) continue;
    var elrId = normaliseElrId(elrCell.v);
    if (!elrId) continue;
    var imgName = state.assignments[elrId];
    if (!imgName) continue;

    var registryRow = state.registry.find(function(row) { return row.elrId === elrId; });
    if (!registryRow) continue;

    var img = state.images.find(function(i) { return i.name === imgName; });
    var renameTo = getTargetFilename(registryRow, img);
    var folder = getTargetFolder(registryRow);
    var finalPath = '/' + folder + '/' + renameTo;
    var meta = getAssignmentMeta(elrId);
    var noteAppend = getAssignmentSourceLabel(meta) + ' from source image ' + imgName;

    if (colFilename !== -1) setCellValue(targetWs, r, colFilename, renameTo);
    if (colUploadStatus !== -1) setCellValue(targetWs, r, colUploadStatus, 'Mapped');
    if (colFinalUrl !== -1) setCellValue(targetWs, r, colFinalUrl, finalPath);
    if (colNotes !== -1) {
      var existing = getCellValue(targetWs, r, colNotes);
      var newNote = existing ? existing + '; ' + noteAppend : noteAppend;
      setCellValue(targetWs, r, colNotes, newNote);
    }
  }

  var outArray = XLSX.write(cloneWb, { bookType: 'xlsx', type: 'array' });
  return new Blob([outArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// ── Export Event Listeners ───────────────────────────────
exportMappingBtn.addEventListener('click', function() {
  if (!state.registry.length) { setStatus('Load a workbook first.', 'error'); return; }
  exportMappingCsv();
});

exportRenameBtn.addEventListener('click', function() {
  if (!state.registry.length) { setStatus('Load a workbook first.', 'error'); return; }
  exportRenameCsv();
});

exportWbBtn.addEventListener('click', function() {
  if (!state.workbookRaw) { setStatus('Load a workbook first.', 'error'); return; }
  exportPatchedWorkbook();
});

exportAssetsZipBtn.addEventListener('click', function() {
  if (!Object.keys(state.assignments).length) { setStatus('No assigned images to export.', 'error'); return; }
  exportAssetsZip();
});

exportPackageZipBtn.addEventListener('click', function() {
  if (!state.workbookRaw) { setStatus('Load a workbook first.', 'error'); return; }
  exportCompletePackageZip();
});

// ── CSV Exports ──────────────────────────────────────────
function exportMappingCsv() {
  var mappingRows = buildMappingRows();
  var headers = [
    'elr_id','seq','asset_title','collection','app_section',
    'storage_folder','workbook_filename','source_image_name','assigned',
    'match_score','confidence','assignment_source','match_reason'
  ];
  var rows = mappingRows.map(function(r) { return headers.map(function(h) { return r[h]; }); });
  downloadCsv([headers].concat(rows), 'elr_mapping.csv');
  setStatus('Mapping CSV exported.', 'success');
}

function exportRenameCsv() {
  var renameRows = buildRenameRows(false);
  var headers = [
    'elr_id','seq','asset_title','source_image_name','source_extension',
    'rename_to','target_folder','final_path','confidence',
    'assignment_source','match_reason','exported_in_zip'
  ];
  var rows = renameRows.map(function(r) { return headers.map(function(h) { return r[h]; }); });
  downloadCsv([headers].concat(rows), 'elr_rename_manifest.csv');
  setStatus('Rename manifest CSV exported.', 'success');
}

function exportPatchedWorkbook() {
  try {
    var blob = buildPatchedWorkbookBlob();
    var baseName = state.workbookFileName.replace(/\.[^.]+$/, '');
    var outName = baseName + '_patched.xlsx';
    triggerDownload(blob, outName);
    setStatus('Patched workbook exported: ' + outName, 'success');
  } catch (err) {
    setStatus('Error exporting workbook: ' + err.message, 'error');
    console.error(err);
  }
}

// ── ZIP Exports ──────────────────────────────────────────

async function buildAssetsZip(zip, prefix) {
  var assignedEntries = Object.entries(state.assignments);
  if (!assignedEntries.length) return 0;

  var addedCount = 0;

  for (var i = 0; i < assignedEntries.length; i++) {
    var elrId = assignedEntries[i][0];
    var imgName = assignedEntries[i][1];
    var row = state.registry.find(function(r) { return r.elrId === elrId; });
    if (!row) continue;
    var img = state.images.find(function(im) { return im.name === imgName; });
    if (!img) continue;

    var folder = getTargetFolder(row);
    var filename = getTargetFilename(row, img);
    var zipPath = prefix ? prefix + '/' + folder + '/' + filename : folder + '/' + filename;

    if (img.file) {
      zip.file(zipPath, img.file);
    } else {
      var resp = await fetch(img.previewUrl);
      var blob = await resp.blob();
      zip.file(zipPath, blob);
    }
    addedCount++;
  }

  return addedCount;
}

async function exportAssetsZip() {
  if (state.isExporting) { setStatus('Export already in progress…', 'info'); return; }
  state.isExporting = true;
  updateButtonStates();

  try {
    setStatus('Building assets ZIP…', 'info');
    var zip = new JSZip();
    var count = await buildAssetsZip(zip, '');
    if (count === 0) {
      setStatus('No assigned images to export.', 'warn');
      return;
    }

    setStatus('Added ' + count + ' assets to ZIP, compressing…', 'info');
    var blob = await zip.generateAsync({ type: 'blob' }, function(meta) {
      if (meta.percent) {
        setStatus('Compressing ZIP: ' + Math.round(meta.percent) + '%…', 'info');
      }
    });

    triggerDownload(blob, 'elr_renamed_assets.zip');
    setStatus('Assets ZIP ready: ' + count + ' renamed files exported.', 'success');
  } catch (err) {
    setStatus('Error building assets ZIP: ' + err.message, 'error');
    console.error(err);
  } finally {
    state.isExporting = false;
    updateButtonStates();
  }
}

async function exportCompletePackageZip() {
  if (state.isExporting) { setStatus('Export already in progress…', 'info'); return; }
  state.isExporting = true;
  updateButtonStates();

  try {
    setStatus('Building complete package ZIP…', 'info');
    var zip = new JSZip();

    // 1. Renamed assets under assets/ folder
    var assetCount = await buildAssetsZip(zip, 'assets');
    setStatus('Added ' + assetCount + ' assets. Building CSVs…', 'info');

    // 2. Mapping CSV
    var mappingHeaders = [
      'elr_id','seq','asset_title','collection','app_section',
      'storage_folder','workbook_filename','source_image_name','assigned',
      'match_score','confidence','assignment_source','match_reason'
    ];
    var mappingRows = buildMappingRows();
    var mappingCsv = buildCsvString([mappingHeaders].concat(
      mappingRows.map(function(r) { return mappingHeaders.map(function(h) { return r[h]; }); })
    ));
    zip.file('exports/elr_mapping.csv', mappingCsv);

    // 3. Rename manifest CSV
    var renameHeaders = [
      'elr_id','seq','asset_title','source_image_name','source_extension',
      'rename_to','target_folder','final_path','confidence',
      'assignment_source','match_reason','exported_in_zip'
    ];
    var renameRows = buildRenameRows(true);
    var renameCsv = buildCsvString([renameHeaders].concat(
      renameRows.map(function(r) { return renameHeaders.map(function(h) { return r[h]; }); })
    ));
    zip.file('exports/elr_rename_manifest.csv', renameCsv);

    // 4. Patched workbook
    setStatus('Building patched workbook…', 'info');
    var wbBlob = buildPatchedWorkbookBlob();
    var baseName = state.workbookFileName.replace(/\.[^.]+$/, '');
    zip.file('exports/' + baseName + '_patched.xlsx', wbBlob);

    setStatus('Compressing complete package…', 'info');
    var pkgBlob = await zip.generateAsync({ type: 'blob' }, function(meta) {
      if (meta.percent) {
        setStatus('Compressing package: ' + Math.round(meta.percent) + '%…', 'info');
      }
    });

    triggerDownload(pkgBlob, 'elr_complete_package.zip');
    setStatus('Complete package ready: ' + assetCount + ' assets + CSVs + patched workbook.', 'success');
  } catch (err) {
    setStatus('Error building complete package: ' + err.message, 'error');
    console.error(err);
  } finally {
    state.isExporting = false;
    updateButtonStates();
  }
}

// ── Workbook Helpers ─────────────────────────────────────
function findColumnIndex(headers, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (headers[keys[i]] !== undefined) return headers[keys[i]];
  }
  return -1;
}

function setCellValue(ws, r, c, value) {
  var addr = XLSX.utils.encode_cell({ r: r, c: c });
  if (!ws[addr]) ws[addr] = {};
  ws[addr].v = value;
  ws[addr].t = 's';
}

function getCellValue(ws, r, c) {
  var addr = XLSX.utils.encode_cell({ r: r, c: c });
  if (!ws[addr]) return '';
  return ws[addr].v != null ? String(ws[addr].v) : '';
}

// ── CSV Helpers ──────────────────────────────────────────
function buildCsvString(rows) {
  return '\uFEFF' + rows.map(function(row) {
    return row.map(function(cell) {
      var str = String(cell == null ? '' : cell);
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',');
  }).join('\r\n');
}

function downloadCsv(rows, filename) {
  var content = buildCsvString(rows);
  var blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  triggerDownload(blob, filename);
}

function triggerDownload(blob, filename) {
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(function() { URL.revokeObjectURL(url); }, 5000);
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
updateVisualThresholdLabel();
