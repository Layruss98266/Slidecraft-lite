// ══════════════════════════════════════════════════════════════════════
// SlideCraft Lite — minimal client
// ══════════════════════════════════════════════════════════════════════
const NUM_SLIDES = window.NUM_SLIDES || 0;
let currentSlide = NUM_SLIDES > 0 ? 1 : 0;
let serverUndo = { can_undo: false, can_redo: false, undo_label: null, redo_label: null };

const slideImg = document.getElementById('slide-img');

// ── Slide zoom ──────────────────────────────────────────────────────────
const ZOOM_MIN = 20, ZOOM_MAX = 200, ZOOM_DEFAULT = 70, ZOOM_STEP = 10;
let zoomPct = ZOOM_DEFAULT;

function setZoom(pct) {
  pct = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, Math.round(+pct)));
  zoomPct = pct;
  const container = document.getElementById('slide-container');
  if (container) container.style.width = pct + '%';
  const label = document.getElementById('zoom-val');
  if (label) label.textContent = pct + '%';
  const range = document.getElementById('zoom-range');
  if (range && +range.value !== pct) range.value = pct;
  try { localStorage.setItem('lite_zoom', pct); } catch (e) {}
}
function zoomIn()    { setZoom(zoomPct + ZOOM_STEP); }
function zoomOut()   { setZoom(zoomPct - ZOOM_STEP); }
function zoomReset() { setZoom(ZOOM_DEFAULT); }
function zoomFit() {
  // Pick the largest zoom where the slide (16:9) fits inside the viewport.
  const vp = document.getElementById('viewport');
  if (!vp) return;
  const padding = 64; // 32 px viewport padding on each side
  const availW = vp.clientWidth - padding;
  const availH = vp.clientHeight - padding;
  // width = vp.clientWidth * (pct/100). slide aspect 16:9 → height = width * 9/16.
  // We need width ≤ availW AND width * 9/16 ≤ availH.
  const maxWidthByH = availH * 16 / 9;
  const targetW = Math.min(availW, maxWidthByH);
  const pct = Math.round((targetW / vp.clientWidth) * 100);
  setZoom(pct);
}

// ── Toast / loading helpers ─────────────────────────────────────────────
let _toastTimer = null;
function showToast(msg, kind = 'info', durationMs = 2400) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = `show ${kind}`;
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => t.className = '', durationMs);
}
function showLoading(msg = 'Processing...') {
  document.getElementById('loading-msg').textContent = msg;
  document.getElementById('loading-overlay').classList.add('active');
}
function hideLoading() {
  document.getElementById('loading-overlay').classList.remove('active');
}

// ── Slide navigation ────────────────────────────────────────────────────
function gotoSlide(n) {
  if (NUM_SLIDES === 0) return;
  if (n < 1) n = 1;
  if (n > NUM_SLIDES) n = NUM_SLIDES;
  currentSlide = n;
  const pad = String(n).padStart(2, '0');
  if (slideImg) {
    slideImg.style.opacity = '0.5';
    slideImg.src = `/sl/slide-${pad}.jpg?t=${Date.now()}`;
    slideImg.onload = () => { slideImg.style.opacity = '1'; };
  }
  document.querySelectorAll('.thumb-item').forEach(t => t.classList.remove('active'));
  document.getElementById(`thumb-${n}`)?.classList.add('active');
  document.getElementById(`thumb-${n}`)?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  const input = document.getElementById('slide-num-input');
  if (input && document.activeElement !== input) input.value = n;
  // Persist current slide across full-page reloads (delete/duplicate trigger one)
  try { sessionStorage.setItem('lite_current_slide', n); } catch (e) {}
}

function reloadAllSlides() {
  const t = Date.now();
  for (let i = 1; i <= NUM_SLIDES; i++) {
    const pad = String(i).padStart(2, '0');
    const thumb = document.querySelector(`#thumb-${i} img`);
    if (thumb) thumb.src = `/sl/slide-${pad}.jpg?t=${t}`;
  }
  if (slideImg && currentSlide > 0) {
    const pad = String(currentSlide).padStart(2, '0');
    slideImg.src = `/sl/slide-${pad}.jpg?t=${t}`;
  }
}

// ── Upload (single + bulk) ──────────────────────────────────────────────
async function uploadPPTX(input) {
  const file = input.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.pptx')) {
    showToast('Only .pptx files supported', 'error');
    input.value = '';
    return;
  }
  showLoading('Converting PPTX — this can take 30-60s...');
  const form = new FormData();
  form.append('file', file);
  try {
    const resp = await fetch('/api/upload', { method: 'POST', body: form });
    hideLoading();
    const data = await resp.json();
    if (data.ok) {
      showToast(`Uploaded ${data.num_slides} slide(s)`, 'success');
      setTimeout(() => location.reload(), 600);
    } else {
      showToast(data.error || 'Upload failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Upload error: ' + e.message, 'error');
  }
  input.value = '';
}

async function bulkRemoveLogo(input) {
  const files = Array.from(input.files);
  if (!files.length) return;
  if (files.length > 10) {
    showToast('Bulk cap is 10 files. Run a second batch for the rest.', 'error', 6000);
    input.value = '';
    return;
  }
  const nonPptx = files.filter(f => !f.name.toLowerCase().endsWith('.pptx'));
  if (nonPptx.length) {
    showToast(`Only .pptx files allowed (${nonPptx.length} rejected)`, 'error');
    input.value = '';
    return;
  }
  const totalBytes = files.reduce((s, f) => s + f.size, 0);
  const totalMB = totalBytes / (1024 * 1024);
  if (totalBytes > 300 * 1024 * 1024) {
    showToast(`Combined size ${totalMB.toFixed(0)} MB exceeds 300 MB cap`, 'error', 6000);
    input.value = '';
    return;
  }

  showLoading(`Uploading ${files.length} file${files.length > 1 ? 's' : ''} (${totalMB.toFixed(0)} MB)...`);
  showProgress(0, files.length, 'Uploading…');
  const form = new FormData();
  files.forEach(f => form.append('files', f));
  try {
    const resp = await fetch('/api/batch/remove-logo', { method: 'POST', body: form });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      hideLoading();
      showToast(data.error || 'Bulk processing failed', 'error', 8000);
      input.value = '';
      return;
    }
    await pollJob(data.job_id, files.length);
  } catch (e) {
    hideLoading();
    showToast('Bulk error: ' + e.message, 'error');
  }
  input.value = '';
}

async function pollJob(jobId, expectedTotal) {
  let lastDone = -1;
  while (true) {
    await new Promise(r => setTimeout(r, 700));
    let state;
    try {
      const r = await fetch(`/api/job/${jobId}`);
      state = await r.json();
    } catch (e) { continue; }
    if (!state || state.error) {
      hideLoading();
      showToast('Job error: ' + (state && state.error || 'unknown'), 'error');
      return;
    }
    const done = state.done || 0;
    const total = state.total || expectedTotal || 1;
    if (done !== lastDone) {
      showProgress(done, total, state.message || '');
      lastDone = done;
    }
    if (state.status === 'done') {
      hideProgress();
      hideLoading();
      const a = document.createElement('a');
      a.href = `/api/job/${jobId}/download`;
      a.download = 'SlideCraft_Bulk_Cleaned.zip';
      document.body.appendChild(a); a.click(); a.remove();
      const okCount = (state.results || []).filter(r => r.status === 'ok').length;
      const errCount = (state.results || []).filter(r => r.status === 'error').length;
      showToast(`Done — ${okCount} cleaned${errCount ? `, ${errCount} failed` : ''}`, 'success', 5000);
      return;
    }
    if (state.status === 'error') {
      hideProgress();
      hideLoading();
      showToast(state.message || 'Bulk processing failed', 'error', 8000);
      return;
    }
  }
}

function showProgress(done, total, label) {
  const shell = document.getElementById('progress-shell');
  const fill = document.getElementById('progress-fill');
  const detail = document.getElementById('progress-detail');
  if (!shell) return;
  shell.style.display = 'block';
  const pct = total > 0 ? Math.round((done / total) * 100) : 0;
  fill.style.width = pct + '%';
  detail.textContent = `${label}  (${done}/${total} • ${pct}%)`;
}
function hideProgress() {
  const shell = document.getElementById('progress-shell');
  if (shell) shell.style.display = 'none';
}

// ── Remove logo on the currently loaded deck ────────────────────────────
async function removeLogos() {
  if (NUM_SLIDES === 0) {
    showToast('Upload a PPTX first', 'info');
    return;
  }
  showLoading('Removing NotebookLM logo from all slides...');
  try {
    const resp = await fetch('/api/remove-logo', { method: 'POST' });
    hideLoading();
    const data = await resp.json();
    if (data.ok) {
      reloadAllSlides();
      refreshUndoState();
      showToast(`Logo removed from ${data.count} slide(s)`, 'success');
    } else {
      showToast(data.error || 'Failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Error: ' + e.message, 'error');
  }
}

// ── Save checkpoint ─────────────────────────────────────────────────────
async function saveCheckpoint() {
  if (NUM_SLIDES === 0) {
    showToast('Nothing to save — upload a PPTX first', 'info');
    return;
  }
  try {
    const resp = await fetch('/api/save', { method: 'POST' });
    const data = await resp.json();
    if (data.ok) {
      refreshUndoState();
      showToast('Checkpoint saved — use Undo to roll back', 'success');
    } else {
      showToast(data.error || 'Save failed', 'error');
    }
  } catch (e) { showToast('Save error: ' + e.message, 'error'); }
}

// ── Reset ──────────────────────────────────────────────────────────────
async function resetAll() {
  if (NUM_SLIDES === 0) {
    showToast('Nothing to reset', 'info');
    return;
  }
  const msg = 'Reset all slides to the version uploaded?\n\n' +
    'A history snapshot is taken first so you can undo this with Ctrl+Z.';
  if (!confirm(msg)) return;
  showLoading('Restoring originals...');
  try {
    const resp = await fetch('/api/reset-all', { method: 'POST' });
    hideLoading();
    const data = await resp.json();
    if (data.ok) {
      reloadAllSlides();
      refreshUndoState();
      showToast(`Restored ${data.slides_restored} slide(s) from ${data.source}`, 'success', 4000);
    } else {
      showToast(data.error || 'Reset failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Reset error: ' + e.message, 'error');
  }
}

// ── Export ─────────────────────────────────────────────────────────────
function toggleExportMenu(event) {
  event.stopPropagation();
  document.getElementById('export-menu').classList.toggle('open');
}
function closeExportMenu() { document.getElementById('export-menu').classList.remove('open'); }
document.addEventListener('click', closeExportMenu);

async function _downloadExport(url, filename) {
  if (NUM_SLIDES === 0) {
    showToast('No slides to export', 'info');
    return;
  }
  showLoading('Exporting...');
  try {
    const resp = await fetch(url, { method: 'POST' });
    hideLoading();
    if (!resp.ok) {
      let msg = 'Export failed';
      try { msg = (await resp.json()).error || msg; } catch (e) {}
      showToast(msg, 'error');
      return;
    }
    const blob = await resp.blob();
    const objUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = objUrl; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(objUrl);
    showToast(`Downloaded ${filename}`, 'success');
  } catch (e) {
    hideLoading();
    showToast('Export error: ' + e.message, 'error');
  }
}
function exportPPTX() { _downloadExport('/api/export', 'SlideCraft_Export.pptx'); }
function exportPDF()  { _downloadExport('/api/export-pdf', 'Slides_Export.pdf'); }

// ── Undo / Redo ────────────────────────────────────────────────────────
async function refreshUndoState() {
  try {
    const resp = await fetch('/api/ops/state');
    serverUndo = await resp.json();
  } catch (e) {
    serverUndo = { can_undo: false, can_redo: false };
  }
  document.getElementById('btn-undo').classList.toggle('disabled', !serverUndo.can_undo);
  document.getElementById('btn-redo').classList.toggle('disabled', !serverUndo.can_redo);
  const u = document.getElementById('btn-undo');
  const r = document.getElementById('btn-redo');
  if (u) u.title = serverUndo.undo_label ? `Undo: ${serverUndo.undo_label}` : 'Undo (Ctrl+Z)';
  if (r) r.title = serverUndo.redo_label ? `Redo: ${serverUndo.redo_label}` : 'Redo (Ctrl+Y)';
  updateAutosaveIndicator(serverUndo.last_saved);
}

// ── Autosave indicator ─────────────────────────────────────────────────
function _humanAgo(iso) {
  if (!iso) return '';
  const then = new Date(iso).getTime();
  if (isNaN(then)) return '';
  const sec = Math.max(0, Math.floor((Date.now() - then) / 1000));
  if (sec < 5) return 'just now';
  if (sec < 60) return sec + 's ago';
  if (sec < 3600) return Math.floor(sec / 60) + 'm ago';
  if (sec < 86400) return Math.floor(sec / 3600) + 'h ago';
  return Math.floor(sec / 86400) + 'd ago';
}
let _lastSavedISO = '';
function updateAutosaveIndicator(iso) {
  if (iso) _lastSavedISO = iso;
  const el = document.getElementById('autosave-text');
  const wrap = document.getElementById('autosave-indicator');
  if (!el || !wrap) return;
  if (!_lastSavedISO) {
    el.textContent = NUM_SLIDES === 0 ? 'No deck' : 'Idle';
    wrap.classList.remove('saved');
    return;
  }
  el.textContent = 'Saved ' + _humanAgo(_lastSavedISO);
  wrap.classList.add('saved');
}
setInterval(() => updateAutosaveIndicator(), 15000);

async function undo() {
  await refreshUndoState();
  if (!serverUndo.can_undo) {
    showToast('Nothing to undo', 'info');
    return;
  }
  showLoading('Undoing...');
  try {
    const resp = await fetch('/api/ops/undo', { method: 'POST' });
    const data = await resp.json();
    hideLoading();
    if (data.ok) {
      reloadAllSlides();
      refreshUndoState();
      showToast(`Undid: ${data.text || data.kind}`, 'success');
    } else if (data.reason) {
      showToast(data.reason, 'info');
    } else {
      showToast(data.error || 'Undo failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Undo error: ' + e.message, 'error');
  }
}

async function redo() {
  await refreshUndoState();
  if (!serverUndo.can_redo) {
    showToast('Nothing to redo', 'info');
    return;
  }
  showLoading('Redoing...');
  try {
    const resp = await fetch('/api/ops/redo', { method: 'POST' });
    const data = await resp.json();
    hideLoading();
    if (data.ok) {
      reloadAllSlides();
      refreshUndoState();
      showToast(`Redid: ${data.text || data.kind}`, 'success');
    } else if (data.reason) {
      showToast(data.reason, 'info');
    } else {
      showToast(data.error || 'Redo failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Redo error: ' + e.message, 'error');
  }
}

// ── Presentation mode ──────────────────────────────────────────────────
let presOn = false;
let presSlide = 1;
function enterPresentation() {
  if (NUM_SLIDES === 0) {
    showToast('No slides to present', 'info');
    return;
  }
  presOn = true;
  presSlide = currentSlide || 1;
  document.getElementById('presentation-mode').classList.add('active');
  updatePresImg();
  if (document.documentElement.requestFullscreen) {
    document.documentElement.requestFullscreen().catch(() => {});
  }
}
function exitPresentation() {
  presOn = false;
  document.getElementById('presentation-mode').classList.remove('active');
  if (document.fullscreenElement && document.exitFullscreen) {
    document.exitFullscreen().catch(() => {});
  }
}
function updatePresImg() {
  const pad = String(presSlide).padStart(2, '0');
  document.getElementById('pres-img').src = `/sl/slide-${pad}.jpg`;
}
function presNext() {
  if (presSlide < NUM_SLIDES) { presSlide++; updatePresImg(); }
}
function presPrev() {
  if (presSlide > 1) { presSlide--; updatePresImg(); }
}

// ── Help modal ──────────────────────────────────────────────────────────
function openHelp() { document.getElementById('help-modal').classList.add('show'); }
function closeHelp() { document.getElementById('help-modal').classList.remove('show'); }

// ── Keyboard shortcuts ──────────────────────────────────────────────────
document.addEventListener('keydown', e => {
  // In input fields, only handle Esc
  const inField = ['INPUT','TEXTAREA'].includes(document.activeElement?.tagName);

  if (e.key === 'Escape') {
    if (presOn) { exitPresentation(); return; }
    document.querySelectorAll('.modal-backdrop.show').forEach(m => m.classList.remove('show'));
    closeExportMenu();
    return;
  }
  if (inField) return;

  if (presOn) {
    if (e.key === 'ArrowRight' || e.key === ' ') { e.preventDefault(); presNext(); return; }
    if (e.key === 'ArrowLeft') { e.preventDefault(); presPrev(); return; }
    return;
  }

  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
    e.preventDefault(); undo(); return;
  }
  if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'y'
       || (e.shiftKey && e.key.toLowerCase() === 'z'))) {
    e.preventDefault(); redo(); return;
  }
  if (e.key === 'ArrowRight') { e.preventDefault(); gotoSlide(currentSlide + 1); }
  if (e.key === 'ArrowLeft')  { e.preventDefault(); gotoSlide(currentSlide - 1); }
  if (e.key === 'Home')       { e.preventDefault(); gotoSlide(1); }
  if (e.key === 'End')        { e.preventDefault(); gotoSlide(NUM_SLIDES); }
  if (e.key === 'Delete' && !e.ctrlKey && !e.metaKey) { e.preventDefault(); deleteSlide(); }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd') {
    e.preventDefault(); duplicateSlide();
  }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
    e.preventDefault(); printDeck();
  }
  if (e.key === 'F5')         { e.preventDefault(); enterPresentation(); }
  if ((e.key.toLowerCase() === 'h' || e.key === '?') && !e.ctrlKey && !e.metaKey) { openHelp(); }
  if (e.key.toLowerCase() === 't' && !e.ctrlKey && !e.metaKey) { toggleTheme(); }
  if (e.key.toLowerCase() === 'e' && !e.ctrlKey && !e.metaKey) { toggleRegionErase(); }
  if (e.key.toLowerCase() === 'r' && !e.ctrlKey && !e.metaKey) { toggleRecentMenu(e); }
  // Zoom shortcuts: + / = (zoom in), - / _ (zoom out), 0 (reset to default), f (fit)
  if (e.key === '+' || e.key === '=') { e.preventDefault(); zoomIn(); }
  if (e.key === '-' || e.key === '_') { e.preventDefault(); zoomOut(); }
  if (e.key === '0') { e.preventDefault(); zoomReset(); }
  if (e.key.toLowerCase() === 'f' && !e.ctrlKey && !e.metaKey) { e.preventDefault(); zoomFit(); }
});

// ── Drag & drop upload ──────────────────────────────────────────────────
let _dragCounter = 0;
window.addEventListener('dragenter', e => {
  if (!e.dataTransfer?.types?.includes('Files')) return;
  _dragCounter++;
  document.getElementById('drop-overlay').classList.add('active');
});
window.addEventListener('dragover', e => { e.preventDefault(); });
window.addEventListener('dragleave', () => {
  _dragCounter = Math.max(0, _dragCounter - 1);
  if (_dragCounter === 0) document.getElementById('drop-overlay').classList.remove('active');
});
window.addEventListener('drop', e => {
  e.preventDefault();
  _dragCounter = 0;
  document.getElementById('drop-overlay').classList.remove('active');
  const files = Array.from(e.dataTransfer.files || []);
  if (files.length === 0) return;
  const pptx = files.filter(f => f.name.toLowerCase().endsWith('.pptx'));
  if (pptx.length === 0) {
    showToast('Only .pptx files supported', 'error');
    return;
  }
  if (pptx.length === 1) {
    const input = document.getElementById('upload-input');
    const dt = new DataTransfer();
    dt.items.add(pptx[0]);
    input.files = dt.files;
    uploadPPTX(input);
  } else {
    const input = document.getElementById('bulk-input');
    const dt = new DataTransfer();
    pptx.forEach(f => dt.items.add(f));
    input.files = dt.files;
    bulkRemoveLogo(input);
  }
});

// ── Init ───────────────────────────────────────────────────────────────
window.addEventListener('load', () => {
  if (NUM_SLIDES > 0) {
    // Restore zoom from a previous session (persisted), else default 70%
    let restoredZoom = ZOOM_DEFAULT;
    try {
      const saved = parseInt(localStorage.getItem('lite_zoom'), 10);
      if (!isNaN(saved) && saved >= ZOOM_MIN && saved <= ZOOM_MAX) restoredZoom = saved;
    } catch (e) {}
    setZoom(restoredZoom);

    // Restore current slide across reload (delete/duplicate/reorder reload the page)
    let startSlide = 1;
    try {
      const saved = parseInt(sessionStorage.getItem('lite_current_slide'), 10);
      if (!isNaN(saved) && saved >= 1 && saved <= NUM_SLIDES) startSlide = saved;
    } catch (e) {}
    gotoSlide(startSlide);

    initThumbDrag();
    _wireSlideNumInput();
  }
  refreshUndoState();
  refreshDeckInfo();
});

// Keep overlay aligned on resize
window.addEventListener('resize', () => {
  // No overlay canvas in lite — just a no-op placeholder so future code can hook in
});

// ── Watermark modal ───────────────────────────────────────────────────
let wmType = 'text';
let wmPosition = 'center';
let wmScope = 'all';
// Custom-position coordinates (normalised 0..1), defaulting to centre
let wmCustomX = 0.5;
let wmCustomY = 0.5;

function openWatermarkModal() {
  if (NUM_SLIDES === 0) {
    showToast('Upload a PPTX first', 'info');
    return;
  }
  document.getElementById('wm-modal').classList.add('show');
  // Point the preview at the current slide
  const pad = String(currentSlide).padStart(2, '0');
  document.getElementById('wm-preview-bg').src = `/sl/slide-${pad}.jpg?t=${Date.now()}`;
  document.getElementById('wm-preview-num').textContent = currentSlide;
  // Wire control labels once, idempotently
  _bindWatermarkControls();
  // First render
  renderWatermarkPreview();
}

function closeWatermarkModal() {
  document.getElementById('wm-modal').classList.remove('show');
}

function setWmType(type) {
  wmType = type;
  document.getElementById('wm-tab-text').classList.toggle('active', type === 'text');
  document.getElementById('wm-tab-image').classList.toggle('active', type === 'image');
  document.getElementById('wm-text-pane').style.display = type === 'text' ? '' : 'none';
  document.getElementById('wm-image-pane').style.display = type === 'image' ? '' : 'none';
  renderWatermarkPreview();
}

function setWmPosition(pos) {
  wmPosition = pos;
  document.querySelectorAll('#wm-pos-grid button').forEach(b => {
    b.classList.toggle('active', b.dataset.pos === pos);
  });
  document.getElementById('wm-tile-spacing-field').style.display = pos === 'tiled' ? '' : 'none';
  // Custom mode adds a click handler + crosshair cursor on the preview
  document.getElementById('wm-preview').classList.toggle('custom-mode', pos === 'custom');
  renderWatermarkPreview();
}

function setWmScope(s) {
  wmScope = s;
  document.getElementById('wm-scope-current').classList.toggle('active', s === 'current');
  document.getElementById('wm-scope-all').classList.toggle('active', s === 'all');
  const btn = document.getElementById('wm-apply-btn');
  btn.textContent = s === 'current' ? 'Apply to current slide'
                                    : `Apply to all ${NUM_SLIDES} slides`;
}

function setWmRotation(deg) {
  document.getElementById('wm-rot').value = deg;
  document.getElementById('wm-rot-val').textContent = deg + '°';
  renderWatermarkPreview();
}

// ── Live preview ────────────────────────────────────────────────────────
// Computes the CSS overlay properties that approximate what the server will
// render. We render at the actual preview-pane dimensions, deriving font/img
// sizes from the same fractions the backend uses (font_scale × slide_width,
// scale × slide_width, etc.).

function renderWatermarkPreview() {
  const preview = document.getElementById('wm-preview');
  if (!preview || !preview.offsetWidth) return;
  const pw = preview.offsetWidth, ph = preview.offsetHeight;

  const textEl  = document.getElementById('wm-preview-text');
  const imgEl   = document.getElementById('wm-preview-img');
  const tileWrap = document.getElementById('wm-preview-tile-wrap');
  const marker  = document.getElementById('wm-preview-marker');

  // Hide everything; the right pane shows itself below
  textEl.style.display = 'none';
  imgEl.style.display  = 'none';
  tileWrap.classList.remove('show');
  tileWrap.innerHTML = '';
  marker.classList.remove('show');

  const opacityPct = +document.getElementById('wm-op').value;
  const opacity    = opacityPct / 100;

  if (wmType === 'text') {
    const text = document.getElementById('wm-text').value || ' ';
    const color = document.getElementById('wm-color').value;
    const fontScale = (+document.getElementById('wm-fs').value) / 100;
    const rotation = +document.getElementById('wm-rot').value;
    const tileSpacing = +document.getElementById('wm-ts').value;
    const fontSizePx = Math.max(8, Math.round(pw * fontScale));

    const baseCss = `
      font-size:${fontSizePx}px;
      color:${color};
      opacity:${opacity};
      transform-origin:50% 50%;
    `;

    if (wmPosition === 'tiled') {
      tileWrap.classList.add('show');
      // Measure a single stamp to figure out tile step
      const probe = document.createElement('span');
      probe.style.cssText = baseCss + 'position:absolute;visibility:hidden;white-space:nowrap';
      probe.textContent = text;
      tileWrap.appendChild(probe);
      const tw = probe.offsetWidth;
      const th = probe.offsetHeight;
      probe.remove();
      const stepX = Math.max(20, (tw + 60) * Math.max(0.3, tileSpacing));
      const stepY = Math.max(20, (th + 40) * Math.max(0.3, tileSpacing));
      for (let ty = 0; ty < ph + stepY; ty += stepY) {
        for (let tx = 0; tx < pw + stepX; tx += stepX) {
          const t = document.createElement('div');
          t.className = 'tile';
          t.textContent = text;
          t.style.cssText = baseCss + `left:${tx}px;top:${ty}px;transform:rotate(${rotation}deg)`;
          tileWrap.appendChild(t);
        }
      }
    } else {
      textEl.style.display = '';
      textEl.textContent = text;
      const { leftPct, topPct, originPct, anchor } = _positionToCss(wmPosition, pw, ph);
      textEl.style.cssText = baseCss + `
        left:${leftPct};
        top:${topPct};
        transform:translate(${anchor.tx}, ${anchor.ty}) rotate(${rotation}deg);
      `;
      if (wmPosition === 'custom') _showMarker(marker, pw, ph);
    }
  } else {
    // image
    const imgInput = document.getElementById('wm-image-input');
    const file = imgInput && imgInput.files[0];
    if (!file) {
      // Nothing to preview yet
      return;
    }
    const scale = (+document.getElementById('wm-scale').value) / 100;
    const targetW = Math.max(8, Math.round(pw * scale));
    if (wmPosition === 'tiled') {
      tileWrap.classList.add('show');
      // Tile via repeating absolute-positioned images of the chosen file
      const objectUrl = URL.createObjectURL(file);
      const probe = new Image();
      probe.onload = () => {
        const aspect = probe.height / Math.max(1, probe.width);
        const tw = targetW, th = Math.max(8, Math.round(targetW * aspect));
        const stepX = tw + 60;
        const stepY = th + 40;
        for (let ty = 0; ty < ph + stepY; ty += stepY) {
          for (let tx = 0; tx < pw + stepX; tx += stepX) {
            const i = document.createElement('img');
            i.src = objectUrl;
            i.style.cssText = `position:absolute;left:${tx}px;top:${ty}px;width:${tw}px;height:${th}px;opacity:${opacity}`;
            tileWrap.appendChild(i);
          }
        }
      };
      probe.src = objectUrl;
    } else {
      imgEl.style.display = '';
      imgEl.src = URL.createObjectURL(file);
      imgEl.style.width = targetW + 'px';
      imgEl.style.height = 'auto';
      imgEl.style.opacity = opacity;
      const { leftPct, topPct, anchor } = _positionToCss(wmPosition, pw, ph);
      imgEl.style.left = leftPct;
      imgEl.style.top = topPct;
      imgEl.style.transform = `translate(${anchor.tx}, ${anchor.ty})`;
      if (wmPosition === 'custom') _showMarker(marker, pw, ph);
    }
  }
}

function _positionToCss(pos, pw, ph) {
  // Returns { leftPct, topPct, anchor:{tx,ty} } so CSS transform-translate
  // anchors the overlay at the correct corner/centre.
  const pad = 4; // % padding from the edges, matching the 20px ~ 2% from server
  switch (pos) {
    case 'top-left':     return { leftPct: pad + '%', topPct: pad + '%',         anchor: { tx: '0', ty: '0' } };
    case 'top-right':    return { leftPct: (100 - pad) + '%', topPct: pad + '%', anchor: { tx: '-100%', ty: '0' } };
    case 'bottom-left':  return { leftPct: pad + '%', topPct: (100 - pad) + '%', anchor: { tx: '0', ty: '-100%' } };
    case 'bottom-right': return { leftPct: (100 - pad) + '%', topPct: (100 - pad) + '%', anchor: { tx: '-100%', ty: '-100%' } };
    case 'custom':       return { leftPct: (wmCustomX * 100) + '%', topPct: (wmCustomY * 100) + '%', anchor: { tx: '-50%', ty: '-50%' } };
    case 'center':
    default:             return { leftPct: '50%', topPct: '50%', anchor: { tx: '-50%', ty: '-50%' } };
  }
}

function _showMarker(marker, pw, ph) {
  marker.style.left = (wmCustomX * 100) + '%';
  marker.style.top  = (wmCustomY * 100) + '%';
  marker.classList.add('show');
}

let _wmBound = false;
function _bindWatermarkControls() {
  if (_wmBound) return;
  _wmBound = true;
  const wire = (id, valId, fmt) => {
    const el = document.getElementById(id);
    const lbl = document.getElementById(valId);
    if (!el || !lbl) return;
    const update = () => {
      lbl.textContent = fmt(el.value);
      renderWatermarkPreview();
    };
    el.addEventListener('input', update);
    update();
  };
  wire('wm-fs',    'wm-fs-val',    v => (+v).toFixed(1) + '%');
  wire('wm-rot',   'wm-rot-val',   v => v + '°');
  wire('wm-op',    'wm-op-val',    v => v + '%');
  wire('wm-ts',    'wm-ts-val',    v => (+v).toFixed(1) + '×');
  wire('wm-scale', 'wm-scale-val', v => v + '%');
  // Free-text controls also re-render the preview
  const txt = document.getElementById('wm-text');
  if (txt) txt.addEventListener('input', renderWatermarkPreview);
  const col = document.getElementById('wm-color');
  if (col) col.addEventListener('input', renderWatermarkPreview);
  // Position buttons
  document.querySelectorAll('#wm-pos-grid button').forEach(b => {
    b.addEventListener('click', () => setWmPosition(b.dataset.pos));
  });
  // Click-to-place on the preview (only when Custom mode is active)
  const preview = document.getElementById('wm-preview');
  if (preview) {
    preview.addEventListener('click', e => {
      if (wmPosition !== 'custom') return;
      const r = preview.getBoundingClientRect();
      wmCustomX = Math.max(0, Math.min(1, (e.clientX - r.left) / r.width));
      wmCustomY = Math.max(0, Math.min(1, (e.clientY - r.top)  / r.height));
      renderWatermarkPreview();
    });
  }
  // Initial scope label
  setWmScope(wmScope);
  // Re-render preview when window resizes (preview pane changes size)
  window.addEventListener('resize', renderWatermarkPreview);
}

function previewWmImage(input) {
  if (!input.files[0]) return;
  const f = input.files[0];
  document.getElementById('wm-image-name').textContent = f.name;
  const url = URL.createObjectURL(f);
  document.getElementById('wm-image-preview').src = url;
  document.getElementById('wm-image-preview-wrap').style.display = '';
  // If user is on the Image tab, also refresh the main preview
  renderWatermarkPreview();
}

async function applyWatermark() {
  if (NUM_SLIDES === 0) { showToast('No slides loaded', 'info'); return; }

  if (wmType === 'image') {
    const file = document.getElementById('wm-image-input').files[0];
    if (!file) { showToast('Pick an image first', 'error'); return; }
    const form = new FormData();
    form.append('image', file);
    form.append('opacity', (+document.getElementById('wm-op').value) / 100);
    form.append('scale',   (+document.getElementById('wm-scale').value) / 100);
    form.append('position', wmPosition);
    form.append('scope', wmScope);
    form.append('slide_num', currentSlide);
    if (wmPosition === 'custom') {
      form.append('custom_x', wmCustomX);
      form.append('custom_y', wmCustomY);
    }
    showLoading(wmScope === 'all'
      ? `Stamping image on ${NUM_SLIDES} slide(s)...`
      : 'Stamping image on this slide...');
    try {
      const resp = await fetch('/api/watermark-image', { method: 'POST', body: form });
      const data = await resp.json();
      hideLoading();
      if (data.ok) {
        closeWatermarkModal();
        reloadAllSlides();
        refreshUndoState();
        showToast(`Watermark applied to ${data.count} slide(s)`, 'success');
      } else {
        showToast(data.error || 'Watermark failed', 'error');
      }
    } catch (e) {
      hideLoading();
      showToast('Watermark error: ' + e.message, 'error');
    }
    return;
  }

  // text
  const text = document.getElementById('wm-text').value.trim();
  if (!text) { showToast('Enter watermark text', 'error'); return; }
  const body = {
    text,
    color:        document.getElementById('wm-color').value,
    opacity:      (+document.getElementById('wm-op').value) / 100,
    font_scale:   (+document.getElementById('wm-fs').value) / 100,
    rotation:     +document.getElementById('wm-rot').value,
    tile_spacing: +document.getElementById('wm-ts').value,
    position:     wmPosition,
    scope:        wmScope,
    slide_num:    currentSlide,
  };
  if (wmPosition === 'custom') {
    body.custom_x = wmCustomX;
    body.custom_y = wmCustomY;
  }
  showLoading(wmScope === 'all'
    ? `Stamping text on ${NUM_SLIDES} slide(s)...`
    : 'Stamping text on this slide...');
  try {
    const resp = await fetch('/api/watermark', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const data = await resp.json();
    hideLoading();
    if (data.ok) {
      closeWatermarkModal();
      reloadAllSlides();
      refreshUndoState();
      showToast(`Watermark applied to ${data.count} slide(s)`, 'success');
    } else {
      showToast(data.error || 'Watermark failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Watermark error: ' + e.message, 'error');
  }
}

// ── Per-slide actions (delete / duplicate / png / print) ────────────────
async function deleteSlide() {
  if (NUM_SLIDES === 0) return;
  if (NUM_SLIDES === 1) {
    showToast("Can't delete the only slide — use Reset to clear the deck", 'info');
    return;
  }
  if (!confirm(`Delete slide ${currentSlide}?\n\nThis is reversible via Ctrl+Z.`)) return;
  showLoading('Deleting slide...');
  try {
    const resp = await fetch(`/api/slide/${currentSlide}/delete`, { method: 'POST' });
    const data = await resp.json();
    hideLoading();
    if (data.ok) {
      // Keep view near the deleted position (clamp)
      const newCurrent = Math.min(currentSlide, data.num_slides);
      try { sessionStorage.setItem('lite_current_slide', newCurrent); } catch (e) {}
      showToast(`Slide deleted. ${data.num_slides} slides remain.`, 'success');
      setTimeout(() => location.reload(), 400);
    } else {
      showToast(data.error || 'Delete failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Delete error: ' + e.message, 'error');
  }
}

async function duplicateSlide() {
  if (NUM_SLIDES === 0) return;
  showLoading('Duplicating slide...');
  try {
    const resp = await fetch(`/api/slide/${currentSlide}/duplicate`, { method: 'POST' });
    const data = await resp.json();
    hideLoading();
    if (data.ok) {
      // Land on the duplicate (one position after the source)
      try { sessionStorage.setItem('lite_current_slide', currentSlide + 1); } catch (e) {}
      showToast(`Slide duplicated. ${data.num_slides} slides total.`, 'success');
      setTimeout(() => location.reload(), 400);
    } else {
      showToast(data.error || 'Duplicate failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Duplicate error: ' + e.message, 'error');
  }
}

function downloadSlidePNG() {
  if (NUM_SLIDES === 0) {
    showToast('No slide to download', 'info');
    return;
  }
  const a = document.createElement('a');
  a.href = `/api/slide/${currentSlide}/download.png`;
  a.download = `slide-${String(currentSlide).padStart(2, '0')}.png`;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function printDeck() {
  if (NUM_SLIDES === 0) {
    showToast('No slides to print', 'info');
    return;
  }
  // Build a single-page-per-slide print document in a new window. This sidesteps
  // the complexity of overriding the current page's layout with print CSS and
  // gives a clean, deterministic print output.
  const w = window.open('', '_blank');
  if (!w) {
    showToast('Pop-up blocked — allow pop-ups to print', 'error');
    return;
  }
  let html = `<!doctype html><html><head><meta charset="utf-8">
    <title>SlideCraft Lite — Print</title>
    <style>
      @page { size: landscape; margin: 0; }
      body { margin: 0; background: white; }
      .slide {
        width: 100vw; height: 100vh;
        display: flex; align-items: center; justify-content: center;
        page-break-after: always; overflow: hidden;
      }
      .slide:last-child { page-break-after: auto; }
      .slide img { max-width: 100%; max-height: 100%; object-fit: contain; }
    </style></head><body>`;
  for (let i = 1; i <= NUM_SLIDES; i++) {
    const pad = String(i).padStart(2, '0');
    html += `<div class="slide"><img src="${location.origin}/sl/slide-${pad}.jpg"></div>`;
  }
  html += `<script>
    let loaded = 0, total = ${NUM_SLIDES};
    document.querySelectorAll('img').forEach(img => {
      img.onload = img.onerror = () => {
        if (++loaded === total) setTimeout(() => window.print(), 250);
      };
    });
  </` + `script></body></html>`;
  w.document.open();
  w.document.write(html);
  w.document.close();
}

// ── Deck info (filename in header) ───────────────────────────────────────
async function refreshDeckInfo() {
  try {
    const resp = await fetch('/api/deck/info');
    const data = await resp.json();
    const pill = document.getElementById('deck-name');
    const txt = document.getElementById('deck-name-text');
    if (!pill || !txt) return;
    if (data.name) {
      txt.textContent = data.name;
      pill.title = `${data.name} — ${data.num_slides} slide${data.num_slides === 1 ? '' : 's'}`;
      pill.classList.add('has-deck');
    } else {
      txt.textContent = 'No deck loaded';
      pill.title = 'No deck loaded';
      pill.classList.remove('has-deck');
    }
  } catch (e) {}
}

// ── Jump-to-slide input ──────────────────────────────────────────────────
function _wireSlideNumInput() {
  const input = document.getElementById('slide-num-input');
  if (!input) return;
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      const n = parseInt(input.value, 10);
      if (!isNaN(n)) gotoSlide(n);
      input.blur();
    } else if (e.key === 'Escape') {
      input.value = currentSlide;
      input.blur();
    }
  });
  input.addEventListener('blur', () => {
    const n = parseInt(input.value, 10);
    if (isNaN(n) || n < 1 || n > NUM_SLIDES) input.value = currentSlide;
  });
}

// ── Drag-and-drop reorder on thumbnails ─────────────────────────────────
let _dragSrcNum = null;

function initThumbDrag() {
  document.querySelectorAll('.thumb-item').forEach(thumb => {
    thumb.draggable = true;
    thumb.addEventListener('dragstart', e => {
      _dragSrcNum = parseInt(thumb.id.replace('thumb-', ''), 10);
      e.dataTransfer.effectAllowed = 'move';
      // Need to set SOME data for Firefox to fire drop, but use a private MIME
      // so our existing file-drop overlay handler ignores it.
      e.dataTransfer.setData('application/slidecraft-thumb', String(_dragSrcNum));
      thumb.classList.add('dragging');
    });
    thumb.addEventListener('dragend', () => {
      thumb.classList.remove('dragging');
      document.querySelectorAll('.thumb-item').forEach(t => {
        t.classList.remove('drag-over-top', 'drag-over-bottom');
      });
      _dragSrcNum = null;
    });
    thumb.addEventListener('dragover', e => {
      if (_dragSrcNum === null) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      const r = thumb.getBoundingClientRect();
      const above = e.clientY < r.top + r.height / 2;
      thumb.classList.toggle('drag-over-top', above);
      thumb.classList.toggle('drag-over-bottom', !above);
    });
    thumb.addEventListener('dragleave', () => {
      thumb.classList.remove('drag-over-top', 'drag-over-bottom');
    });
    thumb.addEventListener('drop', async e => {
      if (_dragSrcNum === null) return;
      e.preventDefault();
      const r = thumb.getBoundingClientRect();
      const above = e.clientY < r.top + r.height / 2;
      const targetNum = parseInt(thumb.id.replace('thumb-', ''), 10);
      const srcNum = _dragSrcNum;
      _dragSrcNum = null;
      document.querySelectorAll('.thumb-item').forEach(t => {
        t.classList.remove('drag-over-top', 'drag-over-bottom', 'dragging');
      });
      if (srcNum === targetNum) return;

      // Build new order: take current 1..N, remove srcNum, insert before/after target
      const order = [];
      for (let i = 1; i <= NUM_SLIDES; i++) order.push(i);
      const [moved] = order.splice(srcNum - 1, 1);
      // After removal, target's new index shifts down by 1 if target was after src
      let insertAt = order.indexOf(targetNum);
      if (insertAt === -1) return;
      if (!above) insertAt++;
      order.splice(insertAt, 0, moved);
      if (order.every((v, i) => v === i + 1)) return; // identity reorder

      showLoading('Reordering...');
      try {
        const resp = await fetch('/api/reorder', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ order }),
        });
        hideLoading();
        const data = await resp.json();
        if (data.ok) {
          // After reorder, the dragged slide lives at its new index
          const newIdx = order.indexOf(moved) + 1;
          try { sessionStorage.setItem('lite_current_slide', newIdx); } catch (e) {}
          showToast(`Reordered to ${NUM_SLIDES} slides`, 'success');
          setTimeout(() => location.reload(), 350);
        } else {
          showToast(data.error || 'Reorder failed', 'error');
        }
      } catch (err) {
        hideLoading();
        showToast('Reorder error: ' + err.message, 'error');
      }
    });
  });
}


// ══════════════════════════════════════════════════════════════════════
// Dark mode toggle (#15)
// ══════════════════════════════════════════════════════════════════════
function applyTheme(t) {
  document.documentElement.setAttribute('data-theme', t);
  const sun = document.getElementById('theme-icon-sun');
  const moon = document.getElementById('theme-icon-moon');
  if (sun && moon) {
    sun.style.display  = t === 'dark' ? 'none' : '';
    moon.style.display = t === 'dark' ? '' : 'none';
  }
}
function toggleTheme() {
  const cur = document.documentElement.getAttribute('data-theme') || 'light';
  const next = cur === 'dark' ? 'light' : 'dark';
  try { localStorage.setItem('sc_theme', next); } catch (e) {}
  applyTheme(next);
  showToast(`${next === 'dark' ? 'Dark' : 'Light'} theme`, 'info', 1200);
}
applyTheme(localStorage.getItem('sc_theme') || 'dark');

// ══════════════════════════════════════════════════════════════════════
// Region eraser (#2): drag a box on the slide to clone-stamp it out
// ══════════════════════════════════════════════════════════════════════
let _regionActive = false;
let _regionDrag = null;

function toggleRegionErase() {
  if (NUM_SLIDES === 0) {
    showToast('Upload a PPTX first', 'info');
    return;
  }
  _regionActive = !_regionActive;
  const layer = document.getElementById('region-erase-layer');
  const btn = document.getElementById('btn-region-erase');
  const hint = document.getElementById('region-hint');
  if (_regionActive) {
    layer?.classList.add('active');
    btn?.classList.add('accent');
    hint?.classList.add('show');
  } else {
    layer?.classList.remove('active');
    btn?.classList.remove('accent');
    hint?.classList.remove('show');
    const box = document.getElementById('region-erase-box');
    if (box) box.style.display = 'none';
    _regionDrag = null;
  }
}

function _setupRegionEraser() {
  const layer = document.getElementById('region-erase-layer');
  const box = document.getElementById('region-erase-box');
  if (!layer || !box) return;

  layer.addEventListener('mousedown', e => {
    if (!_regionActive) return;
    const r = layer.getBoundingClientRect();
    _regionDrag = {
      startX: (e.clientX - r.left) / r.width,
      startY: (e.clientY - r.top)  / r.height,
      curX:   (e.clientX - r.left) / r.width,
      curY:   (e.clientY - r.top)  / r.height,
    };
    box.style.display = 'block';
    _drawRegionBox();
    e.preventDefault();
  });
  layer.addEventListener('mousemove', e => {
    if (!_regionActive || !_regionDrag) return;
    const r = layer.getBoundingClientRect();
    _regionDrag.curX = Math.max(0, Math.min(1, (e.clientX - r.left) / r.width));
    _regionDrag.curY = Math.max(0, Math.min(1, (e.clientY - r.top)  / r.height));
    _drawRegionBox();
  });
  layer.addEventListener('mouseup', async e => {
    if (!_regionActive || !_regionDrag) return;
    const d = _regionDrag;
    const x = Math.min(d.startX, d.curX);
    const y = Math.min(d.startY, d.curY);
    const w = Math.abs(d.curX - d.startX);
    const h = Math.abs(d.curY - d.startY);
    _regionDrag = null;
    box.style.display = 'none';
    if (w < 0.005 || h < 0.005) return;
    showLoading('Erasing region...');
    try {
      const resp = await fetch('/api/region-erase', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ slide_num: currentSlide, x, y, w, h }),
      });
      hideLoading();
      const data = await resp.json();
      if (data.ok) {
        reloadAllSlides();
        refreshUndoState();
        showToast('Region erased', 'success');
      } else {
        showToast(data.error || 'Erase failed', 'error');
      }
    } catch (err) {
      hideLoading();
      showToast('Erase error: ' + err.message, 'error');
    }
  });
}
function _drawRegionBox() {
  const layer = document.getElementById('region-erase-layer');
  const box = document.getElementById('region-erase-box');
  if (!_regionDrag || !layer || !box) return;
  const r = layer.getBoundingClientRect();
  const d = _regionDrag;
  const x = Math.min(d.startX, d.curX) * r.width;
  const y = Math.min(d.startY, d.curY) * r.height;
  const w = Math.abs(d.curX - d.startX) * r.width;
  const h = Math.abs(d.curY - d.startY) * r.height;
  box.style.left   = x + 'px';
  box.style.top    = y + 'px';
  box.style.width  = w + 'px';
  box.style.height = h + 'px';
}
_setupRegionEraser();

// ══════════════════════════════════════════════════════════════════════
// Recent decks (#16)
// ══════════════════════════════════════════════════════════════════════
async function toggleRecentMenu(event) {
  if (event) event.stopPropagation();
  const menu = document.getElementById('recent-menu');
  if (!menu) return;
  if (menu.classList.contains('open')) {
    menu.classList.remove('open');
    return;
  }
  await renderRecent();
  menu.classList.add('open');
}
async function renderRecent() {
  const menu = document.getElementById('recent-menu');
  if (!menu) return;
  let items = [];
  try {
    const r = await fetch('/api/recent');
    items = (await r.json()).items || [];
  } catch (e) {}
  if (!items.length) {
    menu.innerHTML = '<div class="recent-empty">No recent decks yet.</div>';
    return;
  }
  menu.innerHTML = items.map(it => `
    <div class="recent-item">
      <button class="recent-load" onclick="loadRecent('${encodeURIComponent(it.file)}')" title="Load this deck">
        <div class="recent-name">${_escape(it.name || it.file)}</div>
        <div class="recent-meta">${_humanAgo(it.uploaded)}</div>
      </button>
      <button class="recent-del" onclick="deleteRecent('${encodeURIComponent(it.file)}')" title="Remove from recent">×</button>
    </div>
  `).join('');
}
function _escape(s) {
  return String(s || '').replace(/[&<>"']/g,
    c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}
async function loadRecent(encFile) {
  const file = decodeURIComponent(encFile);
  showLoading(`Loading ${file}...`);
  try {
    const r = await fetch('/api/recent/load', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file }),
    });
    const data = await r.json();
    hideLoading();
    if (data.ok) {
      showToast(`Loaded ${data.num_slides} slide(s)`, 'success');
      setTimeout(() => location.reload(), 500);
    } else {
      showToast(data.error || 'Load failed', 'error');
    }
  } catch (e) {
    hideLoading();
    showToast('Load error: ' + e.message, 'error');
  }
}
async function deleteRecent(encFile) {
  const file = decodeURIComponent(encFile);
  try {
    await fetch('/api/recent/delete', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file }),
    });
    await renderRecent();
  } catch (e) {}
}
// Close recent menu when clicking outside
document.addEventListener('click', e => {
  const menu = document.getElementById('recent-menu');
  if (!menu) return;
  if (menu.classList.contains('open') && !e.target.closest('.recent-dropdown')) {
    menu.classList.remove('open');
  }
});
