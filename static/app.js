/* ── Redacter Frontend ───────────────────────────────────────────────────────
   Vanilla JS, no dependencies. Works with Flask backend on localhost:5050.
   ─────────────────────────────────────────────────────────────────────────── */

'use strict';

// ── State ──────────────────────────────────────────────────────────────────
const state = {
  fileId: null,
  filename: null,
  ext: null,
  pageCount: 1,
  currentPage: 0,
  terms: [],          // manual terms
  matches: [],        // [{page, text, rect}]
  scanned: false,
};

// ── Elements ──────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const dropZone       = $('dropZone');
const fileInput      = $('fileInput');
const fileBanner     = $('fileBanner');
const fileBannerName = $('fileBannerName');
const fileBannerMeta = $('fileBannerMeta');
const fileBannerIcon = $('fileBannerIcon');
const clearFileBtn   = $('clearFileBtn');
const workspace      = $('workspace');
const termInput      = $('termInput');
const addTermBtn     = $('addTermBtn');
const termChips      = $('termChips');
const scanBtn        = $('scanBtn');
const redactBtn      = $('redactBtn');
const statusText     = $('statusText');
const statusDot      = document.querySelector('.status-dot');
const prevPage       = $('prevPage');
const nextPage       = $('nextPage');
const pageIndicator  = $('pageIndicator');
const previewPlaceholder  = $('previewPlaceholder');
const previewImageWrap    = $('previewImageWrap');
const previewImg          = $('previewImg');
const overlayCanvas       = $('overlayCanvas');
const matchBadge          = $('matchBadge');
const matchBadgeText      = $('matchBadgeText');
const matchListWrap       = $('matchListWrap');
const matchList           = $('matchList');
const clearMatchesBtn     = $('clearMatchesBtn');
const toast               = $('toast');
const toastText           = $('toastText');
const toastDownloadLink   = $('toastDownloadLink');
const errorToast          = $('errorToast');
const errorToastText      = $('errorToastText');
const nlpPill             = $('nlpPill');
const nlpLabel            = $('nlpLabel');
const namesEngineSub      = $('namesEngineSub');

// ── Init ──────────────────────────────────────────────────────────────────
(async function init() {
  try {
    const res = await fetch('/status');
    const data = await res.json();
    if (data.nlp) {
      nlpPill.classList.add('active');
      nlpLabel.textContent = 'spaCy NLP';
      if (namesEngineSub) namesEngineSub.textContent = '(spaCy)';
    } else {
      nlpPill.classList.add('fallback');
      nlpLabel.textContent = 'Regex mode';
      if (namesEngineSub) namesEngineSub.textContent = '(regex)';
    }
  } catch (e) {
    nlpLabel.textContent = 'Offline?';
  }

  // Style radio
  document.querySelectorAll('input[name="style"]').forEach(radio => {
    radio.addEventListener('change', () => {
      document.querySelectorAll('.style-option').forEach(el => el.classList.remove('selected'));
      radio.closest('.style-option').classList.add('selected');
    });
  });
})();

// ── Status helpers ────────────────────────────────────────────────────────
function setStatus(msg, mode = 'idle') {
  statusText.textContent = msg;
  statusDot.className = 'status-dot ' + mode;
}

function showError(msg) {
  errorToastText.textContent = msg;
  errorToast.classList.remove('hidden');
  setTimeout(() => errorToast.classList.add('hidden'), 6000);
}

function showSuccess(msg, downloadUrl, filename) {
  toastText.textContent = msg;
  if (downloadUrl) {
    toastDownloadLink.href = downloadUrl;
    toastDownloadLink.textContent = `Download ${filename || 'file'}`;
    toastDownloadLink.classList.remove('hidden');
  } else {
    toastDownloadLink.classList.add('hidden');
  }
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), 12000);
}

// ── Drag & Drop ───────────────────────────────────────────────────────────
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});
dropZone.addEventListener('click', e => {
  if (e.target === dropZone || e.target.closest('.drop-icon') || e.target.classList.contains('drop-primary') || e.target.classList.contains('drop-hint')) {
    fileInput.click();
  }
});
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
  fileInput.value = '';
});
clearFileBtn.addEventListener('click', resetAll);

// ── Upload ────────────────────────────────────────────────────────────────
async function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const allowed = ['pdf','docx','doc','xlsx','pptx','txt','png','jpg','jpeg','tiff','tif'];
  if (!allowed.includes(ext)) {
    showError(`Unsupported file type: .${ext}. Accepted: PDF, DOCX, XLSX, PPTX, TXT, PNG, JPG, TIFF.`);
    return;
  }

  setStatus('Uploading… (large files may take a moment)', 'working');
  scanBtn.disabled = true;
  redactBtn.disabled = true;

  const fd = new FormData();
  fd.append('file', file);

  // Use XMLHttpRequest — more reliable than fetch for large binary uploads on iOS Safari
  const data = await new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/upload');
    xhr.timeout = 300000; // 5 min timeout for large files
    xhr.upload.onprogress = e => {
      if (e.lengthComputable) {
        const pct = Math.round((e.loaded / e.total) * 100);
        setStatus(`Uploading… ${pct}%`, 'working');
      }
    };
    xhr.onload = () => {
      try {
        const json = JSON.parse(xhr.responseText);
        if (xhr.status >= 400) reject(new Error(json.error || 'Upload failed'));
        else resolve(json);
      } catch (e) {
        reject(new Error(`Server error (${xhr.status})`));
      }
    };
    xhr.onerror = () => reject(new Error('Network error — check your connection'));
    xhr.ontimeout = () => reject(new Error('Upload timed out — file may be too large for the connection'));
    xhr.send(fd);
  });

  try {

    state.fileId    = data.file_id;
    state.filename  = data.filename;
    state.ext       = data.ext;
    state.pageCount = data.page_count;
    state.currentPage = 0;
    state.matches   = [];
    state.scanned   = false;

    fileBannerName.textContent = data.filename;
    fileBannerMeta.textContent = data.ext === '.pdf'
      ? `PDF · ${data.page_count} page${data.page_count !== 1 ? 's' : ''}`
      : 'Word Document (.docx)';
    fileBannerIcon.textContent = data.ext === '.pdf' ? '📄' : '📝';
    fileBanner.classList.remove('hidden');
    workspace.classList.remove('hidden');

    updatePageNav();
    setStatus('File ready', 'done');
    scanBtn.disabled = false;

    // Load preview
    if (data.ext === '.pdf') {
      loadPreview(0);
    } else {
      // DOCX — no page preview
      previewPlaceholder.innerHTML = `<svg width="48" height="48" viewBox="0 0 48 48" fill="none" opacity="0.4"><rect x="6" y="4" width="30" height="40" rx="3" stroke="#8B949E" stroke-width="1.5"/><rect x="12" y="14" width="18" height="2.5" rx="1.25" fill="#8B949E"/><rect x="12" y="20" width="18" height="2.5" rx="1.25" fill="#8B949E"/><rect x="12" y="26" width="12" height="2.5" rx="1.25" fill="#8B949E"/></svg><p>Word file · Preview scan results below</p>`;
      previewPlaceholder.classList.remove('hidden');
      previewImageWrap.classList.add('hidden');
    }
    clearMatchDisplay();

  } catch (err) {
    setStatus('Upload failed', 'error');
    showError(err.message);
  }
}

// ── Preview ───────────────────────────────────────────────────────────────
function loadPreview(page) {
  if (!state.fileId || state.ext !== '.pdf') return;
  previewPlaceholder.classList.add('hidden');
  previewImageWrap.classList.remove('hidden');

  // Shimmer while loading
  previewImg.style.opacity = '0.4';
  previewImg.src = `/preview/${state.fileId}/${page}?t=${Date.now()}`;
  previewImg.onload = () => {
    previewImg.style.opacity = '1';
    // Redraw overlay for current page
    drawOverlay();
  };
  previewImg.onerror = () => {
    previewImg.style.opacity = '1';
    previewPlaceholder.classList.remove('hidden');
  };
}

function updatePageNav() {
  const n = state.pageCount;
  const p = state.currentPage;
  pageIndicator.textContent = n > 1 ? `Page ${p + 1} / ${n}` : `Page 1 / 1`;
  prevPage.disabled = p === 0;
  nextPage.disabled = p >= n - 1;
}

prevPage.addEventListener('click', () => {
  if (state.currentPage > 0) {
    state.currentPage--;
    updatePageNav();
    loadPreview(state.currentPage);
  }
});
nextPage.addEventListener('click', () => {
  if (state.currentPage < state.pageCount - 1) {
    state.currentPage++;
    updatePageNav();
    loadPreview(state.currentPage);
  }
});

// ── Overlay canvas ────────────────────────────────────────────────────────
function drawOverlay() {
  const img = previewImg;
  const canvas = overlayCanvas;
  canvas.width  = img.naturalWidth  || img.clientWidth;
  canvas.height = img.naturalHeight || img.clientHeight;
  canvas.style.width  = img.clientWidth + 'px';
  canvas.style.height = img.clientHeight + 'px';

  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  // Scale factors: preview is rendered at 1.5x zoom server-side
  // canvas is at natural size, img element may be displayed smaller
  const scaleX = img.clientWidth  / (img.naturalWidth  || img.clientWidth);
  const scaleY = img.clientHeight / (img.naturalHeight || img.clientHeight);

  // Filter matches for current page
  const pageMatches = state.matches.filter(m => m.page === state.currentPage);

  // PDF preview rendered at zoom=1.5
  const PDF_ZOOM = 1.5;

  pageMatches.forEach(m => {
    const [x0, y0, x1, y1] = m.rect;
    // Convert PDF coords (at zoom 1.5) to canvas coords
    const cx0 = x0 * PDF_ZOOM * scaleX;
    const cy0 = y0 * PDF_ZOOM * scaleY;
    const cw  = (x1 - x0) * PDF_ZOOM * scaleX;
    const ch  = (y1 - y0) * PDF_ZOOM * scaleY;

    // Semi-transparent yellow highlight
    ctx.fillStyle = 'rgba(245, 158, 11, 0.35)';
    ctx.fillRect(cx0, cy0, cw, ch);
    ctx.strokeStyle = 'rgba(245, 158, 11, 0.7)';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(cx0, cy0, cw, ch);
  });
}

// Redraw overlay when image resizes
window.addEventListener('resize', drawOverlay);

// ── Terms ─────────────────────────────────────────────────────────────────
function addTerm(term) {
  term = term.trim();
  if (!term || state.terms.includes(term)) return;
  state.terms.push(term);
  renderTermChips();
  state.scanned = false;
}

function removeTerm(term) {
  state.terms = state.terms.filter(t => t !== term);
  renderTermChips();
  state.scanned = false;
}

function renderTermChips() {
  termChips.innerHTML = '';
  state.terms.forEach(term => {
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.innerHTML = `<span class="chip-text" title="${escHtml(term)}">${escHtml(term)}</span>
      <button class="chip-remove" aria-label="Remove ${escHtml(term)}">✕</button>`;
    chip.querySelector('.chip-remove').addEventListener('click', () => removeTerm(term));
    termChips.appendChild(chip);
  });
}

addTermBtn.addEventListener('click', () => {
  addTerm(termInput.value);
  termInput.value = '';
  termInput.focus();
});
termInput.addEventListener('keydown', e => {
  if (e.key === 'Enter') {
    addTerm(termInput.value);
    termInput.value = '';
  }
});

// ── Flags ─────────────────────────────────────────────────────────────────
function getFlags() {
  const flags = {};
  document.querySelectorAll('.flag-check').forEach(cb => {
    flags[cb.dataset.flag] = cb.checked;
  });
  return flags;
}

document.querySelectorAll('.flag-check').forEach(cb => {
  cb.addEventListener('change', () => { state.scanned = false; });
});

// ── Scan ──────────────────────────────────────────────────────────────────
scanBtn.addEventListener('click', async () => {
  if (!state.fileId) return;
  const terms = [...state.terms];
  const flags = getFlags();

  if (terms.length === 0 && !Object.values(flags).some(Boolean)) {
    showError('Add at least one search term or enable an auto-detect option.');
    return;
  }

  setStatus('Scanning document…', 'working');
  scanBtn.disabled = true;
  redactBtn.disabled = true;

  try {
    const res = await fetch('/scan', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_id: state.fileId, terms, flags }),
    });
    if (!res.ok) throw new Error((await res.json()).error || 'Scan failed');
    const data = await res.json();

    state.matches = data.matches;
    state.scanned = true;

    const count = data.matches.length;
    const pageSet = new Set(data.matches.map(m => m.page));

    setStatus(`Found ${count} match${count !== 1 ? 'es' : ''} across ${pageSet.size} page${pageSet.size !== 1 ? 's' : ''}`, 'done');

    // Badge
    matchBadgeText.textContent = `${count} match${count !== 1 ? 'es' : ''}`;
    matchBadge.classList.toggle('hidden', count === 0);

    // Render match list
    renderMatchList(data.matches);

    // Draw overlay on current page
    if (state.ext === '.pdf') drawOverlay();

    redactBtn.disabled = count === 0;

  } catch (err) {
    setStatus('Scan error', 'error');
    showError(err.message);
    scanBtn.disabled = false;
  } finally {
    scanBtn.disabled = false;
  }
});

function renderMatchList(matches) {
  matchList.innerHTML = '';
  if (matches.length === 0) {
    matchListWrap.classList.add('hidden');
    return;
  }
  matchListWrap.classList.remove('hidden');

  // Group by text, show page numbers
  const byText = {};
  matches.forEach(m => {
    if (!byText[m.text]) byText[m.text] = new Set();
    byText[m.text].add(m.page + 1);
  });

  Object.entries(byText).forEach(([text, pages]) => {
    const chip = document.createElement('div');
    chip.className = 'match-chip';
    chip.title = `Found on page${pages.size > 1 ? 's' : ''} ${[...pages].join(', ')}`;
    chip.innerHTML = `<span>${escHtml(text)}</span><span class="match-page">p.${[...pages].join(',')}</span>`;
    // Click jumps to first page containing this match
    chip.addEventListener('click', () => {
      const firstPage = Math.min(...[...pages]) - 1;
      if (firstPage !== state.currentPage) {
        state.currentPage = firstPage;
        updatePageNav();
        loadPreview(firstPage);
      }
    });
    matchList.appendChild(chip);
  });
}

clearMatchesBtn.addEventListener('click', clearMatchDisplay);
function clearMatchDisplay() {
  state.matches = [];
  state.scanned = false;
  matchList.innerHTML = '';
  matchListWrap.classList.add('hidden');
  matchBadge.classList.add('hidden');
  redactBtn.disabled = true;
  if (overlayCanvas) {
    const ctx = overlayCanvas.getContext('2d');
    ctx.clearRect(0, 0, overlayCanvas.width, overlayCanvas.height);
  }
  setStatus('Ready', 'idle');
}

// ── Redact ────────────────────────────────────────────────────────────────
redactBtn.addEventListener('click', async () => {
  if (!state.fileId) return;

  // Auto-run preview first if nothing scanned yet
  if (state.matches.length === 0) {
    await scanBtn.onclick?.() || scanBtn.click();
    if (state.matches.length === 0) {
      showError('No matches found. Add terms or enable auto-detect options above.');
      return;
    }
  }

  const style = document.querySelector('input[name="style"]:checked').value;

  setStatus('Applying redactions…', 'working');
  redactBtn.disabled = true;
  scanBtn.disabled = true;

  try {
    const res = await fetch('/redact', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_id: state.fileId,
        matches: state.matches,
        style,
      }),
    });
    // /redact now returns the file directly as binary
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: 'Redaction failed' }));
      throw new Error(err.error || 'Redaction failed');
    }

    const dlName = res.headers.get('X-Output-Filename')
      || res.headers.get('Content-Disposition')?.match(/filename="?([^"]+)"?/)?.[1]
      || 'redacted_file.pdf';

    const blob = await res.blob();
    const dlUrl = URL.createObjectURL(blob);

    setStatus('Redaction complete ✓', 'done');
    showSuccess(
      `Redacted successfully — tap below to download.`,
      dlUrl,
      dlName,
    );

    // Trigger download
    const a = document.createElement('a');
    a.href = dlUrl;
    a.download = dlName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(dlUrl), 10000);

  } catch (err) {
    setStatus('Redaction failed', 'error');
    showError(err.message);
  } finally {
    redactBtn.disabled = false;
    scanBtn.disabled = false;
  }
});

// ── Reset ─────────────────────────────────────────────────────────────────
function resetAll() {
  state.fileId = null;
  state.filename = null;
  state.ext = null;
  state.pageCount = 1;
  state.currentPage = 0;
  state.terms = [];
  state.matches = [];
  state.scanned = false;

  fileBanner.classList.add('hidden');
  workspace.classList.add('hidden');
  previewPlaceholder.classList.remove('hidden');
  previewImageWrap.classList.add('hidden');
  previewImg.src = '';
  termChips.innerHTML = '';
  matchList.innerHTML = '';
  matchListWrap.classList.add('hidden');
  matchBadge.classList.add('hidden');
  scanBtn.disabled = true;
  redactBtn.disabled = true;
  setStatus('Ready', 'idle');
  toast.classList.add('hidden');
  document.querySelectorAll('.flag-check').forEach(cb => cb.checked = false);
}

// ── Utility ───────────────────────────────────────────────────────────────
function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
