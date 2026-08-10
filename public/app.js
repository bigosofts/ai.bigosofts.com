// ═══════════════════════════════════════════════════════════
// BigOSofts Slide Maker — Frontend Application Logic
// ═══════════════════════════════════════════════════════════

// Generate a unique session ID per browser tab (persists across refreshes)
const SESSION_ID = sessionStorage.getItem('bos_session_id') || (() => {
  const id = 'sess-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 9);
  sessionStorage.setItem('bos_session_id', id);
  return id;
})();

const textarea = document.getElementById('userInput');
const charCount = document.getElementById('charCount');

// Character counter
textarea.addEventListener('input', () => {
  charCount.textContent = textarea.value.length.toLocaleString();
});

// ═══════════════════════════════════════════════════════════
// COLOR PALETTE SELECTOR
// ═══════════════════════════════════════════════════════════

const selectedPalettes = new Set(['forestMoss']);
let selectedMode = 'light';

// Mode toggle
document.getElementById('modeToggle').addEventListener('click', (e) => {
  const btn = e.target.closest('.mode-btn');
  if (btn) switchMode(btn.dataset.mode);
});

function switchMode(mode) {
  selectedMode = mode;

  // Update toggle buttons
  document.querySelectorAll('.mode-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.mode === mode);
  });

  // Slide the toggle indicator
  document.getElementById('modeToggle').dataset.active = mode;

  // Show/hide palette swatches
  document.querySelectorAll('.pal').forEach(swatch => {
    swatch.style.display = swatch.dataset.mode === mode ? '' : 'none';
  });

}

function selectPalette(name) {
  if (selectedPalettes.has(name)) {
    if (selectedPalettes.size > 1) selectedPalettes.delete(name);
  } else if (selectedPalettes.size < 4) {
    selectedPalettes.add(name);
  } else {
    const summary = document.getElementById('paletteSelectionSummary');
    summary.textContent = 'Maximum of 4 palettes selected';
    summary.classList.add('limit-reached');
    setTimeout(() => summary.classList.remove('limit-reached'), 900);
    return;
  }
  document.querySelectorAll('.pal').forEach(s => {
    const isSelected = selectedPalettes.has(s.dataset.palette);
    s.classList.toggle('active', isSelected);
    s.setAttribute('aria-pressed', String(isSelected));
  });
  updatePaletteSummary();
}

document.getElementById('paletteGrid').addEventListener('click', (e) => {
  const swatch = e.target.closest('.pal');
  if (swatch) selectPalette(swatch.dataset.palette);
});

document.getElementById('clearPalettesBtn').addEventListener('click', () => {
  const activePalette = selectedPalettes.values().next().value || 'forestMoss';
  selectedPalettes.clear();
  selectedPalettes.add(activePalette);
  document.querySelectorAll('.pal').forEach(s => {
    const isSelected = selectedPalettes.has(s.dataset.palette);
    s.classList.toggle('active', isSelected);
    s.setAttribute('aria-pressed', String(isSelected));
  });
  updatePaletteSummary();
});

function updatePaletteSummary() {
  const count = selectedPalettes.size;
  document.getElementById('paletteSelectionSummary').textContent = `${count} palette${count === 1 ? '' : 's'} selected`;
  document.getElementById('clearPalettesBtn').disabled = count === 1;
  document.getElementById('btnContent').lastChild.textContent = count === 1 ? ' Generate presentation' : ` Generate ${count} variants`;
}

updatePaletteSummary();

// ═══════════════════════════════════════════════════════════
// IMAGE UPLOAD SECTION
// ═══════════════════════════════════════════════════════════

let imageEntryCounter = 0;
let uploadedImages = [];

// Toggle collapse
document.getElementById('imageUploadToggle').addEventListener('click', () => {
  const body = document.getElementById('imageUploadBody');
  const chevron = document.getElementById('chevronIcon');
  const isVisible = body.style.display !== 'none';
  body.style.display = isVisible ? 'none' : 'block';
  chevron.classList.toggle('open', !isVisible);
});

// Add image button
document.getElementById('addImageBtn').addEventListener('click', addImageEntry);

function addImageEntry() {
  const id = ++imageEntryCounter;
  const container = document.getElementById('imageEntries');

  const entry = document.createElement('div');
  entry.className = 'image-entry-row';
  entry.id = `image-entry-${id}`;
  entry.innerHTML = `
    <input type="file" accept="image/*" id="fileInput-${id}" style="max-width:150px;" />
    <input type="text" id="urlField-${id}" placeholder="Or paste URL" />
    <input type="text" id="descField-${id}" placeholder="Description (e.g. 'Company logo')" />
    <button type="button" class="remove-img-btn" onclick="removeImageEntry(${id})" title="Remove">✕</button>
  `;
  container.appendChild(entry);

  document.getElementById(`fileInput-${id}`).addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (file) await uploadFile(id, file);
  });

  document.getElementById(`urlField-${id}`).addEventListener('blur', (e) => {
    const url = e.target.value.trim();
    if (url && (url.startsWith('http://') || url.startsWith('https://'))) {
      updateImageRecord(id, { staticUrl: url, type: 'url' });
    }
  });
}

async function uploadFile(id, file) {
  const urlField = document.getElementById(`urlField-${id}`);
  urlField.value = 'Uploading...';

  const formData = new FormData();
  formData.append('image', file);

  try {
    // Send session ID header so server namespaces file to this user
    const response = await fetch('/api/upload-image', {
      method: 'POST',
      headers: { 'X-Session-ID': SESSION_ID },
      body: formData
    });
    const data = await response.json();
    if (data.success) {
      urlField.value = data.staticUrl;
      urlField.readOnly = true;
      updateImageRecord(id, { serverPath: data.filePath, staticUrl: data.staticUrl, type: 'upload' });
    } else {
      urlField.value = 'Error: ' + (data.error || 'Upload failed');
    }
  } catch (err) {
    urlField.value = 'Error: ' + err.message;
  }
}

function updateImageRecord(id, data) {
  uploadedImages = uploadedImages.filter(img => img.id !== id);
  uploadedImages.push({ id, ...data });
}

function removeImageEntry(id) {
  const el = document.getElementById(`image-entry-${id}`);
  if (el) {
    el.style.opacity = '0';
    el.style.transform = 'translateX(20px)';
    el.style.transition = 'all 0.25s ease';
    setTimeout(() => el.remove(), 250);
  }
  uploadedImages = uploadedImages.filter(img => img.id !== id);
}

function collectImageData() {
  const images = [];
  document.querySelectorAll('.image-entry-row').forEach(entry => {
    const id = parseInt(entry.id.replace('image-entry-', ''));
    const urlField = document.getElementById(`urlField-${id}`);
    const descField = document.getElementById(`descField-${id}`);
    const url = urlField ? urlField.value.trim() : '';
    const desc = descField ? descField.value.trim() : '';
    if (url && !url.startsWith('Error') && !url.startsWith('Uploading')) {
      const record = uploadedImages.find(img => img.id === id);
      images.push({
        url: url,
        description: desc || 'No description provided',
        type: record ? record.type : 'url',
        serverPath: record ? record.serverPath : null
      });
    }
  });
  return images;
}

// Reset image uploads
function resetImageUploads() {
  document.getElementById('imageEntries').innerHTML = '';
  uploadedImages = [];
  imageEntryCounter = 0;
  // Collapse the section
  document.getElementById('imageUploadBody').style.display = 'none';
  document.getElementById('chevronIcon').classList.remove('open');
}

// ═══════════════════════════════════════════════════════════
// SPA VIEW STATE MANAGEMENT
// ═══════════════════════════════════════════════════════════

function setViewState(stateId) {
  document.querySelectorAll('.state-panel').forEach(panel => {
    panel.classList.remove('active');
  });
  const panel = document.getElementById(stateId);
  if (panel) {
    panel.style.display = 'flex';
    // Force reflow then add active class for animation
    void panel.offsetWidth;
    panel.classList.add('active');
  }
}

function setProgress(percent, title, subtitle) {
  document.getElementById('progressBar').style.width = percent + '%';
  if (title) document.getElementById('progressTitle').textContent = title;
  if (subtitle) document.getElementById('progressSubtitle').textContent = subtitle;
}

// This is an activity feed, not private model reasoning. The backend currently
// returns one final response, so these messages describe the user-visible work
// being performed while that request is in progress.
const designActivities = [
  'Reading your topic and identifying the core narrative',
  'Matching your content to the most suitable slide layouts',
  'Planning the visual hierarchy, typography, and color balance',
  'Organizing ideas into a clear presentation storyline',
  'Preparing cards, charts, and visual details for the deck'
];
let activityTimer = null;

function setActivityMessage(message) {
  const activityMessage = document.getElementById('activityMessage');
  if (activityMessage) activityMessage.textContent = message;
}

function startActivityFeed() {
  stopActivityFeed();
  let activityIndex = 0;
  setActivityMessage(designActivities[activityIndex]);
  activityTimer = setInterval(() => {
    activityIndex = (activityIndex + 1) % designActivities.length;
    setActivityMessage(designActivities[activityIndex]);
  }, 2400);
}

function stopActivityFeed() {
  if (activityTimer) {
    clearInterval(activityTimer);
    activityTimer = null;
  }
}

function activateStep(stepNum) {
  for (let i = 1; i <= 3; i++) {
    const step = document.getElementById('step' + i);
    if (i < stepNum) {
      step.classList.remove('active');
      step.classList.add('done');
    } else if (i === stepNum) {
      step.classList.add('active');
      step.classList.remove('done');
    } else {
      step.classList.remove('active', 'done');
    }
  }
}

// Reads the server's real generation milestones. These are operational updates
// from the pipeline, not hidden model reasoning or a client-side timer.
async function readGenerationProgress(response) {
  if (!response.body) throw new Error('The server did not provide a progress stream.');

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let result = null;

  const handleEvent = (event) => {
    if (event.type === 'status') {
      activateStep(event.step);
      setProgress(event.progress, event.title, event.subtitle);
      setActivityMessage(event.activity);
    } else if (event.type === 'error') {
      throw new Error(event.error || 'Generation failed');
    } else if (event.type === 'result') {
      result = event;
    }
  };

  while (true) {
    const { value, done } = await reader.read();
    buffer += decoder.decode(value || new Uint8Array(), { stream: !done });
    const lines = buffer.split('\n');
    buffer = lines.pop();
    lines.filter(Boolean).forEach((line) => handleEvent(JSON.parse(line)));
    if (done) break;
  }

  if (buffer.trim()) handleEvent(JSON.parse(buffer));
  if (!result || !result.success) throw new Error(result?.error || 'Generation failed');
  return result;
}

function resetUI() {
  const btn = document.getElementById('generateBtn');
  btn.disabled = false;
  document.getElementById('btnContent').style.display = 'flex';
  document.getElementById('btnLoading').style.display = 'none';

  // Reset progress
  setProgress(0);
  activateStep(1);
  stopActivityFeed();
  setActivityMessage(designActivities[0]);

  // Reset images
  resetImageUploads();

  // Show idle state
  setViewState('state-idle');
}

function renderDownloads(files) {
  const downloadsList = document.getElementById('downloadsList');
  downloadsList.innerHTML = '';
  const safeFiles = files && files.length ? files : [];
  document.getElementById('successTitle').textContent = safeFiles.length > 1
    ? `${safeFiles.length} presentation variants ready`
    : 'Presentation ready';

  safeFiles.forEach((file) => {
    const link = document.createElement('a');
    link.className = 'download-variant';
    link.href = '/api/download/' + file.fileId;
    link.download = file.fileName || 'presentation.pptx';
    link.innerHTML = `
      <span class="variant-swatch" data-palette="${file.palette || ''}"></span>
      <span><strong>${(file.palette || 'Presentation').replace(/([A-Z])/g, ' $1').trim()}</strong><small>Editable .pptx</small></span>
      <span class="variant-download">Download ↓</span>`;
    downloadsList.appendChild(link);
  });
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ═══════════════════════════════════════════════════════════
// MAIN GENERATION FUNCTION
// ═══════════════════════════════════════════════════════════

async function generatePresentation() {
  const content = textarea.value.trim();
  if (!content) {
    textarea.focus();
    textarea.style.borderColor = '#EF4444';
    setTimeout(() => { textarea.style.borderColor = ''; }, 2000);
    return;
  }

  // Disable button, show loading
  const btn = document.getElementById('generateBtn');
  btn.disabled = true;
  document.getElementById('btnContent').style.display = 'none';
  document.getElementById('btnLoading').style.display = 'flex';

  // Switch to loading view
  setViewState('state-loading');
  stopActivityFeed();
  activateStep(1);
  setProgress(2, 'Contacting the presentation service...', 'Preparing your request');
  setActivityMessage('Waiting for the generation pipeline to begin');

  try {
    // Collect image data
    const images = collectImageData();

    const response = await fetch('/api/generate', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/x-ndjson'
      },
      body: JSON.stringify({
        content,
        palettes: [...selectedPalettes],
        images: images
      })
    });

    if (!response.ok) {
      // Preserve actionable server feedback instead of replacing it with a
      // generic status code in the error panel.
      const responseText = await response.text();
      let message = `Server error: ${response.status}`;
      try {
        const parsed = JSON.parse(responseText);
        message = parsed.error || message;
      } catch (_) {
        message = responseText.trim() || message;
      }
      throw new Error(message);
    }

    const data = await readGenerationProgress(response);
    setProgress(100, 'Complete!', 'Your presentation is ready');
    setActivityMessage('Your PowerPoint is ready to download');

    // Render one download action for each selected palette variant.
    const downloadBtn = document.getElementById('downloadBtn');
    downloadBtn.href = '/api/download/' + data.fileId;
    downloadBtn.download = data.fileName || 'presentation.pptx';
    renderDownloads(data.files || [data]);

    // Show success
    setViewState('state-success');
    stopActivityFeed();

    // Re-enable button
    btn.disabled = false;
    document.getElementById('btnContent').style.display = 'flex';
    document.getElementById('btnLoading').style.display = 'none';

    // Keep downloads explicit so browsers do not block multi-file batches.
    setTimeout(() => resetImageUploads(), 800);

  } catch (err) {
    console.error('Generation error:', err);
    stopActivityFeed();
    document.getElementById('errorMessage').textContent = err.message || 'Failed to generate. Please try again.';
    setViewState('state-error');

    btn.disabled = false;
    document.getElementById('btnContent').style.display = 'flex';
    document.getElementById('btnLoading').style.display = 'none';
  }
}
