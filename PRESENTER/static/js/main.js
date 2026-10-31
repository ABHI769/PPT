// State
let presentations = [];
let currentPresId = 'pres_5750c2f9';
let currentSlideIndex = 0;
const syncChannel = new BroadcastChannel('python_ppt_sync');

// DOM Elements
const dashboardView = document.getElementById('dashboard-view');
const editorView = document.getElementById('editor-view');
const navDashboardBtn = document.getElementById('nav-dashboard-btn');
const navEditorBtn = document.getElementById('nav-editor-btn');
const launchPresentationBtn = document.getElementById('launch-pres-btn');
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const presentationsGrid = document.getElementById('presentations-grid');
const reviewModal = document.getElementById('review-modal');
const reviewTableBody = document.getElementById('review-table-body');

// Harmonious purple-violet accent colors for slide preview badges
const ACCENT_COLORS = ['#a855f7', '#8b5cf6', '#c084fc', '#7c3aed', '#6366f1', '#d946ef'];

// Fetch Presentations
async function loadPresentations() {
  try {
    const res = await fetch('/api/presentations');
    presentations = await res.json();
    if (!presentations.some(p => p.id === currentPresId)) {
      currentPresId = presentations.length > 0 ? presentations[0].id : null;
    }
    renderDashboard();
    renderEditor();
  } catch (err) {
    console.error('Failed to load presentations:', err);
  }
}

// Render Dashboard Presentations Grid
function renderDashboard() {
  if (!presentationsGrid) return;
  presentationsGrid.innerHTML = '';

  if (presentations.length === 0) {
    presentationsGrid.innerHTML = `
      <div style="grid-column: 1 / -1; text-align: center; padding: 48px 24px; background: rgba(22, 10, 42, 0.55); backdrop-filter: blur(16px); -webkit-backdrop-filter: blur(16px); border-radius: 14px; border: 1px dashed rgba(168, 85, 247, 0.4); box-shadow: 0 10px 25px rgba(0,0,0,0.5);">
        <div style="font-size: 36px; margin-bottom: 12px;">📂</div>
        <h3 style="font-size: 17px; font-weight: 700; color: #fff; margin-bottom: 8px;">No Presentations Available</h3>
        <p style="font-size: 13px; color: var(--text-muted); max-width: 450px; margin: 0 auto;">
          Upload a PowerPoint presentation (.pptx or .ppt) above to automatically export high-definition slides and configure voice tags!
        </p>
      </div>
    `;
    return;
  }

  presentations.forEach((p) => {
    const isSelected = p.id === currentPresId;
    const totalLabels = (p.slides || []).reduce((acc, s) => acc + (s.labels ? s.labels.length : 0), 0);
    const coverThumb = (p.slides && p.slides[0] && p.slides[0].image_url)
      ? `<div style="width: 100%; aspect-ratio: 16/9; background: #000; border-radius: 8px; overflow: hidden; margin-bottom: 14px; border: 1px solid rgba(168, 85, 247, 0.28); display: flex; align-items: center; justify-content: center; box-shadow: 0 8px 18px rgba(0,0,0,0.6);">
           <img src="${p.slides[0].image_url}" alt="Cover Thumbnail" style="max-width: 100%; max-height: 100%; object-fit: contain;">
         </div>`
      : '';

    const card = document.createElement('div');
    card.style = `
      background: ${isSelected ? 'rgba(38, 16, 72, 0.72)' : 'rgba(20, 9, 40, 0.55)'};
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: ${isSelected ? '2px solid rgba(192, 132, 252, 0.7)' : '1px solid rgba(168, 85, 247, 0.22)'};
      border-radius: 14px;
      padding: 20px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      box-shadow: ${isSelected ? '0 12px 35px rgba(0,0,0,0.7), 0 0 25px rgba(168, 85, 247, 0.35)' : '0 10px 25px rgba(0,0,0,0.6), 0 0 15px rgba(147, 51, 234, 0.08)'};
      transition: all 0.2s ease;
    `;

    card.innerHTML = `
      <div>
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
          <span style="background-color: rgba(168, 85, 247, 0.2); color: #d8b4fe; border: 1px solid rgba(168, 85, 247, 0.35); padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 700;">
            ${p.slide_count} Slides
          </span>
          <button onclick="deletePresentation('${p.id}', event)" title="Remove presentation" style="background: none; border: none; color: #a855f7; cursor: pointer; font-size: 16px; padding: 4px;">
            ✕
          </button>
        </div>
        ${coverThumb}
        <h3 style="font-size: 17px; font-weight: 700; color: #f8fafc; margin-bottom: 6px;">${p.title}</h3>
        <p style="font-size: 12px; color: #c4b5fd; margin-bottom: 16px;">✓ ${totalLabels} Active Voice Labels</p>
      </div>

      <div style="display: flex; gap: 8px; padding-top: 16px; border-top: 1px solid rgba(168, 85, 247, 0.2);">
        <button onclick="selectAndEdit('${p.id}')" class="btn-secondary" style="flex: 1; justify-content: center;">
          Edit Labels
        </button>
        <button onclick="startPresentation('${p.id}')" class="btn-primary" style="flex: 1.2; justify-content: center;">
          ▶ Present
        </button>
      </div>
    `;

    presentationsGrid.appendChild(card);
  });
}

// Upload Handling
if (dropZone && fileInput) {
  dropZone.addEventListener('click', () => fileInput.click());

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });

  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      uploadFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) {
      uploadFile(fileInput.files[0]);
    }
  });
}

async function uploadFile(file) {
  const fileNameLower = file.name.toLowerCase();
  if (!fileNameLower.endsWith('.pptx') && !fileNameLower.endsWith('.ppt')) {
    alert('Please upload a valid Microsoft PowerPoint presentation (.pptx or .ppt).');
    return;
  }

  const formData = new FormData();
  formData.append('file', file);

  const dropText = document.getElementById('drop-text');
  if (dropText) dropText.innerText = 'Processing PowerPoint presentation & rendering 1080p slide stages...';

  try {
    const res = await fetch('/api/upload', {
      method: 'POST',
      body: formData,
    });
    const result = await res.json();
    if (result.error) {
      alert(result.error);
    } else {
      currentPresId = result.id;
      currentSlideIndex = 0;
      await loadPresentations();
      showView('editor');
    }
  } catch (err) {
    alert('Failed to upload presentation.');
  } finally {
    if (dropText) dropText.innerText = 'Click or Drag & Drop your PowerPoint presentation here';
    if (fileInput) fileInput.value = '';
  }
}

// Select & Edit
function selectAndEdit(id) {
  currentPresId = id;
  currentSlideIndex = 0;
  showView('editor');
  renderEditor();
}

// Delete Presentation
async function deletePresentation(id, event) {
  event.stopPropagation();
  if (!confirm('Are you sure you want to remove this presentation?')) return;
  try {
    await fetch(`/api/presentation/${id}`, { method: 'DELETE' });
    if (currentPresId === id) {
      currentPresId = null;
    }
    await loadPresentations();
  } catch (err) {
    console.error('Failed to delete presentation:', err);
  }
}

// Show View
function showView(view) {
  if (view === 'dashboard') {
    dashboardView.style.display = 'block';
    editorView.style.display = 'none';
    navDashboardBtn?.classList.add('active');
    navEditorBtn?.classList.remove('active');
  } else {
    dashboardView.style.display = 'none';
    editorView.style.display = 'grid';
    navDashboardBtn?.classList.remove('active');
    navEditorBtn?.classList.add('active');
    renderEditor();
  }
}

navDashboardBtn?.addEventListener('click', () => showView('dashboard'));
navEditorBtn?.addEventListener('click', () => showView('editor'));

// Render Slide Label Editor
function renderEditor() {
  const currentPres = presentations.find((p) => p.id === currentPresId);
  const thumbsContainer = document.getElementById('editor-thumbs-list');
  const previewImg = document.getElementById('preview-slide-img');
  const previewCard = document.getElementById('preview-card-fallback');
  const activeLabelsContainer = document.getElementById('active-labels-container');
  const titleEl = document.getElementById('editor-pres-title');
  const slideCounterEl = document.getElementById('preview-slide-counter');
  const labelHeader = document.getElementById('label-slide-header');

  if (!currentPres || !currentPres.slides || currentPres.slides.length === 0) {
    if (titleEl) titleEl.innerText = 'No presentation selected';
    if (thumbsContainer) {
      thumbsContainer.innerHTML = '<div style="padding: 24px; color: var(--text-muted); font-size: 13px; text-align: center;">No presentation selected.<br><br><a href="#" onclick="showView(\'dashboard\')" style="color: var(--accent-light);">Go to Dashboard</a></div>';
    }
    if (previewImg) previewImg.style.display = 'none';
    if (previewCard) previewCard.style.display = 'none';
    if (activeLabelsContainer) activeLabelsContainer.innerHTML = '';
    if (slideCounterEl) slideCounterEl.innerText = 'Slide 0 of 0';
    if (labelHeader) labelHeader.innerText = 'No active slide';
    return;
  }

  currentSlideIndex = Math.max(0, Math.min(currentSlideIndex, currentPres.slides.length - 1));

  if (titleEl) {
    titleEl.innerText = `${currentPres.title} (${currentPres.slide_count} slides)`;
  }

  // Left column thumbnails
  if (thumbsContainer) {
    thumbsContainer.innerHTML = '';
    currentPres.slides.forEach((slide) => {
      const isSelected = slide.index === currentSlideIndex;
      const card = document.createElement('div');
      card.className = `slide-thumb-card ${isSelected ? 'active' : ''}`;
      card.onclick = () => {
        currentSlideIndex = slide.index;
        renderEditor();
      };

      const labelChips = (slide.labels || [])
        .slice(0, 3)
        .map((l) => `<span class="slide-tag">${l}</span>`)
        .join(' ');

      const thumbImg = slide.image_url
        ? `<div style="width: 100%; aspect-ratio: 16/9; background: #000; border-radius: 6px; overflow: hidden; margin-bottom: 8px; border: 1px solid rgba(168, 85, 247, 0.25); display: flex; align-items: center; justify-content: center;">
             <img src="${slide.image_url}" alt="Slide thumbnail" style="max-width: 100%; max-height: 100%; object-fit: contain;">
           </div>`
        : '';

      card.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
          <span style="font-size: 11px; font-weight: 700; background: ${isSelected ? '#a855f7' : 'rgba(124, 58, 237, 0.3)'}; color: #fff; padding: 2px 7px; border-radius: 4px; border: 1px solid ${isSelected ? 'rgba(255,255,255,0.4)' : 'rgba(168,85,247,0.3)'};">
            Slide ${slide.number}
          </span>
          <span style="font-size: 11px; color: #c4b5fd;">${slide.labels ? slide.labels.length : 0} labels</span>
        </div>
        ${thumbImg}
        <div style="font-size: 12px; font-weight: 600; color: ${isSelected ? '#fff' : '#e2e8f0'}; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; margin-bottom: 6px;">
          ${slide.title}
        </div>
        <div style="display: flex; flex-wrap: wrap; gap: 4px;">
          ${labelChips}
        </div>
      `;

      thumbsContainer.appendChild(card);
    });
  }

  // Center stage slide preview
  const currentSlide = currentPres.slides[currentSlideIndex];
  if (!currentSlide) return;

  if (previewImg && previewCard) {
    previewImg.onerror = () => {
      previewImg.style.display = 'none';
      previewCard.style.display = 'flex';
    };

    if (currentSlide.image_url) {
      previewImg.src = currentSlide.image_url;
      previewImg.style.display = 'block';
      previewCard.style.display = 'none';
    } else {
      previewImg.style.display = 'none';
      previewCard.style.display = 'flex';
    }

    const accent = ACCENT_COLORS[currentSlideIndex % ACCENT_COLORS.length];
    const pill = document.getElementById('preview-pill-badge');
    if (pill) {
      pill.innerText = `SLIDE 0${currentSlide.number}`;
      pill.style.background = `linear-gradient(135deg, ${accent} 0%, #4c1d95 100%)`;
      pill.style.borderColor = 'rgba(255,255,255,0.3)';
    }

    const titleBox = document.getElementById('preview-slide-title');
    if (titleBox) titleBox.innerText = currentSlide.title;

    const pointsList = document.getElementById('preview-points-list');
    if (pointsList) {
      pointsList.innerHTML = '';
      const items = currentSlide.points || (currentSlide.raw_text ? currentSlide.raw_text.split('\n') : []);
      items.forEach((pt) => {
        if (pt.trim()) {
          const li = document.createElement('li');
          li.innerText = pt.trim();
          pointsList.appendChild(li);
        }
      });
    }
  }

  if (slideCounterEl) {
    slideCounterEl.innerText = `Slide ${currentSlide.number} of ${currentPres.slide_count}`;
  }

  // Right column label manager
  if (labelHeader) {
    labelHeader.innerText = `Slide ${currentSlide.number}: ${currentSlide.title}`;
  }
  renderLabelChips(currentSlide.labels || []);
}

function renderLabelChips(labels) {
  const container = document.getElementById('active-labels-container');
  if (!container) return;
  container.innerHTML = '';

  if (labels.length === 0) {
    container.innerHTML = `
      <div style="background-color: rgba(20, 9, 40, 0.5); border: 1px dashed rgba(168, 85, 247, 0.35); border-radius: 8px; padding: 20px; text-align: center; color: #c4b5fd; font-size: 12px;">
        No labels assigned yet.<br>Type "Introduction", "Technology Used", or "Conclusion" above and press Enter.
      </div>
    `;
    return;
  }

  labels.forEach((label) => {
    const chip = document.createElement('div');
    chip.style = `
      display: flex;
      align-items: center;
      justify-content: space-between;
      background-color: rgba(30, 14, 56, 0.65);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border: 1px solid rgba(168, 85, 247, 0.3);
      padding: 8px 12px;
      border-radius: 8px;
      margin-bottom: 8px;
      box-shadow: 0 4px 10px rgba(0,0,0,0.3);
    `;
    chip.innerHTML = `
      <span style="font-size: 13px; font-weight: 600; color: #f8fafc;">🏷️ ${label}</span>
      <button onclick="removeLabel('${encodeURIComponent(label)}')" title="Remove tag" style="background: none; border: none; color: #d8b4fe; cursor: pointer; padding: 2px 6px; font-size: 13px;">
        ✕
      </button>
    `;
    container.appendChild(chip);
  });
}

// Add Label
document.getElementById('add-label-form')?.addEventListener('submit', async (e) => {
  e.preventDefault();
  const input = document.getElementById('new-label-input');
  const label = input.value.trim();
  if (!label) return;

  const currentPres = presentations.find((p) => p.id === currentPresId);
  if (!currentPres) return;

  const currentSlide = currentPres.slides[currentSlideIndex];
  const currentLabels = currentSlide.labels || [];
  if (!currentLabels.includes(label)) {
    const updated = [...currentLabels, label];
    await saveLabelsToBackend(currentPres.id, currentSlideIndex, updated);
  }
  input.value = '';
});

// Remove Label
async function removeLabel(encodedLabel) {
  const labelToRemove = decodeURIComponent(encodedLabel);
  const currentPres = presentations.find((p) => p.id === currentPresId);
  if (!currentPres) return;

  const currentSlide = currentPres.slides[currentSlideIndex];
  const updated = (currentSlide.labels || []).filter((l) => l !== labelToRemove);
  await saveLabelsToBackend(currentPres.id, currentSlideIndex, updated);
}

// Quick Label Suggestion
function addSuggestedLabel(label) {
  const currentPres = presentations.find((p) => p.id === currentPresId);
  if (!currentPres) return;

  const currentSlide = currentPres.slides[currentSlideIndex];
  const currentLabels = currentSlide.labels || [];
  if (!currentLabels.includes(label)) {
    saveLabelsToBackend(currentPres.id, currentSlideIndex, [...currentLabels, label]);
  }
}

async function saveLabelsToBackend(presId, slideIndex, labels) {
  try {
    const res = await fetch(`/api/presentation/${presId}/labels`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ slide_index: slideIndex, labels }),
    });
    const result = await res.json();
    if (result.success) {
      const pres = presentations.find((p) => p.id === presId);
      if (pres) pres.slides = result.slides;
      renderEditor();
    }
  } catch (err) {
    console.error('Failed to save labels:', err);
  }
}

// Navigation buttons on preview
document.getElementById('prev-preview-btn')?.addEventListener('click', () => {
  if (currentSlideIndex > 0) {
    currentSlideIndex--;
    renderEditor();
  }
});

document.getElementById('next-preview-btn')?.addEventListener('click', () => {
  const currentPres = presentations.find((p) => p.id === currentPresId);
  if (currentPres && currentSlideIndex < currentPres.slide_count - 1) {
    currentSlideIndex++;
    renderEditor();
  }
});

// Pre-Presentation Review Modal
document.getElementById('review-labels-btn')?.addEventListener('click', () => {
  const currentPres = presentations.find((p) => p.id === currentPresId);
  if (!currentPres) return;

  document.getElementById('review-pres-title').innerText = `${currentPres.title} • ${currentPres.slide_count} Slides`;
  reviewTableBody.innerHTML = '';

  currentPres.slides.forEach((slide) => {
    const tr = document.createElement('tr');
    tr.style = 'border-bottom: 1px solid rgba(168, 85, 247, 0.2);';

    const tags = (slide.labels && slide.labels.length > 0)
      ? slide.labels.map((l) => `<span class="slide-tag" style="margin-right: 4px;">"${l}"</span>`).join('')
      : '<span style="color: #f43f5e; font-style: italic; font-size: 12px;">⚠️ No labels</span>';

    const thumbHtml = slide.image_url
      ? `<div style="width: 80px; aspect-ratio: 16/9; background: #000; border-radius: 4px; overflow: hidden; border: 1px solid rgba(168, 85, 247, 0.3); display: flex; align-items: center; justify-content: center;">
           <img src="${slide.image_url}" alt="Slide thumbnail" style="max-width: 100%; max-height: 100%; object-fit: contain;">
         </div>`
      : '';

    tr.innerHTML = `
      <td style="padding: 12px;"><span style="background: rgba(124, 58, 237, 0.25); color: #c084fc; border: 1px solid rgba(168, 85, 247, 0.35); padding: 4px 8px; border-radius: 4px; font-weight: 700;">#${slide.number}</span></td>
      <td style="padding: 12px; width: 90px;">${thumbHtml}</td>
      <td style="padding: 12px; font-weight: 600; color: #fff;">${slide.title}</td>
      <td style="padding: 12px;">${tags}</td>
    `;
    reviewTableBody.appendChild(tr);
  });

  reviewModal.style.display = 'flex';
});

document.getElementById('close-review-modal')?.addEventListener('click', () => {
  reviewModal.style.display = 'none';
});

// Presentation Launching
function startPresentation(id) {
  const targetId = id || currentPresId;
  if (!targetId) {
    alert('Please upload or select a presentation first.');
    return;
  }
  // Open Audience Window in secondary popup window
  window.open(`/audience?id=${targetId}`, 'AudienceWindow', 'width=1280,height=720,menubar=no,toolbar=no');
  // Navigate current tab to Presenter HUD
  window.location.href = `/presenter?id=${targetId}`;
}

launchPresentationBtn?.addEventListener('click', () => startPresentation());
document.getElementById('modal-start-pres-btn')?.addEventListener('click', () => startPresentation());

// Initialize
loadPresentations();
