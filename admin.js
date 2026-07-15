const API_URL = 'https://script.google.com/macros/s/AKfycbzVuUAyFVE1cM3x2_JG2j4xBVtJ9_Ib_8WaFrThVd7HQbnaJ4mSFsk5i8y7kFQTshY69Q/exec';

const loginView = document.getElementById('loginView');
const dashboardView = document.getElementById('dashboardView');
const loginForm = document.getElementById('loginForm');
const loginMessage = document.getElementById('loginMessage');
const demoForm = document.getElementById('demoForm');
const settingsForm = document.getElementById('settingsForm');
const packageForm = document.getElementById('packageForm');
const demoImageInput = document.getElementById('demoImage');
const demoImagePreview = document.getElementById('demoImagePreview');
const demoImagePreviewImg = document.getElementById('demoImagePreviewImg');
const packageCoverImageInput = document.getElementById('packageCoverImage');
const packageSampleImagesInput = document.getElementById('packageSampleImages');
const packageCoverPreview = document.getElementById('packageCoverPreview');
const packageSamplePreviewGrid = document.getElementById('packageSamplePreviewGrid');

let token = sessionStorage.getItem('asp_admin_token') || '';
let adminData = { demos: [], packages: [], leads: [], reviews: [], metaAdsRequests: [], settings: {} };
let packageSampleUrls = [];

function escapeHtml(value = '') {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function normalizeUrl(url) {
  const value = String(url || '').trim();
  if (!value) return '';
  if (/^https?:\/\//i.test(value)) return value;
  return `https://${value}`;
}

function parseImageUrls(value) {
  if (Array.isArray(value)) return value.filter(Boolean).map(String);
  const raw = String(value || '').trim();
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter(Boolean).map(String) : [];
  } catch (error) {
    return raw.split(/\r?\n|,/).map(item => item.trim()).filter(Boolean);
  }
}

async function apiGet(action, params = {}) {
  if (API_URL.includes('PASTE_YOUR')) throw new Error('API URL is not configured yet.');
  const url = new URL(API_URL);
  url.searchParams.set('action', action);
  Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
  const response = await fetch(url.toString(), { redirect: 'follow' });
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || 'Request failed.');
  return data;
}

async function apiPost(action, payload = {}) {
  if (API_URL.includes('PASTE_YOUR')) throw new Error('API URL is not configured yet.');
  const body = new URLSearchParams({ action, payload: JSON.stringify(payload) });
  const response = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
    body,
    redirect: 'follow'
  });
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || 'Request failed.');
  return data;
}

function showMessage(element, message, type = '') {
  element.className = `form-message ${type}`;
  element.textContent = message;
}

async function login(password) {
  const data = await apiPost('adminLogin', { password });
  token = data.token;
  sessionStorage.setItem('asp_admin_token', token);
  await loadAdminData();
}

async function loadAdminData() {
  try {
    const data = await apiGet('getAdminData', { token });
    adminData = {
      demos: data.demos || [],
      packages: data.packages || [],
      leads: data.leads || [],
      reviews: data.reviews || [],
      metaAdsRequests: data.metaAdsRequests || [],
      settings: data.settings || {}
    };
    loginView.classList.add('hidden');
    dashboardView.classList.remove('hidden');
    renderAll();
  } catch (error) {
    sessionStorage.removeItem('asp_admin_token');
    token = '';
    loginView.classList.remove('hidden');
    dashboardView.classList.add('hidden');
    showMessage(loginMessage, error.message, 'error');
  }
}

function renderAll() {
  document.getElementById('demoCount').textContent = adminData.demos.length;
  document.getElementById('leadCount').textContent = adminData.leads.length;
  document.getElementById('newLeadCount').textContent = adminData.leads.filter(item => item.status === 'New').length;
  document.getElementById('pendingReviewCount').textContent = adminData.reviews.filter(item => item.status === 'Pending').length;
  document.getElementById('packageCount').textContent = adminData.packages.length;
  document.getElementById('metaAdsRequestCount').textContent = adminData.metaAdsRequests.length;
  document.getElementById('newMetaAdsCount').textContent = adminData.metaAdsRequests.filter(item => item.status === 'New').length;
  renderDemos();
  renderPackages();
  renderLeads();
  renderMetaAdsRequests();
  renderReviews();
  fillSettings();
}

function renderDemos() {
  const container = document.getElementById('adminDemoList');
  if (!adminData.demos.length) {
    container.innerHTML = '<div class="empty-state">No demo websites yet.</div>';
    return;
  }

  container.innerHTML = adminData.demos.map(item => `
    <div style="padding:16px 0;border-bottom:1px solid var(--line)">
      <b style="display:block;color:var(--text)">${escapeHtml(item.title)}</b>
      <span>${escapeHtml(item.category)} · ${escapeHtml(item.description)}</span>
      <div style="margin-top:10px">
        <button class="small-button" onclick="editDemo('${item.id}')">Edit</button>
        <button class="small-button danger" onclick="deleteDemo('${item.id}')">Delete</button>
      </div>
    </div>
  `).join('');
}

function renderPackages() {
  const container = document.getElementById('adminPackageList');
  if (!container) return;

  if (!adminData.packages.length) {
    container.innerHTML = '<div class="empty-state">No service packages yet.</div>';
    return;
  }

  container.innerHTML = adminData.packages.map(item => {
    const sampleCount = parseImageUrls(item.sampleImageUrls).length;
    return `
      <div class="package-admin-item">
        <b style="display:block;color:var(--text)">${escapeHtml(item.name)}</b>
        ${item.coverImageUrl ? `<img class="package-admin-cover" src="${escapeHtml(item.coverImageUrl)}" alt="" />` : ''}
        <div class="package-admin-meta">
          ${escapeHtml(item.category)} ·
          <span class="package-admin-price">${escapeHtml(item.price)}</span> ·
          ${escapeHtml(item.active === 'Yes' ? 'Public' : 'Hidden')} ·
          ${sampleCount} sample${sampleCount === 1 ? '' : 's'}
          ${item.featured === 'Yes' ? ' · Featured' : ''}
        </div>
        <div style="margin-top:10px">
          <button class="small-button" onclick="editPackage('${item.id}')">Edit</button>
          <button class="small-button danger" onclick="deletePackage('${item.id}')">Delete</button>
        </div>
      </div>
    `;
  }).join('');
}

function renderLeads() {
  const body = document.getElementById('leadTableBody');
  if (!adminData.leads.length) {
    body.innerHTML = '<tr><td colspan="7">No website or creative inquiries yet.</td></tr>';
    return;
  }

  body.innerHTML = adminData.leads.map(item => `
    <tr>
      <td>${escapeHtml(item.createdAt)}</td>
      <td><b>${escapeHtml(item.name)}</b><br>${escapeHtml(item.business || '')}</td>
      <td>${escapeHtml(item.contact)}</td>
      <td>
        <b>${escapeHtml(item.service)}</b>
        ${item.package ? `<br><span>${escapeHtml(item.package)}</span>` : ''}
      </td>
      <td>${escapeHtml(item.budget || '')}</td>
      <td>${escapeHtml(item.message)}</td>
      <td>
        <select class="status-select" onchange="updateLeadStatus('${item.id}', this.value)">
          ${['New','Contacted','Quoted','Closed','Not Proceeding'].map(status =>
            `<option ${status === item.status ? 'selected' : ''}>${status}</option>`
          ).join('')}
        </select>
      </td>
    </tr>
  `).join('');
}

function detailBlock(label, value, wide = false, asLink = false) {
  const cleanValue = String(value || '').trim();
  let body = '<p>—</p>';
  if (cleanValue) {
    if (asLink) {
      const url = normalizeUrl(cleanValue);
      body = `<a href="${escapeHtml(url)}" target="_blank" rel="noopener">${escapeHtml(cleanValue)}</a>`;
    } else {
      body = `<p>${escapeHtml(cleanValue)}</p>`;
    }
  }
  return `<div class="meta-detail ${wide ? 'wide' : ''}"><span>${escapeHtml(label)}</span>${body}</div>`;
}

function renderMetaAdsRequests() {
  const container = document.getElementById('metaAdsRequestList');
  if (!container) return;

  if (!adminData.metaAdsRequests.length) {
    container.innerHTML = '<div class="empty-state">No detailed Meta Ads requests yet.</div>';
    return;
  }

  const statuses = ['New','Reviewing','Waiting for Client','Waiting for Access','Ready to Launch','Active','Completed','Not Proceeding'];

  container.innerHTML = adminData.metaAdsRequests.map(item => `
    <article class="meta-request-card">
      <div class="meta-request-head">
        <div>
          <h3>${escapeHtml(item.business || item.name)}</h3>
          <p><b>${escapeHtml(item.name)}</b> · ${escapeHtml(item.email)} · ${escapeHtml(item.contact)}</p>
          ${item.package ? `<p>Selected package: ${escapeHtml(item.package)}</p>` : ''}
          <p>Submitted: ${escapeHtml(item.createdAt)}</p>
        </div>
        <select class="status-select" onchange="updateMetaAdsStatus('${item.id}', this.value)">
          ${statuses.map(status => `<option ${status === item.status ? 'selected' : ''}>${status}</option>`).join('')}
        </select>
      </div>
      <div class="meta-request-details">
        ${detailBlock('Campaign Objective', item.objective)}
        ${detailBlock('Daily Budget', item.dailyBudget)}
        ${detailBlock('Run Duration', item.runDuration)}
        ${detailBlock('Meta Business Portfolio', item.businessPortfolioStatus)}
        ${detailBlock('Ad Account', item.adsAccountStatus)}
        ${detailBlock('Previous Ads Experience', item.previousAds)}
        ${detailBlock('Facebook Page', item.facebookPageUrl, false, true)}
        ${detailBlock('Website / Store', item.websiteUrl, false, true)}
        ${detailBlock('Reference / Product Files', item.referenceLink, false, true)}
        ${detailBlock('Product / Service to Advertise', item.productService, true)}
        ${detailBlock('Offer / Price / Promo', item.offer, true)}
        ${detailBlock('Target Locations', item.targetLocations)}
        ${detailBlock('Ideal Customer / Audience', item.targetAudience, true)}
        ${detailBlock('Creative Status', item.creativeStatus, true)}
        ${detailBlock('Additional Details', item.notes, true)}
      </div>
    </article>
  `).join('');
}

function renderReviews() {
  const body = document.getElementById('reviewTableBody');
  if (!body) return;

  if (!adminData.reviews.length) {
    body.innerHTML = '<tr><td colspan="7">No customer reviews yet.</td></tr>';
    return;
  }

  body.innerHTML = adminData.reviews.map(item => `
    <tr>
      <td>${escapeHtml(item.createdAt)}</td>
      <td><b>${escapeHtml(item.name)}</b><br>${escapeHtml(item.business || '')}</td>
      <td>${escapeHtml(item.email)}</td>
      <td>${'★'.repeat(Math.max(1, Math.min(5, Number(item.rating) || 5)))}<br>${escapeHtml(item.service || '')}</td>
      <td>${escapeHtml(item.review)}</td>
      <td>
        <select class="status-select" onchange="updateReviewStatus('${item.id}', this.value)">
          ${['Pending','Approved','Rejected'].map(status =>
            `<option ${status === item.status ? 'selected' : ''}>${status}</option>`
          ).join('')}
        </select>
      </td>
      <td><button class="small-button danger" onclick="deleteReview('${item.id}')">Delete</button></td>
    </tr>
  `).join('');
}

function showDemoImagePreview(url) {
  if (!url) {
    demoImagePreview.classList.add('hidden');
    demoImagePreviewImg.removeAttribute('src');
    return;
  }
  demoImagePreviewImg.src = url;
  demoImagePreview.classList.remove('hidden');
}

function showPackageCoverPreview(url) {
  if (!url) {
    packageCoverPreview.innerHTML = '';
    packageCoverPreview.classList.add('hidden');
    return;
  }
  packageCoverPreview.innerHTML = `<img src="${escapeHtml(url)}" alt="Package cover preview" />`;
  packageCoverPreview.classList.remove('hidden');
}

function renderPackageSamplePreviews() {
  packageForm.elements.sampleImageUrls.value = JSON.stringify(packageSampleUrls);
  if (!packageSampleUrls.length) {
    packageSamplePreviewGrid.innerHTML = '<span class="notice">No saved sample images yet.</span>';
    return;
  }

  packageSamplePreviewGrid.innerHTML = packageSampleUrls.map((url, index) => `
    <div class="package-sample-preview">
      <img src="${escapeHtml(url)}" alt="Package sample ${index + 1}" />
      <button type="button" onclick="removePackageSample(${index})" aria-label="Remove sample ${index + 1}">×</button>
    </div>
  `).join('');
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
    reader.onerror = () => reject(new Error('Unable to read the selected image.'));
    reader.readAsDataURL(file);
  });
}

async function uploadImage(file, action) {
  if (!file) return '';
  if (!file.type.startsWith('image/')) throw new Error('Please select valid image files only.');
  if (file.size > 4 * 1024 * 1024) throw new Error(`${file.name} is larger than 4 MB.`);

  const base64 = await fileToBase64(file);
  const data = await apiPost(action, {
    token,
    fileName: file.name,
    mimeType: file.type,
    base64
  });
  return data.imageUrl;
}

function fillSettings() {
  Object.entries(adminData.settings || {}).forEach(([key, value]) => {
    const field = settingsForm.elements[key];
    if (field) field.value = value || '';
  });
}

function resetPackageForm() {
  packageForm.reset();
  packageForm.elements.active.value = 'Yes';
  packageForm.elements.featured.value = 'No';
  packageForm.elements.sortOrder.value = '0';
  packageForm.elements.coverImageUrl.value = '';
  packageSampleUrls = [];
  renderPackageSamplePreviews();
  showPackageCoverPreview('');
  packageCoverImageInput.value = '';
  packageSampleImagesInput.value = '';
  document.getElementById('packageFormTitle').textContent = 'Add Service Package';
  document.getElementById('cancelEditPackage').classList.add('hidden');
}

window.editDemo = function(id) {
  const demo = adminData.demos.find(item => item.id === id);
  if (!demo) return;
  Object.entries(demo).forEach(([key, value]) => {
    const field = demoForm.elements[key];
    if (field) field.value = value || '';
  });
  showDemoImagePreview(demo.imageUrl || '');
  demoImageInput.value = '';
  document.getElementById('demoFormTitle').textContent = 'Edit Demo Website';
  document.getElementById('cancelEditDemo').classList.remove('hidden');
  demoForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
};

window.deleteDemo = async function(id) {
  if (!confirm('Delete this demo website?')) return;
  try {
    await apiPost('deleteDemo', { token, id });
    await loadAdminData();
  } catch (error) {
    alert(error.message);
  }
};

window.editPackage = function(id) {
  const item = adminData.packages.find(packageItem => packageItem.id === id);
  if (!item) return;
  Object.entries(item).forEach(([key, value]) => {
    const field = packageForm.elements[key];
    if (field && !['sampleImageUrls'].includes(key)) field.value = value || '';
  });
  packageSampleUrls = parseImageUrls(item.sampleImageUrls);
  packageForm.elements.sampleImageUrls.value = JSON.stringify(packageSampleUrls);
  showPackageCoverPreview(item.coverImageUrl || '');
  renderPackageSamplePreviews();
  packageCoverImageInput.value = '';
  packageSampleImagesInput.value = '';
  document.getElementById('packageFormTitle').textContent = 'Edit Service Package';
  document.getElementById('cancelEditPackage').classList.remove('hidden');
  packageForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
};

window.removePackageSample = function(index) {
  packageSampleUrls.splice(index, 1);
  renderPackageSamplePreviews();
};

window.deletePackage = async function(id) {
  if (!confirm('Delete this service package?')) return;
  try {
    await apiPost('deletePackage', { token, id });
    await loadAdminData();
  } catch (error) {
    alert(error.message);
  }
};

window.updateLeadStatus = async function(id, status) {
  try {
    await apiPost('updateLeadStatus', { token, id, status });
    await loadAdminData();
  } catch (error) {
    alert(error.message);
  }
};

window.updateMetaAdsStatus = async function(id, status) {
  try {
    await apiPost('updateMetaAdsStatus', { token, id, status });
    await loadAdminData();
  } catch (error) {
    alert(error.message);
  }
};

window.updateReviewStatus = async function(id, status) {
  try {
    await apiPost('updateReviewStatus', { token, id, status });
    await loadAdminData();
  } catch (error) {
    alert(error.message);
  }
};

window.deleteReview = async function(id) {
  if (!confirm('Delete this customer review?')) return;
  try {
    await apiPost('deleteReview', { token, id });
    await loadAdminData();
  } catch (error) {
    alert(error.message);
  }
};

loginForm.addEventListener('submit', async event => {
  event.preventDefault();
  showMessage(loginMessage, 'Logging in...');
  try {
    await login(document.getElementById('adminPassword').value);
  } catch (error) {
    showMessage(loginMessage, error.message, 'error');
  }
});

document.getElementById('logoutButton').addEventListener('click', () => {
  sessionStorage.removeItem('asp_admin_token');
  location.reload();
});

document.querySelectorAll('.admin-tabs button').forEach(button => {
  button.addEventListener('click', () => {
    document.querySelectorAll('.admin-tabs button').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.admin-panel').forEach(panel => panel.classList.remove('active'));
    button.classList.add('active');
    document.getElementById(button.dataset.tab).classList.add('active');
  });
});

demoForm.addEventListener('submit', async event => {
  event.preventDefault();
  const message = document.getElementById('demoMessage');
  const submitButton = demoForm.querySelector('button[type="submit"]');
  const payload = Object.fromEntries(new FormData(demoForm).entries());
  payload.token = token;
  delete payload.demoImage;

  showMessage(message, demoImageInput.files[0] ? 'Uploading image to Google Drive...' : 'Saving demo...');
  submitButton.disabled = true;
  submitButton.textContent = 'Saving...';

  try {
    if (demoImageInput.files[0]) {
      payload.imageUrl = await uploadImage(demoImageInput.files[0], 'uploadDemoImage');
      showMessage(message, 'Image uploaded. Saving demo...');
    }

    if (!payload.imageUrl) throw new Error('Please upload a demo screenshot/image.');

    await apiPost('saveDemo', payload);
    showMessage(message, 'Demo and image saved successfully.', 'success');
    demoForm.reset();
    showDemoImagePreview('');
    document.getElementById('demoFormTitle').textContent = 'Add Demo Website';
    document.getElementById('cancelEditDemo').classList.add('hidden');
    await loadAdminData();
  } catch (error) {
    showMessage(message, error.message, 'error');
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = 'Save Demo';
  }
});

document.getElementById('cancelEditDemo').addEventListener('click', () => {
  demoForm.reset();
  showDemoImagePreview('');
  document.getElementById('demoFormTitle').textContent = 'Add Demo Website';
  document.getElementById('cancelEditDemo').classList.add('hidden');
});

packageForm.addEventListener('submit', async event => {
  event.preventDefault();
  const message = document.getElementById('packageMessage');
  const submitButton = packageForm.querySelector('button[type="submit"]');
  const payload = Object.fromEntries(new FormData(packageForm).entries());
  payload.token = token;

  showMessage(message, 'Preparing package...');
  submitButton.disabled = true;
  submitButton.textContent = 'Saving...';

  try {
    if (packageCoverImageInput.files[0]) {
      showMessage(message, 'Uploading package cover to Google Drive...');
      payload.coverImageUrl = await uploadImage(packageCoverImageInput.files[0], 'uploadPackageImage');
    }

    const newSamples = Array.from(packageSampleImagesInput.files || []);
    if (packageSampleUrls.length + newSamples.length > 8) {
      throw new Error('A package can have a maximum of 8 sample images.');
    }

    for (let index = 0; index < newSamples.length; index += 1) {
      showMessage(message, `Uploading sample ${index + 1} of ${newSamples.length} to Google Drive...`);
      const imageUrl = await uploadImage(newSamples[index], 'uploadPackageImage');
      packageSampleUrls.push(imageUrl);
    }

    payload.sampleImageUrls = JSON.stringify(packageSampleUrls);
    showMessage(message, 'Images ready. Saving package...');
    await apiPost('savePackage', payload);
    showMessage(message, 'Service package, cover, and samples saved successfully.', 'success');
    resetPackageForm();
    await loadAdminData();
  } catch (error) {
    showMessage(message, error.message, 'error');
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = 'Save Package';
  }
});

document.getElementById('cancelEditPackage').addEventListener('click', resetPackageForm);

settingsForm.addEventListener('submit', async event => {
  event.preventDefault();
  const message = document.getElementById('settingsMessage');
  const payload = Object.fromEntries(new FormData(settingsForm).entries());
  payload.token = token;

  showMessage(message, 'Saving...');
  try {
    await apiPost('saveSettings', payload);
    showMessage(message, 'Website settings saved.', 'success');
    await loadAdminData();
  } catch (error) {
    showMessage(message, error.message, 'error');
  }
});

demoImageInput.addEventListener('change', () => {
  const file = demoImageInput.files[0];
  if (!file) return;
  if (!file.type.startsWith('image/')) {
    showMessage(document.getElementById('demoMessage'), 'Please select a valid image file.', 'error');
    demoImageInput.value = '';
    return;
  }
  if (file.size > 4 * 1024 * 1024) {
    showMessage(document.getElementById('demoMessage'), 'Image must be 4 MB or smaller.', 'error');
    demoImageInput.value = '';
    return;
  }
  showDemoImagePreview(URL.createObjectURL(file));
});

packageCoverImageInput.addEventListener('change', () => {
  const file = packageCoverImageInput.files[0];
  if (!file) return;
  if (!file.type.startsWith('image/') || file.size > 4 * 1024 * 1024) {
    showMessage(document.getElementById('packageMessage'), 'Package cover must be a valid image up to 4 MB.', 'error');
    packageCoverImageInput.value = '';
    return;
  }
  showPackageCoverPreview(URL.createObjectURL(file));
});

packageSampleImagesInput.addEventListener('change', () => {
  const files = Array.from(packageSampleImagesInput.files || []);
  if (packageSampleUrls.length + files.length > 8) {
    showMessage(document.getElementById('packageMessage'), 'Maximum 8 sample images per package.', 'error');
    packageSampleImagesInput.value = '';
    return;
  }
  const invalid = files.find(file => !file.type.startsWith('image/') || file.size > 4 * 1024 * 1024);
  if (invalid) {
    showMessage(document.getElementById('packageMessage'), `${invalid.name} must be a valid image up to 4 MB.`, 'error');
    packageSampleImagesInput.value = '';
    return;
  }
  showMessage(document.getElementById('packageMessage'), `${files.length} new sample image${files.length === 1 ? '' : 's'} selected. They will upload when you save the package.`);
});

renderPackageSamplePreviews();
if (token) loadAdminData();
