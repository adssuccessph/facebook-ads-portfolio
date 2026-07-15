const API_URL = 'https://script.google.com/macros/s/AKfycbzVuUAyFVE1cM3x2_JG2j4xBVtJ9_Ib_8WaFrThVd7HQbnaJ4mSFsk5i8y7kFQTshY69Q/exec';

const demoGrid = document.getElementById('demoGrid');
const demoFilter = document.getElementById('demoFilter');
const packageGrid = document.getElementById('packageGrid');
const packageTabs = document.getElementById('packageTabs');
const leadForm = document.getElementById('leadForm');
const formMessage = document.getElementById('formMessage');
const reviewGrid = document.getElementById('reviewGrid');
const reviewForm = document.getElementById('reviewForm');
const reviewMessage = document.getElementById('reviewMessage');
const selectedPackageNote = document.getElementById('selectedPackageNote');
const selectedPackageName = document.getElementById('selectedPackageName');
const selectedPackageInput = document.getElementById('selectedPackageInput');
const clearSelectedPackage = document.getElementById('clearSelectedPackage');
const metaAdsRequestPanel = document.getElementById('metaAdsRequestPanel');
const metaAdsForm = document.getElementById('metaAdsForm');
const metaAdsMessage = document.getElementById('metaAdsMessage');
const selectedMetaAdsPackageNote = document.getElementById('selectedMetaAdsPackageNote');
const selectedMetaAdsPackageName = document.getElementById('selectedMetaAdsPackageName');
const selectedMetaAdsPackageInput = document.getElementById('selectedMetaAdsPackageInput');
const clearSelectedMetaAdsPackage = document.getElementById('clearSelectedMetaAdsPackage');
const sampleGalleryModal = document.getElementById('sampleGalleryModal');
const sampleGalleryTitle = document.getElementById('sampleGalleryTitle');
const sampleGalleryDescription = document.getElementById('sampleGalleryDescription');
const sampleGalleryGrid = document.getElementById('sampleGalleryGrid');
const closeSampleGallery = document.getElementById('closeSampleGallery');
const menuButton = document.getElementById('menuButton');
const siteNav = document.getElementById('siteNav');

let demos = [];
let reviews = [];
let packages = [];
let activeFilter = 'All';
let activePackageFilter = 'Website';

document.getElementById('year').textContent = new Date().getFullYear();

menuButton.addEventListener('click', () => siteNav.classList.toggle('open'));
siteNav.querySelectorAll('a').forEach(link => link.addEventListener('click', () => siteNav.classList.remove('open')));

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
  if (!value) return '#';
  if (/^(https?:|mailto:|tel:)/i.test(value)) return value;
  return `https://${value}`;
}

function parseImageUrls(value) {
  if (Array.isArray(value)) return value.filter(Boolean);
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

function getClientCooldownRemaining(key, seconds) {
  try {
    const value = Number(localStorage.getItem(`asp_submit_${key}`) || 0);
    if (!value) return 0;
    return Math.max(0, seconds - Math.floor((Date.now() - value) / 1000));
  } catch (error) {
    return 0;
  }
}

function markClientSubmission(key) {
  try {
    localStorage.setItem(`asp_submit_${key}`, String(Date.now()));
  } catch (error) {
    // Server-side anti-spam and duplicate checks still apply.
  }
}

function showCooldownMessage(element, seconds, label) {
  element.className = 'form-message error';
  element.textContent = `Please wait about ${seconds} more second${seconds === 1 ? '' : 's'} before sending another ${label}.`;
}

function renderFilters() {
  const categories = ['All', ...new Set(demos.map(item => item.category).filter(Boolean))];
  demoFilter.innerHTML = categories.map(category => `
    <button class="${category === activeFilter ? 'active' : ''}" data-filter="${escapeHtml(category)}">
      ${escapeHtml(category)}
    </button>
  `).join('');

  demoFilter.querySelectorAll('button').forEach(button => {
    button.addEventListener('click', () => {
      activeFilter = button.dataset.filter;
      renderFilters();
      renderDemos();
    });
  });
}

function renderDemos() {
  const filtered = activeFilter === 'All' ? demos : demos.filter(item => item.category === activeFilter);

  if (!filtered.length) {
    demoGrid.innerHTML = '<div class="empty-state">No demo websites in this category yet.</div>';
    return;
  }

  demoGrid.innerHTML = filtered.map(item => `
    <article class="demo-card">
      <div class="demo-image-wrap">
        <img class="demo-image" src="${escapeHtml(item.imageUrl || 'https://placehold.co/800x500/101d31/f7f9fc?text=Website+Demo')}" alt="${escapeHtml(item.title)}" loading="lazy" />
        <span class="demo-badge">${escapeHtml(item.category || 'Website')}</span>
      </div>
      <div class="demo-body">
        <h3>${escapeHtml(item.title)}</h3>
        <p>${escapeHtml(item.description || '')}</p>
        <div class="demo-actions">
          <a class="button primary" href="${escapeHtml(normalizeUrl(item.demoUrl))}" target="_blank" rel="noopener">Try Demo</a>
          <a class="button secondary" href="#avail">Build Similar</a>
        </div>
      </div>
    </article>
  `).join('');
}

function normalizePackageCategory(value = '') {
  const category = String(value).trim();
  return ['Website', 'Creative Images', 'Meta Ads'].includes(category) ? category : 'Website';
}

function renderPackageTabs() {
  if (!packageTabs) return;
  packageTabs.querySelectorAll('button').forEach(button => {
    button.classList.toggle('active', button.dataset.packageFilter === activePackageFilter);
  });
  metaAdsRequestPanel?.classList.toggle('hidden', activePackageFilter !== 'Meta Ads');
}

function renderPackages() {
  if (!packageGrid) return;
  const filtered = packages.filter(item => normalizePackageCategory(item.category) === activePackageFilter);

  if (!filtered.length) {
    packageGrid.innerHTML = `<div class="empty-state">No ${escapeHtml(activePackageFilter)} packages added yet. You can still request a custom quote.</div>`;
    return;
  }

  packageGrid.innerHTML = filtered.map(item => {
    const inclusions = String(item.inclusions || '')
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean);
    const samples = parseImageUrls(item.sampleImageUrls);
    const cover = String(item.coverImageUrl || '').trim();

    return `
      <article class="package-card ${String(item.featured).toLowerCase() === 'yes' ? 'featured' : ''}">
        ${cover ? `
          <div class="package-cover-wrap">
            <img class="package-cover" src="${escapeHtml(cover)}" alt="${escapeHtml(item.name)} package preview" loading="lazy" />
          </div>
        ` : ''}
        <div class="package-card-content">
          ${item.badge ? `<span class="package-badge">${escapeHtml(item.badge)}</span>` : '<span class="package-badge">Service Package</span>'}
          <h3>${escapeHtml(item.name)}</h3>
          <div class="package-price">${escapeHtml(item.price || 'Custom Quote')}</div>
          <p class="package-description">${escapeHtml(item.description || '')}</p>
          <ul class="package-inclusions">
            ${inclusions.map(line => `<li>${escapeHtml(line)}</li>`).join('')}
          </ul>
          <div class="package-card-actions">
            ${samples.length ? `
              <button class="button secondary full package-samples-button" type="button" data-package-id="${escapeHtml(item.id)}">
                View ${samples.length} Sample${samples.length === 1 ? '' : 's'}
              </button>
            ` : ''}
            <button class="button primary full package-avail-button" type="button" data-package-id="${escapeHtml(item.id)}">
              ${escapeHtml(item.ctaLabel || 'Avail This Package')}
            </button>
          </div>
        </div>
      </article>
    `;
  }).join('');

  packageGrid.querySelectorAll('.package-avail-button').forEach(button => {
    button.addEventListener('click', () => selectPackage(button.dataset.packageId));
  });
  packageGrid.querySelectorAll('.package-samples-button').forEach(button => {
    button.addEventListener('click', () => openSampleGallery(button.dataset.packageId));
  });
}

function setActivePackageFilter(category) {
  activePackageFilter = normalizePackageCategory(category);
  renderPackageTabs();
  renderPackages();
}

function selectPackage(id) {
  const selected = packages.find(item => item.id === id);
  if (!selected) return;

  const packageLabel = `${selected.name}${selected.price ? ` — ${selected.price}` : ''}`;

  if (normalizePackageCategory(selected.category) === 'Meta Ads') {
    setActivePackageFilter('Meta Ads');
    selectedMetaAdsPackageInput.value = packageLabel;
    selectedMetaAdsPackageName.textContent = packageLabel;
    selectedMetaAdsPackageNote.classList.remove('hidden');
    metaAdsForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
    return;
  }

  selectedPackageInput.value = packageLabel;
  selectedPackageName.textContent = packageLabel;
  selectedPackageNote.classList.remove('hidden');

  const serviceField = leadForm.elements.service;
  if (selected.category === 'Creative Images') {
    serviceField.value = selected.name.toLowerCase().includes('product')
      ? 'Product Posting Design'
      : 'Custom Ad Creatives';
  } else if (selected.category === 'Website') {
    serviceField.value = selected.name.toLowerCase().includes('admin')
      ? 'Website with Admin System'
      : 'Business Website';
  }

  document.getElementById('avail').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function clearPackageSelection() {
  selectedPackageInput.value = '';
  selectedPackageName.textContent = '';
  selectedPackageNote.classList.add('hidden');
}

function clearMetaAdsPackageSelection() {
  selectedMetaAdsPackageInput.value = '';
  selectedMetaAdsPackageName.textContent = '';
  selectedMetaAdsPackageNote.classList.add('hidden');
}

function openSampleGallery(id) {
  const selected = packages.find(item => item.id === id);
  if (!selected) return;
  const samples = parseImageUrls(selected.sampleImageUrls);
  if (!samples.length) return;

  sampleGalleryTitle.textContent = selected.name;
  sampleGalleryDescription.textContent = selected.description || '';
  sampleGalleryGrid.innerHTML = samples.map((url, index) => `
    <a class="sample-gallery-item" href="${escapeHtml(url)}" target="_blank" rel="noopener">
      <img src="${escapeHtml(url)}" alt="${escapeHtml(selected.name)} sample ${index + 1}" loading="lazy" />
      <span>Sample ${index + 1}</span>
    </a>
  `).join('');
  sampleGalleryModal.classList.remove('hidden');
  document.body.classList.add('modal-open');
}

function closeSamples() {
  sampleGalleryModal.classList.add('hidden');
  sampleGalleryGrid.innerHTML = '';
  document.body.classList.remove('modal-open');
}

if (packageTabs) {
  packageTabs.querySelectorAll('button').forEach(button => {
    button.addEventListener('click', () => setActivePackageFilter(button.dataset.packageFilter));
  });
}

document.querySelectorAll('[data-package-tab-link]').forEach(link => {
  link.addEventListener('click', () => setActivePackageFilter(link.dataset.packageTabLink));
});

clearSelectedPackage?.addEventListener('click', clearPackageSelection);
clearSelectedMetaAdsPackage?.addEventListener('click', clearMetaAdsPackageSelection);
closeSampleGallery?.addEventListener('click', closeSamples);
sampleGalleryModal?.querySelectorAll('[data-close-samples]').forEach(element => element.addEventListener('click', closeSamples));
document.addEventListener('keydown', event => {
  if (event.key === 'Escape' && !sampleGalleryModal.classList.contains('hidden')) closeSamples();
});

function renderReviews() {
  if (!reviewGrid) return;

  if (!reviews.length) {
    reviewGrid.innerHTML = '<div class="empty-state">No approved customer reviews yet. Be the first client to leave a review.</div>';
    return;
  }

  reviewGrid.innerHTML = reviews.map(item => {
    const rating = Math.max(1, Math.min(5, Number(item.rating) || 5));
    return `
      <article class="review-card">
        <div class="review-stars" aria-label="${rating} out of 5 stars">${'★'.repeat(rating)}${'☆'.repeat(5 - rating)}</div>
        <p class="review-text">“${escapeHtml(item.review || '')}”</p>
        <div class="review-author">
          <b>${escapeHtml(item.name || 'Customer')}</b>
          <span>${escapeHtml([item.business, item.service].filter(Boolean).join(' · '))}</span>
        </div>
      </article>
    `;
  }).join('');
}

function applySettings(settings = {}) {
  const brandName = settings.brandName || 'Ads Success PH';
  document.getElementById('brandName').textContent = brandName;
  document.getElementById('heroTitle').textContent = settings.heroTitle || 'Digital services built to help small businesses get noticed.';
  document.getElementById('heroSubtitle').textContent = settings.heroSubtitle || 'Custom websites, AI-assisted advertising creatives, product posting designs, and Meta Ads services—planned around your business goals and budget.';

  const logo = document.getElementById('brandLogo');
  const brandMark = document.querySelector('.brand-mark');
  if (settings.logoUrl) {
    logo.src = settings.logoUrl;
    logo.classList.remove('hidden');
    brandMark.classList.add('hidden');
  }

  const messenger = document.getElementById('messengerButton');
  if (settings.messengerUrl) {
    messenger.href = normalizeUrl(settings.messengerUrl);
    messenger.classList.remove('hidden');
  }
}

async function loadPublicData() {
  try {
    const data = await apiGet('getPublicData');
    demos = data.demos || [];
    packages = data.packages || [];
    reviews = data.reviews || [];
    applySettings(data.settings || {});
    renderFilters();
    renderDemos();
    renderPackageTabs();
    renderPackages();
    renderReviews();
  } catch (error) {
    console.error(error);
    demos = [
      {
        title: 'Job Hiring Website',
        category: 'Business System',
        description: 'Hiring landing page with applicant form and admin dashboard.',
        imageUrl: 'https://placehold.co/900x560/6d5dfc/ffffff?text=Job+Hiring+Demo',
        demoUrl: '#'
      },
      {
        title: 'Online Quotation Maker',
        category: 'Custom Tool',
        description: 'Customer quotation builder with product catalog and admin management.',
        imageUrl: 'https://placehold.co/900x560/39d9b8/07111f?text=Quotation+Demo',
        demoUrl: '#'
      }
    ];
    packages = [
      {
        id: 'fallback-website',
        category: 'Website',
        name: 'Basic Business Website',
        price: 'Starts at ₱5,000',
        description: 'A mobile-friendly website for a small business.',
        inclusions: 'Custom landing page or business website\nContact form\nMobile-friendly design\nBasic deployment',
        badge: 'Starter',
        ctaLabel: 'Request This Package',
        active: 'Yes',
        featured: 'Yes',
        sampleImageUrls: '[]'
      },
      {
        id: 'fallback-creative',
        category: 'Creative Images',
        name: 'Custom Ad Creatives',
        price: 'Custom Package',
        description: 'AI-assisted creative images directed and customized for your offer.',
        inclusions: 'Custom ad creative concepts\nProduct or service-focused layout\nHuman review and refinement',
        badge: 'Creative Service',
        ctaLabel: 'Request Creatives',
        active: 'Yes',
        featured: 'No',
        sampleImageUrls: '[]'
      },
      {
        id: 'fallback-meta',
        category: 'Meta Ads',
        name: 'Meta Ads Campaign Review',
        price: 'Custom Quote',
        description: 'Submit your campaign details for review before setup.',
        inclusions: 'Business and offer review\nCampaign objective review\nBudget and targeting discussion',
        badge: 'Detailed Intake',
        ctaLabel: 'Complete Ads Form',
        active: 'Yes',
        featured: 'No',
        sampleImageUrls: '[]'
      }
    ];
    reviews = [];
    renderFilters();
    renderDemos();
    renderPackageTabs();
    renderPackages();
    renderReviews();
  }
}

leadForm.addEventListener('submit', async event => {
  event.preventDefault();
  const button = leadForm.querySelector('button[type="submit"]');
  const payload = Object.fromEntries(new FormData(leadForm).entries());
  const cooldown = getClientCooldownRemaining('lead', 30);
  if (cooldown > 0) {
    showCooldownMessage(formMessage, cooldown, 'service request');
    return;
  }

  formMessage.className = 'form-message';
  formMessage.textContent = 'Sending your request...';
  button.disabled = true;
  button.textContent = 'Sending...';

  try {
    await apiPost('submitLead', payload);
    formMessage.className = 'form-message success';
    formMessage.textContent = 'Request sent successfully! Ads Success PH will contact you soon.';
    markClientSubmission('lead');
    leadForm.reset();
    clearPackageSelection();
  } catch (error) {
    formMessage.className = 'form-message error';
    formMessage.textContent = error.message.includes('configured')
      ? 'Website demo mode: connect the Apps Script API to receive real inquiries.'
      : error.message;
  } finally {
    button.disabled = false;
    button.textContent = 'Send Service Request';
  }
});

metaAdsForm?.addEventListener('submit', async event => {
  event.preventDefault();
  const button = metaAdsForm.querySelector('button[type="submit"]');
  const payload = Object.fromEntries(new FormData(metaAdsForm).entries());
  const cooldown = getClientCooldownRemaining('metaAds', 60);
  if (cooldown > 0) {
    showCooldownMessage(metaAdsMessage, cooldown, 'Meta Ads request');
    return;
  }

  metaAdsMessage.className = 'form-message';
  metaAdsMessage.textContent = 'Submitting your detailed Meta Ads request...';
  button.disabled = true;
  button.textContent = 'Submitting...';

  try {
    await apiPost('submitMetaAdsRequest', payload);
    metaAdsMessage.className = 'form-message success';
    metaAdsMessage.textContent = 'Meta Ads request submitted successfully! Your campaign details are now ready for admin review.';
    markClientSubmission('metaAds');
    metaAdsForm.reset();
    clearMetaAdsPackageSelection();
  } catch (error) {
    metaAdsMessage.className = 'form-message error';
    metaAdsMessage.textContent = error.message.includes('configured')
      ? 'Website demo mode: connect the Apps Script API to receive real Meta Ads requests.'
      : error.message;
  } finally {
    button.disabled = false;
    button.textContent = 'Submit Detailed Meta Ads Request';
  }
});

if (reviewForm) {
  reviewForm.addEventListener('submit', async event => {
    event.preventDefault();
    const button = reviewForm.querySelector('button[type="submit"]');
    const payload = Object.fromEntries(new FormData(reviewForm).entries());
    const cooldown = getClientCooldownRemaining('review', 60);
    if (cooldown > 0) {
      showCooldownMessage(reviewMessage, cooldown, 'review');
      return;
    }

    reviewMessage.className = 'form-message';
    reviewMessage.textContent = 'Sending your review...';
    button.disabled = true;
    button.textContent = 'Submitting...';

    try {
      await apiPost('submitReview', payload);
      reviewMessage.className = 'form-message success';
      reviewMessage.textContent = 'Thank you! Your review was submitted and is waiting for admin approval.';
      markClientSubmission('review');
      reviewForm.reset();
    } catch (error) {
      reviewMessage.className = 'form-message error';
      reviewMessage.textContent = error.message.includes('configured')
        ? 'Website demo mode: connect the Apps Script API to receive customer reviews.'
        : error.message;
    } finally {
      button.disabled = false;
      button.textContent = 'Submit Customer Review';
    }
  });
}

loadPublicData();
